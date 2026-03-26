from __future__ import annotations

import argparse
import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List

from src.analysis_v14.loader import build_report_record, collect_sample_files
from src.analysis_v14.parser import extract_text
from src.analysis_v14.review import build_review_item_for_parse_failure, build_review_queue_from_tags
from src.analysis_v14.run import strip_tag_rows
from src.analysis_v14.schema import DECISION_VALUES, MODEL_REASON_FAILED, PARSE_STATUS_FAILED, PARSE_STATUS_SUCCESS
from src.analysis_v14.tagger import Tagger

from .normalize import (
    build_dashboard_snapshot,
    build_evidence_facts,
    build_insight_cards,
    build_report_facts,
    build_review_tasks,
    build_sales_monthly_rollup,
    load_review_decisions,
)
from .owner import build_owner_registry, extract_owner_hint
from .parser import segment_text_with_owner
from .reporter import build_summary_markdown, build_workbench_html, write_csv_template, write_json, write_jsonl, write_markdown


def parse_args() -> argparse.Namespace:
    """解析 v1.5 离线入口参数。"""
    parser = argparse.ArgumentParser(description="v1.5 AI 一线情报工作台离线分析入口")
    parser.add_argument("--samples", required=True, help="样本目录路径")
    parser.add_argument("--annotations", required=True, help="标注或复核目录路径")
    parser.add_argument("--out", required=True, help="输出目录路径")
    parser.add_argument("--model-mode", choices=["mock", "real"], default="mock", help="模型模式")
    return parser.parse_args()


def ensure_input_dirs(samples_dir: Path, annotations_dir: Path) -> None:
    """检查输入目录是否存在，并确保标注目录可写。"""
    if not samples_dir.exists():
        raise FileNotFoundError(f"样本目录不存在: {samples_dir}")
    annotations_dir.mkdir(parents=True, exist_ok=True)


def run_pipeline(samples_dir: Path, annotations_dir: Path, out_dir: Path, model_mode: str) -> Dict[str, int]:
    """执行 v1.5 第一阶段离线构建，产出 extracted + normalized + reports + review + web。"""
    ensure_input_dirs(samples_dir, annotations_dir)
    sample_files = collect_sample_files(samples_dir)
    if not sample_files:
        raise RuntimeError(f"样本目录为空: {samples_dir}")

    run_id = datetime.now().strftime("v15_%Y%m%d_%H%M%S")
    extracted_dir = out_dir / "extracted"
    normalized_dir = out_dir / "normalized"
    reports_dir = out_dir / "reports"
    review_dir = out_dir / "review"
    web_dir = out_dir / "web"
    for path in (extracted_dir, normalized_dir, reports_dir, review_dir, web_dir):
        path.mkdir(parents=True, exist_ok=True)

    tagger = Tagger(mode=model_mode)
    default_model_name = "mock-rule-engine" if model_mode == "mock" else (tagger.model_name or "unknown")

    report_rows: List[Dict[str, object]] = []
    tag_rows: List[Dict[str, object]] = []
    evidence_rows: List[Dict[str, object]] = []
    parse_failure_reviews: List[Dict[str, object]] = []

    for path in sample_files:
        report_row = build_report_record(path)
        text, parse_error, parse_reason_code = extract_text(path)
        fallback_owner = extract_owner_hint(str(report_row["file_path"]))
        segment_items = segment_text_with_owner(text, fallback_owner) if not parse_error else []

        parse_status = PARSE_STATUS_FAILED if parse_error else PARSE_STATUS_SUCCESS
        report_row["run_id"] = run_id
        report_row["model_mode"] = model_mode
        report_row["model_name"] = default_model_name
        report_row["text_status"] = parse_status
        report_row["parse_status"] = parse_status
        report_row["parse_reason_code"] = parse_reason_code if parse_error else ""
        report_row["segment_count"] = len(segment_items)
        report_rows.append(report_row)

        if parse_error:
            parse_failure_reviews.append(build_review_item_for_parse_failure(report_row, parse_error))
            continue

        for idx, segment_item in enumerate(segment_items, start=1):
            segment_id = f"S{idx:03d}"
            segment = str(segment_item.get("source_text", ""))
            segment_owner_hint = str(segment_item.get("owner_hint", "")).strip()
            cls = tagger.classify(segment, context=report_row)
            cls = _normalize_classification(cls)

            if not cls["is_ai_hit"] and cls["decision_status"] == "confirmed":
                continue

            tag_id = _stable_tag_id(str(report_row["report_id"]), segment_id, segment)
            tag_row = {
                "tag_id": tag_id,
                "report_id": report_row["report_id"],
                "segment_id": segment_id,
                "is_ai_hit": cls["is_ai_hit"],
                "business_line": cls["business_line"],
                "ai_actor": cls["ai_actor"],
                "actor_primary": cls["actor_primary"],
                "actor_subtype": cls["actor_subtype"],
                "ai_scope": cls["ai_scope"],
                "interaction_outcome": cls["interaction_outcome"],
                "certainty_level": cls["certainty_level"],
                "review_reason_code": cls["review_reason_code"],
                "decision_status": cls["decision_status"],
                "confidence": cls["confidence"],
                "reason": cls["reason"],
                "run_id": run_id,
                "model_mode": cls.get("model_mode", model_mode),
                "model_name": cls.get("model_name", default_model_name),
                "parse_status": parse_status,
                "source_text": segment,
                "segment_owner_hint": segment_owner_hint,
                "file_path": report_row["file_path"],
            }
            tag_rows.append(tag_row)

            if cls["is_ai_hit"]:
                evidence_rows.append(
                    {
                        "evidence_id": _stable_tag_id(str(report_row["report_id"]), segment_id, "evidence"),
                        "report_id": report_row["report_id"],
                        "segment_id": segment_id,
                        "source_text": segment,
                        "business_line": cls["business_line"],
                        "ai_actor": cls["ai_actor"],
                        "actor_primary": cls["actor_primary"],
                        "actor_subtype": cls["actor_subtype"],
                        "ai_scope": cls["ai_scope"],
                        "interaction_outcome": cls["interaction_outcome"],
                        "certainty_level": cls["certainty_level"],
                        "decision_status": cls["decision_status"],
                        "run_id": run_id,
                        "model_mode": cls.get("model_mode", model_mode),
                        "model_name": cls.get("model_name", default_model_name),
                        "segment_owner_hint": segment_owner_hint,
                        "file_path": report_row["file_path"],
                    }
                )

    review_rows = build_review_queue_from_tags(tag_rows)
    review_rows.extend(parse_failure_reviews)
    tag_owner_map = {(str(row.get("report_id", "")), str(row.get("segment_id", ""))): str(row.get("segment_owner_hint", "")) for row in tag_rows}
    for row in review_rows:
        key = (str(row.get("report_id", "")), str(row.get("segment_id", "")))
        row["segment_owner_hint"] = tag_owner_map.get(key, "")

    review_decisions = load_review_decisions(annotations_dir)
    extra_owner_hints = [str(row.get("segment_owner_hint", "")) for row in tag_rows if str(row.get("segment_owner_hint", "")).strip()]
    owner_registry = build_owner_registry(report_rows, extra_owner_hints=extra_owner_hints)
    report_facts = build_report_facts(report_rows, owner_registry)
    evidence_facts = build_evidence_facts(evidence_rows, report_facts, review_decisions, owner_registry)
    review_tasks = build_review_tasks(review_rows, report_facts, review_decisions, owner_registry)
    sales_rollup = build_sales_monthly_rollup(evidence_facts, report_facts)
    insight_cards = build_insight_cards(evidence_facts)
    dashboard_snapshot = build_dashboard_snapshot(report_facts, evidence_facts, sales_rollup, insight_cards, review_tasks)

    write_jsonl(extracted_dir / "report_index.jsonl", report_rows)
    write_jsonl(extracted_dir / "tag_result.jsonl", _strip_tag_rows_v15(tag_rows))
    write_jsonl(extracted_dir / "evidence_span.jsonl", evidence_rows)
    write_jsonl(extracted_dir / "review_queue.jsonl", review_rows)

    write_jsonl(normalized_dir / "owner_registry.jsonl", owner_registry)
    write_jsonl(normalized_dir / "report_facts.jsonl", report_facts)
    write_jsonl(normalized_dir / "evidence_facts.jsonl", evidence_facts)
    write_jsonl(normalized_dir / "sales_monthly_rollup.jsonl", sales_rollup)
    write_jsonl(normalized_dir / "insight_cards.jsonl", insight_cards)
    write_jsonl(normalized_dir / "review_tasks.jsonl", review_tasks)
    write_jsonl(normalized_dir / "review_decisions.jsonl", review_decisions.values())
    write_json(normalized_dir / "dashboard_snapshot.json", dashboard_snapshot)

    summary = build_summary_markdown(dashboard_snapshot, insight_cards, review_tasks)
    write_markdown(reports_dir / "AI情报工作台摘要.md", summary)
    write_markdown(web_dir / "AI情报工作台.html", build_workbench_html(dashboard_snapshot, sales_rollup, insight_cards, review_tasks))
    write_csv_template(review_dir / "review_result_template.csv")
    _write_review_result_csv(review_dir / "review_result.csv", review_tasks)

    return {
        "run_id": run_id,
        "reports": len(report_rows),
        "tag_rows": len(tag_rows),
        "evidence_rows": len(evidence_rows),
        "review_rows": len(review_rows),
        "normalized_rows": len(evidence_facts),
        "insight_cards": len(insight_cards),
    }


def _stable_tag_id(report_id: str, segment_id: str, payload: str) -> str:
    from src.analysis_v14.schema import stable_hash

    return stable_hash(report_id, segment_id, payload)


def _strip_tag_rows_v15(rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    base_rows = strip_tag_rows(rows)
    kept: List[Dict[str, object]] = []
    for base, original in zip(base_rows, rows):
        item = dict(base)
        item["segment_owner_hint"] = original.get("segment_owner_hint", "")
        kept.append(item)
    return kept


def _write_review_result_csv(path: Path, review_tasks: List[Dict[str, object]]) -> None:
    fieldnames = [
        "sample_id",
        "task_id",
        "report_id",
        "segment_id",
        "salesperson_id",
        "review_reason_code",
        "current_labels",
        "reviewed_fields",
        "final_labels",
        "is_pass",
        "review_comment",
        "reviewer",
        "reviewed_at",
        "need_rule_update",
        "need_skill_update",
        "need_annotation_update",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for task in review_tasks:
            writer.writerow(
                {
                    "sample_id": "",
                    "task_id": task.get("task_id", ""),
                    "report_id": task.get("report_id", ""),
                    "segment_id": task.get("segment_id", ""),
                    "salesperson_id": task.get("salesperson_id", ""),
                    "review_reason_code": task.get("review_reason_code", ""),
                    "current_labels": json.dumps(task.get("current_fields", {}), ensure_ascii=False),
                    "reviewed_fields": "",
                    "final_labels": "",
                    "is_pass": "",
                    "review_comment": task.get("review_comment", ""),
                    "reviewer": task.get("reviewer", ""),
                    "reviewed_at": task.get("reviewed_at", ""),
                    "need_rule_update": "",
                    "need_skill_update": "",
                    "need_annotation_update": "",
                }
            )


def _normalize_classification(cls: Dict[str, object]) -> Dict[str, object]:
    normalized = dict(cls)
    decision_status = str(normalized.get("decision_status", "")).strip()
    if decision_status in DECISION_VALUES:
        return normalized
    reason_code = str(normalized.get("review_reason_code", "")).strip()
    normalized["decision_status"] = "pending_human_review"
    normalized["certainty_level"] = "low"
    merged_codes = [code for code in reason_code.split(";") if code]
    merged_codes.append(MODEL_REASON_FAILED)
    normalized["review_reason_code"] = ";".join(sorted(set(merged_codes)))
    normalized["reason"] = f"{normalized.get('reason', '分类结果非法')}；decision_status非法"
    return normalized


def main() -> None:
    args = parse_args()
    result = run_pipeline(Path(args.samples), Path(args.annotations), Path(args.out), args.model_mode)
    print(
        "v1.5 分析完成: run_id={run_id}, reports={reports}, tags={tag_rows}, evidence={evidence_rows}, review={review_rows}, normalized={normalized_rows}, insight_cards={insight_cards}".format(
            **result
        )
    )


if __name__ == "__main__":
    main()
