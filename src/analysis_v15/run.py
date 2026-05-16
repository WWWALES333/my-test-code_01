from __future__ import annotations

import argparse
import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from src.analysis_v14.loader import build_report_record, collect_sample_files
from src.analysis_v14.parser import extract_text
from src.analysis_v14.review import build_review_item_for_parse_failure, build_review_queue_from_tags
from src.analysis_v14.run import strip_tag_rows
from src.analysis_v14.schema import DECISION_VALUES, MODEL_REASON_FAILED, PARSE_STATUS_FAILED, PARSE_STATUS_SUCCESS
from src.analysis_v14.tagger import Tagger

from .insights import build_insight_tree, flatten_insight_tree
from .metrics import build_dashboard_snapshot, build_trend_cube, build_trend_explanations
from .normalize import build_evidence_facts, build_report_facts, build_sales_monthly_rollup
from .owner import build_owner_registry, extract_owner_hint
from .parser import segment_text_with_owner
from .profiles import build_region_sales_rollup, build_salesperson_profiles
from .reporter import build_summary_markdown, build_workbench_pages, write_csv_template, write_json, write_jsonl, write_markdown, write_web_pages
from .review_state import (
    build_review_audit_log,
    build_review_batch_summaries,
    build_review_learning_candidates,
    build_review_learning_summary,
    build_system_review_tasks,
    build_review_tasks,
    load_review_decisions,
)
from .roster import discover_roster_file, load_sales_roster


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="v1.5 AI 一线情报工作台离线分析入口")
    parser.add_argument("--samples", required=True, help="样本目录路径")
    parser.add_argument("--annotations", required=True, help="标注或复核目录路径")
    parser.add_argument("--out", required=True, help="输出目录路径")
    parser.add_argument("--model-mode", choices=["mock", "real"], default="mock", help="模型模式")
    parser.add_argument("--llm-concurrency", type=int, default=4, help="real 模式下 needs_llm 样本的并发度")
    parser.add_argument("--llm-chunk-size", type=int, default=50, help="real 模式下每轮送入 LLM 的分块大小")
    parser.add_argument("--roster", default="", help="花名册 Excel 路径；为空时自动从 data/input/v1.5/roster 发现")
    return parser.parse_args()


def ensure_input_dirs(samples_dir: Path, annotations_dir: Path) -> None:
    if not samples_dir.exists():
        raise FileNotFoundError(f"样本目录不存在: {samples_dir}")
    annotations_dir.mkdir(parents=True, exist_ok=True)


def run_pipeline(
    samples_dir: Path,
    annotations_dir: Path,
    out_dir: Path,
    model_mode: str,
    llm_concurrency: int = 4,
    llm_chunk_size: int = 50,
    roster_path: Path | None = None,
) -> Dict[str, int]:
    ensure_input_dirs(samples_dir, annotations_dir)
    sample_files = collect_sample_files(samples_dir)
    if not sample_files:
        raise RuntimeError(f"样本目录为空: {samples_dir}")

    run_id = datetime.now().strftime("v15_%Y%m%d_%H%M%S")
    extracted_dir = out_dir / "extracted"
    normalized_dir = out_dir / "normalized"
    reports_dir = out_dir / "reports"
    review_dir = out_dir / "review"
    runtime_dir = out_dir / "runtime"
    web_dir = out_dir / "web"
    for path in (extracted_dir, normalized_dir, reports_dir, review_dir, runtime_dir, web_dir):
        path.mkdir(parents=True, exist_ok=True)

    roster_file = discover_roster_file(roster_path)
    sales_roster = load_sales_roster(roster_file) if roster_file else []

    tagger = Tagger(mode=model_mode)
    default_model_name = "mock-rule-engine" if model_mode == "mock" else (tagger.model_name or "unknown")

    report_rows: List[Dict[str, object]] = []
    tag_rows: List[Dict[str, object]] = []
    evidence_rows: List[Dict[str, object]] = []
    parse_failure_reviews: List[Dict[str, object]] = []
    segment_tasks: List[Dict[str, object]] = []

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
            segment_tasks.append(
                {
                    "report_row": report_row,
                    "segment_id": f"S{idx:03d}",
                    "task_key": _stable_tag_id(str(report_row["report_id"]), f"S{idx:03d}", str(segment_item.get("source_text", ""))),
                    "segment": str(segment_item.get("source_text", "")),
                    "segment_owner_hint": str(segment_item.get("owner_hint", "")).strip(),
                    "parse_status": parse_status,
                }
            )

    classifications = _classify_segment_tasks(
        tagger=tagger,
        segment_tasks=segment_tasks,
        model_mode=model_mode,
        llm_concurrency=llm_concurrency,
        llm_chunk_size=llm_chunk_size,
        runtime_dir=runtime_dir,
    )
    for task, cls in zip(segment_tasks, classifications):
        report_row = dict(task["report_row"])
        segment_id = str(task["segment_id"])
        segment = str(task["segment"])
        segment_owner_hint = str(task["segment_owner_hint"])
        parse_status = str(task["parse_status"])
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
            "triage_status": cls.get("triage_status", ""),
            "used_label_gap": cls.get("used_label_gap", False),
            "llm_invoked": cls.get("llm_invoked", False),
            "llm_failed": cls.get("llm_failed", False),
            "rule_baseline": cls.get("rule_baseline", {}),
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
                    "triage_status": cls.get("triage_status", ""),
                    "used_label_gap": cls.get("used_label_gap", False),
                    "llm_invoked": cls.get("llm_invoked", False),
                    "llm_failed": cls.get("llm_failed", False),
                    "rule_baseline": cls.get("rule_baseline", {}),
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

    review_decisions = load_review_decisions(annotations_dir, review_dir=review_dir)
    extra_owner_hints = [str(row.get("segment_owner_hint", "")) for row in tag_rows if str(row.get("segment_owner_hint", "")).strip()]
    owner_registry = build_owner_registry(report_rows, extra_owner_hints=extra_owner_hints, sales_roster=sales_roster)
    report_facts = build_report_facts(report_rows, owner_registry)
    evidence_facts = build_evidence_facts(evidence_rows, report_facts, review_decisions, owner_registry)
    sales_rollup = build_sales_monthly_rollup(evidence_facts, report_facts)
    salesperson_profiles = build_salesperson_profiles(evidence_facts, sales_roster)
    region_rollups = build_region_sales_rollup(salesperson_profiles)
    trend_cube = build_trend_cube(evidence_facts, sales_roster)
    trend_explanations = build_trend_explanations(trend_cube, salesperson_profiles)
    insight_tree = build_insight_tree(evidence_facts)
    insight_cards = flatten_insight_tree(insight_tree)
    review_tasks = build_review_tasks(review_rows, tag_rows, report_facts, review_decisions, owner_registry)
    system_review_tasks = build_system_review_tasks(review_rows, tag_rows, report_facts, review_decisions, owner_registry)
    review_batch_summaries = build_review_batch_summaries(review_tasks)
    review_audit_log = build_review_audit_log(review_decisions)
    review_candidates = build_review_learning_candidates(review_decisions)
    review_learning_summary = build_review_learning_summary(review_decisions, review_candidates)
    evidence_index = _build_evidence_index(evidence_facts)
    dashboard_snapshot = build_dashboard_snapshot(
        report_facts=report_facts,
        evidence_facts=evidence_facts,
        trend_cube=trend_cube,
        trend_explanations=trend_explanations,
        salesperson_profiles=salesperson_profiles,
        region_rollups=region_rollups,
        insight_cards=insight_cards,
        review_tasks=review_tasks,
        sales_roster=sales_roster,
    )

    write_jsonl(extracted_dir / "report_index.jsonl", report_rows)
    write_jsonl(extracted_dir / "tag_result.jsonl", _strip_tag_rows_v15(tag_rows))
    write_jsonl(extracted_dir / "evidence_span.jsonl", evidence_rows)
    write_jsonl(extracted_dir / "review_queue.jsonl", review_rows)

    write_jsonl(normalized_dir / "sales_roster.jsonl", sales_roster)
    write_jsonl(normalized_dir / "owner_registry.jsonl", owner_registry)
    write_jsonl(normalized_dir / "report_facts.jsonl", report_facts)
    write_jsonl(normalized_dir / "evidence_facts.jsonl", evidence_facts)
    write_jsonl(normalized_dir / "sales_monthly_rollup.jsonl", sales_rollup)
    write_json(normalized_dir / "trend_cube.json", trend_cube)
    write_json(normalized_dir / "trend_explanations.json", trend_explanations)
    write_jsonl(normalized_dir / "salesperson_profile.jsonl", salesperson_profiles)
    write_jsonl(normalized_dir / "region_sales_rollup.jsonl", region_rollups)
    write_json(normalized_dir / "insight_tree.json", insight_tree)
    write_jsonl(normalized_dir / "insight_cards.jsonl", insight_cards)
    write_jsonl(normalized_dir / "review_tasks.jsonl", review_tasks)
    write_jsonl(normalized_dir / "system_review_tasks.jsonl", system_review_tasks)
    write_json(normalized_dir / "review_batch_summaries.json", review_batch_summaries)
    write_jsonl(normalized_dir / "review_decisions.jsonl", review_decisions.values())
    write_jsonl(normalized_dir / "review_audit_log.jsonl", review_audit_log)
    write_jsonl(normalized_dir / "review_candidates.jsonl", review_candidates)
    write_json(normalized_dir / "review_learning_summary.json", review_learning_summary)
    write_jsonl(normalized_dir / "evidence_index.jsonl", evidence_index)
    write_json(normalized_dir / "dashboard_snapshot.json", dashboard_snapshot)

    write_jsonl(review_dir / "review_decisions.jsonl", review_decisions.values())
    write_jsonl(review_dir / "review_audit_log.jsonl", review_audit_log)
    write_jsonl(review_dir / "system_review_tasks.jsonl", system_review_tasks)
    write_json(review_dir / "review_batch_summaries.json", review_batch_summaries)
    write_jsonl(review_dir / "review_candidates.jsonl", review_candidates)
    write_json(review_dir / "review_learning_summary.json", review_learning_summary)
    write_json(
        out_dir / "run_manifest.json",
        {
            "samples": str(samples_dir.resolve()),
            "annotations": str(annotations_dir.resolve()),
            "out": str(out_dir.resolve()),
            "model_mode": model_mode,
            "llm_concurrency": llm_concurrency,
            "llm_chunk_size": llm_chunk_size,
            "roster": str(roster_file.resolve()) if roster_file else "",
            "generated_at": datetime.now().isoformat(timespec="seconds"),
        },
    )

    summary = build_summary_markdown(
        snapshot=dashboard_snapshot,
        trend_explanations=trend_explanations,
        salesperson_profiles=salesperson_profiles,
        insight_cards=insight_cards,
        review_tasks=review_tasks,
        region_rollups=region_rollups,
    )
    write_markdown(reports_dir / "AI情报工作台摘要.md", summary)
    write_web_pages(
        web_dir,
        build_workbench_pages(
            snapshot=dashboard_snapshot,
            trend_cube=trend_cube,
            trend_explanations=trend_explanations,
            salesperson_profiles=salesperson_profiles,
            region_rollups=region_rollups,
            insight_tree=insight_tree,
            review_tasks=review_tasks,
            evidence_index=evidence_index,
            review_learning_summary=review_learning_summary,
            review_candidates=review_candidates,
            review_batch_summaries=review_batch_summaries,
            interactive_review=False,
        ),
    )
    write_csv_template(review_dir / "review_result_template.csv")
    _write_review_result_csv(review_dir / "review_result.csv", review_tasks)

    return {
        "run_id": run_id,
        "reports": len(report_rows),
        "tag_rows": len(tag_rows),
        "evidence_rows": len(evidence_rows),
        "review_rows": len(review_rows),
        "system_review_rows": len(system_review_tasks),
        "normalized_rows": len(evidence_facts),
        "insight_cards": len(insight_cards),
        "roster_sales": len(sales_roster),
        "profiles": len(salesperson_profiles),
        "llm_concurrency": llm_concurrency,
        "llm_chunk_size": llm_chunk_size,
        "llm_invoked": sum(1 for row in tag_rows if bool(row.get("llm_invoked", False))),
    }


def _build_evidence_index(evidence_facts: List[Dict[str, object]]) -> List[Dict[str, object]]:
    rows = []
    for row in evidence_facts:
        rows.append(
            {
                "evidence_id": row.get("evidence_id", ""),
                "report_id": row.get("report_id", ""),
                "segment_id": row.get("segment_id", ""),
                "year": row.get("year", 0),
                "month": row.get("month", 0),
                "week_of_month": row.get("week_of_month", 0),
                "salesperson_id": row.get("salesperson_id", ""),
                "salesperson_name": row.get("salesperson_name", ""),
                "battle_zone_name": row.get("battle_zone_name", ""),
                "region_name": row.get("region_name", ""),
                "business_line": row.get("business_line", ""),
                "actor_primary": row.get("actor_primary", ""),
                "ai_scope": row.get("ai_scope", ""),
                "decision_status": row.get("decision_status", ""),
                "triage_status": row.get("triage_status", ""),
                "used_label_gap": row.get("used_label_gap", False),
                "llm_invoked": row.get("llm_invoked", False),
                "llm_failed": row.get("llm_failed", False),
                "review_status": row.get("review_status", ""),
                "review_comment": row.get("review_comment", ""),
                "source_text": row.get("source_text", ""),
                "file_path": row.get("file_path", ""),
            }
        )
    return rows


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
        "error_reason_primary",
        "review_necessity",
        "actionability",
        "action_bucket",
        "learning_note",
        "need_rule_update",
        "need_prompt_update",
        "need_annotation_update",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for task in review_tasks:
            learning_fields = dict(task.get("learning_fields", {}))
            writer.writerow(
                {
                    "sample_id": "",
                    "task_id": task.get("task_id", ""),
                    "report_id": task.get("report_id", ""),
                    "segment_id": task.get("segment_id", ""),
                    "salesperson_id": task.get("salesperson_id", ""),
                    "review_reason_code": task.get("review_reason_code", ""),
                    "current_labels": json.dumps(task.get("current_fields", {}), ensure_ascii=False),
                    "reviewed_fields": json.dumps(task.get("edited_fields", {}), ensure_ascii=False) if task.get("edited_fields") else "",
                    "final_labels": json.dumps(task.get("edited_fields", {}), ensure_ascii=False) if task.get("edited_fields") else "",
                    "is_pass": "1" if str(task.get("task_status", "")) == "reviewed" else "",
                    "review_comment": task.get("review_comment", ""),
                    "reviewer": task.get("reviewer", ""),
                    "reviewed_at": task.get("reviewed_at", ""),
                    "error_reason_primary": learning_fields.get("error_reason_primary", ""),
                    "review_necessity": learning_fields.get("review_necessity", ""),
                    "actionability": learning_fields.get("actionability", ""),
                    "action_bucket": learning_fields.get("action_bucket", ""),
                    "learning_note": learning_fields.get("learning_note", ""),
                    "need_rule_update": "1" if bool(learning_fields.get("need_rule_update", False)) else "",
                    "need_prompt_update": "1" if bool(learning_fields.get("need_prompt_update", False)) else "",
                    "need_annotation_update": "1" if bool(learning_fields.get("need_annotation_update", False)) else "",
                }
            )


def _normalize_classification(cls: Dict[str, object]) -> Dict[str, object]:
    item = dict(cls)
    item.setdefault("is_ai_hit", False)
    item.setdefault("business_line", "待判断")
    item.setdefault("ai_actor", "")
    item.setdefault("actor_primary", item.get("ai_actor", ""))
    item.setdefault("actor_subtype", "")
    item.setdefault("ai_scope", "product_ai")
    item.setdefault("interaction_outcome", "not_applicable")
    item.setdefault("certainty_level", "medium")
    item.setdefault("review_reason_code", "")
    item.setdefault("triage_status", "")
    item["used_label_gap"] = bool(item.get("used_label_gap", False) or str(item.get("actor_primary", "")).strip() == "label_gap")
    item.setdefault("llm_invoked", False)
    item.setdefault("llm_failed", False)
    item.setdefault("rule_baseline", {})
    decision_status = str(item.get("decision_status", "")).strip() or "pending_human_review"
    if decision_status not in DECISION_VALUES:
        decision_status = "pending_human_review"
    item["decision_status"] = decision_status
    if decision_status == "uncertain" and not item.get("review_reason_code"):
        item["review_reason_code"] = "SCOPE_AMBIGUOUS"
    if decision_status == "pending_human_review" and not item.get("review_reason_code"):
        item["review_reason_code"] = "ACTOR_OVERLAP"
    reason = str(item.get("reason", "")).strip()
    if not reason and decision_status != "confirmed":
        item["reason"] = "需要人工复核"
    return item


def _classify_segment_tasks(
    tagger: Tagger,
    segment_tasks: List[Dict[str, object]],
    model_mode: str,
    llm_concurrency: int,
    llm_chunk_size: int,
    runtime_dir: Path,
) -> List[Dict[str, object]]:
    if not segment_tasks:
        return []
    cache_path = runtime_dir / "real_llm_cache.jsonl"
    progress_path = runtime_dir / "real_llm_progress.json"
    cached_results = _load_classification_cache(cache_path) if model_mode == "real" else {}
    results: List[Dict[str, object] | None] = [None] * len(segment_tasks)
    pending_indexes: List[int] = []

    for idx, task in enumerate(segment_tasks):
        task_key = str(task.get("task_key", ""))
        if model_mode == "real" and task_key in cached_results:
            results[idx] = dict(cached_results[task_key])
        else:
            pending_indexes.append(idx)

    if model_mode == "real":
        _write_llm_progress(
            progress_path,
            total=len(segment_tasks),
            completed=len(segment_tasks) - len(pending_indexes),
            pending=len(pending_indexes),
            llm_concurrency=llm_concurrency,
            llm_chunk_size=llm_chunk_size,
            cache_path=cache_path,
        )

    chunk_size = max(1, llm_chunk_size)
    for offset in range(0, len(pending_indexes), chunk_size):
        chunk_indexes = pending_indexes[offset : offset + chunk_size]
        chunk_inputs = [
            (str(segment_tasks[idx]["segment"]), dict(segment_tasks[idx]["report_row"]))
            for idx in chunk_indexes
        ]
        chunk_results = tagger.classify_batch(chunk_inputs, llm_concurrency=llm_concurrency)
        new_rows = []
        for idx, cls in zip(chunk_indexes, chunk_results):
            task_key = str(segment_tasks[idx].get("task_key", ""))
            result = dict(cls)
            results[idx] = result
            if model_mode == "real":
                new_rows.append({"task_key": task_key, "classification": result})
                cached_results[task_key] = result
        if new_rows:
            _append_classification_cache(cache_path, new_rows)
            _write_llm_progress(
                progress_path,
                total=len(segment_tasks),
                completed=sum(1 for item in results if item is not None),
                pending=sum(1 for item in results if item is None),
                llm_concurrency=llm_concurrency,
                llm_chunk_size=llm_chunk_size,
                cache_path=cache_path,
            )
            print(
                "real 批次进度: {done}/{total}（chunk={chunk}, 并发={concurrency}）".format(
                    done=sum(1 for item in results if item is not None),
                    total=len(segment_tasks),
                    chunk=chunk_size,
                    concurrency=llm_concurrency,
                )
            )

    final_results: List[Dict[str, object]] = []
    for item in results:
        if item is None:
            raise RuntimeError("存在未完成的分类任务，无法继续生成产物")
        final_results.append(dict(item))
    return final_results


def _load_classification_cache(path: Path) -> Dict[str, Dict[str, object]]:
    if not path.exists():
        return {}
    cache: Dict[str, Dict[str, object]] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        row = json.loads(line)
        task_key = str(row.get("task_key", "")).strip()
        classification = row.get("classification", {})
        if task_key and isinstance(classification, dict):
            cache[task_key] = classification
    return cache


def _append_classification_cache(path: Path, rows: List[Dict[str, object]]) -> None:
    with path.open("a", encoding="utf-8") as handle:
        for row in rows:
            handle.write(json.dumps(row, ensure_ascii=False) + "\n")


def _write_llm_progress(
    path: Path,
    total: int,
    completed: int,
    pending: int,
    llm_concurrency: int,
    llm_chunk_size: int,
    cache_path: Path,
) -> None:
    payload = {
        "total": total,
        "completed": completed,
        "pending": pending,
        "llm_concurrency": llm_concurrency,
        "llm_chunk_size": llm_chunk_size,
        "cache_path": str(cache_path.resolve()),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    args = parse_args()
    run_pipeline(
        samples_dir=Path(args.samples),
        annotations_dir=Path(args.annotations),
        out_dir=Path(args.out),
        model_mode=args.model_mode,
        llm_concurrency=args.llm_concurrency,
        llm_chunk_size=args.llm_chunk_size,
        roster_path=Path(args.roster) if args.roster else None,
    )


if __name__ == "__main__":
    main()
