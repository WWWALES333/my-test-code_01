from __future__ import annotations

import argparse
import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List

from .loader import build_report_record, collect_sample_files
from .parser import extract_text, segment_text
from .reporter import build_summary_markdown, write_markdown
from .review import build_review_item_for_parse_failure, build_review_queue_from_tags
from .schema import (
    DECISION_VALUES,
    MODEL_REASON_FAILED,
    PARSE_STATUS_FAILED,
    PARSE_STATUS_SUCCESS,
    stable_hash,
)
from .tagger import Tagger


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="v1.4 AI 专题离线分析入口")
    parser.add_argument("--samples", required=True, help="样本目录路径")
    parser.add_argument("--annotations", required=True, help="标注目录路径")
    parser.add_argument("--out", required=True, help="输出目录路径")
    parser.add_argument("--model-mode", choices=["mock", "real"], default="mock", help="模型模式")
    return parser.parse_args()


def ensure_input_dirs(samples_dir: Path, annotations_dir: Path) -> None:
    if not samples_dir.exists():
        raise FileNotFoundError(f"样本目录不存在: {samples_dir}")
    annotations_dir.mkdir(parents=True, exist_ok=True)


def run_pipeline(samples_dir: Path, annotations_dir: Path, out_dir: Path, model_mode: str) -> Dict[str, int]:
    ensure_input_dirs(samples_dir, annotations_dir)
    sample_files = collect_sample_files(samples_dir)
    if not sample_files:
        raise RuntimeError(f"样本目录为空: {samples_dir}")

    run_id = datetime.now().strftime("v14_%Y%m%d_%H%M%S")
    extracted_dir = out_dir / "extracted"
    report_dir = out_dir / "reports"
    review_dir = out_dir / "review"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)
    review_dir.mkdir(parents=True, exist_ok=True)

    tagger = Tagger(mode=model_mode)
    default_model_name = "mock-rule-engine" if model_mode == "mock" else (tagger.model_name or "unknown")

    report_rows: List[Dict[str, object]] = []
    tag_rows: List[Dict[str, object]] = []
    evidence_rows: List[Dict[str, object]] = []
    parse_failure_reviews: List[Dict[str, object]] = []

    for path in sample_files:
        report_row = build_report_record(path)
        text, parse_error, parse_reason_code = extract_text(path)
        segments = segment_text(text) if not parse_error else []

        parse_status = PARSE_STATUS_FAILED if parse_error else PARSE_STATUS_SUCCESS
        report_row["run_id"] = run_id
        report_row["model_mode"] = model_mode
        report_row["model_name"] = default_model_name
        report_row["text_status"] = parse_status
        report_row["parse_status"] = parse_status
        report_row["parse_reason_code"] = parse_reason_code if parse_error else ""
        report_row["segment_count"] = len(segments)
        report_rows.append(report_row)

        if parse_error:
            parse_failure_reviews.append(build_review_item_for_parse_failure(report_row, parse_error))
            continue

        for idx, segment in enumerate(segments, start=1):
            segment_id = f"S{idx:03d}"
            cls = tagger.classify(segment, context=report_row)
            cls = _normalize_classification(cls)

            if not cls["is_ai_hit"] and cls["decision_status"] == "confirmed":
                continue

            tag_id = stable_hash(str(report_row["report_id"]), segment_id, segment)
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
                "file_path": report_row["file_path"],
            }
            tag_rows.append(tag_row)

            if cls["is_ai_hit"]:
                evidence_rows.append(
                    {
                        "evidence_id": stable_hash(str(report_row["report_id"]), segment_id, "evidence"),
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
                        "file_path": report_row["file_path"],
                    }
                )

    review_rows = build_review_queue_from_tags(tag_rows)
    review_rows.extend(parse_failure_reviews)
    review_rows = sorted(
        review_rows,
        key=lambda row: _priority_order(str(row.get("review_priority", "low"))),
    )

    write_jsonl(extracted_dir / "report_index.jsonl", report_rows)
    write_jsonl(extracted_dir / "tag_result.jsonl", strip_tag_rows(tag_rows))
    write_jsonl(extracted_dir / "evidence_span.jsonl", evidence_rows)
    write_jsonl(extracted_dir / "review_queue.jsonl", review_rows)
    write_review_csv(report_dir / "review_queue.csv", review_rows)
    write_review_csv(review_dir / "review_queue.csv", review_rows)
    write_review_result_template(review_dir / "review_result_template.csv", review_rows)

    summary = build_summary_markdown(
        report_rows=report_rows,
        tag_rows=tag_rows,
        evidence_rows=evidence_rows,
        review_rows=review_rows,
        run_meta={
            "run_id": run_id,
            "model_mode": model_mode,
            "model_name": default_model_name,
            "samples_dir": str(samples_dir.resolve()),
        },
    )
    write_markdown(report_dir / "AI专题摘要.md", summary)

    return {
        "run_id": run_id,
        "reports": len(report_rows),
        "tag_rows": len(tag_rows),
        "evidence_rows": len(evidence_rows),
        "review_rows": len(review_rows),
    }


def strip_tag_rows(rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    kept: List[Dict[str, object]] = []
    for row in rows:
        kept.append(
            {
                "tag_id": row["tag_id"],
                "report_id": row["report_id"],
                "segment_id": row["segment_id"],
                "is_ai_hit": row["is_ai_hit"],
                "business_line": row["business_line"],
                "ai_actor": row["ai_actor"],
                "actor_primary": row["actor_primary"],
                "actor_subtype": row["actor_subtype"],
                "ai_scope": row["ai_scope"],
                "interaction_outcome": row["interaction_outcome"],
                "certainty_level": row["certainty_level"],
                "review_reason_code": row["review_reason_code"],
                "decision_status": row["decision_status"],
                "confidence": row["confidence"],
                "reason": row["reason"],
                "run_id": row["run_id"],
                "model_mode": row["model_mode"],
                "model_name": row["model_name"],
                "parse_status": row["parse_status"],
                "file_path": row["file_path"],
            }
        )
    return kept


def write_jsonl(path: Path, rows: Iterable[Dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")


def write_review_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    fieldnames = [
        "review_id",
        "report_id",
        "segment_id",
        "review_reason",
        "review_reason_code",
        "review_priority",
        "decision_status",
        "current_decision_status",
        "model_mode",
        "model_name",
        "parse_status",
        "source_text",
        "file_path",
    ]
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})


def write_review_result_template(path: Path, review_rows: List[Dict[str, object]]) -> None:
    fieldnames = [
        "sample_id",
        "report_id",
        "segment_id",
        "expected",
        "actual",
        "is_pass",
        "review_comment",
        "need_rule_update",
        "need_prompt_update",
        "need_annotation_update",
    ]
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in review_rows:
            writer.writerow(
                {
                    "sample_id": "",
                    "report_id": row.get("report_id", ""),
                    "segment_id": row.get("segment_id", ""),
                    "expected": "",
                    "actual": row.get("current_decision_status", row.get("decision_status", "")),
                    "is_pass": "",
                    "review_comment": "",
                    "need_rule_update": "",
                    "need_prompt_update": "",
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


def _priority_order(priority: str) -> int:
    mapping = {"high": 0, "medium": 1, "low": 2}
    return mapping.get(priority, 3)


def main() -> None:
    args = parse_args()
    result = run_pipeline(
        samples_dir=Path(args.samples),
        annotations_dir=Path(args.annotations),
        out_dir=Path(args.out),
        model_mode=args.model_mode,
    )
    print(
        "v1.4 分析完成: run_id={run_id}, reports={reports}, tags={tag_rows}, evidence={evidence_rows}, review={review_rows}".format(
            **result
        )
    )


if __name__ == "__main__":
    main()
