from __future__ import annotations

from typing import Dict, List

from .schema import DECISION_CONFIRMED, stable_hash


def build_review_queue_from_tags(tag_rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    review_items: List[Dict[str, object]] = []
    for row in tag_rows:
        if row["decision_status"] == DECISION_CONFIRMED:
            continue
        review_reason = _infer_review_reason(row)
        review_items.append(
            {
                "review_id": stable_hash(str(row["tag_id"]), "review"),
                "report_id": row["report_id"],
                "segment_id": row["segment_id"],
                "review_reason": review_reason,
                "decision_status": row["decision_status"],
                "source_text": row["source_text"],
                "file_path": row["file_path"],
            }
        )
    return review_items


def build_review_item_for_parse_failure(report_row: Dict[str, object], reason: str) -> Dict[str, object]:
    return {
        "review_id": stable_hash(str(report_row["report_id"]), "parse_failed"),
        "report_id": report_row["report_id"],
        "segment_id": "",
        "review_reason": f"解析失败: {reason}",
        "decision_status": "pending_human_review",
        "source_text": "",
        "file_path": report_row["file_path"],
    }


def _infer_review_reason(tag_row: Dict[str, object]) -> str:
    reasons = []
    if tag_row["business_line"] == "待判断":
        reasons.append("业务线不稳定")
    if tag_row["ai_actor"] == "待判断":
        reasons.append("主体不稳定")
    if tag_row["decision_status"] == "uncertain":
        reasons.append("置信度不足")
    return "；".join(reasons) if reasons else "建议人工复核"

