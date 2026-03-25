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
                "review_reason_code": row.get("review_reason_code", ""),
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
        "review_reason_code": "PARSE_FAILED",
        "decision_status": "pending_human_review",
        "source_text": "",
        "file_path": report_row["file_path"],
    }


def _infer_review_reason(tag_row: Dict[str, object]) -> str:
    code_text = str(tag_row.get("review_reason_code", ""))
    codes = [c for c in code_text.split(";") if c]
    if not codes:
        return "建议人工复核"

    code_map = {
        "ACTOR_OVERLAP": "主体复合或归属不稳定",
        "SCOPE_AMBIGUOUS": "范围口径不稳定",
        "BROAD_STATEMENT": "表达宽泛，需人工确认",
        "BUSINESSLINE_LOW_SIGNAL": "业务线信号不足",
        "OUTCOME_UNCLEAR": "互动结果不清晰",
        "PARSE_FAILED": "解析失败",
    }
    reasons = [code_map.get(code, code) for code in codes]
    return "；".join(reasons)
