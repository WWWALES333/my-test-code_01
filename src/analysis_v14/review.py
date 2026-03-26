from __future__ import annotations

from typing import Dict, List

from .schema import DECISION_CONFIRMED, stable_hash


def build_review_queue_from_tags(tag_rows: List[Dict[str, object]]) -> List[Dict[str, object]]:
    review_items: List[Dict[str, object]] = []
    for row in tag_rows:
        if row["decision_status"] == DECISION_CONFIRMED:
            continue
        review_reason = _infer_review_reason(row)
        reason_code = str(row.get("review_reason_code", ""))
        review_items.append(
            {
                "review_id": stable_hash(str(row["tag_id"]), "review"),
                "report_id": row["report_id"],
                "segment_id": row["segment_id"],
                "review_reason": review_reason,
                "review_reason_code": reason_code,
                "review_priority": _infer_priority(reason_code),
                "decision_status": row["decision_status"],
                "current_decision_status": row["decision_status"],
                "model_mode": row.get("model_mode", ""),
                "model_name": row.get("model_name", ""),
                "parse_status": row.get("parse_status", ""),
                "source_text": row["source_text"],
                "file_path": row["file_path"],
            }
        )
    return review_items


def build_review_item_for_parse_failure(report_row: Dict[str, object], reason: str) -> Dict[str, object]:
    parse_reason_code = str(report_row.get("parse_reason_code", "PARSE_FAILED"))
    return {
        "review_id": stable_hash(str(report_row["report_id"]), "parse_failed"),
        "report_id": report_row["report_id"],
        "segment_id": "",
        "review_reason": f"解析失败: {reason}",
        "review_reason_code": parse_reason_code,
        "review_priority": _infer_priority(parse_reason_code),
        "decision_status": "pending_human_review",
        "current_decision_status": "pending_human_review",
        "model_mode": report_row.get("model_mode", ""),
        "model_name": report_row.get("model_name", ""),
        "parse_status": report_row.get("text_status", "failed"),
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
        "PARSE_FAILED_DOC": "DOC 解析失败",
        "PARSE_FAILED_PDF": "PDF 解析失败",
        "PARSER_TOOL_MISSING": "解析工具缺失",
        "MODEL_CALL_FAILED": "模型调用失败",
    }
    reasons = [code_map.get(code, code) for code in codes]
    return "；".join(reasons)


def _infer_priority(code_text: str) -> str:
    codes = [c for c in code_text.split(";") if c]
    if any(code in {"MODEL_CALL_FAILED", "PARSER_TOOL_MISSING"} for code in codes):
        return "high"
    if any(code in {"PARSE_FAILED", "PARSE_FAILED_DOC", "PARSE_FAILED_PDF"} for code in codes):
        return "high"
    if any(code in {"ACTOR_OVERLAP", "SCOPE_AMBIGUOUS"} for code in codes):
        return "medium"
    return "low"
