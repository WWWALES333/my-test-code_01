from __future__ import annotations

import json
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

from src.analysis_v14.schema import stable_hash

from .schema import (
    ACTIONABILITY_LABELS,
    BUSINESS_QUESTION_LABELS,
    DOCTOR_ACCEPTANCE_LABELS,
    DOCTOR_NEED_LABELS,
)

DEFAULT_BATCH_SIZE = 20


def load_review_decisions(path: Path) -> Dict[str, Dict[str, object]]:
    if not path.exists():
        return {}
    decisions: Dict[str, Dict[str, object]] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        row = json.loads(line)
        task_id = str(row.get("task_id", "")).strip()
        if task_id:
            decisions[task_id] = row
    return decisions


def apply_review_decisions_to_facts(
    business_facts: Sequence[Dict[str, object]],
    decisions: Dict[str, Dict[str, object]],
) -> List[Dict[str, object]]:
    """把人工复核后的结果回流到业务问题层。

    这里不自动修改规则或 Prompt，只让本轮输出消费已确认的复核结论。
    """
    if not decisions:
        return [dict(row) for row in business_facts]

    rows: List[Dict[str, object]] = []
    for row in business_facts:
        item = dict(row)
        task_id = stable_hash(str(item.get("business_fact_id", "")), "v16-review-task")
        decision = decisions.get(task_id)
        if not decision:
            rows.append(item)
            continue

        final_fields = dict(decision.get("final_fields", {}))
        final_question = str(final_fields.get("final_business_question", "")).strip()
        if final_question:
            item["business_question"] = final_question
            item["business_question_label"] = BUSINESS_QUESTION_LABELS.get(final_question, final_question)
        for key in ("doctor_acceptance_level", "doctor_need_type", "sales_ai_usage_type", "competitor_signal_type"):
            if str(final_fields.get(key, "")).strip():
                item[key] = final_fields[key]

        report_worthy = str(final_fields.get("is_report_worthy", "")).strip()
        business_value = str(final_fields.get("business_value", "")).strip()
        if report_worthy == "yes":
            item["actionability"] = "report_ready"
        elif report_worthy == "no" or business_value == "low":
            item["actionability"] = "no_action"
        elif report_worthy == "observe":
            item["actionability"] = "observe"

        item["review_status"] = "reviewed"
        item["reviewed_at"] = decision.get("reviewed_at", "")
        item["reviewer"] = decision.get("reviewer", "")
        item["review_comment"] = decision.get("review_comment", "")
        item["should_review"] = False
        item["confidence"] = max(float(item.get("confidence", 0.0) or 0.0), 0.85)
        rows.append(item)
    return rows


def build_review_batch(
    business_facts: Sequence[Dict[str, object]],
    decisions: Dict[str, Dict[str, object]] | None = None,
    *,
    batch_size: int = DEFAULT_BATCH_SIZE,
) -> List[Dict[str, object]]:
    decisions = decisions or {}
    candidates = [_build_task(row, decisions) for row in business_facts]
    open_candidates = [row for row in candidates if str(row.get("task_status", "")) == "open"]
    ranked = sorted(open_candidates, key=_batch_priority, reverse=True)[:batch_size]
    for idx, row in enumerate(ranked, start=1):
        row["batch_id"] = f"v16_review_batch_{datetime.now().strftime('%Y%m%d')}_001"
        row["batch_position"] = idx
        row["batch_size"] = batch_size
    reviewed = [row for row in candidates if str(row.get("task_status", "")) == "reviewed"]
    return [*ranked, *reviewed]


def build_learning_outputs(
    decisions: Dict[str, Dict[str, object]],
    review_batch: Sequence[Dict[str, object]],
) -> Tuple[str, List[Dict[str, object]], List[Dict[str, object]], List[Dict[str, object]], List[Dict[str, object]]]:
    enriched = _attach_current_tasks(decisions, review_batch)
    rule_candidates = _build_candidates(enriched, "rule")
    prompt_candidates = _build_candidates(enriched, "prompt")
    label_candidates = _build_candidates(enriched, "label_gap")
    golden_set = _build_golden_set(enriched)
    summary = _build_learning_summary(enriched, rule_candidates, prompt_candidates, label_candidates)
    return summary, rule_candidates, prompt_candidates, label_candidates, golden_set


def validate_review_payload(task: Dict[str, object], reviewed_fields: Dict[str, object]) -> List[str]:
    errors: List[str] = []
    if not str(reviewed_fields.get("business_value", "")).strip():
        errors.append("必须判断这条内容是否有业务价值")
    if not str(reviewed_fields.get("final_business_question", "")).strip():
        errors.append("必须确认业务问题归属")
    if not str(reviewed_fields.get("is_report_worthy", "")).strip():
        errors.append("必须判断是否值得进入报告或后续行动")
    return errors


def apply_review_decision(
    review_dir: Path,
    task: Dict[str, object],
    reviewed_fields: Dict[str, object],
    *,
    reviewer: str = "wales",
    reviewed_at: str | None = None,
    review_comment: str = "",
) -> Dict[str, object]:
    review_dir.mkdir(parents=True, exist_ok=True)
    current_fields = dict(task.get("current_fields", {}))
    final_fields = _normalize_reviewed_fields(current_fields, reviewed_fields)
    decision = {
        "task_id": str(task.get("task_id", "")),
        "evidence_id": str(task.get("evidence_id", "")),
        "report_id": str(task.get("report_id", "")),
        "segment_id": str(task.get("segment_id", "")),
        "reviewer": reviewer,
        "reviewed_at": reviewed_at or datetime.now().isoformat(timespec="seconds"),
        "review_comment": review_comment,
        "current_fields": current_fields,
        "final_fields": final_fields,
        "system_inferred_error_reason": _infer_error_reason(current_fields, final_fields),
        "system_inferred_update_type": _infer_update_type(current_fields, final_fields),
        "change_diff": _build_change_diff(current_fields, final_fields),
        "source_text": str(task.get("source_text", "")),
        "file_path": str(task.get("file_path", "")),
    }
    with (review_dir / "review_decisions.jsonl").open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(decision, ensure_ascii=False) + "\n")
    return decision


def _build_task(row: Dict[str, object], decisions: Dict[str, Dict[str, object]]) -> Dict[str, object]:
    task_id = stable_hash(str(row.get("business_fact_id", "")), "v16-review-task")
    decision = decisions.get(task_id)
    current_fields = {
        "business_question": row.get("business_question", ""),
        "doctor_acceptance_level": row.get("doctor_acceptance_level", ""),
        "doctor_need_type": row.get("doctor_need_type", ""),
        "sales_ai_usage_type": row.get("sales_ai_usage_type", ""),
        "competitor_signal_type": row.get("competitor_signal_type", ""),
        "actionability": row.get("actionability", ""),
        "confidence": row.get("confidence", 0),
        "should_review": row.get("should_review", False),
    }
    return {
        "task_id": task_id,
        "task_status": "reviewed" if decision else "open",
        "evidence_id": row.get("evidence_id", ""),
        "business_fact_id": row.get("business_fact_id", ""),
        "report_id": row.get("report_id", ""),
        "segment_id": row.get("segment_id", ""),
        "salesperson_name": row.get("salesperson_name", ""),
        "battle_zone_name": row.get("battle_zone_name", ""),
        "region_name": row.get("region_name", ""),
        "year": row.get("year", 0),
        "month": row.get("month", 0),
        "week_of_month": row.get("week_of_month", 0),
        "selection_reason": _selection_reason(row),
        "review_guidance": "只判断业务价值、业务问题归属、是否可进报告；错因和优化方式由系统自动归因。",
        "current_fields": current_fields,
        "final_fields": dict(decision.get("final_fields", {})) if decision else {},
        "business_question_label": BUSINESS_QUESTION_LABELS.get(str(row.get("business_question", "")), str(row.get("business_question", ""))),
        "doctor_acceptance_label": DOCTOR_ACCEPTANCE_LABELS.get(str(row.get("doctor_acceptance_level", "")), str(row.get("doctor_acceptance_level", ""))),
        "doctor_need_label": DOCTOR_NEED_LABELS.get(str(row.get("doctor_need_type", "")), str(row.get("doctor_need_type", ""))),
        "actionability_label": ACTIONABILITY_LABELS.get(str(row.get("actionability", "")), str(row.get("actionability", ""))),
        "source_text": row.get("source_text", ""),
        "file_path": row.get("file_path", ""),
        "review_comment": str(decision.get("review_comment", "")) if decision else "",
        "reviewed_at": str(decision.get("reviewed_at", "")) if decision else "",
        "reviewer": str(decision.get("reviewer", "")) if decision else "",
    }


def _batch_priority(row: Dict[str, object]) -> Tuple[float, int, int, int]:
    fields = dict(row.get("current_fields", {}))
    confidence = float(fields.get("confidence", 0.0) or 0.0)
    should_review = 1 if bool(fields.get("should_review", False)) else 0
    actionable = 1 if str(fields.get("actionability", "")) in {"report_ready", "actionable"} else 0
    text_len = min(len(str(row.get("source_text", ""))), 220)
    low_conf_bonus = 1.0 - confidence
    return (should_review + actionable + low_conf_bonus, text_len, int(row.get("year", 0)), int(row.get("month", 0)))


def _selection_reason(row: Dict[str, object]) -> str:
    reasons = []
    if bool(row.get("should_review", False)):
        reasons.append("系统判断需要复核")
    if str(row.get("actionability", "")) in {"report_ready", "actionable"}:
        reasons.append("可能有报告或行动价值")
    if float(row.get("confidence", 0.0) or 0.0) < 0.65:
        reasons.append("置信度偏低")
    return "；".join(reasons) if reasons else "代表性样本"


def _attach_current_tasks(decisions: Dict[str, Dict[str, object]], review_batch: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    task_map = {str(row.get("task_id", "")): row for row in review_batch}
    rows: List[Dict[str, object]] = []
    for task_id, decision in decisions.items():
        item = dict(decision)
        item["task"] = task_map.get(task_id, {})
        rows.append(item)
    return rows


def _build_candidates(rows: Sequence[Dict[str, object]], update_type: str) -> List[Dict[str, object]]:
    grouped: Dict[Tuple[str, str, str], Dict[str, object]] = {}
    for row in rows:
        inferred = str(row.get("system_inferred_update_type", ""))
        if update_type == "label_gap":
            include = "label_gap" in str(row.get("system_inferred_error_reason", ""))
        else:
            include = inferred == update_type
        if not include:
            continue
        current = dict(row.get("current_fields", {}))
        final = dict(row.get("final_fields", {}))
        key = (
            update_type,
            str(row.get("system_inferred_error_reason", "")),
            json.dumps({"current": current, "final": final}, ensure_ascii=False, sort_keys=True),
        )
        bucket = grouped.setdefault(
            key,
            {
                "candidate_id": stable_hash(*key),
                "update_type": update_type,
                "count": 0,
                "error_reason": key[1],
                "current_pattern": current,
                "final_pattern": final,
                "sample_task_ids": [],
                "sample_texts": [],
                "priority": "observe",
            },
        )
        bucket["count"] = int(bucket["count"]) + 1
        bucket["sample_task_ids"].append(str(row.get("task_id", "")))
        if len(bucket["sample_texts"]) < 3:
            bucket["sample_texts"].append(str(row.get("source_text", ""))[:220])
    for bucket in grouped.values():
        bucket["priority"] = "high" if int(bucket["count"]) >= 3 else "observe"
    return sorted(grouped.values(), key=lambda item: (-int(item.get("count", 0)), str(item.get("candidate_id", ""))))


def _build_golden_set(rows: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    golden = []
    for row in rows:
        final_fields = dict(row.get("final_fields", {}))
        if str(final_fields.get("business_value", "")) in {"high", "medium"} or str(final_fields.get("is_report_worthy", "")) == "yes":
            golden.append(
                {
                    "golden_id": stable_hash(str(row.get("task_id", "")), "golden"),
                    "task_id": row.get("task_id", ""),
                    "source_text": row.get("source_text", ""),
                    "expected_fields": final_fields,
                    "review_comment": row.get("review_comment", ""),
                }
            )
    return golden


def _build_learning_summary(
    rows: Sequence[Dict[str, object]],
    rule_candidates: Sequence[Dict[str, object]],
    prompt_candidates: Sequence[Dict[str, object]],
    label_candidates: Sequence[Dict[str, object]],
) -> str:
    error_counts = Counter(str(row.get("system_inferred_error_reason", "")) for row in rows)
    lines = [
        "# v1.6 复核学习摘要",
        "",
        f"- 已复核样本数：{len(rows)}",
        f"- 规则候选：{len(rule_candidates)}",
        f"- Prompt 候选：{len(prompt_candidates)}",
        f"- 标签扩展候选：{len(label_candidates)}",
        "",
        "## 主要错因",
    ]
    if error_counts:
        for reason, count in error_counts.most_common(8):
            lines.append(f"- {reason or '未归因'}：{count}")
    else:
        lines.append("- 暂无复核样本。")
    lines.extend(["", "## 下一步建议"])
    if prompt_candidates:
        lines.append("- 优先检查 Prompt 边界判定，避免把业务问题识别留给硬规则。")
    if rule_candidates:
        lines.append("- 对高频明确错例补充规则拦截，但不要继续无限堆关键词。")
    if label_candidates:
        lines.append("- 对高频 label_gap 样本评估是否需要扩展标签体系。")
    if not rows:
        lines.append("- 先完成第一轮 20 条复核。")
    return "\n".join(lines)


def _normalize_reviewed_fields(current: Dict[str, object], reviewed: Dict[str, object]) -> Dict[str, object]:
    final = dict(current)
    final.update({key: value for key, value in reviewed.items() if value not in {None, ""}})
    return final


def _infer_error_reason(current: Dict[str, object], final: Dict[str, object]) -> str:
    if current.get("business_question") != final.get("final_business_question", final.get("business_question")):
        return "business_definition_gap"
    if current.get("doctor_acceptance_level") != final.get("doctor_acceptance_level"):
        return "prompt_issue"
    if current.get("doctor_need_type") != final.get("doctor_need_type"):
        return "prompt_issue"
    if str(final.get("final_business_question", final.get("business_question", ""))) == "label_gap":
        return "label_gap"
    if str(final.get("business_value", "")) == "low":
        return "low_value_noise"
    return "other"


def _infer_update_type(current: Dict[str, object], final: Dict[str, object]) -> str:
    reason = _infer_error_reason(current, final)
    if reason in {"prompt_issue", "business_definition_gap"}:
        return "prompt"
    if reason == "label_gap":
        return "annotation"
    if reason == "low_value_noise":
        return "rule"
    return "observe"


def _build_change_diff(current: Dict[str, object], final: Dict[str, object]) -> Dict[str, object]:
    diff = {}
    for key in sorted(set(current) | set(final)):
        if current.get(key) != final.get(key):
            diff[key] = {"before": current.get(key, ""), "after": final.get(key, "")}
    return diff
