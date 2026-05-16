from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

from .owner import build_owner_lookup, resolve_owner
from .schema import (
    ACTIONABILITY_VALUES,
    ACTION_BUCKET_VALUES,
    DECISION_CONFIRMED,
    LEARNING_ERROR_REASON_VALUES,
    REVIEW_NECESSITY_VALUES,
    TASK_STATUS_OPEN,
    TASK_STATUS_REVIEWED,
    stable_hash,
)

REVIEW_EDITABLE_FIELDS = (
    "is_ai_hit",
    "business_line",
    "actor_primary",
    "ai_scope",
    "decision_status",
    "review_comment",
)

REVIEW_LEARNING_FIELDS = (
    "error_reason_primary",
    "review_necessity",
    "actionability",
    "action_bucket",
    "need_rule_update",
    "need_prompt_update",
    "need_annotation_update",
    "learning_note",
)

REVIEW_UPDATE_FLAG_FIELDS = (
    "need_rule_update",
    "need_prompt_update",
    "need_annotation_update",
)

REVIEW_BATCH_SIZE = 20
SYSTEM_REVIEW_REASON_CODES = {
    "MODEL_CALL_FAILED",
    "PARSER_TOOL_MISSING",
    "PARSE_FAILED",
    "PARSE_FAILED_DOC",
    "PARSE_FAILED_PDF",
}
QUEUE_TYPE_BUSINESS = "business"
QUEUE_TYPE_SYSTEM = "system"


def load_review_decisions(
    annotations_dir: Path,
    review_dir: Path | None = None,
) -> Dict[Tuple[str, str], Dict[str, object]]:
    """读取人工复核回写结果，支持历史 CSV 和 v1.5 review jsonl。"""
    decisions: Dict[Tuple[str, str], Dict[str, object]] = {}
    if annotations_dir.exists():
        for path in sorted(annotations_dir.rglob("*.csv")):
            _consume_review_csv(path, decisions)
    if review_dir and review_dir.exists():
        jsonl_path = review_dir / "review_decisions.jsonl"
        if jsonl_path.exists():
            _consume_review_jsonl(jsonl_path, decisions)
    for key, value in list(decisions.items()):
        decisions[key] = _normalize_loaded_decision(value)
    return decisions


def build_review_tasks(
    review_rows: Iterable[Dict[str, object]],
    tag_rows: Sequence[Dict[str, object]],
    report_facts: Sequence[Dict[str, object]],
    review_decisions: Dict[Tuple[str, str], Dict[str, object]],
    owner_registry: Sequence[Dict[str, object]],
) -> List[Dict[str, object]]:
    """构建业务复核任务对象。系统失败任务会单独拆出。"""
    sorted_tasks = _sort_and_batch_tasks(_build_task_objects(review_rows, tag_rows, report_facts, review_decisions, owner_registry, queue_type=QUEUE_TYPE_BUSINESS))
    return sorted_tasks


def build_system_review_tasks(
    review_rows: Iterable[Dict[str, object]],
    tag_rows: Sequence[Dict[str, object]],
    report_facts: Sequence[Dict[str, object]],
    review_decisions: Dict[Tuple[str, str], Dict[str, object]],
    owner_registry: Sequence[Dict[str, object]],
) -> List[Dict[str, object]]:
    """构建系统异常任务对象，不进入业务复核工作台。"""
    tasks = _build_task_objects(review_rows, tag_rows, report_facts, review_decisions, owner_registry, queue_type=QUEUE_TYPE_SYSTEM)
    return sorted(
        tasks,
        key=lambda item: (
            str(item.get("task_status", "")),
            -_review_priority(item),
            int(item.get("year", 0)),
            int(item.get("month", 0)),
            str(item.get("task_id", "")),
        ),
    )


def _build_task_objects(
    review_rows: Iterable[Dict[str, object]],
    tag_rows: Sequence[Dict[str, object]],
    report_facts: Sequence[Dict[str, object]],
    review_decisions: Dict[Tuple[str, str], Dict[str, object]],
    owner_registry: Sequence[Dict[str, object]],
    queue_type: str,
) -> List[Dict[str, object]]:
    report_map = {str(row["report_id"]): row for row in report_facts}
    tag_map = {(str(row.get("report_id", "")), str(row.get("segment_id", ""))): row for row in tag_rows}
    owner_lookup = build_owner_lookup(owner_registry)
    tasks: List[Dict[str, object]] = []
    for row in review_rows:
        review_reason_code = str(row.get("review_reason_code", ""))
        task_queue_type = _infer_queue_type(review_reason_code)
        if task_queue_type != queue_type:
            continue
        report_id = str(row.get("report_id", ""))
        segment_id = str(row.get("segment_id", ""))
        tag_row = tag_map.get((report_id, segment_id), {})
        report = report_map.get(report_id, {})
        decision = review_decisions.get((report_id, segment_id))
        owner_hint = str(tag_row.get("segment_owner_hint", "")).strip() or str(row.get("segment_owner_hint", "")).strip() or str(report.get("report_owner_hint", "")).strip()
        owner = resolve_owner(owner_hint, owner_lookup)

        current_fields = {
            "is_ai_hit": bool(tag_row.get("is_ai_hit", True)),
            "business_line": str(tag_row.get("business_line", row.get("business_line", ""))),
            "actor_primary": str(tag_row.get("actor_primary", tag_row.get("ai_actor", ""))),
            "ai_scope": str(tag_row.get("ai_scope", "")),
            "decision_status": str(tag_row.get("decision_status", row.get("current_decision_status", row.get("decision_status", "")))),
            "review_comment": str(decision.get("review_comment", "")) if decision else "",
        }
        edited_labels = dict(decision.get("final_labels", {})) if decision else {}
        learning_fields = dict(decision.get("learning_fields", _default_learning_fields())) if decision else _default_learning_fields()
        current_labels = (
            _normalize_reviewed_fields(dict(decision.get("current_labels", current_fields)))
            if decision
            else _normalize_reviewed_fields(current_fields)
        )
        tasks.append(
            {
                "task_id": str(row.get("review_id", stable_hash(report_id, segment_id, "task"))),
                "review_id": str(row.get("review_id", "")),
                "queue_type": task_queue_type,
                "report_id": report_id,
                "segment_id": segment_id,
                "salesperson_id": str(owner.get("salesperson_id", report.get("report_owner_id", ""))),
                "salesperson_name": str(owner.get("display_name", owner.get("salesperson_name", report.get("report_owner_name", "")))),
                "owner_type": str(owner.get("owner_type", report.get("report_owner_type", ""))),
                "battle_zone_name": str(owner.get("battle_zone_name", "")),
                "region_name": str(owner.get("region_name", "")),
                "owner_hint": owner_hint,
                "report_owner_name": str(report.get("report_owner_name", "")),
                "year": int(report.get("year", 0)),
                "month": int(report.get("month", 0)),
                "week_of_month": int(report.get("week_of_month", 0)),
                "review_reason_code": review_reason_code,
                "review_reason": str(row.get("review_reason", "")),
                "task_status": TASK_STATUS_REVIEWED if decision else TASK_STATUS_OPEN,
                "current_fields": current_fields,
                "current_labels": current_labels,
                "edited_fields": edited_labels,
                "learning_fields": learning_fields,
                "source_text": str(row.get("source_text", tag_row.get("source_text", ""))),
                "source_context": str(row.get("source_text", tag_row.get("source_text", ""))),
                "file_path": str(row.get("file_path", tag_row.get("file_path", ""))),
                "review_comment": str(decision.get("review_comment", "")) if decision else "",
                "reviewer": str(decision.get("reviewer", "")) if decision else "",
                "reviewed_at": str(decision.get("reviewed_at", "")) if decision else "",
                "change_diff": _build_change_diff(current_labels, edited_labels, _default_learning_fields(), learning_fields),
            }
        )
    return tasks


def _sort_and_batch_tasks(tasks: List[Dict[str, object]]) -> List[Dict[str, object]]:
    sorted_tasks = sorted(
        tasks,
        key=lambda item: (
            str(item.get("task_status", "")),
            -_review_priority(item),
            int(item.get("year", 0)),
            int(item.get("month", 0)),
            str(item.get("task_id", "")),
        ),
    )
    for index, task in enumerate(sorted_tasks, start=1):
        batch_number = ((index - 1) // REVIEW_BATCH_SIZE) + 1
        task["batch_id"] = f"review_batch_{batch_number:03d}"
        task["batch_number"] = batch_number
        task["batch_position"] = index - ((batch_number - 1) * REVIEW_BATCH_SIZE)
        task["batch_size"] = REVIEW_BATCH_SIZE
    return sorted_tasks


def build_review_audit_log(review_decisions: Dict[Tuple[str, str], Dict[str, object]]) -> List[Dict[str, object]]:
    """产出复核提交审计日志。"""
    rows: List[Dict[str, object]] = []
    for decision in review_decisions.values():
        final_labels = dict(decision.get("final_labels", {}))
        rows.append(
            {
                "task_id": str(decision.get("task_id", stable_hash(str(decision.get("report_id", "")), str(decision.get("segment_id", "")), "task"))),
                "batch_id": str(decision.get("batch_id", "")),
                "report_id": str(decision.get("report_id", "")),
                "segment_id": str(decision.get("segment_id", "")),
                "reviewer": str(decision.get("reviewer", "")),
                "reviewed_at": str(decision.get("reviewed_at", "")),
                "review_comment": str(decision.get("review_comment", "")),
                "final_labels": final_labels,
                "learning_fields": dict(decision.get("learning_fields", {})),
                "change_diff": dict(decision.get("change_diff", {})),
                "review_reason_code": str(decision.get("review_reason_code", "")),
                "source_file": str(decision.get("source_file", "")),
            }
        )
    return sorted(rows, key=lambda item: (str(item.get("reviewed_at", "")), str(item.get("task_id", ""))))


def build_review_learning_candidates(review_decisions: Dict[Tuple[str, str], Dict[str, object]]) -> List[Dict[str, object]]:
    grouped: Dict[Tuple[str, str, str, str, str], Dict[str, object]] = {}
    for decision in review_decisions.values():
        learning_fields = dict(decision.get("learning_fields", {}))
        current_labels = dict(decision.get("current_labels", {}))
        final_labels = dict(decision.get("final_labels", {}))
        review_reason_code = str(decision.get("review_reason_code", "")).strip()
        error_reason_primary = str(learning_fields.get("error_reason_primary", "")).strip()
        for update_type in _resolve_update_types(decision):
            key = (
                update_type,
                error_reason_primary,
                review_reason_code,
                json.dumps(current_labels, ensure_ascii=False, sort_keys=True),
                json.dumps(final_labels, ensure_ascii=False, sort_keys=True),
            )
            bucket = grouped.setdefault(
                key,
                {
                    "candidate_id": stable_hash(*key),
                    "update_type": update_type,
                    "count": 0,
                    "error_reason_primary": error_reason_primary,
                    "review_reason_code": review_reason_code,
                    "current_labels": current_labels,
                    "final_labels": final_labels,
                    "sample_task_ids": [],
                    "sample_texts": [],
                    "latest_reviewed_at": "",
                },
            )
            bucket["count"] = int(bucket.get("count", 0)) + 1
            if len(bucket["sample_task_ids"]) < 5:
                bucket["sample_task_ids"].append(str(decision.get("task_id", "")))
            sample_text = str(decision.get("source_text", "")).strip()
            if sample_text and len(bucket["sample_texts"]) < 3:
                bucket["sample_texts"].append(sample_text)
            reviewed_at = str(decision.get("reviewed_at", ""))
            if reviewed_at >= str(bucket.get("latest_reviewed_at", "")):
                bucket["latest_reviewed_at"] = reviewed_at
    candidates = list(grouped.values())
    for row in candidates:
        row["priority_level"] = "high" if int(row.get("count", 0)) >= 3 else "observe"
    return sorted(
        candidates,
        key=lambda item: (-int(item.get("count", 0)), str(item.get("update_type", "")), str(item.get("candidate_id", ""))),
    )


def build_review_learning_summary(
    review_decisions: Dict[Tuple[str, str], Dict[str, object]],
    candidates: Sequence[Dict[str, object]],
) -> Dict[str, object]:
    error_reason_counter: Dict[str, int] = {}
    review_necessity_counter: Dict[str, int] = {}
    actionability_counter: Dict[str, int] = {}
    update_counter = {flag: 0 for flag in REVIEW_UPDATE_FLAG_FIELDS}
    for decision in review_decisions.values():
        learning_fields = dict(decision.get("learning_fields", {}))
        error_reason = str(learning_fields.get("error_reason_primary", "")).strip()
        review_necessity = str(learning_fields.get("review_necessity", "")).strip()
        actionability = str(learning_fields.get("actionability", "")).strip()
        if error_reason:
            error_reason_counter[error_reason] = error_reason_counter.get(error_reason, 0) + 1
        if review_necessity:
            review_necessity_counter[review_necessity] = review_necessity_counter.get(review_necessity, 0) + 1
        if actionability:
            actionability_counter[actionability] = actionability_counter.get(actionability, 0) + 1
        for update_type in _resolve_update_types(decision):
            if update_type == "rule":
                update_counter["need_rule_update"] += 1
            elif update_type == "prompt":
                update_counter["need_prompt_update"] += 1
            elif update_type == "annotation":
                update_counter["need_annotation_update"] += 1
    return {
        "reviewed_decision_count": len(review_decisions),
        "top_error_reasons": sorted(error_reason_counter.items(), key=lambda item: (-item[1], item[0])),
        "review_necessity_breakdown": sorted(review_necessity_counter.items(), key=lambda item: (-item[1], item[0])),
        "actionability_breakdown": sorted(actionability_counter.items(), key=lambda item: (-item[1], item[0])),
        "update_request_count": update_counter,
        "candidate_count": len(candidates),
        "high_priority_candidate_count": sum(1 for row in candidates if str(row.get("priority_level", "")) == "high"),
    }


def build_review_batch_summaries(tasks: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    grouped: Dict[str, List[Dict[str, object]]] = {}
    for task in tasks:
        grouped.setdefault(str(task.get("batch_id", "")), []).append(task)
    rows: List[Dict[str, object]] = []
    for batch_id, items in sorted(grouped.items()):
        open_items = [item for item in items if str(item.get("task_status", "")) == TASK_STATUS_OPEN]
        reviewed_items = [item for item in items if str(item.get("task_status", "")) == TASK_STATUS_REVIEWED]
        error_reason_counter: Dict[str, int] = {}
        review_reason_counter: Dict[str, int] = {}
        update_counter = {flag: 0 for flag in REVIEW_UPDATE_FLAG_FIELDS}
        for item in reviewed_items:
            learning_fields = dict(item.get("learning_fields", {}))
            error_reason = str(learning_fields.get("error_reason_primary", "")).strip()
            review_reason = str(item.get("review_reason_code", "")).strip()
            if error_reason:
                error_reason_counter[error_reason] = error_reason_counter.get(error_reason, 0) + 1
            if review_reason:
                review_reason_counter[review_reason] = review_reason_counter.get(review_reason, 0) + 1
            for update_type in _resolve_update_types(item):
                if update_type == "rule":
                    update_counter["need_rule_update"] += 1
                elif update_type == "prompt":
                    update_counter["need_prompt_update"] += 1
                elif update_type == "annotation":
                    update_counter["need_annotation_update"] += 1
        rows.append(
            {
                "batch_id": batch_id,
                "batch_number": int(items[0].get("batch_number", 0)) if items else 0,
                "batch_size": int(items[0].get("batch_size", REVIEW_BATCH_SIZE)) if items else REVIEW_BATCH_SIZE,
                "task_count": len(items),
                "reviewed_count": len(reviewed_items),
                "open_count": len(open_items),
                "completion_rate": round(len(reviewed_items) / len(items), 4) if items else 0.0,
                "ready_for_optimization": len(reviewed_items) >= min(REVIEW_BATCH_SIZE, len(items)),
                "top_error_reasons": sorted(error_reason_counter.items(), key=lambda item: (-item[1], item[0]))[:5],
                "top_review_reasons": sorted(review_reason_counter.items(), key=lambda item: (-item[1], item[0]))[:5],
                "update_request_count": update_counter,
                "sample_open_task_ids": [str(item.get("task_id", "")) for item in open_items[:5]],
                "sample_reviewed_task_ids": [str(item.get("task_id", "")) for item in reviewed_items[:5]],
            }
        )
    return rows


def validate_review_submission(
    task: Dict[str, object],
    reviewed_fields: Dict[str, object],
    learning_fields: Dict[str, object],
) -> List[str]:
    errors: List[str] = []
    current_fields = dict(task.get("current_fields", {}))
    merged_reviewed_fields = dict(current_fields)
    merged_reviewed_fields.update({key: reviewed_fields.get(key) for key in REVIEW_EDITABLE_FIELDS if key in reviewed_fields})
    normalized_learning_fields = _normalize_learning_fields(learning_fields)

    if not str(normalized_learning_fields.get("review_necessity", "")).strip():
        errors.append("提交复核时必须填写 review_necessity。")
    if bool(merged_reviewed_fields.get("is_ai_hit")) and not str(normalized_learning_fields.get("actionability", "")).strip():
        errors.append("AI 命中条目必须填写 actionability。")
    if str(normalized_learning_fields.get("actionability", "")).strip() == "actionable" and not str(normalized_learning_fields.get("action_bucket", "")).strip():
        errors.append("actionability=actionable 时必须填写 action_bucket。")
    return errors


def apply_review_decision(
    review_dir: Path,
    task: Dict[str, object],
    reviewed_fields: Dict[str, object],
    reviewer: str,
    reviewed_at: str,
    review_comment: str,
    learning_fields: Dict[str, object] | None = None,
) -> Dict[str, object]:
    """将单条复核结果写回 review/review_decisions.jsonl。"""
    review_dir.mkdir(parents=True, exist_ok=True)
    jsonl_path = review_dir / "review_decisions.jsonl"
    existing = load_review_decisions(Path("/nonexistent"), review_dir=review_dir)
    key = (str(task.get("report_id", "")), str(task.get("segment_id", "")))
    current = dict(existing.get(key, {}))
    normalized_reviewed_fields = _normalize_reviewed_fields(reviewed_fields)
    normalized_learning_fields = _normalize_learning_fields(learning_fields or {})
    current_labels = _normalize_reviewed_fields(dict(task.get("current_fields", {})))
    if not str(normalized_learning_fields.get("error_reason_primary", "")).strip():
        normalized_learning_fields["error_reason_primary"] = _infer_error_reason(current_labels, normalized_reviewed_fields)
    current_learning_fields = dict(task.get("learning_fields", _default_learning_fields()))
    current.update(
        {
            "task_id": str(task.get("task_id", "")),
            "batch_id": str(task.get("batch_id", "")),
            "report_id": key[0],
            "segment_id": key[1],
            "review_comment": review_comment,
            "reviewer": reviewer,
            "reviewed_at": reviewed_at,
            "review_reason_code": str(task.get("review_reason_code", "")),
            "current_labels": current_labels,
            "final_labels": normalized_reviewed_fields,
            "learning_fields": normalized_learning_fields,
            "change_diff": _build_change_diff(current_labels, normalized_reviewed_fields, current_learning_fields, normalized_learning_fields),
            "source_text": str(task.get("source_text", "")),
            "source_file": str(jsonl_path),
        }
    )
    existing[key] = current
    with jsonl_path.open("w", encoding="utf-8") as file:
        for row in existing.values():
            file.write(json.dumps(row, ensure_ascii=False) + "\n")
    return current


def _consume_review_csv(path: Path, decisions: Dict[Tuple[str, str], Dict[str, object]]) -> None:
    with path.open("r", encoding="utf-8", newline="") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames or "report_id" not in reader.fieldnames or "segment_id" not in reader.fieldnames:
            return
        for row in reader:
            report_id = str(row.get("report_id", "")).strip()
            segment_id = str(row.get("segment_id", "")).strip()
            if not report_id:
                continue
            decisions[(report_id, segment_id)] = {
                "task_id": str(row.get("task_id", "")).strip(),
                "batch_id": str(row.get("batch_id", "")).strip(),
                "report_id": report_id,
                "segment_id": segment_id,
                "review_comment": str(row.get("review_comment", "")).strip(),
                "reviewer": str(row.get("reviewer", "")).strip(),
                "reviewed_at": str(row.get("reviewed_at", "")).strip(),
                "review_reason_code": str(row.get("review_reason_code", "")).strip(),
                "current_labels": _parse_json_dict(str(row.get("current_labels", "")).strip()),
                "final_labels": _parse_final_labels(row),
                "learning_fields": _parse_learning_fields(row),
                "change_diff": {},
                "source_text": str(row.get("source_text", "")).strip(),
                "raw_row": dict(row),
                "source_file": str(path),
            }


def _consume_review_jsonl(path: Path, decisions: Dict[Tuple[str, str], Dict[str, object]]) -> None:
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        row = json.loads(line)
        report_id = str(row.get("report_id", "")).strip()
        segment_id = str(row.get("segment_id", "")).strip()
        if not report_id:
            continue
        decisions[(report_id, segment_id)] = {
            "task_id": str(row.get("task_id", "")).strip(),
            "batch_id": str(row.get("batch_id", "")).strip(),
            "report_id": report_id,
            "segment_id": segment_id,
            "review_comment": str(row.get("review_comment", "")).strip(),
            "reviewer": str(row.get("reviewer", "")).strip(),
            "reviewed_at": str(row.get("reviewed_at", "")).strip(),
            "final_labels": dict(row.get("final_labels", {})),
            "review_reason_code": str(row.get("review_reason_code", "")).strip(),
            "current_labels": dict(row.get("current_labels", {})),
            "learning_fields": _normalize_learning_fields(row.get("learning_fields", {})),
            "change_diff": dict(row.get("change_diff", {})),
            "source_text": str(row.get("source_text", "")).strip(),
            "source_file": str(path),
        }


def _parse_final_labels(row: Dict[str, str]) -> Dict[str, object]:
    raw = str(row.get("final_labels", "")).strip()
    if raw:
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass
    final_labels: Dict[str, object] = {}
    for key in ("is_ai_hit", "business_line", "actor_primary", "ai_scope", "decision_status"):
        value = str(row.get(key, "")).strip()
        if value:
            final_labels[key] = value if key != "is_ai_hit" else value in {"1", "true", "True", "yes", "YES"}
    actual = str(row.get("actual", "")).strip()
    if actual and "decision_status" not in final_labels:
        final_labels["decision_status"] = actual
    return final_labels


def _build_change_diff(
    current_fields: Dict[str, object],
    edited_fields: Dict[str, object],
    current_learning_fields: Dict[str, object] | None = None,
    learning_fields: Dict[str, object] | None = None,
) -> Dict[str, Dict[str, Dict[str, object]]]:
    diff: Dict[str, Dict[str, Dict[str, object]]] = {"final_labels": {}, "learning_fields": {}}
    for key in REVIEW_EDITABLE_FIELDS:
        old = current_fields.get(key)
        new = edited_fields.get(key)
        if new is not None and old != new:
            diff["final_labels"][key] = {"before": old, "after": new}
    current_learning = current_learning_fields or _default_learning_fields()
    next_learning = learning_fields or _default_learning_fields()
    for key in REVIEW_LEARNING_FIELDS:
        old = current_learning.get(key)
        new = next_learning.get(key)
        if new is not None and old != new:
            diff["learning_fields"][key] = {"before": old, "after": new}
    return {key: value for key, value in diff.items() if value}


def _default_learning_fields() -> Dict[str, object]:
    return {
        "error_reason_primary": "",
        "review_necessity": "",
        "actionability": "",
        "action_bucket": "",
        "need_rule_update": False,
        "need_prompt_update": False,
        "need_annotation_update": False,
        "learning_note": "",
    }


def _resolve_update_types(decision: Dict[str, object]) -> List[str]:
    learning_fields = dict(decision.get("learning_fields", {}))
    explicit: List[str] = []
    if _parse_boolish(learning_fields.get("need_rule_update")):
        explicit.append("rule")
    if _parse_boolish(learning_fields.get("need_prompt_update", learning_fields.get("need_skill_update"))):
        explicit.append("prompt")
    if _parse_boolish(learning_fields.get("need_annotation_update")):
        explicit.append("annotation")
    if explicit:
        return explicit
    inferred = _infer_update_types(decision)
    return inferred


def _infer_update_types(decision: Dict[str, object]) -> List[str]:
    learning_fields = dict(decision.get("learning_fields", {}))
    error_reason = str(learning_fields.get("error_reason_primary", "")).strip()
    if error_reason == "label_gap":
        return ["annotation"]
    if error_reason in {"low_signal_noise", "rule_threshold_issue", "parser_or_segmentation_error"}:
        return ["rule"]
    if error_reason in {"context_loss", "model_misread"}:
        return ["prompt"]
    if error_reason in {"actor_boundary", "business_line_boundary", "ai_scope_boundary"}:
        return ["annotation"]
    if error_reason == "other":
        return []
    return []


def _infer_queue_type(review_reason_code: str) -> str:
    codes = {code.strip() for code in str(review_reason_code).split(";") if code.strip()}
    if codes & SYSTEM_REVIEW_REASON_CODES:
        return QUEUE_TYPE_SYSTEM
    return QUEUE_TYPE_BUSINESS


def _infer_error_reason(current_labels: Dict[str, object], final_labels: Dict[str, object]) -> str:
    current_is_ai = bool(current_labels.get("is_ai_hit"))
    final_is_ai = bool(final_labels.get("is_ai_hit"))
    current_actor = str(current_labels.get("actor_primary", "")).strip()
    final_actor = str(final_labels.get("actor_primary", "")).strip()
    current_line = str(current_labels.get("business_line", "")).strip()
    final_line = str(final_labels.get("business_line", "")).strip()
    current_scope = str(current_labels.get("ai_scope", "")).strip()
    final_scope = str(final_labels.get("ai_scope", "")).strip()
    current_status = str(current_labels.get("decision_status", "")).strip()
    final_status = str(final_labels.get("decision_status", "")).strip()

    if final_is_ai and not final_actor:
        return "label_gap"
    if current_actor != final_actor:
        return "actor_boundary"
    if current_line != final_line:
        return "business_line_boundary"
    if current_scope != final_scope:
        return "ai_scope_boundary"
    if current_is_ai and not final_is_ai:
        return "low_signal_noise"
    if not current_is_ai and final_is_ai:
        return "rule_threshold_issue"
    if current_status != final_status:
        return "context_loss"
    return "other"


def _normalize_reviewed_fields(reviewed_fields: Dict[str, object]) -> Dict[str, object]:
    normalized: Dict[str, object] = {}
    for key in REVIEW_EDITABLE_FIELDS:
        if key == "review_comment" or key not in reviewed_fields:
            continue
        value = reviewed_fields.get(key)
        if key == "is_ai_hit":
            normalized[key] = bool(value)
        else:
            normalized[key] = value
    return normalized


def _normalize_learning_fields(learning_fields: Dict[str, object]) -> Dict[str, object]:
    normalized = _default_learning_fields()
    if not isinstance(learning_fields, dict):
        return normalized
    error_reason = str(learning_fields.get("error_reason_primary", "")).strip()
    if error_reason in LEARNING_ERROR_REASON_VALUES:
        normalized["error_reason_primary"] = error_reason
    review_necessity = str(learning_fields.get("review_necessity", "")).strip()
    if review_necessity in REVIEW_NECESSITY_VALUES:
        normalized["review_necessity"] = review_necessity
    actionability = str(learning_fields.get("actionability", "")).strip()
    if actionability in ACTIONABILITY_VALUES:
        normalized["actionability"] = actionability
    action_bucket = str(learning_fields.get("action_bucket", "")).strip()
    if action_bucket in ACTION_BUCKET_VALUES:
        normalized["action_bucket"] = action_bucket
    normalized["need_rule_update"] = _parse_boolish(learning_fields.get("need_rule_update"))
    normalized["need_prompt_update"] = _parse_boolish(
        learning_fields.get("need_prompt_update", learning_fields.get("need_skill_update"))
    )
    normalized["need_annotation_update"] = _parse_boolish(learning_fields.get("need_annotation_update"))
    normalized["learning_note"] = str(learning_fields.get("learning_note", "")).strip()
    return normalized


def _parse_learning_fields(row: Dict[str, str]) -> Dict[str, object]:
    raw = str(row.get("learning_fields", "")).strip()
    if raw:
        parsed = _parse_json_dict(raw)
        if parsed:
            return _normalize_learning_fields(parsed)
    return _normalize_learning_fields(
        {
            "error_reason_primary": row.get("error_reason_primary", ""),
            "review_necessity": row.get("review_necessity", ""),
            "actionability": row.get("actionability", ""),
            "action_bucket": row.get("action_bucket", ""),
            "need_rule_update": row.get("need_rule_update", ""),
            "need_prompt_update": row.get("need_prompt_update", row.get("need_skill_update", "")),
            "need_annotation_update": row.get("need_annotation_update", ""),
            "learning_note": row.get("learning_note", ""),
        }
    )


def _parse_json_dict(raw: str) -> Dict[str, object]:
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _parse_boolish(value: object) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    return text in {"1", "true", "yes", "y", "on"}


def _normalize_loaded_decision(decision: Dict[str, object]) -> Dict[str, object]:
    normalized = dict(decision)
    normalized["current_labels"] = _normalize_reviewed_fields(dict(normalized.get("current_labels", {})))
    normalized["final_labels"] = _normalize_reviewed_fields(dict(normalized.get("final_labels", {})))
    normalized["learning_fields"] = _normalize_learning_fields(normalized.get("learning_fields", {}))
    normalized["change_diff"] = dict(normalized.get("change_diff", {}))
    return normalized
    return diff


def _review_priority(task: Dict[str, object]) -> int:
    reason = str(task.get("review_reason_code", ""))
    if reason == "ACTOR_OVERLAP":
        return 3
    if reason.startswith("PARSE_FAILED"):
        return 2
    if reason == "BUSINESSLINE_LOW_SIGNAL":
        return 1
    return 0
