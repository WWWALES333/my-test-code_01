from __future__ import annotations

import csv
import json
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

from .owner import build_owner_registry, extract_owner_hint
from .schema import (
    DECISION_CONFIRMED,
    TASK_STATUS_OPEN,
    TASK_STATUS_REVIEWED,
    stable_hash,
)


def load_review_decisions(annotations_dir: Path) -> Dict[Tuple[str, str], Dict[str, object]]:
    """读取人工复核回写结果，按报告和片段定位。"""
    decisions: Dict[Tuple[str, str], Dict[str, object]] = {}
    if not annotations_dir.exists():
        return decisions

    for path in sorted(annotations_dir.rglob("*.csv")):
        with path.open("r", encoding="utf-8", newline="") as file:
            reader = csv.DictReader(file)
            if not reader.fieldnames:
                continue
            if "report_id" not in reader.fieldnames or "segment_id" not in reader.fieldnames:
                continue
            for row in reader:
                report_id = str(row.get("report_id", "")).strip()
                segment_id = str(row.get("segment_id", "")).strip()
                if not report_id:
                    continue
                key = (report_id, segment_id)
                decisions[key] = {
                    "report_id": report_id,
                    "segment_id": segment_id,
                    "review_comment": str(row.get("review_comment", "")).strip(),
                    "reviewer": str(row.get("reviewer", "")).strip(),
                    "reviewed_at": str(row.get("reviewed_at", "")).strip(),
                    "final_labels": _parse_final_labels(row),
                    "raw_row": dict(row),
                    "source_file": str(path),
                }
    return decisions


def build_report_facts(
    report_rows: Iterable[Dict[str, object]],
    owner_registry: List[Dict[str, object]],
) -> List[Dict[str, object]]:
    """构建报告级事实表，统一销售对象和时间信息。"""
    owner_map = {str(item["owner_hint"]): item for item in owner_registry}
    rows: List[Dict[str, object]] = []
    for row in report_rows:
        owner_hint = extract_owner_hint(str(row.get("file_path", "")))
        owner = owner_map.get(owner_hint, {})
        rows.append(
            {
                "report_id": row.get("report_id", ""),
                "run_id": row.get("run_id", ""),
                "report_type": row.get("report_type", ""),
                "year": int(row.get("year", 0)),
                "month": int(row.get("month", 0)),
                "week_of_month": int(row.get("week_of_month", 0)),
                "report_owner_id": owner.get("salesperson_id", ""),
                "report_owner_name": owner.get("salesperson_name", owner_hint),
                "report_owner_type": owner.get("owner_type", "unknown"),
                "report_owner_hint": owner_hint,
                "file_path": row.get("file_path", ""),
                "parse_status": row.get("parse_status", row.get("text_status", "")),
                "parse_reason_code": row.get("parse_reason_code", ""),
                "segment_count": int(row.get("segment_count", 0)),
                "model_mode": row.get("model_mode", ""),
                "model_name": row.get("model_name", ""),
            }
        )
    return rows


def build_evidence_facts(
    evidence_rows: Iterable[Dict[str, object]],
    report_facts: Iterable[Dict[str, object]],
    review_decisions: Dict[Tuple[str, str], Dict[str, object]],
    owner_registry: List[Dict[str, object]],
) -> List[Dict[str, object]]:
    """构建证据级事实表，并应用人工复核覆盖。"""
    report_map = {str(row["report_id"]): row for row in report_facts}
    owner_map = {str(item["owner_hint"]): item for item in owner_registry}
    rows: List[Dict[str, object]] = []
    for row in evidence_rows:
        report_id = str(row.get("report_id", ""))
        segment_id = str(row.get("segment_id", ""))
        report = report_map.get(report_id, {})
        decision = review_decisions.get((report_id, segment_id))
        final_labels = dict(decision.get("final_labels", {})) if decision else {}
        segment_owner_hint = str(row.get("segment_owner_hint", "")).strip() or str(report.get("report_owner_hint", "")).strip()
        owner = owner_map.get(segment_owner_hint, {})

        business_line = str(final_labels.get("business_line", row.get("business_line", "")))
        actor_primary = str(final_labels.get("actor_primary", row.get("actor_primary", row.get("ai_actor", ""))))
        ai_scope = str(final_labels.get("ai_scope", row.get("ai_scope", "")))
        decision_status = str(final_labels.get("decision_status", row.get("decision_status", "")))
        if not decision_status:
            decision_status = DECISION_CONFIRMED

        rows.append(
            {
                "evidence_id": row.get("evidence_id", stable_hash(report_id, segment_id, "evidence")),
                "report_id": report_id,
                "segment_id": segment_id,
                "year": int(report.get("year", 0)),
                "month": int(report.get("month", 0)),
                "week_of_month": int(report.get("week_of_month", 0)),
                "salesperson_id": owner.get("salesperson_id", report.get("report_owner_id", "")),
                "salesperson_name": owner.get("salesperson_name", report.get("report_owner_name", "")),
                "owner_type": owner.get("owner_type", report.get("report_owner_type", "")),
                "owner_hint": segment_owner_hint,
                "report_owner_name": report.get("report_owner_name", ""),
                "business_line": business_line,
                "actor_primary": actor_primary,
                "ai_scope": ai_scope,
                "decision_status": decision_status,
                "source_text": row.get("source_text", ""),
                "file_path": row.get("file_path", ""),
                "review_status": TASK_STATUS_REVIEWED if decision else ("open" if decision_status != DECISION_CONFIRMED else "not_needed"),
                "review_comment": decision.get("review_comment", "") if decision else "",
                "reviewer": decision.get("reviewer", "") if decision else "",
                "reviewed_at": decision.get("reviewed_at", "") if decision else "",
            }
        )
    return rows


def build_review_tasks(
    review_rows: Iterable[Dict[str, object]],
    report_facts: Iterable[Dict[str, object]],
    review_decisions: Dict[Tuple[str, str], Dict[str, object]],
    owner_registry: List[Dict[str, object]],
) -> List[Dict[str, object]]:
    """构建复核任务表，供 v1.5 工作台和回写逻辑消费。"""
    report_map = {str(row["report_id"]): row for row in report_facts}
    owner_map = {str(item["owner_hint"]): item for item in owner_registry}
    tasks: List[Dict[str, object]] = []
    for row in review_rows:
        report_id = str(row.get("report_id", ""))
        segment_id = str(row.get("segment_id", ""))
        report = report_map.get(report_id, {})
        decision = review_decisions.get((report_id, segment_id))
        segment_owner_hint = str(row.get("segment_owner_hint", "")).strip() or str(report.get("report_owner_hint", "")).strip()
        owner = owner_map.get(segment_owner_hint, {})
        tasks.append(
            {
                "task_id": row.get("review_id", stable_hash(report_id, segment_id, "task")),
                "review_id": row.get("review_id", ""),
                "report_id": report_id,
                "segment_id": segment_id,
                "salesperson_id": owner.get("salesperson_id", report.get("report_owner_id", "")),
                "salesperson_name": owner.get("salesperson_name", report.get("report_owner_name", "")),
                "owner_type": owner.get("owner_type", report.get("report_owner_type", "")),
                "owner_hint": segment_owner_hint,
                "report_owner_name": report.get("report_owner_name", ""),
                "year": int(report.get("year", 0)),
                "month": int(report.get("month", 0)),
                "week_of_month": int(report.get("week_of_month", 0)),
                "review_reason_code": row.get("review_reason_code", ""),
                "review_reason": row.get("review_reason", ""),
                "task_status": TASK_STATUS_REVIEWED if decision else TASK_STATUS_OPEN,
                "current_decision_status": row.get("current_decision_status", row.get("decision_status", "")),
                "current_fields": {
                    "decision_status": row.get("current_decision_status", row.get("decision_status", "")),
                    "review_reason_code": row.get("review_reason_code", ""),
                },
                "source_text": row.get("source_text", ""),
                "file_path": row.get("file_path", ""),
                "review_comment": decision.get("review_comment", "") if decision else "",
                "reviewer": decision.get("reviewer", "") if decision else "",
                "reviewed_at": decision.get("reviewed_at", "") if decision else "",
            }
        )
    return sorted(tasks, key=lambda item: (str(item["task_status"]), int(item["year"]), int(item["month"])))


def build_sales_monthly_rollup(
    evidence_facts: Iterable[Dict[str, object]],
    report_facts: Iterable[Dict[str, object]],
) -> List[Dict[str, object]]:
    """按销售个人和月份聚合业务工作台所需的核心指标。"""
    evidence_owner_weeks: Dict[Tuple[str, int, int], set[str]] = defaultdict(set)
    evidence_owner_reports: Dict[Tuple[str, int, int], set[str]] = defaultdict(set)
    agg: Dict[Tuple[str, int, int], Dict[str, object]] = {}
    for row in evidence_facts:
        salesperson_id = str(row.get("salesperson_id", ""))
        if not salesperson_id:
            continue
        year = int(row.get("year", 0))
        month = int(row.get("month", 0))
        key = (salesperson_id, year, month)
        evidence_owner_weeks[key].add(str(row.get("week_of_month", 0)))
        evidence_owner_reports[key].add(str(row.get("report_id", "")))
        if key not in agg:
            agg[key] = {
                "salesperson_id": salesperson_id,
                "salesperson_name": row.get("salesperson_name", ""),
                "owner_type": row.get("owner_type", ""),
                "year": year,
                "month": month,
                "ai_mentions": 0,
                "confirmed_mentions": 0,
                "pending_mentions": 0,
                "doctor_feedback_mentions": 0,
                "sales_intro_mentions": 0,
                "self_use_mentions": 0,
                "opportunity_mentions": 0,
                "cloud_clinic_mentions": 0,
                "cloud_steward_mentions": 0,
                "mixed_mentions": 0,
                "review_open_count": 0,
            }

        item = agg[key]
        item["ai_mentions"] = int(item["ai_mentions"]) + 1
        if str(row.get("decision_status", "")) == DECISION_CONFIRMED:
            item["confirmed_mentions"] = int(item["confirmed_mentions"]) + 1
        else:
            item["pending_mentions"] = int(item["pending_mentions"]) + 1
        if str(row.get("review_status", "")) == TASK_STATUS_OPEN:
            item["review_open_count"] = int(item["review_open_count"]) + 1
        actor = str(row.get("actor_primary", ""))
        if actor == "医生反馈":
            item["doctor_feedback_mentions"] = int(item["doctor_feedback_mentions"]) + 1
        elif actor == "销售对外介绍":
            item["sales_intro_mentions"] = int(item["sales_intro_mentions"]) + 1
        elif actor == "销售自用":
            item["self_use_mentions"] = int(item["self_use_mentions"]) + 1
        elif actor == "潜在 AI 机会":
            item["opportunity_mentions"] = int(item["opportunity_mentions"]) + 1

        line = str(row.get("business_line", ""))
        if line == "云诊室":
            item["cloud_clinic_mentions"] = int(item["cloud_clinic_mentions"]) + 1
        elif line == "云管家":
            item["cloud_steward_mentions"] = int(item["cloud_steward_mentions"]) + 1
        elif line == "混合":
            item["mixed_mentions"] = int(item["mixed_mentions"]) + 1

    rows: List[Dict[str, object]] = []
    for key in sorted(agg.keys(), key=lambda item: (item[1], item[2], str(item[0]))):
        item = agg[key]
        item["active_weeks"] = len(evidence_owner_weeks.get(key, set()))
        item["active_reports"] = len(evidence_owner_reports.get(key, set()))
        rows.append(item)
    return rows


def build_insight_cards(evidence_facts: Iterable[Dict[str, object]]) -> List[Dict[str, object]]:
    """按主题和业务线生成最小可用的结论卡对象。"""
    grouped: Dict[Tuple[str, str], List[Dict[str, object]]] = defaultdict(list)
    for row in evidence_facts:
        topic = _infer_topic(row)
        grouped[(topic, str(row.get("business_line", "待判断")))].append(row)

    cards: List[Dict[str, object]] = []
    for (topic, business_line), rows in grouped.items():
        high_value_rows = [row for row in rows if _is_high_value_evidence(row)]
        selected = high_value_rows[:3] if high_value_rows else rows[:3]
        owner_refs = sorted(
            {
                str(row.get("salesperson_name", ""))
                for row in selected
                if str(row.get("salesperson_name", "")) and str(row.get("owner_type", "")) == "person" and _looks_like_salesperson_name(str(row.get("salesperson_name", "")))
            }
        )
        confidence = "high" if len(high_value_rows) >= 2 else ("medium" if len(selected) >= 2 else "low")
        needs_review = any(str(row.get("review_status", "")) == TASK_STATUS_OPEN for row in rows)
        cards.append(
            {
                "insight_id": stable_hash(topic, business_line, str(len(rows))),
                "topic": topic,
                "title": _build_card_title(topic, business_line),
                "summary": _build_card_summary(topic, business_line, selected),
                "business_line": business_line,
                "signal_type": _topic_to_signal_type(topic),
                "evidence_count": len(rows),
                "open_review_count": sum(1 for row in rows if str(row.get("review_status", "")) == TASK_STATUS_OPEN),
                "evidence_refs": [
                    {
                        "report_id": row.get("report_id", ""),
                        "segment_id": row.get("segment_id", ""),
                        "source_text": row.get("source_text", ""),
                    }
                    for row in selected
                ],
                "owner_refs": owner_refs[:5],
                "needs_review": needs_review,
                "confidence": confidence,
            }
        )
    return sorted(cards, key=lambda item: (_topic_priority(str(item.get("topic", ""))), -int(item.get("evidence_count", 0)), str(item.get("business_line", ""))))


def build_dashboard_snapshot(
    report_facts: Iterable[Dict[str, object]],
    evidence_facts: Iterable[Dict[str, object]],
    sales_rollup: Iterable[Dict[str, object]],
    insight_cards: Iterable[Dict[str, object]],
    review_tasks: Iterable[Dict[str, object]],
) -> Dict[str, object]:
    """构建首页和趋势页直接可消费的快照对象。"""
    report_rows = list(report_facts)
    evidence_rows = list(evidence_facts)
    sales_rows = list(sales_rollup)
    insight_rows = list(insight_cards)
    task_rows = list(review_tasks)

    latest_month = _latest_year_month(report_rows)
    active_sales = {
        str(row.get("salesperson_id", ""))
        for row in evidence_rows
        if str(row.get("salesperson_id", ""))
    }
    person_sales = {
        str(row.get("salesperson_id", ""))
        for row in evidence_rows
        if str(row.get("owner_type", "")) == "person" and str(row.get("salesperson_id", ""))
    }
    group_sales = {
        str(row.get("salesperson_id", ""))
        for row in evidence_rows
        if str(row.get("owner_type", "")) == "group" and str(row.get("salesperson_id", ""))
    }
    yoy = _build_yoy_summary(evidence_rows)
    compare_year = int(yoy.get("year_b", 0)) if yoy.get("has_compare") else _latest_year(report_rows)
    people_rows = [row for row in sales_rows if str(row.get("owner_type", "")) == "person"]
    group_rows = [row for row in sales_rows if str(row.get("owner_type", "")) != "person"]
    top_sales = _aggregate_owner_rows(people_rows, compare_year)[:10]
    if len(top_sales) < 10:
        top_sales.extend(_aggregate_owner_rows(group_rows, compare_year)[: 10 - len(top_sales)])
    review_reason_counter = Counter(str(task.get("review_reason_code", "")) for task in task_rows if str(task.get("review_reason_code", "")))
    return {
        "latest_year_month": latest_month,
        "total_reports": len(report_rows),
        "total_ai_mentions": len(evidence_rows),
        "confirmed_mentions": sum(1 for row in evidence_rows if str(row.get("decision_status", "")) == DECISION_CONFIRMED),
        "open_review_tasks": sum(1 for row in task_rows if str(row.get("task_status", "")) == TASK_STATUS_OPEN),
        "reviewed_tasks": sum(1 for row in task_rows if str(row.get("task_status", "")) == TASK_STATUS_REVIEWED),
        "active_sales_count": len(active_sales),
        "active_person_count": len(person_sales),
        "active_group_count": len(group_sales),
        "insight_card_count": len(insight_rows),
        "yoy_summary": yoy,
        "breadth_depth_summary": _build_breadth_depth_summary(yoy),
        "actor_yoy": _build_compare_breakdown(evidence_rows, "actor_primary"),
        "business_line_yoy": _build_compare_breakdown(evidence_rows, "business_line"),
        "top_sales": top_sales,
        "top_people_latest": _aggregate_owner_rows(people_rows, compare_year)[:12],
        "top_groups_latest": _aggregate_owner_rows(group_rows, compare_year)[:12],
        "feedback_examples": _select_example_rows(evidence_rows, actor_primary="医生反馈", year=compare_year),
        "opportunity_examples": _select_example_rows(evidence_rows, actor_primary="潜在 AI 机会", year=compare_year),
        "sales_intro_examples": _select_example_rows(evidence_rows, actor_primary="销售对外介绍", year=compare_year),
        "top_review_reasons": review_reason_counter.most_common(5),
        "priority_review_tasks": _select_priority_review_tasks(task_rows),
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
    for key in ("business_line", "actor_primary", "ai_scope", "decision_status"):
        value = str(row.get(key, "")).strip()
        if value:
            final_labels[key] = value
    actual = str(row.get("actual", "")).strip()
    if actual and "decision_status" not in final_labels:
        final_labels["decision_status"] = actual
    return final_labels


def _infer_topic(row: Dict[str, object]) -> str:
    actor = str(row.get("actor_primary", ""))
    scope = str(row.get("ai_scope", ""))
    if actor == "医生反馈":
        return "doctor_feedback"
    if actor == "销售对外介绍":
        return "sales_behavior"
    if actor == "潜在 AI 机会":
        return "product_opportunity"
    if scope in {"market_trend", "competitor_ai"}:
        return "market_watch"
    return "adoption_signal"


def _topic_to_signal_type(topic: str) -> str:
    mapping = {
        "doctor_feedback": "feedback",
        "sales_behavior": "behavior",
        "product_opportunity": "opportunity",
        "market_watch": "trend",
        "adoption_signal": "trend",
    }
    return mapping.get(topic, "trend")


def _build_card_title(topic: str, business_line: str) -> str:
    mapping = {
        "doctor_feedback": f"{business_line}医生反馈信号",
        "sales_behavior": f"{business_line}销售介绍行为",
        "product_opportunity": f"{business_line}产品机会线索",
        "market_watch": f"{business_line}市场与竞品观察",
        "adoption_signal": f"{business_line}AI 使用趋势信号",
    }
    return mapping.get(topic, f"{business_line}业务结论")


def _build_card_summary(topic: str, business_line: str, rows: List[Dict[str, object]]) -> str:
    owner_names = sorted(
        {
            str(row.get("salesperson_name", ""))
            for row in rows
            if str(row.get("salesperson_name", ""))
            and str(row.get("owner_type", "")) == "person"
            and _looks_like_salesperson_name(str(row.get("salesperson_name", "")))
        }
    )
    owner_text = "、".join(owner_names[:3]) if owner_names else "多位销售"
    positive_keywords = ("认可", "感兴趣", "不错", "愿意", "方便", "好用", "接受", "信任", "体验")
    concern_keywords = ("担心", "不能", "不允许", "太初级", "帮倒忙", "诈骗", "严格", "质疑", "替代")
    positive_count = _count_rows_with_keywords(rows, positive_keywords)
    concern_count = _count_rows_with_keywords(rows, concern_keywords)
    if topic == "doctor_feedback":
        if positive_count and concern_count:
            return f"{owner_text}在 {business_line} 语境中记录到医生对 AI 的反馈已经出现明显分化，既有认可和试用意愿，也有对准确度、政策限制或替代风险的顾虑。"
        if positive_count > 0:
            return f"{owner_text}在 {business_line} 语境中记录到的医生反馈以正向认可为主，效率提升、病历整理和辅助诊疗更容易触发接受。"
        return f"{owner_text}在 {business_line} 语境中已经出现真实医生反馈，但当前更多仍是探索和观察，适合继续跟踪复购与复访信号。"
    if topic == "sales_behavior":
        return f"{owner_text}在 {business_line} 语境中的 AI 表达以“介绍功能、演示流程、引导体验”为主，说明 AI 已进入一线常规沟通动作，而不只是临时卖点。"
    if topic == "product_opportunity":
        return f"{owner_text}在 {business_line} 语境中反复暴露出可继续产品化的 AI 机会，当前高频方向集中在 {_top_feature_keywords(rows)}。"
    if topic == "market_watch":
        return f"{owner_text}记录到 {business_line} 相关的市场趋势或竞品 AI 动态，适合作为风险和机会观察输入。"
    if concern_count > positive_count:
        return f"{owner_text}在 {business_line} 语境中的 AI 提及已经形成连续信号，但当前以顾虑和待确认表达为主，说明趋势已出现、口径仍需复核。"
    return f"{owner_text}在 {business_line} 语境中的 AI 提及已经从零散表达变成持续信号，值得继续观察它是否能沉淀为稳定话术和客户反馈。"


def _is_high_value_evidence(row: Dict[str, object]) -> bool:
    text = str(row.get("source_text", ""))
    weak_keywords = ("资料整理", "模板", "申报", "很重要", "审核资料整理", "工作应该做到", "分享会", "两会")
    return len(text) >= 20 and not any(keyword in text for keyword in weak_keywords)


def _sales_sort_key(item: Dict[str, object]) -> tuple[int, int, int, str]:
    return (
        -int(item.get("ai_mentions", 0)),
        -int(item.get("confirmed_mentions", 0)),
        -int(item.get("active_weeks", 0)),
        str(item.get("salesperson_name", "")),
    )


def _topic_priority(topic: str) -> int:
    priority = {
        "doctor_feedback": 0,
        "product_opportunity": 1,
        "sales_behavior": 2,
        "market_watch": 3,
        "adoption_signal": 4,
    }
    return priority.get(topic, 9)


def _count_rows_with_keywords(rows: Iterable[Dict[str, object]], keywords: Tuple[str, ...]) -> int:
    count = 0
    for row in rows:
        text = str(row.get("source_text", ""))
        if any(keyword in text for keyword in keywords):
            count += 1
    return count


def _top_feature_keywords(rows: Iterable[Dict[str, object]]) -> str:
    mapping = {
        "病历整理": ("病历", "整理"),
        "辅助诊疗": ("辅助诊疗", "AI诊疗", "Ai开处方", "AI开方"),
        "问诊提效": ("问诊", "效率", "提效"),
        "话术支持": ("话术", "沟通"),
        "随访管理": ("随访", "诊后"),
        "搜索与查询": ("搜索", "查询", "Deepseek", "deepseek"),
        "系统兼容": ("鸿蒙", "不能用", "兼容"),
    }
    counter: Counter[str] = Counter()
    for row in rows:
        text = str(row.get("source_text", ""))
        for label, keywords in mapping.items():
            if any(keyword in text for keyword in keywords):
                counter[label] += 1
    if not counter:
        return "辅助诊疗、效率提升等场景"
    labels = [label for label, _ in counter.most_common(2)]
    return "、".join(labels)


def _looks_like_salesperson_name(name: str) -> bool:
    invalid_keywords = ("人员", "部门", "心得", "完成", "意向", "副本", "售后", "北京", "中医", "其他")
    token = name.strip()
    if not token:
        return False
    if any(keyword in token for keyword in invalid_keywords):
        return False
    return True


def _aggregate_owner_rows(rows: Iterable[Dict[str, object]], target_year: int) -> List[Dict[str, object]]:
    grouped: Dict[str, Dict[str, object]] = {}
    for row in rows:
        if target_year and int(row.get("year", 0)) != target_year:
            continue
        salesperson_id = str(row.get("salesperson_id", ""))
        if not salesperson_id:
            continue
        if salesperson_id not in grouped:
            grouped[salesperson_id] = {
                "salesperson_id": salesperson_id,
                "salesperson_name": row.get("salesperson_name", ""),
                "owner_type": row.get("owner_type", ""),
                "year": target_year,
                "month": 0,
                "ai_mentions": 0,
                "confirmed_mentions": 0,
                "pending_mentions": 0,
                "doctor_feedback_mentions": 0,
                "sales_intro_mentions": 0,
                "self_use_mentions": 0,
                "opportunity_mentions": 0,
                "cloud_clinic_mentions": 0,
                "cloud_steward_mentions": 0,
                "mixed_mentions": 0,
                "review_open_count": 0,
                "active_weeks": 0,
                "active_reports": 0,
            }
        item = grouped[salesperson_id]
        for key in (
            "ai_mentions",
            "confirmed_mentions",
            "pending_mentions",
            "doctor_feedback_mentions",
            "sales_intro_mentions",
            "self_use_mentions",
            "opportunity_mentions",
            "cloud_clinic_mentions",
            "cloud_steward_mentions",
            "mixed_mentions",
            "review_open_count",
            "active_weeks",
            "active_reports",
        ):
            item[key] = int(item.get(key, 0)) + int(row.get(key, 0))
    return sorted(grouped.values(), key=_sales_sort_key)


def _latest_year_month(report_rows: Iterable[Dict[str, object]]) -> str:
    pairs = sorted({(int(row.get("year", 0)), int(row.get("month", 0))) for row in report_rows if int(row.get("year", 0)) > 0})
    if not pairs:
        return ""
    year, month = pairs[-1]
    return f"{year:04d}-{month:02d}"


def _latest_year(report_rows: Iterable[Dict[str, object]]) -> int:
    years = sorted({int(row.get("year", 0)) for row in report_rows if int(row.get("year", 0)) > 0})
    return years[-1] if years else 0


def _build_yoy_summary(evidence_rows: Iterable[Dict[str, object]]) -> Dict[str, object]:
    month_counter: Dict[Tuple[int, int], int] = defaultdict(int)
    sales_counter: Dict[Tuple[int, int], set[str]] = defaultdict(set)
    for row in evidence_rows:
        year = int(row.get("year", 0))
        month = int(row.get("month", 0))
        if year <= 0 or month <= 0:
            continue
        month_counter[(year, month)] += 1
        salesperson_id = str(row.get("salesperson_id", ""))
        if salesperson_id:
            sales_counter[(year, month)].add(salesperson_id)

    years = sorted({year for year, _ in month_counter.keys()})
    if len(years) < 2:
        return {"has_compare": False}
    year_a = years[-2]
    year_b = years[-1]
    compare_months = sorted({month for (year, month) in month_counter.keys() if year == year_b})
    a_mentions = sum(month_counter.get((year_a, month), 0) for month in compare_months)
    b_mentions = sum(month_counter.get((year_b, month), 0) for month in compare_months)
    a_sales = len({sid for month in compare_months for sid in sales_counter.get((year_a, month), set())})
    b_sales = len({sid for month in compare_months for sid in sales_counter.get((year_b, month), set())})
    return {
        "has_compare": True,
        "year_a": year_a,
        "year_b": year_b,
        "months": compare_months,
        "mentions_a": a_mentions,
        "mentions_b": b_mentions,
        "sales_a": a_sales,
        "sales_b": b_sales,
    }


def _build_breadth_depth_summary(yoy: Dict[str, object]) -> Dict[str, object]:
    if not yoy.get("has_compare"):
        return {"has_compare": False}
    mentions_a = int(yoy.get("mentions_a", 0))
    mentions_b = int(yoy.get("mentions_b", 0))
    sales_a = int(yoy.get("sales_a", 0))
    sales_b = int(yoy.get("sales_b", 0))
    avg_a = round(mentions_a / sales_a, 2) if sales_a else 0.0
    avg_b = round(mentions_b / sales_b, 2) if sales_b else 0.0
    sales_growth = ((sales_b - sales_a) / sales_a) if sales_a else 0.0
    avg_growth = ((avg_b - avg_a) / avg_a) if avg_a else 0.0
    if sales_growth >= 0.1 and avg_growth >= 0.1:
        judgement = "广度和深度都在提升"
    elif sales_growth >= 0.1:
        judgement = "覆盖销售对象在扩大，但单人提及强度提升有限"
    elif avg_growth >= 0.1:
        judgement = "主要由少数销售更高频地推动，而不是更多销售同时加入"
    else:
        judgement = "整体变化有限，仍需继续观察"
    return {
        "has_compare": True,
        "avg_mentions_per_sales_a": avg_a,
        "avg_mentions_per_sales_b": avg_b,
        "judgement": judgement,
    }


def _build_compare_breakdown(evidence_rows: Iterable[Dict[str, object]], field_name: str) -> List[Dict[str, object]]:
    rows = list(evidence_rows)
    yoy = _build_yoy_summary(rows)
    if not yoy.get("has_compare"):
        return []
    compare_months = set(int(item) for item in yoy.get("months", []))
    year_a = int(yoy.get("year_a", 0))
    year_b = int(yoy.get("year_b", 0))
    a_counter: Counter[str] = Counter()
    b_counter: Counter[str] = Counter()
    for row in rows:
        month = int(row.get("month", 0))
        if month not in compare_months:
            continue
        key = str(row.get(field_name, "")).strip() or "未标注"
        year = int(row.get("year", 0))
        if year == year_a:
            a_counter[key] += 1
        elif year == year_b:
            b_counter[key] += 1
    items: List[Dict[str, object]] = []
    for key in sorted(set(a_counter) | set(b_counter)):
        items.append(
            {
                "label": key,
                "year_a": a_counter[key],
                "year_b": b_counter[key],
                "delta": b_counter[key] - a_counter[key],
            }
        )
    return sorted(items, key=lambda item: (-int(item["delta"]), -int(item["year_b"]), str(item["label"])))


def _select_example_rows(
    evidence_rows: Iterable[Dict[str, object]],
    actor_primary: str,
    year: int,
    limit: int = 5,
) -> List[Dict[str, object]]:
    selected: List[Dict[str, object]] = []
    seen: set[tuple[str, str]] = set()
    rows = [
        row
        for row in evidence_rows
        if str(row.get("actor_primary", "")) == actor_primary and (year <= 0 or int(row.get("year", 0)) == year)
    ]
    rows.sort(
        key=lambda row: (
            str(row.get("review_status", "")) == TASK_STATUS_OPEN,
            str(row.get("decision_status", "")) != DECISION_CONFIRMED,
            str(row.get("owner_type", "")) != "person",
            not _is_high_value_evidence(row),
            -len(str(row.get("source_text", ""))),
        )
    )
    for row in rows:
        key = (str(row.get("salesperson_name", "")), str(row.get("source_text", ""))[:40])
        if key in seen:
            continue
        seen.add(key)
        selected.append(
            {
                "salesperson_name": row.get("salesperson_name", ""),
                "business_line": row.get("business_line", ""),
                "decision_status": row.get("decision_status", ""),
                "source_text": row.get("source_text", ""),
                "file_path": row.get("file_path", ""),
            }
        )
        if len(selected) >= limit:
            break
    return selected


def _select_priority_review_tasks(tasks: Iterable[Dict[str, object]], limit: int = 8) -> List[Dict[str, object]]:
    open_tasks = [task for task in tasks if str(task.get("task_status", "")) == TASK_STATUS_OPEN]
    open_tasks.sort(
        key=lambda task: (
            "ACTOR_OVERLAP" not in str(task.get("review_reason_code", "")),
            str(task.get("owner_type", "")) != "person",
            -len(str(task.get("source_text", ""))),
        )
    )
    return open_tasks[:limit]
