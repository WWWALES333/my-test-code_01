from __future__ import annotations

from collections import Counter, defaultdict
from typing import Dict, Iterable, List, Sequence, Tuple

from .schema import DECISION_CONFIRMED, TASK_STATUS_OPEN


def build_trend_cube(
    evidence_facts: Sequence[Dict[str, object]],
    sales_roster: Sequence[Dict[str, object]],
) -> List[Dict[str, object]]:
    """构建周/月趋势指标层。"""
    roster_people = [row for row in sales_roster if str(row.get("employment_status", "")) == "active"]
    roster_count = len(roster_people)
    roster_regions = {str(row.get("region_name", "")) for row in roster_people if str(row.get("region_name", ""))}
    total_regions = len(roster_regions)

    monthly: Dict[Tuple[int, int], Dict[str, object]] = {}
    weekly: Dict[Tuple[int, int, int], Dict[str, object]] = {}

    for row in evidence_facts:
        year = int(row.get("year", 0))
        month = int(row.get("month", 0))
        week = int(row.get("week_of_month", 0))
        salesperson_id = str(row.get("salesperson_id", ""))
        region_name = str(row.get("region_name", ""))
        actor = str(row.get("actor_primary", "")) or "未标注"
        business_line = str(row.get("business_line", "")) or "待判断"
        decision_status = str(row.get("decision_status", ""))
        review_open = str(row.get("review_status", "")) == TASK_STATUS_OPEN

        month_key = (year, month)
        if month_key not in monthly:
            monthly[month_key] = _new_bucket("month", year, month, 0)
        week_key = (year, month, week)
        if week > 0 and week_key not in weekly:
            weekly[week_key] = _new_bucket("week", year, month, week)

        buckets = [monthly[month_key]]
        if week > 0:
            buckets.append(weekly[week_key])
        for bucket in buckets:
            bucket["ai_mentions"] = int(bucket["ai_mentions"]) + 1
            if decision_status == DECISION_CONFIRMED:
                bucket["confirmed_mentions"] = int(bucket["confirmed_mentions"]) + 1
            if review_open:
                bucket["pending_review_mentions"] = int(bucket["pending_review_mentions"]) + 1
            if salesperson_id:
                bucket["_sales_ids"].add(salesperson_id)
            if region_name:
                bucket["_regions"].add(region_name)
            bucket["_actors"][actor] += 1
            bucket["_business_lines"][business_line] += 1

    rows: List[Dict[str, object]] = []
    for bucket in sorted([*monthly.values(), *weekly.values()], key=lambda item: (item["grain"], int(item["year"]), int(item["month"]), int(item["week_of_month"]))):
        active_sales = len(bucket.pop("_sales_ids"))
        active_regions = len(bucket.pop("_regions"))
        actor_breakdown = dict(bucket.pop("_actors"))
        business_line_breakdown = dict(bucket.pop("_business_lines"))
        ai_mentions = int(bucket["ai_mentions"])
        pending_mentions = int(bucket["pending_review_mentions"])
        bucket["active_sales_count"] = active_sales
        bucket["sales_penetration_rate"] = round(active_sales / roster_count, 4) if roster_count else 0.0
        bucket["active_region_count"] = active_regions
        bucket["region_coverage_rate"] = round(active_regions / total_regions, 4) if total_regions else 0.0
        bucket["pending_review_rate"] = round(pending_mentions / ai_mentions, 4) if ai_mentions else 0.0
        bucket["actor_breakdown"] = actor_breakdown
        bucket["business_line_breakdown"] = business_line_breakdown
        rows.append(bucket)
    return rows


def build_trend_explanations(
    trend_cube: Sequence[Dict[str, object]],
    salesperson_profiles: Sequence[Dict[str, object]],
) -> List[Dict[str, object]]:
    """把指标层转换为解释型趋势。"""
    monthly = [row for row in trend_cube if str(row.get("grain", "")) == "month"]
    if not monthly:
        return []
    monthly_sorted = sorted(monthly, key=lambda row: (int(row.get("year", 0)), int(row.get("month", 0))))
    latest = monthly_sorted[-1]
    previous = monthly_sorted[-2] if len(monthly_sorted) >= 2 else {}
    yoy_pair = _find_yoy_pair(monthly_sorted)
    explanations: List[Dict[str, object]] = []

    breadth_text = _build_breadth_text(latest, previous)
    confidence = _confidence_from_pending(float(latest.get("pending_review_rate", 0.0)))
    explanations.append(
        {
            "metric_name": "breadth_depth",
            "period": _period_label(latest),
            "metric_value": latest.get("ai_mentions", 0),
            "baseline_value": previous.get("ai_mentions", 0),
            "delta_value": int(latest.get("ai_mentions", 0)) - int(previous.get("ai_mentions", 0)),
            "change_type": "mom",
            "confidence_level": confidence,
            "explanation": breadth_text,
            "drilldown_target_ids": [str(item.get("salesperson_id", "")) for item in salesperson_profiles[:5] if str(item.get("salesperson_id", ""))],
        }
    )

    if yoy_pair:
        a_row, b_row = yoy_pair
        yoy_text = _build_yoy_text(a_row, b_row)
        explanations.append(
            {
                "metric_name": "year_over_year",
                "period": _period_label(b_row),
                "metric_value": b_row.get("ai_mentions", 0),
                "baseline_value": a_row.get("ai_mentions", 0),
                "delta_value": int(b_row.get("ai_mentions", 0)) - int(a_row.get("ai_mentions", 0)),
                "change_type": "yoy",
                "confidence_level": _confidence_from_pending(max(float(a_row.get("pending_review_rate", 0.0)), float(b_row.get("pending_review_rate", 0.0)))),
                "explanation": yoy_text,
                "drilldown_target_ids": [],
            }
        )

    latest_actor = dict(latest.get("actor_breakdown", {}))
    dominant_actor = max(latest_actor.items(), key=lambda item: item[1])[0] if latest_actor else "未标注"
    explanations.append(
        {
            "metric_name": "structure_actor",
            "period": _period_label(latest),
            "metric_value": dominant_actor,
            "baseline_value": "",
            "delta_value": 0,
            "change_type": "structure",
            "confidence_level": confidence,
            "explanation": f"当前主体结构里，{dominant_actor} 占比最高，需要结合销售画像判断是普遍扩散还是少数人反复提及。",
            "drilldown_target_ids": [],
        }
    )
    return explanations


def build_dashboard_snapshot(
    report_facts: Sequence[Dict[str, object]],
    evidence_facts: Sequence[Dict[str, object]],
    trend_cube: Sequence[Dict[str, object]],
    trend_explanations: Sequence[Dict[str, object]],
    salesperson_profiles: Sequence[Dict[str, object]],
    region_rollups: Sequence[Dict[str, object]],
    insight_cards: Sequence[Dict[str, object]],
    review_tasks: Sequence[Dict[str, object]],
    sales_roster: Sequence[Dict[str, object]],
) -> Dict[str, object]:
    """构建首页/多页工作台可直接消费的快照。"""
    monthly = [row for row in trend_cube if str(row.get("grain", "")) == "month"]
    weekly = [row for row in trend_cube if str(row.get("grain", "")) == "week"]
    latest_month = monthly[-1] if monthly else {}
    latest_week = weekly[-1] if weekly else {}
    yoy = _build_yoy_summary(monthly)
    active_person_count = len(
        {
            str(row.get("salesperson_id", ""))
            for row in evidence_facts
            if str(row.get("owner_type", "")) == "person" and str(row.get("salesperson_id", ""))
        }
    )
    active_group_count = len(
        {
            str(row.get("salesperson_id", ""))
            for row in evidence_facts
            if str(row.get("owner_type", "")) != "person" and str(row.get("salesperson_id", ""))
        }
    )
    top_profiles = sorted(
        [row for row in salesperson_profiles if str(row.get("segment", "")) != "长期未提及者"],
        key=lambda item: (-int(item.get("total_mentions", 0)), -int(item.get("doctor_feedback_mentions", 0)), str(item.get("display_name", ""))),
    )
    top_regions = sorted(
        region_rollups,
        key=lambda item: (-float(item.get("dominance_ratio", 0.0)), -int(item.get("active_sales_count", 0)), str(item.get("region_name", ""))),
    )
    top_review_reasons = Counter(str(item.get("review_reason_code", "")) for item in review_tasks if str(item.get("task_status", "")) == TASK_STATUS_OPEN).most_common(5)
    priority_review_tasks = sorted(
        [item for item in review_tasks if str(item.get("task_status", "")) == TASK_STATUS_OPEN],
        key=lambda item: (_review_priority(item), -len(str(item.get("source_text", "")))),
        reverse=True,
    )[:8]

    stable_cards = [row for row in insight_cards if str(row.get("confidence_level", "")) in {"high", "medium"}][:5]
    risk_cards = [row for row in insight_cards if int(row.get("pending_review_count", 0)) > 0][:5]
    available_months = [_period_label(row) for row in monthly]
    available_weeks = [_period_label(row) for row in weekly]
    data_range = {
        "start_month": _period_label(monthly[0]) if monthly else "",
        "end_month": _period_label(latest_month) if latest_month else "",
        "start_week": _period_label(weekly[0]) if weekly else "",
        "end_week": _period_label(latest_week) if latest_week else "",
    }

    latest_year_month = _period_label(latest_month)
    latest_year_week = _period_label(latest_week)
    open_review_tasks = sum(1 for row in review_tasks if str(row.get("task_status", "")) == TASK_STATUS_OPEN)
    reviewed_tasks = sum(1 for row in review_tasks if str(row.get("task_status", "")) == "reviewed")
    active_sales_count = len({str(row.get("salesperson_id", "")) for row in evidence_facts if str(row.get("salesperson_id", ""))})

    return {
        "latest_year_month": latest_year_month,
        "latest_month": latest_year_month,
        "latest_period": latest_month,
        "latest_year_week": latest_year_week,
        "latest_week": latest_year_week,
        "latest_week_period": latest_week,
        "data_range": data_range,
        "available_months": available_months,
        "available_weeks": available_weeks,
        "total_reports": len(report_facts),
        "total_ai_mentions": len(evidence_facts),
        "confirmed_mentions": sum(1 for row in evidence_facts if str(row.get("decision_status", "")) == DECISION_CONFIRMED),
        "open_review_tasks": open_review_tasks,
        "review_open_count": open_review_tasks,
        "reviewed_tasks": reviewed_tasks,
        "review_reviewed_count": reviewed_tasks,
        "active_sales_count": active_sales_count,
        "active_salespeople": active_sales_count,
        "active_person_count": active_person_count,
        "active_individuals": active_person_count,
        "active_group_count": active_group_count,
        "roster_active_count": len([row for row in sales_roster if str(row.get("employment_status", "")) == "active"]),
        "insight_card_count": len(insight_cards),
        "yoy_summary": yoy,
        "trend_explanations": list(trend_explanations),
        "headline_judgements": [str(row.get("explanation", "")) for row in trend_explanations[:2] if str(row.get("explanation", ""))],
        "top_sales_profiles": top_profiles[:12],
        "silent_sales_profiles": [row for row in salesperson_profiles if str(row.get("segment", "")) == "长期未提及者"][:12],
        "top_region_rollups": top_regions[:10],
        "insight_cards": list(insight_cards),
        "stable_insights": stable_cards,
        "risk_insights": risk_cards,
        "top_review_reasons": top_review_reasons,
        "priority_review_tasks": priority_review_tasks,
        "time_scope_note": (
            f"当前默认月度口径：{data_range['end_month']}；当前默认周度口径：{data_range['end_week']}。"
            if data_range["end_month"] or data_range["end_week"]
            else "当前还没有可用时间窗口。"
        ),
    }


def _new_bucket(grain: str, year: int, month: int, week: int) -> Dict[str, object]:
    return {
        "grain": grain,
        "year": year,
        "month": month,
        "week_of_month": week,
        "ai_mentions": 0,
        "confirmed_mentions": 0,
        "pending_review_mentions": 0,
        "_sales_ids": set(),
        "_regions": set(),
        "_actors": Counter(),
        "_business_lines": Counter(),
    }


def _build_breadth_text(latest: Dict[str, object], previous: Dict[str, object]) -> str:
    latest_mentions = int(latest.get("ai_mentions", 0))
    previous_mentions = int(previous.get("ai_mentions", 0))
    latest_sales = int(latest.get("active_sales_count", 0))
    previous_sales = int(previous.get("active_sales_count", 0))
    latest_avg = round(latest_mentions / latest_sales, 2) if latest_sales else 0.0
    previous_avg = round(previous_mentions / previous_sales, 2) if previous_sales else 0.0
    if latest_sales > previous_sales and latest_avg > previous_avg:
        judgement = "覆盖和单人强度都在上升"
    elif latest_sales > previous_sales:
        judgement = "更多销售开始加入，但单人提及强度提升有限"
    elif latest_avg > previous_avg:
        judgement = "主要由少数销售更高频地推动，而不是更多销售同时加入"
    else:
        judgement = "整体变化有限，仍需继续观察"
    return f"{_period_label(latest)} 的 AI 证据 {previous_mentions}→{latest_mentions}，活跃销售 {previous_sales}→{latest_sales}，单人强度 {previous_avg}→{latest_avg}。当前判断：{judgement}。"


def _build_yoy_text(a_row: Dict[str, object], b_row: Dict[str, object]) -> str:
    a_mentions = int(a_row.get("ai_mentions", 0))
    b_mentions = int(b_row.get("ai_mentions", 0))
    a_sales = int(a_row.get("active_sales_count", 0))
    b_sales = int(b_row.get("active_sales_count", 0))
    return (
        f"同比 {_period_label(a_row)} vs {_period_label(b_row)}：AI 证据 {a_mentions}→{b_mentions}，"
        f"活跃销售 {a_sales}→{b_sales}。"
    )


def _find_yoy_pair(monthly_rows: Sequence[Dict[str, object]]) -> Tuple[Dict[str, object], Dict[str, object]] | None:
    if len(monthly_rows) < 2:
        return None
    latest = monthly_rows[-1]
    latest_month = int(latest.get("month", 0))
    latest_year = int(latest.get("year", 0))
    for row in reversed(monthly_rows[:-1]):
        if int(row.get("month", 0)) == latest_month and int(row.get("year", 0)) == latest_year - 1:
            return row, latest
    return None


def _confidence_from_pending(rate: float) -> str:
    if rate >= 0.4:
        return "low"
    if rate >= 0.2:
        return "medium"
    return "high"


def _period_label(row: Dict[str, object]) -> str:
    year = int(row.get("year", 0))
    month = int(row.get("month", 0))
    week = int(row.get("week_of_month", 0))
    if week:
        return f"{year:04d}-{month:02d}-W{week}"
    return f"{year:04d}-{month:02d}" if year and month else ""


def _build_yoy_summary(monthly_rows: Sequence[Dict[str, object]]) -> Dict[str, object]:
    pair = _find_yoy_pair(monthly_rows)
    if not pair:
        return {"has_compare": False}
    a_row, b_row = pair
    return {
        "has_compare": True,
        "year_a": int(a_row.get("year", 0)),
        "year_b": int(b_row.get("year", 0)),
        "month": int(b_row.get("month", 0)),
        "mentions_a": int(a_row.get("ai_mentions", 0)),
        "mentions_b": int(b_row.get("ai_mentions", 0)),
        "sales_a": int(a_row.get("active_sales_count", 0)),
        "sales_b": int(b_row.get("active_sales_count", 0)),
    }


def _review_priority(task: Dict[str, object]) -> int:
    reason = str(task.get("review_reason_code", ""))
    if reason == "ACTOR_OVERLAP":
        return 3
    if reason.startswith("PARSE_FAILED"):
        return 2
    if reason == "BUSINESSLINE_LOW_SIGNAL":
        return 1
    return 0
