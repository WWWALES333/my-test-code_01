from __future__ import annotations

from collections import Counter, defaultdict
from typing import Dict, Iterable, List, Sequence, Tuple

from .schema import TASK_STATUS_OPEN


def build_salesperson_profiles(
    evidence_facts: Sequence[Dict[str, object]],
    sales_roster: Sequence[Dict[str, object]],
) -> List[Dict[str, object]]:
    """生成销售画像和沉默销售基线。"""
    week_keys = sorted(
        {
            (int(row.get("year", 0)), int(row.get("month", 0)), int(row.get("week_of_month", 0)))
            for row in evidence_facts
            if int(row.get("year", 0)) > 0 and int(row.get("month", 0)) > 0
        }
    )
    recent_weeks = set(week_keys[-12:])
    latest_months = sorted({(year, month) for year, month, _ in week_keys})[-12:]

    evidence_by_sales: Dict[str, List[Dict[str, object]]] = defaultdict(list)
    for row in evidence_facts:
        salesperson_id = str(row.get("salesperson_id", ""))
        if salesperson_id:
            evidence_by_sales[salesperson_id].append(row)

    profiles: List[Dict[str, object]] = []
    roster_ids = set()
    for roster_row in sales_roster:
        if str(roster_row.get("employment_status", "")) != "active":
            continue
        salesperson_id = str(roster_row.get("salesperson_id", ""))
        roster_ids.add(salesperson_id)
        evidence_rows = evidence_by_sales.get(salesperson_id, [])
        profiles.append(_build_profile(roster_row, evidence_rows, recent_weeks, latest_months))

    for salesperson_id, rows in evidence_by_sales.items():
        if salesperson_id in roster_ids:
            continue
        first = rows[0]
        pseudo_roster = {
            "salesperson_id": salesperson_id,
            "display_name": str(first.get("salesperson_name", "")) or str(first.get("owner_hint", "")) or "历史对象",
            "flower_name": str(first.get("salesperson_name", "")),
            "battle_zone_name": str(first.get("battle_zone_name", "")),
            "region_name": str(first.get("region_name", "")),
            "employment_status": "historical_unmatched",
            "department_name": "",
            "job_title": "",
            "team_name": "",
            "org_full_name": "",
        }
        profiles.append(_build_profile(pseudo_roster, rows, recent_weeks, latest_months))

    return sorted(profiles, key=lambda item: (_segment_rank(str(item.get("segment", ""))), -int(item.get("total_mentions", 0)), str(item.get("display_name", ""))))


def build_region_sales_rollup(profiles: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    """按区域汇总销售构成，识别区域是否依赖少数销售。"""
    grouped: Dict[Tuple[str, str], List[Dict[str, object]]] = defaultdict(list)
    for row in profiles:
        region_name = str(row.get("region_name", "")) or "未识别区域"
        battle_zone_name = str(row.get("battle_zone_name", "")) or "未识别战区"
        grouped[(battle_zone_name, region_name)].append(row)

    rows: List[Dict[str, object]] = []
    for (battle_zone_name, region_name), items in grouped.items():
        active_items = [row for row in items if int(row.get("total_mentions", 0)) > 0]
        silent_items = [row for row in items if int(row.get("total_mentions", 0)) == 0 and str(row.get("employment_status", "")) == "active"]
        sorted_active = sorted(active_items, key=lambda item: (-int(item.get("total_mentions", 0)), str(item.get("display_name", ""))))
        top_mentions = int(sorted_active[0].get("total_mentions", 0)) if sorted_active else 0
        total_mentions = sum(int(item.get("total_mentions", 0)) for item in active_items)
        dominance_ratio = round(top_mentions / total_mentions, 4) if total_mentions else 0.0
        if not active_items:
            maturity = "未启动"
        elif dominance_ratio >= 0.6:
            maturity = "依赖少数销售"
        elif len(active_items) >= 3:
            maturity = "相对均衡"
        else:
            maturity = "起步中"
        rows.append(
            {
                "battle_zone_name": battle_zone_name,
                "region_name": region_name,
                "active_sales_count": len(active_items),
                "silent_sales_count": len(silent_items),
                "total_sales_count": len(items),
                "total_mentions": total_mentions,
                "top_salespeople": [
                    {
                        "salesperson_id": row.get("salesperson_id", ""),
                        "display_name": row.get("display_name", ""),
                        "total_mentions": row.get("total_mentions", 0),
                    }
                    for row in sorted_active[:5]
                ],
                "dominance_ratio": dominance_ratio,
                "maturity_judgement": maturity,
            }
        )
    return sorted(rows, key=lambda item: (-int(item.get("total_mentions", 0)), -int(item.get("active_sales_count", 0)), str(item.get("region_name", ""))))


def _build_profile(
    roster_row: Dict[str, object],
    evidence_rows: Sequence[Dict[str, object]],
    recent_weeks: set[Tuple[int, int, int]],
    latest_months: Sequence[Tuple[int, int]],
) -> Dict[str, object]:
    display_name = str(roster_row.get("display_name", "")) or str(roster_row.get("flower_name", ""))
    total_mentions = len(evidence_rows)
    first_seen = _format_period(min((_period_tuple(row) for row in evidence_rows), default=None))
    last_seen = _format_period(max((_period_tuple(row) for row in evidence_rows), default=None))
    recent_rows = [row for row in evidence_rows if _period_tuple(row, include_week=True) in recent_weeks]
    active_weeks = len({_period_tuple(row, include_week=True) for row in recent_rows})
    monthly_counter = Counter((_period_tuple(row) for row in evidence_rows))
    history = [
        {"period": f"{year:04d}-{month:02d}", "mentions": monthly_counter[(year, month)]}
        for year, month in latest_months
    ]

    actor_counter = Counter(str(row.get("actor_primary", "")) or "未标注" for row in evidence_rows)
    line_counter = Counter(str(row.get("business_line", "")) or "待判断" for row in evidence_rows)
    doctor_feedback_mentions = sum(1 for row in evidence_rows if str(row.get("actor_primary", "")) == "医生反馈")
    opportunity_mentions = sum(1 for row in evidence_rows if str(row.get("actor_primary", "")) == "潜在 AI 机会")
    review_open_count = sum(1 for row in evidence_rows if str(row.get("review_status", "")) == TASK_STATUS_OPEN)
    high_quality_evidence = [
        {
            "report_id": str(row.get("report_id", "")),
            "segment_id": str(row.get("segment_id", "")),
            "source_text": str(row.get("source_text", "")),
        }
        for row in sorted(evidence_rows, key=lambda item: (-len(str(item.get("source_text", ""))), str(item.get("report_id", ""))))[:3]
        if int(total_mentions) > 0
    ]

    segment = _segment_profile(total_mentions, active_weeks)
    if str(roster_row.get("employment_status", "")) == "active" and total_mentions == 0:
        segment = "长期未提及者"
    recommended_case = total_mentions >= 6 and doctor_feedback_mentions >= 2 and review_open_count <= 1

    return {
        "salesperson_id": str(roster_row.get("salesperson_id", "")),
        "display_name": display_name,
        "flower_name": str(roster_row.get("flower_name", "")),
        "battle_zone_name": str(roster_row.get("battle_zone_name", "")),
        "region_name": str(roster_row.get("region_name", "")),
        "employment_status": str(roster_row.get("employment_status", "")),
        "team_name": str(roster_row.get("team_name", "")),
        "department_name": str(roster_row.get("department_name", "")),
        "job_title": str(roster_row.get("job_title", "")),
        "segment": segment,
        "first_seen_period": first_seen,
        "last_seen_period": last_seen,
        "total_mentions": total_mentions,
        "recent_active_weeks": active_weeks,
        "history": history,
        "actor_breakdown": dict(actor_counter),
        "business_line_breakdown": dict(line_counter),
        "doctor_feedback_mentions": doctor_feedback_mentions,
        "opportunity_mentions": opportunity_mentions,
        "review_open_count": review_open_count,
        "high_quality_evidence": high_quality_evidence,
        "recommended_case": recommended_case,
    }


def _segment_profile(total_mentions: int, active_weeks: int) -> str:
    if total_mentions >= 10 or active_weeks >= 6:
        return "高频使用者"
    if total_mentions >= 4 or active_weeks >= 3:
        return "中频使用者"
    if total_mentions >= 1:
        return "偶发使用者"
    return "长期未提及者"


def _format_period(period: Tuple[int, int] | None) -> str:
    if not period:
        return ""
    year, month = period
    return f"{year:04d}-{month:02d}"


def _period_tuple(row: Dict[str, object], include_week: bool = False) -> Tuple[int, int] | Tuple[int, int, int]:
    if include_week:
        return int(row.get("year", 0)), int(row.get("month", 0)), int(row.get("week_of_month", 0))
    return int(row.get("year", 0)), int(row.get("month", 0))


def _segment_rank(segment: str) -> int:
    order = {
        "高频使用者": 0,
        "中频使用者": 1,
        "偶发使用者": 2,
        "长期未提及者": 3,
    }
    return order.get(segment, 9)
