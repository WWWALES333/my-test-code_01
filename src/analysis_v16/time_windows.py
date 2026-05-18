from __future__ import annotations

import os
from datetime import date, timedelta
from typing import Dict, Sequence


def compute_time_context(trend_cube: Sequence[Dict[str, object]], today: date | None = None) -> Dict[str, object]:
    today = today or _today()
    target_month_date = _previous_month(today)
    target_month = _month_key(target_month_date.year, target_month_date.month)
    mom_date = _previous_month(target_month_date)
    mom_month = _month_key(mom_date.year, mom_date.month)
    yoy_month = _month_key(target_month_date.year - 1, target_month_date.month)

    target_week_date = today - timedelta(days=7)
    target_week = _week_key(target_week_date.year, target_week_date.month, _week_of_month(target_week_date))
    previous_week_date = target_week_date - timedelta(days=7)
    wow_week = _week_key(previous_week_date.year, previous_week_date.month, _week_of_month(previous_week_date))
    yoy_week_date = target_week_date.replace(year=target_week_date.year - 1)
    yoy_week = _week_key(yoy_week_date.year, yoy_week_date.month, _week_of_month(yoy_week_date))

    months = [row for row in trend_cube if str(row.get("grain", "")) == "month"]
    weeks = [row for row in trend_cube if str(row.get("grain", "")) == "week"]
    month_map = {_row_month_key(row): row for row in months}
    week_map = {_row_week_key(row): row for row in weeks}
    latest_month = max(month_map) if month_map else ""
    latest_week = max(week_map) if week_map else ""

    return {
        "generated_on": today.isoformat(),
        "target_month": target_month,
        "target_week": target_week,
        "latest_available_month": latest_month,
        "latest_available_week": latest_week,
        "month_observation": _comparison(
            target_month,
            month_map.get(target_month),
            {
                "mom": (mom_month, month_map.get(mom_month)),
                "yoy": (yoy_month, month_map.get(yoy_month)),
            },
            latest_month,
        ),
        "week_observation": _comparison(
            target_week,
            week_map.get(target_week),
            {
                "wow": (wow_week, week_map.get(wow_week)),
                "yoy": (yoy_week, week_map.get(yoy_week)),
            },
            latest_week,
        ),
        "series": {
            "months": [month_map[key] for key in sorted(month_map)[-12:]],
            "weeks": [week_map[key] for key in sorted(week_map)[-12:]],
        },
    }


def _today() -> date:
    raw = os.environ.get("ANALYSIS_TODAY", "").strip()
    if raw:
        return date.fromisoformat(raw)
    return date.today()


def _previous_month(value: date) -> date:
    if value.month == 1:
        return date(value.year - 1, 12, 1)
    return date(value.year, value.month - 1, 1)


def _week_of_month(value: date) -> int:
    return (value.day - 1) // 7 + 1


def _month_key(year: int, month: int) -> str:
    return f"{year}-{month:02d}"


def _week_key(year: int, month: int, week: int) -> str:
    return f"{year}-{month:02d}-W{week}"


def _row_month_key(row: Dict[str, object]) -> str:
    return _month_key(int(row.get("year", 0) or 0), int(row.get("month", 0) or 0))


def _row_week_key(row: Dict[str, object]) -> str:
    return _week_key(
        int(row.get("year", 0) or 0),
        int(row.get("month", 0) or 0),
        int(row.get("week_of_month", 0) or 0),
    )


def _comparison(
    target_period: str,
    target_row: Dict[str, object] | None,
    baselines: Dict[str, tuple[str, Dict[str, object] | None]],
    latest_available: str,
) -> Dict[str, object]:
    available = target_row is not None
    payload = {
        "period": target_period,
        "available": available,
        "latest_available_period": latest_available,
        "row": target_row or {},
        "status_note": (
            f"目标周期 {target_period} 已有数据。"
            if available
            else f"目标周期 {target_period} 暂无产物；当前数据最新到 {latest_available or '无'}。"
        ),
        "baselines": {},
    }
    for name, (period, row) in baselines.items():
        payload["baselines"][name] = _baseline_payload(target_row, period, row)
    return payload


def _baseline_payload(target_row: Dict[str, object] | None, baseline_period: str, baseline_row: Dict[str, object] | None) -> Dict[str, object]:
    if not target_row or not baseline_row:
        return {
            "period": baseline_period,
            "available": False,
            "row": baseline_row or {},
            "delta_ai_mentions": None,
            "delta_active_sales": None,
            "note": f"基准 {baseline_period} 或目标周期缺少数据，暂不输出对比结论。",
        }
    target_mentions = int(target_row.get("ai_mentions", 0) or 0)
    baseline_mentions = int(baseline_row.get("ai_mentions", 0) or 0)
    target_sales = int(target_row.get("active_sales_count", 0) or 0)
    baseline_sales = int(baseline_row.get("active_sales_count", 0) or 0)
    return {
        "period": baseline_period,
        "available": True,
        "row": baseline_row,
        "delta_ai_mentions": target_mentions - baseline_mentions,
        "delta_active_sales": target_sales - baseline_sales,
        "note": _delta_note(target_mentions, baseline_mentions, target_sales, baseline_sales),
    }


def _delta_note(target_mentions: int, baseline_mentions: int, target_sales: int, baseline_sales: int) -> str:
    mention_delta = target_mentions - baseline_mentions
    sales_delta = target_sales - baseline_sales
    mention_direction = "增加" if mention_delta > 0 else "减少" if mention_delta < 0 else "持平"
    sales_direction = "增加" if sales_delta > 0 else "减少" if sales_delta < 0 else "持平"
    return f"AI 证据{mention_direction} {abs(mention_delta)} 条，活跃销售{sales_direction} {abs(sales_delta)} 人。"
