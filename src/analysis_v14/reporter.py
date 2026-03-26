from __future__ import annotations

import html
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, List, Set, Tuple


def build_summary_markdown(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
    run_meta: Dict[str, str] | None = None,
) -> str:
    report_map = {str(r["report_id"]): r for r in report_rows}
    ai_hits = [r for r in tag_rows if bool(r["is_ai_hit"])]
    confirmed_ai_hits = [r for r in ai_hits if r.get("decision_status") == "confirmed"]
    pending_ai_hits = [r for r in ai_hits if r.get("decision_status") in {"uncertain", "pending_human_review"}]

    actor_counter = Counter(str(r["actor_primary"]) for r in confirmed_ai_hits)
    line_counter = Counter(str(r["business_line"]) for r in confirmed_ai_hits)
    scope_counter = Counter(str(r["ai_scope"]) for r in confirmed_ai_hits)
    main_kpi_hits = [
        r
        for r in confirmed_ai_hits
        if r["ai_scope"] == "product_ai" or r["actor_primary"] == "潜在 AI 机会"
    ]
    market_radar_hits = [r for r in ai_hits if r["ai_scope"] in {"market_trend", "competitor_ai"}]
    market_scope_counter = Counter(str(r["ai_scope"]) for r in market_radar_hits)

    parse_status_counter = Counter(str(r.get("parse_status", r.get("text_status", "unknown"))) for r in report_rows)
    parse_reason_counter = Counter(
        str(r.get("parse_reason_code", ""))
        for r in report_rows
        if str(r.get("parse_reason_code", ""))
    )
    review_priority_counter = Counter(str(r.get("review_priority", "low")) for r in review_rows)
    review_reason_counter = Counter()
    for row in review_rows:
        for code in str(row.get("review_reason_code", "")).split(";"):
            if code:
                review_reason_counter[code] += 1

    quality_status, quality_reason = _evaluate_quality(
        report_count=len(report_rows),
        parse_failed_count=parse_status_counter.get("failed", 0),
        ai_hit_count=len(ai_hits),
        pending_count=len(pending_ai_hits),
    )
    self_check = _run_self_checks(report_rows, tag_rows, evidence_rows, review_rows)
    product_examples = _pick_evidence_examples(evidence_rows, preferred_scope={"product_ai"}, limit=5)
    market_examples = _pick_evidence_examples(
        evidence_rows,
        preferred_scope={"market_trend", "competitor_ai"},
        limit=3,
    )

    trend = _build_year_trend(report_map, ai_hits)
    opp_examples = _pick_opportunity_examples(confirmed_ai_hits, report_map, limit=5)

    lines: List[str] = []
    lines.append("# AI专题业务验收简报")
    lines.append("")
    lines.append("## 一页结论（先给业务看）")
    lines.append(f"- 本轮处理文档：{len(report_rows)} 份；AI 命中片段：{len(ai_hits)} 条。")
    lines.append(f"- 可直接参考（已确认）片段：{len(confirmed_ai_hits)} 条。")
    lines.append(f"- 待人工复核片段：{len(pending_ai_hits)} 条（队列总量：{len(review_rows)}）。")
    lines.append(f"- 主业务 AI 信号（product_ai + 潜在机会）：{len(main_kpi_hits)} 条。")
    lines.append(f"- 市场雷达信号（趋势+竞品）：{len(market_radar_hits)} 条。")
    lines.append(f"- 试运行可用性结论：**{quality_status}**（{quality_reason}）。")
    lines.append("")
    lines.append("## 现状如何")
    lines.append(f"- 当前已确认的 AI 信号中，销售对外介绍占比最高（{actor_counter.get('销售对外介绍', 0)} 条）。")
    lines.append(f"- 医生反馈相关信号：{actor_counter.get('医生反馈', 0)} 条；潜在机会信号：{actor_counter.get('潜在 AI 机会', 0)} 条。")
    lines.append(
        f"- 业务线分布：云诊室 {line_counter.get('云诊室', 0)}，"
        f"云管家 {line_counter.get('云管家', 0)}，混合 {line_counter.get('混合', 0)}。"
    )
    lines.append("")
    lines.append("## 趋势如何")
    if trend["has_compare"]:
        lines.append(
            f"- 对比 {trend['year_a']} 同期({trend['compare_months']}月) -> {trend['year_b']} 同期：AI 提及片段 "
            f"{trend['ai_mentions_a']} -> {trend['ai_mentions_b']}（{trend['ai_mentions_change_text']}）。"
        )
        lines.append(
            f"- 对比 {trend['year_a']} 同期({trend['compare_months']}月) -> {trend['year_b']} 同期：AI 提及的销售主体数 "
            f"{trend['sales_entities_a']} -> {trend['sales_entities_b']}（{trend['sales_entities_change_text']}）。"
        )
    else:
        lines.append("- 当前输入仅覆盖单一年份，无法自动给出跨年趋势；建议将 2025+2026 一起跑批。")
    lines.append(
        f"- 市场雷达分布：market_trend={market_scope_counter.get('market_trend', 0)}，"
        f"competitor_ai={market_scope_counter.get('competitor_ai', 0)}。"
    )
    lines.append("")
    lines.append("## 可反哺业务的机会点")
    if opp_examples:
        lines.append("- 建议优先跟进以下高价值线索（含可追溯原文）：")
        for row in opp_examples:
            lines.append(
                f"- [{row['report_id']}/{row['segment_id']}] {row['owner_hint']} | "
                f"{row['ai_scope']} | {trim_text(str(row['source_text']), 92)}"
            )
    else:
        lines.append("- 当前未筛出明显机会线索。")
    lines.append("")
    lines.append("## 给产品负责人的重点")
    lines.append("- 主体分布（仅已确认片段）：")
    if actor_counter:
        for key, val in actor_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无可用主体信号")
    lines.append("- 主业务 AI 重点证据：")
    if product_examples:
        for row in product_examples:
            lines.append(f"- [{row['report_id']}/{row['segment_id']}] {trim_text(str(row['source_text']), 96)}")
    else:
        lines.append("- 暂无")
    lines.append("")
    lines.append("## 给销售管理者的重点")
    lines.append("- 业务线分布（仅已确认片段）：")
    if line_counter:
        for key, val in line_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无")
    lines.append("- 建议动作：")
    lines.append(f"- 优先处理高优先级复核：{review_priority_counter.get('high', 0)} 条。")
    lines.append(f"- 可直接复用片段池（已确认）：{len(confirmed_ai_hits)} 条。")
    lines.append(f"- 趋势/竞品观察片段：{len(market_examples)} 条（用于周会风险提示）。")
    if market_examples:
        lines.append("- 市场雷达代表证据：")
        for row in market_examples:
            lines.append(f"- [{row['report_id']}/{row['segment_id']}] {trim_text(str(row['source_text']), 96)}")
    lines.append("")
    lines.append("## 人工复核工作台")
    lines.append(f"- 高优先级：{review_priority_counter.get('high', 0)} 条")
    lines.append(f"- 中优先级：{review_priority_counter.get('medium', 0)} 条")
    lines.append(f"- 低优先级：{review_priority_counter.get('low', 0)} 条")
    lines.append("- 主要进入复核原因：")
    if review_reason_counter:
        for code, count in review_reason_counter.most_common(6):
            lines.append(f"- {code}：{count}（{_review_reason_label(code)}）")
    else:
        lines.append("- 暂无")
    lines.append("")
    lines.append("## 系统自测结果（本轮自动验收）")
    lines.append(f"- 解析成功率：{_ratio(parse_status_counter.get('success', 0), len(report_rows))}")
    lines.append(f"- 证据可追溯率：{_ratio(self_check['traceable_evidence'], self_check['evidence_total'])}")
    lines.append(f"- 复核状态合规率：{_ratio(self_check['review_status_valid'], self_check['review_total'])}")
    lines.append(f"- 自测结论：**{'通过' if self_check['all_passed'] else '需关注'}**")
    if self_check["issues"]:
        lines.append("- 自测告警：")
        for issue in self_check["issues"]:
            lines.append(f"- {issue}")
    lines.append("")
    lines.append("## 运行元信息（技术附录）")
    if run_meta:
        lines.append(f"- run_id：{run_meta.get('run_id', '')}")
        lines.append(f"- model_mode：{run_meta.get('model_mode', '')}")
        lines.append(f"- model_name：{run_meta.get('model_name', '')}")
        lines.append(f"- samples_dir：{run_meta.get('samples_dir', '')}")
    else:
        lines.append("- run_meta 未提供")
    lines.append("")
    lines.append("## 附录统计")
    lines.append("- 范围分布（已确认片段）：")
    if scope_counter:
        for key, val in scope_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无")
    lines.append("- 解析失败原因：")
    if parse_reason_counter:
        for key, val in parse_reason_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 无")
    lines.append("")
    lines.append("## 待复核示例（前10条）")
    if review_rows:
        for row in review_rows[:10]:
            lines.append(
                f"- [{row['report_id']}/{row['segment_id']}] {row['review_reason']} | "
                f"{trim_text(str(row['source_text']), 80)}"
            )
    else:
        lines.append("- 暂无待复核项")
    lines.append("")
    lines.append("## 追溯与口径说明")
    lines.append("- 所有结论应追溯到 `evidence_span.jsonl` 的原文片段。")
    lines.append("- 所有证据应追溯到原始文档路径。")
    lines.append("- `product_ai` 与 `market_trend/competitor_ai` 分开展示，避免口径混淆。")
    return "\n".join(lines)


def build_business_tables(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
) -> Dict[str, List[Dict[str, object]]]:
    report_map = {str(r["report_id"]): r for r in report_rows}
    report_ai_mentions: Dict[str, int] = defaultdict(int)
    report_confirmed_mentions: Dict[str, int] = defaultdict(int)
    report_pending_mentions: Dict[str, int] = defaultdict(int)
    report_has_ai: Set[str] = set()
    report_has_confirmed_ai: Set[str] = set()

    for row in tag_rows:
        report_id = str(row.get("report_id", ""))
        if not bool(row.get("is_ai_hit", False)):
            continue
        report_ai_mentions[report_id] += 1
        report_has_ai.add(report_id)
        if str(row.get("decision_status", "")) == "confirmed":
            report_confirmed_mentions[report_id] += 1
            report_has_confirmed_ai.add(report_id)
        else:
            report_pending_mentions[report_id] += 1

    weekly = _aggregate_period_rows(report_rows, report_ai_mentions, report_confirmed_mentions, report_pending_mentions, level="weekly")
    monthly = _aggregate_period_rows(report_rows, report_ai_mentions, report_confirmed_mentions, report_pending_mentions, level="monthly")

    actor_trend_counter: Dict[Tuple[int, int, str], int] = defaultdict(int)
    business_trend_counter: Dict[Tuple[int, int, str], int] = defaultdict(int)
    for row in tag_rows:
        if not bool(row.get("is_ai_hit", False)):
            continue
        report = report_map.get(str(row.get("report_id", "")))
        if not report:
            continue
        y = int(report.get("year", 0))
        m = int(report.get("month", 0))
        actor = str(row.get("actor_primary", "待判断"))
        line = str(row.get("business_line", "待判断"))
        actor_trend_counter[(y, m, actor)] += 1
        business_trend_counter[(y, m, line)] += 1

    actor_trend_rows: List[Dict[str, object]] = []
    for (y, m, actor), count in sorted(actor_trend_counter.items()):
        actor_trend_rows.append({"year": y, "month": m, "actor_primary": actor, "mentions": count})

    business_trend_rows: List[Dict[str, object]] = []
    for (y, m, line), count in sorted(business_trend_counter.items()):
        business_trend_rows.append({"year": y, "month": m, "business_line": line, "mentions": count})

    opportunity_rows: List[Dict[str, object]] = []
    for row in tag_rows:
        if not bool(row.get("is_ai_hit", False)):
            continue
        if str(row.get("actor_primary", "")) != "潜在 AI 机会" and str(row.get("ai_scope", "")) not in {"market_trend", "competitor_ai"}:
            continue
        report = report_map.get(str(row.get("report_id", "")))
        if not report:
            continue
        file_path = str(row.get("file_path", ""))
        opportunity_rows.append(
            {
                "year": int(report.get("year", 0)),
                "month": int(report.get("month", 0)),
                "week_of_month": int(report.get("week_of_month", 0)),
                "report_id": str(row.get("report_id", "")),
                "segment_id": str(row.get("segment_id", "")),
                "owner_hint": _extract_owner_hint(file_path),
                "ai_scope": str(row.get("ai_scope", "")),
                "actor_primary": str(row.get("actor_primary", "")),
                "business_line": str(row.get("business_line", "")),
                "decision_status": str(row.get("decision_status", "")),
                "confidence": row.get("confidence", ""),
                "source_text": str(row.get("source_text", "")),
                "file_path": file_path,
            }
        )

    trace_rows: List[Dict[str, object]] = []
    for row in evidence_rows:
        report = report_map.get(str(row.get("report_id", "")))
        file_path = str(row.get("file_path", ""))
        trace_rows.append(
            {
                "year": int(report.get("year", 0)) if report else 0,
                "month": int(report.get("month", 0)) if report else 0,
                "week_of_month": int(report.get("week_of_month", 0)) if report else 0,
                "report_id": str(row.get("report_id", "")),
                "segment_id": str(row.get("segment_id", "")),
                "owner_hint": _extract_owner_hint(file_path),
                "business_line": str(row.get("business_line", "")),
                "actor_primary": str(row.get("actor_primary", "")),
                "ai_scope": str(row.get("ai_scope", "")),
                "decision_status": str(row.get("decision_status", "")),
                "source_text": str(row.get("source_text", "")),
                "file_path": file_path,
            }
        )

    review_rows_export: List[Dict[str, object]] = []
    for row in review_rows:
        report = report_map.get(str(row.get("report_id", "")))
        file_path = str(row.get("file_path", ""))
        review_rows_export.append(
            {
                "year": int(report.get("year", 0)) if report else 0,
                "month": int(report.get("month", 0)) if report else 0,
                "week_of_month": int(report.get("week_of_month", 0)) if report else 0,
                "report_id": str(row.get("report_id", "")),
                "segment_id": str(row.get("segment_id", "")),
                "owner_hint": _extract_owner_hint(file_path),
                "review_priority": str(row.get("review_priority", "")),
                "review_reason_code": str(row.get("review_reason_code", "")),
                "review_reason": str(row.get("review_reason", "")),
                "decision_status": str(row.get("decision_status", "")),
                "source_text": str(row.get("source_text", "")),
                "file_path": file_path,
            }
        )

    return {
        "dashboard_weekly": weekly,
        "dashboard_monthly": monthly,
        "dashboard_actor_trend": actor_trend_rows,
        "dashboard_business_line_trend": business_trend_rows,
        "opportunity_backlog": opportunity_rows,
        "evidence_trace": trace_rows,
        "review_worklist": review_rows_export,
    }


def write_markdown(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def trim_text(text: str, max_len: int) -> str:
    return text if len(text) <= max_len else text[: max_len - 1] + "…"


def _ratio(numerator: int, denominator: int) -> str:
    if denominator <= 0:
        return "N/A"
    return f"{(numerator / denominator) * 100:.1f}% ({numerator}/{denominator})"


def _evaluate_quality(
    report_count: int,
    parse_failed_count: int,
    ai_hit_count: int,
    pending_count: int,
) -> Tuple[str, str]:
    if report_count <= 0:
        return "不可用", "无输入文档"
    parse_fail_rate = parse_failed_count / report_count
    pending_rate = (pending_count / ai_hit_count) if ai_hit_count else 0.0
    if parse_fail_rate > 0.08:
        return "需谨慎使用", "解析失败比例偏高"
    if pending_rate > 0.55:
        return "需谨慎使用", "待复核比例较高"
    if pending_rate > 0.35:
        return "可用于试运行", "结论可看，但需优先处理复核队列"
    return "可用于业务试读", "结论稳定性处于可用区间"


def _pick_evidence_examples(
    evidence_rows: List[Dict[str, object]],
    preferred_scope: Set[str],
    limit: int,
) -> List[Dict[str, object]]:
    picked: List[Dict[str, object]] = []
    for row in evidence_rows:
        if str(row.get("ai_scope", "")) not in preferred_scope:
            continue
        picked.append(row)
        if len(picked) >= limit:
            break
    return picked


def _pick_opportunity_examples(
    tag_rows: List[Dict[str, object]],
    report_map: Dict[str, Dict[str, object]],
    limit: int,
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    for row in tag_rows:
        scope = str(row.get("ai_scope", ""))
        actor = str(row.get("actor_primary", ""))
        if actor != "潜在 AI 机会" and scope not in {"market_trend", "competitor_ai"}:
            continue
        file_path = str(row.get("file_path", ""))
        rows.append(
            {
                "report_id": str(row.get("report_id", "")),
                "segment_id": str(row.get("segment_id", "")),
                "ai_scope": scope,
                "source_text": str(row.get("source_text", "")),
                "owner_hint": _extract_owner_hint(file_path),
                "file_path": file_path,
            }
        )
        if len(rows) >= limit:
            break
    return rows


def _review_reason_label(code: str) -> str:
    mapping = {
        "ACTOR_OVERLAP": "主体语义重叠，需人工确认",
        "SCOPE_AMBIGUOUS": "范围口径边界不清",
        "BROAD_STATEMENT": "表达宽泛，缺少业务动作",
        "BUSINESSLINE_LOW_SIGNAL": "业务线线索不足",
        "OUTCOME_UNCLEAR": "反馈结果不明确",
        "PARSE_FAILED": "解析失败",
        "PARSE_FAILED_DOC": "DOC 解析失败",
        "PARSE_FAILED_PDF": "PDF 解析失败",
        "PARSER_TOOL_MISSING": "解析工具缺失",
        "MODEL_CALL_FAILED": "模型调用失败",
    }
    return mapping.get(code, "未定义原因")


def _run_self_checks(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
) -> Dict[str, object]:
    issues: List[str] = []
    tag_key_set = {
        (str(row.get("report_id", "")), str(row.get("segment_id", "")))
        for row in tag_rows
    }
    traceable_evidence = 0
    for row in evidence_rows:
        if (
            str(row.get("source_text", "")).strip()
            and str(row.get("file_path", "")).strip()
            and (str(row.get("report_id", "")), str(row.get("segment_id", ""))) in tag_key_set
        ):
            traceable_evidence += 1
    review_status_valid = 0
    for row in review_rows:
        if str(row.get("decision_status", "")) in {"uncertain", "pending_human_review"}:
            review_status_valid += 1
    if review_status_valid != len(review_rows):
        issues.append("复核队列中存在非复核状态项")
    parse_failed = sum(1 for row in report_rows if str(row.get("parse_status", "")) == "failed")
    parse_failed_in_review = sum(
        1
        for row in review_rows
        if str(row.get("review_reason_code", "")).startswith("PARSE_FAILED")
        or str(row.get("review_reason_code", "")) == "PARSER_TOOL_MISSING"
    )
    if parse_failed > parse_failed_in_review:
        issues.append("部分解析失败文档未进入复核队列")
    if evidence_rows and traceable_evidence < len(evidence_rows):
        issues.append("存在无法回链到标签结果的证据项")
    return {
        "evidence_total": len(evidence_rows),
        "traceable_evidence": traceable_evidence,
        "review_total": len(review_rows),
        "review_status_valid": review_status_valid,
        "all_passed": len(issues) == 0,
        "issues": issues,
    }


def _extract_owner_hint(file_path: str) -> str:
    stem = Path(file_path).stem
    zone_patterns = [
        r"[一二三四五六七八九十0-9]+战区[（(][^）)]+[）)]",
        r"[一二三四五六七八九十0-9]+战区[\u4e00-\u9fa5]{0,8}",
        r"线上战区",
        r"[\u4e00-\u9fa5]{2,8}区域",
    ]
    for pattern in zone_patterns:
        m = re.search(pattern, stem)
        if m:
            return m.group(0).strip()

    m = re.search(r"[-_—]\s*([\u4e00-\u9fa5]{2,6})\s*$", stem)
    if m:
        return m.group(1).strip()

    parts = re.split(r"[-_—]", stem)
    for part in reversed(parts):
        token = _clean_owner_token(part)
        if token:
            return token
    token = _clean_owner_token(stem)
    return token if token else stem[:30]


def _clean_owner_token(raw: str) -> str:
    token = raw.strip()
    token = re.sub(r"20\d{2}", "", token)
    token = re.sub(r"\d{1,2}月", "", token)
    token = re.sub(r"第?\d{1,2}周", "", token)
    token = re.sub(r"[（(].*?[）)]", "", token)
    token = re.sub(r"(将军汤|工作周报|周报|月报|部门|区域|战区|整理|提交|线上|模板)", "", token)
    token = re.sub(r"[0-9. ]+", "", token)
    token = token.strip(" -")
    if not token:
        return ""
    if len(token) > 20:
        return token[:20]
    return token


def _build_year_trend(report_map: Dict[str, Dict[str, object]], ai_hits: List[Dict[str, object]]) -> Dict[str, object]:
    year_month_mentions: Dict[Tuple[int, int], int] = defaultdict(int)
    year_month_entities: Dict[Tuple[int, int], Set[str]] = defaultdict(set)
    year_months: Dict[int, Set[int]] = defaultdict(set)

    for row in ai_hits:
        report = report_map.get(str(row.get("report_id", "")))
        if not report:
            continue
        year = int(report.get("year", 0))
        month = int(report.get("month", 0))
        if year <= 0:
            continue
        if month <= 0:
            continue
        year_month_mentions[(year, month)] += 1
        year_months[year].add(month)
        file_path = str(row.get("file_path", ""))
        year_month_entities[(year, month)].add(_extract_owner_hint(file_path))

    years = sorted(year_months.keys())
    if len(years) < 2:
        return {"has_compare": False}

    year_a = years[-2]
    year_b = years[-1]
    compare_months = sorted(year_months[year_b])
    if not compare_months:
        return {"has_compare": False}

    a_mentions = 0
    b_mentions = 0
    a_entities_set: Set[str] = set()
    b_entities_set: Set[str] = set()
    for month in compare_months:
        a_mentions += year_month_mentions.get((year_a, month), 0)
        b_mentions += year_month_mentions.get((year_b, month), 0)
        a_entities_set.update(year_month_entities.get((year_a, month), set()))
        b_entities_set.update(year_month_entities.get((year_b, month), set()))

    a_entities = len(a_entities_set)
    b_entities = len(b_entities_set)
    month_text = ",".join(str(m) for m in compare_months)
    return {
        "has_compare": True,
        "year_a": year_a,
        "year_b": year_b,
        "compare_months": month_text,
        "ai_mentions_a": a_mentions,
        "ai_mentions_b": b_mentions,
        "ai_mentions_change_text": _delta_text(a_mentions, b_mentions),
        "sales_entities_a": a_entities,
        "sales_entities_b": b_entities,
        "sales_entities_change_text": _delta_text(a_entities, b_entities),
    }


def _delta_text(base: int, current: int) -> str:
    if base <= 0:
        return f"新增 {current}"
    delta = current - base
    pct = (delta / base) * 100
    if delta >= 0:
        return f"+{delta}（+{pct:.1f}%）"
    return f"{delta}（{pct:.1f}%）"


def _aggregate_period_rows(
    report_rows: List[Dict[str, object]],
    report_ai_mentions: Dict[str, int],
    report_confirmed_mentions: Dict[str, int],
    report_pending_mentions: Dict[str, int],
    level: str,
) -> List[Dict[str, object]]:
    agg: Dict[Tuple[int, int, int], Dict[str, object]] = {}

    for report in report_rows:
        report_id = str(report.get("report_id", ""))
        year = int(report.get("year", 0))
        month = int(report.get("month", 0))
        week = int(report.get("week_of_month", 0)) if level == "weekly" else 0
        key = (year, month, week)
        if key not in agg:
            agg[key] = {
                "year": year,
                "month": month,
                "week_of_month": week,
                "reports": 0,
                "ai_reports": 0,
                "ai_mentions": 0,
                "confirmed_mentions": 0,
                "pending_mentions": 0,
                "unique_owner_hints_with_ai": set(),
            }
        row = agg[key]
        row["reports"] = int(row["reports"]) + 1

        ai_mentions = report_ai_mentions.get(report_id, 0)
        confirmed = report_confirmed_mentions.get(report_id, 0)
        pending = report_pending_mentions.get(report_id, 0)
        if ai_mentions > 0:
            row["ai_reports"] = int(row["ai_reports"]) + 1
            row["unique_owner_hints_with_ai"].add(_extract_owner_hint(str(report.get("file_path", ""))))
        row["ai_mentions"] = int(row["ai_mentions"]) + ai_mentions
        row["confirmed_mentions"] = int(row["confirmed_mentions"]) + confirmed
        row["pending_mentions"] = int(row["pending_mentions"]) + pending

    rows: List[Dict[str, object]] = []
    for key in sorted(agg.keys()):
        row = agg[key]
        owners = row.pop("unique_owner_hints_with_ai")
        row["unique_owner_hints_with_ai"] = len(owners)
        reports = int(row["reports"])
        ai_reports = int(row["ai_reports"])
        row["ai_report_rate"] = f"{(ai_reports / reports * 100):.1f}%" if reports else "0.0%"
        rows.append(row)
    return rows


def build_dashboard_html(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
    business_tables: Dict[str, List[Dict[str, object]]],
    run_meta: Dict[str, str] | None = None,
) -> str:
    report_map = {str(r["report_id"]): r for r in report_rows}
    ai_hits = [r for r in tag_rows if bool(r["is_ai_hit"])]
    confirmed_ai_hits = [r for r in ai_hits if str(r.get("decision_status", "")) == "confirmed"]
    pending_ai_hits = [r for r in ai_hits if str(r.get("decision_status", "")) in {"uncertain", "pending_human_review"}]
    actor_counter = Counter(str(r.get("actor_primary", "待判断")) for r in confirmed_ai_hits)
    line_counter = Counter(str(r.get("business_line", "待判断")) for r in confirmed_ai_hits)
    review_reason_counter = Counter()
    for row in review_rows:
        for code in str(row.get("review_reason_code", "")).split(";"):
            if code:
                review_reason_counter[code] += 1

    parse_status_counter = Counter(str(r.get("parse_status", r.get("text_status", "unknown"))) for r in report_rows)
    quality_status, quality_reason = _evaluate_quality(
        report_count=len(report_rows),
        parse_failed_count=parse_status_counter.get("failed", 0),
        ai_hit_count=len(ai_hits),
        pending_count=len(pending_ai_hits),
    )
    trend = _build_year_trend(report_map, ai_hits)

    monthly_rows_all = business_tables.get("dashboard_monthly", [])
    monthly_rows = [r for r in monthly_rows_all if int(r.get("year", 0)) >= 2025]
    monthly_rows = sorted(monthly_rows, key=lambda r: (int(r.get("year", 0)), int(r.get("month", 0))))
    max_month_mentions = max([int(r.get("ai_mentions", 0)) for r in monthly_rows], default=1)

    opportunity_rows = business_tables.get("opportunity_backlog", [])[:20]
    review_rows_top = business_tables.get("review_worklist", [])[:20]
    evidence_rows_top = business_tables.get("evidence_trace", [])[:30]

    cards = [
        ("处理文档", str(len(report_rows)), "本轮扫描到的周报/月报总量"),
        ("AI命中", str(len(ai_hits)), "提及AI的片段总数"),
        ("已确认", str(len(confirmed_ai_hits)), "可直接参考的片段"),
        ("待复核", str(len(review_rows)), "需要业务人工确认"),
        ("解析成功率", _ratio(parse_status_counter.get("success", 0), len(report_rows)).split(" ")[0], "文档解析可用性"),
        ("试运行状态", quality_status, quality_reason),
    ]

    actor_bar = _render_bar_list(actor_counter, max_items=6)
    line_bar = _render_bar_list(line_counter, max_items=6)
    review_bar = _render_bar_list(review_reason_counter, max_items=8)
    monthly_table = _render_monthly_rows(monthly_rows, max_month_mentions)
    opp_table = _render_generic_table(
        rows=opportunity_rows,
        columns=[
            ("year", "年"),
            ("month", "月"),
            ("owner_hint", "销售/战区"),
            ("ai_scope", "范围"),
            ("actor_primary", "主体"),
            ("source_text", "原文片段"),
        ],
    )
    review_table = _render_generic_table(
        rows=review_rows_top,
        columns=[
            ("year", "年"),
            ("month", "月"),
            ("owner_hint", "销售/战区"),
            ("review_priority", "优先级"),
            ("review_reason_code", "原因码"),
            ("source_text", "待确认片段"),
        ],
    )
    trace_table = _render_generic_table(
        rows=evidence_rows_top,
        columns=[
            ("year", "年"),
            ("month", "月"),
            ("owner_hint", "销售/战区"),
            ("ai_scope", "范围"),
            ("actor_primary", "主体"),
            ("source_text", "证据原文"),
        ],
    )

    trend_html = ""
    if trend.get("has_compare"):
        trend_html = (
            f"<p><strong>同比趋势（{trend['year_a']} 同期 {trend['compare_months']}月 vs {trend['year_b']} 同期）</strong></p>"
            f"<p>AI提及片段：{trend['ai_mentions_a']} → {trend['ai_mentions_b']}（{trend['ai_mentions_change_text']}）</p>"
            f"<p>提及主体数：{trend['sales_entities_a']} → {trend['sales_entities_b']}（{trend['sales_entities_change_text']}）</p>"
        )
    else:
        trend_html = "<p>当前缺少可比年度数据，暂无法自动给出同比趋势。</p>"

    run_id = html.escape(str((run_meta or {}).get("run_id", "")))
    samples_dir = html.escape(str((run_meta or {}).get("samples_dir", "")))
    model_mode = html.escape(str((run_meta or {}).get("model_mode", "")))

    card_html = "".join(
        [
            (
                "<div class='card'>"
                f"<div class='k'>{html.escape(title)}</div>"
                f"<div class='v'>{html.escape(value)}</div>"
                f"<div class='d'>{html.escape(desc)}</div>"
                "</div>"
            )
            for title, value, desc in cards
        ]
    )

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>AI专题看板 - v1.4</title>
  <style>
    :root {{
      --bg1:#f4f8ff;
      --bg2:#f9f3ea;
      --card:#ffffff;
      --ink:#1f2a37;
      --muted:#5f6b7a;
      --line:#d9e0ea;
      --brand:#2266dd;
      --accent:#e58a2b;
      --good:#138a43;
      --warn:#a25f00;
    }}
    *{{box-sizing:border-box}}
    body {{
      margin:0;
      font-family:"PingFang SC","Noto Sans SC","Helvetica Neue",sans-serif;
      color:var(--ink);
      background:linear-gradient(140deg,var(--bg1),var(--bg2));
    }}
    .wrap {{
      max-width:1280px;
      margin:0 auto;
      padding:24px 28px 48px;
    }}
    .hero {{
      background:linear-gradient(120deg,#0f2f6b,#15419a);
      color:#fff;
      border-radius:16px;
      padding:24px;
      box-shadow:0 12px 28px rgba(15,47,107,.2);
    }}
    .hero h1 {{
      margin:0 0 8px;
      font-size:30px;
      letter-spacing:.5px;
    }}
    .hero p {{margin:4px 0;color:#dce7ff}}
    .meta {{
      margin-top:8px;
      font-size:13px;
      color:#bfd2ff;
      word-break:break-all;
    }}
    .grid {{
      margin-top:18px;
      display:grid;
      grid-template-columns:repeat(3,minmax(0,1fr));
      gap:12px;
    }}
    .card {{
      background:var(--card);
      border:1px solid var(--line);
      border-radius:14px;
      padding:14px 16px;
      box-shadow:0 4px 12px rgba(15,47,107,.06);
    }}
    .k {{font-size:13px;color:var(--muted)}}
    .v {{font-size:28px;font-weight:700;margin-top:6px}}
    .d {{font-size:12px;color:var(--muted);margin-top:4px;line-height:1.4}}
    .section {{
      background:var(--card);
      border:1px solid var(--line);
      border-radius:14px;
      margin-top:16px;
      padding:16px;
    }}
    .section h2 {{
      margin:0 0 10px;
      font-size:18px;
      color:#102a52;
    }}
    .two {{
      display:grid;
      grid-template-columns:1fr 1fr;
      gap:14px;
    }}
    .barlist {{display:grid;gap:8px}}
    .barrow {{display:grid;grid-template-columns:120px 1fr 52px;gap:8px;align-items:center}}
    .barlabel {{font-size:13px;color:#2a3a4f;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
    .bartrack {{height:10px;background:#edf2fb;border-radius:20px;overflow:hidden}}
    .bar {{height:100%;background:linear-gradient(90deg,var(--brand),#55a1ff)}}
    .barvalue {{font-size:12px;color:var(--muted);text-align:right}}
    table {{
      width:100%;
      border-collapse:collapse;
      font-size:13px;
    }}
    th,td {{
      border-bottom:1px solid var(--line);
      padding:8px 10px;
      text-align:left;
      vertical-align:top;
    }}
    th {{background:#f7f9fd;color:#24364f;position:sticky;top:0}}
    .table-wrap {{max-height:360px;overflow:auto;border:1px solid var(--line);border-radius:10px}}
    .chip {{
      display:inline-block;
      font-size:12px;
      border-radius:999px;
      padding:2px 8px;
      background:#eef3ff;
      color:#2558b7;
      margin-right:6px;
    }}
    .ok {{color:var(--good)}}
    .warn {{color:var(--warn)}}
    .monthbar {{
      height:8px;border-radius:999px;background:#eef2f9;overflow:hidden;min-width:120px
    }}
    .monthbar > span {{
      display:block;height:100%;background:linear-gradient(90deg,var(--accent),#ffd08f)
    }}
    @media (max-width: 960px) {{
      .grid {{grid-template-columns:1fr 1fr}}
      .two {{grid-template-columns:1fr}}
      .barrow {{grid-template-columns:100px 1fr 46px}}
    }}
    @media (max-width: 640px) {{
      .wrap {{padding:14px}}
      .grid {{grid-template-columns:1fr}}
      .hero h1 {{font-size:24px}}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <h1>AI专题业务看板（v1.4）</h1>
      <p>面向产品与销售管理的试运行报告：现状、趋势、机会、复核闭环</p>
      <div class="meta">run_id={run_id} ｜ model_mode={model_mode} ｜ samples={samples_dir}</div>
    </div>

    <div class="grid">{card_html}</div>

    <div class="section">
      <h2>趋势结论</h2>
      {trend_html}
      <p class="meta">说明：同比按“2026已覆盖月份”回看 2025 同期，避免跨月误读。</p>
    </div>

    <div class="two">
      <div class="section">
        <h2>主体分布（已确认）</h2>
        <div class="barlist">{actor_bar}</div>
      </div>
      <div class="section">
        <h2>业务线分布（已确认）</h2>
        <div class="barlist">{line_bar}</div>
      </div>
    </div>

    <div class="section">
      <h2>月度趋势看板（2025+）</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>月份</th><th>文档数</th><th>AI命中</th><th>AI覆盖率</th><th>已确认</th><th>待复核</th><th>趋势条</th></tr>
          </thead>
          <tbody>
            {monthly_table}
          </tbody>
        </table>
      </div>
    </div>

    <div class="section">
      <h2>机会池（可反哺业务）</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>年</th><th>月</th><th>销售/战区</th><th>范围</th><th>主体</th><th>原文片段</th></tr></thead>
          <tbody>{opp_table}</tbody>
        </table>
      </div>
    </div>

    <div class="section">
      <h2>复核工作台</h2>
      <p><span class="chip">总待复核 {len(review_rows)}</span><span class="chip">中高优先级 {sum(1 for r in review_rows if str(r.get('review_priority','')) in ('high','medium'))}</span></p>
      <div class="barlist">{review_bar}</div>
      <div class="table-wrap" style="margin-top:12px">
        <table>
          <thead><tr><th>年</th><th>月</th><th>销售/战区</th><th>优先级</th><th>原因码</th><th>待确认片段</th></tr></thead>
          <tbody>{review_table}</tbody>
        </table>
      </div>
    </div>

    <div class="section">
      <h2>证据溯源（前30条）</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>年</th><th>月</th><th>销售/战区</th><th>范围</th><th>主体</th><th>证据原文</th></tr></thead>
          <tbody>{trace_table}</tbody>
        </table>
      </div>
    </div>
  </div>
</body>
</html>"""


def _render_bar_list(counter: Counter, max_items: int) -> str:
    if not counter:
        return "<div class='meta'>暂无数据</div>"
    top_items = counter.most_common(max_items)
    max_value = max(v for _, v in top_items) if top_items else 1
    chunks: List[str] = []
    for label, value in top_items:
        width = 0 if max_value <= 0 else int(value / max_value * 100)
        chunks.append(
            "<div class='barrow'>"
            f"<div class='barlabel'>{html.escape(str(label))}</div>"
            f"<div class='bartrack'><div class='bar' style='width:{width}%'></div></div>"
            f"<div class='barvalue'>{value}</div>"
            "</div>"
        )
    return "".join(chunks)


def _render_monthly_rows(rows: List[Dict[str, object]], max_month_mentions: int) -> str:
    if not rows:
        return "<tr><td colspan='7'>暂无数据</td></tr>"
    rendered: List[str] = []
    for r in rows:
        year = int(r.get("year", 0))
        month = int(r.get("month", 0))
        reports = int(r.get("reports", 0))
        mentions = int(r.get("ai_mentions", 0))
        confirmed = int(r.get("confirmed_mentions", 0))
        pending = int(r.get("pending_mentions", 0))
        rate = str(r.get("ai_report_rate", "0.0%"))
        width = 0 if max_month_mentions <= 0 else int(mentions / max_month_mentions * 100)
        rendered.append(
            "<tr>"
            f"<td>{year}-{month:02d}</td>"
            f"<td>{reports}</td>"
            f"<td>{mentions}</td>"
            f"<td>{html.escape(rate)}</td>"
            f"<td>{confirmed}</td>"
            f"<td>{pending}</td>"
            f"<td><div class='monthbar'><span style='width:{width}%'></span></div></td>"
            "</tr>"
        )
    return "".join(rendered)


def _render_generic_table(rows: List[Dict[str, object]], columns: List[Tuple[str, str]]) -> str:
    if not rows:
        return f"<tr><td colspan='{len(columns)}'>暂无数据</td></tr>"
    rendered: List[str] = []
    for row in rows:
        cells: List[str] = []
        for key, _ in columns:
            val = row.get(key, "")
            text = html.escape(str(val))
            if key in {"source_text"} and len(text) > 180:
                text = text[:180] + "..."
            cells.append(f"<td>{text}</td>")
        rendered.append("<tr>" + "".join(cells) + "</tr>")
    return "".join(rendered)
