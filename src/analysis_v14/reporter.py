from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Dict, List, Set, Tuple


def build_summary_markdown(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
    run_meta: Dict[str, str] | None = None,
) -> str:
    ai_hits = [r for r in tag_rows if bool(r["is_ai_hit"])]
    confirmed_ai_hits = [r for r in ai_hits if r.get("decision_status") == "confirmed"]
    pending_ai_hits = [r for r in ai_hits if r.get("decision_status") in {"uncertain", "pending_human_review"}]

    actor_counter = Counter(str(r["actor_primary"]) for r in confirmed_ai_hits)
    line_counter = Counter(str(r["business_line"]) for r in confirmed_ai_hits)
    scope_counter = Counter(str(r["ai_scope"]) for r in confirmed_ai_hits)
    main_kpi_hits = [r for r in confirmed_ai_hits if r["ai_scope"] == "product_ai" or r["actor_primary"] == "潜在 AI 机会"]
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
    lines.append("## 给产品负责人的重点")
    lines.append("- 目标：判断 AI 专题是否值得继续投入，以及下一步应优化哪类能力。")
    lines.append("- 主体分布（仅已确认片段）：")
    if actor_counter:
        for key, val in actor_counter.most_common():
            lines.append(f"  - {key}：{val}")
    else:
        lines.append("  - 暂无可用主体信号")
    lines.append(f"- 潜在 AI 机会：{actor_counter.get('潜在 AI 机会', 0)} 条。")
    lines.append(
        f"- 市场雷达：`market_trend={market_scope_counter.get('market_trend', 0)}`，"
        f"`competitor_ai={market_scope_counter.get('competitor_ai', 0)}`。"
    )
    lines.append("- 主业务 AI 重点证据：")
    if product_examples:
        for row in product_examples:
            lines.append(f"  - [{row['report_id']}/{row['segment_id']}] {trim_text(str(row['source_text']), 96)}")
    else:
        lines.append("  - 暂无")
    lines.append("")
    lines.append("## 给销售管理者的重点")
    lines.append("- 目标：快速识别可复用话术线索与高风险待确认项。")
    lines.append("- 业务线分布（仅已确认片段）：")
    if line_counter:
        for key, val in line_counter.most_common():
            lines.append(f"  - {key}：{val}")
    else:
        lines.append("  - 暂无")
    lines.append("- 建议动作：")
    lines.append(f"  - 优先处理高优先级复核：{review_priority_counter.get('high', 0)} 条。")
    lines.append(f"  - 可直接复用片段池（已确认）：{len(confirmed_ai_hits)} 条。")
    lines.append(f"  - 趋势/竞品观察片段：{len(market_examples)} 条（用于周会风险提示）。")
    if market_examples:
        lines.append("- 市场雷达代表证据：")
        for row in market_examples:
            lines.append(f"  - [{row['report_id']}/{row['segment_id']}] {trim_text(str(row['source_text']), 96)}")
    lines.append("")
    lines.append("## 人工复核工作台")
    lines.append(f"- 高优先级：{review_priority_counter.get('high', 0)} 条")
    lines.append(f"- 中优先级：{review_priority_counter.get('medium', 0)} 条")
    lines.append(f"- 低优先级：{review_priority_counter.get('low', 0)} 条")
    lines.append("- 主要进入复核原因：")
    if review_reason_counter:
        for code, count in review_reason_counter.most_common(6):
            lines.append(f"  - {code}：{count}（{_review_reason_label(code)}）")
    else:
        lines.append("  - 暂无")
    lines.append("")
    lines.append("## 系统自测结果（本轮自动验收）")
    lines.append(f"- 解析成功率：{_ratio(parse_status_counter.get('success', 0), len(report_rows))}")
    lines.append(f"- 证据可追溯率：{_ratio(self_check['traceable_evidence'], self_check['evidence_total'])}")
    lines.append(f"- 复核状态合规率：{_ratio(self_check['review_status_valid'], self_check['review_total'])}")
    lines.append(f"- 自测结论：**{'通过' if self_check['all_passed'] else '需关注'}**")
    if self_check["issues"]:
        lines.append("- 自测告警：")
        for issue in self_check["issues"]:
            lines.append(f"  - {issue}")
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
            lines.append(f"  - {key}：{val}")
    else:
        lines.append("  - 暂无")
    lines.append("- 解析失败原因：")
    if parse_reason_counter:
        for key, val in parse_reason_counter.most_common():
            lines.append(f"  - {key}：{val}")
    else:
        lines.append("  - 无")
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
