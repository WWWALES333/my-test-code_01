from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Dict, List


def build_summary_markdown(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
    run_meta: Dict[str, str] | None = None,
) -> str:
    ai_hits = [r for r in tag_rows if bool(r["is_ai_hit"])]
    actor_counter = Counter(str(r["actor_primary"]) for r in ai_hits)
    line_counter = Counter(str(r["business_line"]) for r in ai_hits)
    scope_counter = Counter(str(r["ai_scope"]) for r in ai_hits)
    main_kpi_hits = [r for r in ai_hits if r["ai_scope"] == "product_ai" or r["actor_primary"] == "潜在 AI 机会"]
    market_radar_hits = [r for r in ai_hits if r["ai_scope"] in {"market_trend", "competitor_ai"}]
    market_scope_counter = Counter(str(r["ai_scope"]) for r in market_radar_hits)
    parse_status_counter = Counter(str(r.get("parse_status", r.get("text_status", "unknown"))) for r in report_rows)
    parse_reason_counter = Counter(
        str(r.get("parse_reason_code", ""))
        for r in report_rows
        if str(r.get("parse_reason_code", ""))
    )

    lines: List[str] = []
    lines.append("# AI专题摘要")
    lines.append("")
    if run_meta:
        lines.append("## 运行信息")
        lines.append(f"- run_id：{run_meta.get('run_id', '')}")
        lines.append(f"- model_mode：{run_meta.get('model_mode', '')}")
        lines.append(f"- model_name：{run_meta.get('model_name', '')}")
        lines.append(f"- samples_dir：{run_meta.get('samples_dir', '')}")
        lines.append("")
    lines.append("## 结论摘要")
    lines.append(f"- 文档总数：{len(report_rows)}")
    lines.append(f"- AI命中片段数：{len(ai_hits)}")
    lines.append(f"- 主业务AI结论片段数：{len(main_kpi_hits)}")
    lines.append(f"- 市场雷达片段数：{len(market_radar_hits)}")
    lines.append(f"- 待复核项：{len(review_rows)}")
    lines.append("")
    lines.append("## 双视图统计")
    lines.append("- 主业务AI结论：`ai_scope=product_ai` 为主，`actor_primary=潜在 AI 机会` 可纳入。")
    lines.append("- 市场雷达结论：`ai_scope in {market_trend, competitor_ai}`。")
    lines.append("")
    lines.append("## 主体分布")
    if actor_counter:
        for key, val in actor_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无命中")
    lines.append("")
    lines.append("## 范围分布")
    if scope_counter:
        for key, val in scope_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无命中")
    lines.append("")
    lines.append("## 市场雷达分布")
    if market_scope_counter:
        for key, val in market_scope_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无市场雷达命中")
    lines.append("")
    lines.append("## 解析情况")
    for key, val in parse_status_counter.most_common():
        lines.append(f"- {key}：{val}")
    if parse_reason_counter:
        lines.append("- 失败原因分布：")
        for key, val in parse_reason_counter.most_common():
            lines.append(f"  - {key}：{val}")
    lines.append("")
    lines.append("## 业务线分布")
    if line_counter:
        for key, val in line_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无命中")
    lines.append("")
    lines.append("## 重点证据")
    if evidence_rows:
        for row in evidence_rows[:10]:
            lines.append(
                f"- [{row['report_id']}/{row['segment_id']}] {trim_text(str(row['source_text']), 80)}"
            )
    else:
        lines.append("- 暂无证据命中")
    lines.append("")
    lines.append("## 待复核项")
    if review_rows:
        for row in review_rows[:10]:
            lines.append(
                f"- [{row['report_id']}/{row['segment_id']}] {row['review_reason']} | {trim_text(str(row['source_text']), 80)}"
            )
    else:
        lines.append("- 暂无待复核项")
    lines.append("")
    lines.append("## 追溯说明")
    lines.append("- 所有结论应追溯到 `evidence_span.jsonl` 的原文片段。")
    lines.append("- 所有证据应追溯到原始文档路径。")
    lines.append("- `product_ai` 与 `market_trend/competitor_ai` 需分开展示，避免口径混淆。")

    return "\n".join(lines)


def write_markdown(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def trim_text(text: str, max_len: int) -> str:
    return text if len(text) <= max_len else text[: max_len - 1] + "…"
