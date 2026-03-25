from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Dict, List


def build_summary_markdown(
    report_rows: List[Dict[str, object]],
    tag_rows: List[Dict[str, object]],
    evidence_rows: List[Dict[str, object]],
    review_rows: List[Dict[str, object]],
) -> str:
    ai_hits = [r for r in tag_rows if bool(r["is_ai_hit"])]
    actor_counter = Counter(str(r["ai_actor"]) for r in ai_hits)
    line_counter = Counter(str(r["business_line"]) for r in ai_hits)

    lines: List[str] = []
    lines.append("# AI专题摘要")
    lines.append("")
    lines.append("## 结论摘要")
    lines.append(f"- 文档总数：{len(report_rows)}")
    lines.append(f"- AI命中片段数：{len(ai_hits)}")
    lines.append(f"- 待复核项：{len(review_rows)}")
    lines.append("")
    lines.append("## 主体分布")
    if actor_counter:
        for key, val in actor_counter.most_common():
            lines.append(f"- {key}：{val}")
    else:
        lines.append("- 暂无命中")
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

    return "\n".join(lines)


def write_markdown(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def trim_text(text: str, max_len: int) -> str:
    return text if len(text) <= max_len else text[: max_len - 1] + "…"

