from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Dict, Iterable, List


def write_jsonl(path: Path, rows: Iterable[Dict[str, object]]) -> None:
    """把结构化对象写入 JSONL 文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        for row in rows:
            file.write(json.dumps(row, ensure_ascii=False) + "\n")


def write_json(path: Path, payload: Dict[str, object]) -> None:
    """把快照对象写入 JSON 文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_markdown(path: Path, content: str) -> None:
    """把 Markdown 报告写入磁盘。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_csv_template(path: Path) -> None:
    """输出 v1.5 复核回写模板。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "sample_id,task_id,report_id,segment_id,salesperson_id,review_reason_code,current_labels,reviewed_fields,final_labels,is_pass,review_comment,reviewer,reviewed_at,need_rule_update,need_skill_update,need_annotation_update\n",
        encoding="utf-8",
    )


def build_summary_markdown(
    snapshot: Dict[str, object],
    insight_cards: List[Dict[str, object]],
    review_tasks: List[Dict[str, object]],
) -> str:
    """生成更偏业务判断的 v1.5 摘要。"""
    yoy = snapshot.get("yoy_summary", {})
    breadth = snapshot.get("breadth_depth_summary", {})
    actor_yoy = list(snapshot.get("actor_yoy", []))
    line_yoy = list(snapshot.get("business_line_yoy", []))
    top_people = list(snapshot.get("top_people_latest", []))
    feedback_examples = list(snapshot.get("feedback_examples", []))
    opportunity_examples = list(snapshot.get("opportunity_examples", []))
    review_reasons = list(snapshot.get("top_review_reasons", []))
    priority_tasks = list(snapshot.get("priority_review_tasks", []))

    lines: List[str] = ["# AI 一线情报工作台简报", "", "## 一句话判断"]
    if yoy.get("has_compare"):
        lines.append(
            f"- {yoy.get('year_b')} 年同期 AI 证据由 {yoy.get('mentions_a', 0)} 条上升到 {yoy.get('mentions_b', 0)} 条，"
            f"涉及销售对象由 {yoy.get('sales_a', 0)} 个上升到 {yoy.get('sales_b', 0)} 个。当前判断：{breadth.get('judgement', '趋势已出现，但仍需继续复核')}。"
        )
    else:
        lines.append(
            f"- 当前已沉淀 {snapshot.get('total_ai_mentions', 0)} 条 AI 相关证据，"
            f"覆盖 {snapshot.get('active_sales_count', 0)} 个销售对象，但暂无稳定同比窗口。"
        )

    lines.extend(
        [
            "",
            "## 1. 现状如何",
            f"- 最新统计月份：{snapshot.get('latest_year_month', '未知')}",
            f"- 当前 AI 证据：{snapshot.get('total_ai_mentions', 0)} 条，其中已确认 {snapshot.get('confirmed_mentions', 0)} 条。",
            f"- 当前识别到的销售对象：{snapshot.get('active_sales_count', 0)} 个，其中个人销售 {snapshot.get('active_person_count', 0)} 个，组织兜底对象 {snapshot.get('active_group_count', 0)} 个。",
            f"- 当前待复核任务：{snapshot.get('open_review_tasks', 0)} 条，已沉淀结论卡：{snapshot.get('insight_card_count', 0)} 张。",
            "",
            "## 2. 趋势如何",
        ]
    )
    if yoy.get("has_compare"):
        months = ",".join(str(item) for item in yoy.get("months", []))
        lines.append(
            f"- 对比 {yoy.get('year_a')} 年同期（{months}月）与 {yoy.get('year_b')} 年同期：AI 证据 {yoy.get('mentions_a', 0)} -> {yoy.get('mentions_b', 0)}，销售对象 {yoy.get('sales_a', 0)} -> {yoy.get('sales_b', 0)}。"
        )
        lines.append(
            f"- 单个活跃销售对象平均提及强度：{breadth.get('avg_mentions_per_sales_a', 0)} -> {breadth.get('avg_mentions_per_sales_b', 0)}。"
        )
    else:
        lines.append("- 当前缺少可比年度数据，暂无法输出稳定同比结论。")

    if actor_yoy:
        lines.append("- 主体结构变化：")
        for item in actor_yoy[:4]:
            lines.append(
                f"  - {item.get('label', '')}：{yoy.get('year_a', '')} 年 {item.get('year_a', 0)} 条，{yoy.get('year_b', '')} 年 {item.get('year_b', 0)} 条，变化 {item.get('delta', 0):+d}。"
            )
    if line_yoy:
        lines.append("- 业务线变化：")
        for item in line_yoy[:4]:
            lines.append(
                f"  - {item.get('label', '')}：{yoy.get('year_a', '')} 年 {item.get('year_a', 0)} 条，{yoy.get('year_b', '')} 年 {item.get('year_b', 0)} 条，变化 {item.get('delta', 0):+d}。"
            )

    lines.extend(["", "## 3. 谁在推动变化"])
    if top_people:
        for row in top_people[:8]:
            lines.append(
                f"- {row.get('salesperson_name', '')}：AI 提及 {row.get('ai_mentions', 0)} 条，医生反馈 {row.get('doctor_feedback_mentions', 0)} 条，潜在机会 {row.get('opportunity_mentions', 0)} 条，涉及 {row.get('active_weeks', 0)} 个活跃周。"
            )
    else:
        lines.append("- 当前仍缺少稳定的个人销售识别结果，暂以组织对象兜底。")

    lines.extend(["", "## 4. 一线反馈与机会", "### 医生反馈样例"])
    if feedback_examples:
        for item in feedback_examples[:5]:
            lines.append(
                f"- {item.get('salesperson_name', '未识别')}｜{item.get('business_line', '待判断')}：{trim_text(str(item.get('source_text', '')), 88)}"
            )
    else:
        lines.append("- 当前没有足够稳定的医生反馈样例。")

    lines.extend(["", "### 产品机会样例"])
    if opportunity_examples:
        for item in opportunity_examples[:5]:
            lines.append(
                f"- {item.get('salesperson_name', '未识别')}｜{item.get('business_line', '待判断')}：{trim_text(str(item.get('source_text', '')), 88)}"
            )
    else:
        lines.append("- 当前没有足够稳定的产品机会样例。")

    lines.extend(["", "## 5. 重点结论卡"])
    for card in insight_cards[:6]:
        lines.append(
            f"- {card.get('title', '')}：{card.get('summary', '')}（证据 {card.get('evidence_count', 0)} 条，待复核 {card.get('open_review_count', 0)} 条）"
        )

    lines.extend(["", "## 6. 复核优先级"])
    if review_reasons:
        for reason, count in review_reasons:
            lines.append(f"- {reason}：{count} 条")
    if priority_tasks:
        lines.append("- 优先复核样例：")
        for task in priority_tasks[:5]:
            lines.append(
                f"  - {task.get('salesperson_name', '未识别')}｜{task.get('review_reason', '')}｜{trim_text(str(task.get('source_text', '')), 80)}"
            )

    lines.extend(
        [
            "",
            "## 7. 追溯与验收入口",
            "- 趋势和结论必须能回到销售对象与证据片段。",
            "- 证据必须能回到原始文件路径。",
            "- 待复核项优先处理会直接影响趋势判断和结论质量。",
        ]
    )
    return "\n".join(lines)


def build_workbench_html(
    snapshot: Dict[str, object],
    sales_rollup: List[Dict[str, object]],
    insight_cards: List[Dict[str, object]],
    review_tasks: List[Dict[str, object]],
) -> str:
    """生成更偏业务决策视角的 v1.5 工作台页面。"""
    yoy = snapshot.get("yoy_summary", {})
    breadth = snapshot.get("breadth_depth_summary", {})
    top_people = list(snapshot.get("top_people_latest", []))
    top_groups = list(snapshot.get("top_groups_latest", []))
    feedback_examples = list(snapshot.get("feedback_examples", []))
    opportunity_examples = list(snapshot.get("opportunity_examples", []))
    review_reasons = list(snapshot.get("top_review_reasons", []))
    priority_tasks = list(snapshot.get("priority_review_tasks", []))
    actor_yoy = list(snapshot.get("actor_yoy", []))
    line_yoy = list(snapshot.get("business_line_yoy", []))
    trend_html = (
        f"<p>同比：{yoy.get('year_a')} 年同期 vs {yoy.get('year_b')} 年同期，AI 证据 {yoy.get('mentions_a', 0)} → {yoy.get('mentions_b', 0)}，销售对象 {yoy.get('sales_a', 0)} → {yoy.get('sales_b', 0)}，单个活跃销售平均提及强度 {breadth.get('avg_mentions_per_sales_a', 0)} → {breadth.get('avg_mentions_per_sales_b', 0)}。</p>"
        if yoy.get("has_compare")
        else "<p>当前缺少可比年度数据，暂无法展示稳定同比趋势。</p>"
    )
    headline = (
        f"{yoy.get('year_b')} 年同期 AI 证据和涉及销售对象都高于 {yoy.get('year_a')} 年，当前判断：{breadth.get('judgement', '趋势已出现，但仍需继续复核')}。"
        if yoy.get("has_compare")
        else "当前已能稳定沉淀 AI 证据，但仍缺少足够的同比窗口。"
    )
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>AI 一线情报工作台 - v1.5</title>
  <style>
    :root {{
      --bg:#f7f2e8;
      --ink:#18222f;
      --muted:#586576;
      --panel:#fffdfa;
      --line:#d7d0c6;
      --brand:#b44a2e;
      --brand2:#295b80;
      --good:#1d7a43;
    }}
    *{{box-sizing:border-box}}
    body{{margin:0;font-family:"PingFang SC","Noto Serif SC","Helvetica Neue",sans-serif;background:linear-gradient(180deg,#f6efe4,#fbf8f1);color:var(--ink)}}
    .wrap{{max-width:1360px;margin:0 auto;padding:24px}}
    .hero{{background:linear-gradient(120deg,#25170f,#6f2f1f 55%,#9d6a2d);color:#fff;border-radius:20px;padding:26px 28px;box-shadow:0 16px 40px rgba(37,23,15,.18)}}
    .hero h1{{margin:0;font-size:34px}}
    .hero p{{margin:8px 0 0;color:#f9dfc4}}
    .grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:18px}}
    .card,.section{{background:var(--panel);border:1px solid var(--line);border-radius:16px;box-shadow:0 4px 14px rgba(37,23,15,.05)}}
    .card{{padding:16px}}
    .k{{font-size:13px;color:var(--muted)}}
    .v{{font-size:30px;font-weight:700;margin-top:6px}}
    .section{{padding:18px;margin-top:16px}}
    .section h2{{margin:0 0 10px;font-size:19px}}
    .two{{display:grid;grid-template-columns:1.2fr .8fr;gap:14px}}
    .list{{display:grid;gap:10px}}
    .item{{padding:12px;border:1px solid var(--line);border-radius:12px;background:#fff}}
    .title{{font-weight:700}}
    .meta{{font-size:12px;color:var(--muted);margin-top:4px}}
    table{{width:100%;border-collapse:collapse;font-size:13px}}
    th,td{{padding:10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}}
    th{{background:#faf4eb}}
    .table-wrap{{max-height:360px;overflow:auto;border:1px solid var(--line);border-radius:12px}}
    .pill{{display:inline-block;padding:2px 8px;border-radius:999px;background:#f3e5d8;color:#7a351f;font-size:12px;margin-right:6px}}
    .ok{{color:var(--good)}}
    @media (max-width: 980px) {{
      .grid{{grid-template-columns:repeat(2,minmax(0,1fr))}}
      .two{{grid-template-columns:1fr}}
    }}
    @media (max-width: 640px) {{
      .grid{{grid-template-columns:1fr}}
      .wrap{{padding:14px}}
      .hero h1{{font-size:26px}}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <h1>AI 一线情报工作台（v1.5）</h1>
      <p>{html.escape(headline)}</p>
    </div>
    <div class="grid">
      {_metric_card("AI证据", snapshot.get("total_ai_mentions", 0), "当前可追溯的 AI 相关证据总量")}
      {_metric_card("活跃销售", snapshot.get("active_sales_count", 0), "在当前统计窗口内出现 AI 信号的销售对象数")}
      {_metric_card("个人销售", snapshot.get("active_person_count", 0), "已从正文中识别出的个人销售对象数")}
      {_metric_card("待复核", snapshot.get("open_review_tasks", 0), "仍需人工确认的任务总量")}
      {_metric_card("结论卡", snapshot.get("insight_card_count", 0), "当前已沉淀的结论卡数量")}
    </div>
    <div class="section">
      <h2>核心判断</h2>
      <div class="list">
        <div class="item">{html.escape(headline)}</div>
        <div class="item">当前待复核任务 {html.escape(str(snapshot.get("open_review_tasks", 0)))} 条，说明趋势已经能看，但口径稳定性仍受复核积压影响。</div>
      </div>
    </div>
    <div class="section">
      <h2>趋势中心</h2>
      {trend_html}
      <p class="meta">当前月份：{html.escape(str(snapshot.get("latest_year_month", "")))} ｜ 个人销售对象：{html.escape(str(snapshot.get("active_person_count", 0)))} ｜ 组织对象：{html.escape(str(snapshot.get("active_group_count", 0)))}</p>
      <div class="two">
        <div>
          <h3>主体结构变化</h3>
          <div class="table-wrap">
            <table>
              <thead><tr><th>主体</th><th>{html.escape(str(yoy.get("year_a", "")))}</th><th>{html.escape(str(yoy.get("year_b", "")))}</th><th>变化</th></tr></thead>
              <tbody>{_compare_rows(actor_yoy)}</tbody>
            </table>
          </div>
        </div>
        <div>
          <h3>业务线变化</h3>
          <div class="table-wrap">
            <table>
              <thead><tr><th>业务线</th><th>{html.escape(str(yoy.get("year_a", "")))}</th><th>{html.escape(str(yoy.get("year_b", "")))}</th><th>变化</th></tr></thead>
              <tbody>{_compare_rows(line_yoy)}</tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
    <div class="two">
      <div class="section">
        <h2>销售分析中心（个人优先）</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th>销售对象</th><th>AI提及</th><th>医生反馈</th><th>潜在机会</th><th>活跃周</th><th>待复核</th></tr>
            </thead>
            <tbody>
              {_sales_rows(top_people, empty_text="暂未识别出足够的个人销售对象")}
            </tbody>
          </table>
        </div>
        <p class="meta">这里只展示当前同比年份内的个人销售聚合结果，不再按单月切碎展示。</p>
      </div>
      <div class="section">
        <h2>复核工作台</h2>
        <p><span class="pill">开放任务 {html.escape(str(snapshot.get("open_review_tasks", 0)))}</span><span class="pill">已复核 {html.escape(str(snapshot.get("reviewed_tasks", 0)))}</span></p>
        <div class="list">
          {_review_reason_items(review_reasons)}
          {_review_items(priority_tasks[:6])}
        </div>
      </div>
    </div>
    <div class="section">
      <h2>一线反馈与机会</h2>
      <div class="two">
        <div>
          <h3>医生反馈样例</h3>
          <div class="list">{_example_items(feedback_examples, "暂无稳定医生反馈样例")}</div>
        </div>
        <div>
          <h3>产品机会样例</h3>
          <div class="list">{_example_items(opportunity_examples, "暂无稳定产品机会样例")}</div>
        </div>
      </div>
    </div>
    <div class="section">
      <h2>结论中心</h2>
      <div class="list">
        {_insight_items(insight_cards[:8])}
      </div>
    </div>
    <div class="section">
      <h2>组织与区域视角（兜底）</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr><th>组织对象</th><th>AI提及</th><th>医生反馈</th><th>潜在机会</th><th>活跃周</th><th>待复核</th></tr>
          </thead>
          <tbody>
            {_sales_rows(top_groups, empty_text="暂无组织级对象")}
          </tbody>
        </table>
      </div>
    </div>
  </div>
</body>
</html>
"""


def trim_text(text: str, max_len: int) -> str:
    """截断过长文本，避免报表过宽。"""
    return text if len(text) <= max_len else text[: max_len - 1] + "…"


def _metric_card(title: str, value: object, desc: str) -> str:
    return (
        "<div class='card'>"
        f"<div class='k'>{html.escape(str(title))}</div>"
        f"<div class='v'>{html.escape(str(value))}</div>"
        f"<div class='meta'>{html.escape(str(desc))}</div>"
        "</div>"
    )


def _sales_rows(rows: List[Dict[str, object]], empty_text: str) -> str:
    if not rows:
        return f"<tr><td colspan='6'>{html.escape(empty_text)}</td></tr>"
    html_rows: List[str] = []
    for row in rows:
        html_rows.append(
            "<tr>"
            f"<td>{html.escape(str(row.get('salesperson_name', '')))}</td>"
            f"<td>{html.escape(str(row.get('ai_mentions', 0)))}</td>"
            f"<td>{html.escape(str(row.get('doctor_feedback_mentions', 0)))}</td>"
            f"<td>{html.escape(str(row.get('opportunity_mentions', 0)))}</td>"
            f"<td>{html.escape(str(row.get('active_weeks', 0)))}</td>"
            f"<td>{html.escape(str(row.get('review_open_count', 0)))}</td>"
            "</tr>"
        )
    return "".join(html_rows)


def _review_items(rows: List[Dict[str, object]]) -> str:
    if not rows:
        return "<div class='item'>暂无待复核任务</div>"
    items: List[str] = []
    for row in rows:
        items.append(
            "<div class='item'>"
            f"<div class='title'>{html.escape(str(row.get('salesperson_name', '未识别对象')))}</div>"
            f"<div class='meta'>{html.escape(str(row.get('review_reason', '')))}</div>"
            f"<div>{html.escape(trim_text(str(row.get('source_text', '')), 120))}</div>"
            "</div>"
        )
    return "".join(items)


def _review_reason_items(rows: List[object]) -> str:
    if not rows:
        return ""
    items: List[str] = []
    for reason, count in rows[:4]:
        items.append(
            "<div class='item'>"
            f"<div class='title'>{html.escape(str(reason))}</div>"
            f"<div class='meta'>待复核 {html.escape(str(count))} 条</div>"
            "</div>"
        )
    return "".join(items)


def _insight_items(rows: List[Dict[str, object]]) -> str:
    if not rows:
        return "<div class='item'>暂无结论卡</div>"
    items: List[str] = []
    for row in rows:
        owner_refs = "、".join(str(item) for item in row.get("owner_refs", []))
        evidence_refs = row.get("evidence_refs", [])
        evidence_preview = ""
        if evidence_refs:
            evidence_preview = html.escape(trim_text(str(evidence_refs[0].get("source_text", "")), 100))
        items.append(
            "<div class='item'>"
            f"<div class='title'>{html.escape(str(row.get('title', '')))}</div>"
            f"<div>{html.escape(str(row.get('summary', '')))}</div>"
            f"<div class='meta'>涉及对象：{html.escape(owner_refs)} ｜ 置信度：{html.escape(str(row.get('confidence', '')))} ｜ 证据 {html.escape(str(row.get('evidence_count', 0)))} 条 ｜ 待复核 {html.escape(str(row.get('open_review_count', 0)))} 条</div>"
            f"<div class='meta'>代表证据：{evidence_preview}</div>"
            "</div>"
        )
    return "".join(items)


def _compare_rows(rows: List[Dict[str, object]]) -> str:
    if not rows:
        return "<tr><td colspan='4'>暂无对比数据</td></tr>"
    html_rows: List[str] = []
    for row in rows[:6]:
        delta = int(row.get("delta", 0))
        html_rows.append(
            "<tr>"
            f"<td>{html.escape(str(row.get('label', '')))}</td>"
            f"<td>{html.escape(str(row.get('year_a', 0)))}</td>"
            f"<td>{html.escape(str(row.get('year_b', 0)))}</td>"
            f"<td>{html.escape(f'{delta:+d}')}</td>"
            "</tr>"
        )
    return "".join(html_rows)


def _example_items(rows: List[Dict[str, object]], empty_text: str) -> str:
    if not rows:
        return f"<div class='item'>{html.escape(empty_text)}</div>"
    items: List[str] = []
    for row in rows[:5]:
        items.append(
            "<div class='item'>"
            f"<div class='title'>{html.escape(str(row.get('salesperson_name', '未识别对象')))} ｜ {html.escape(str(row.get('business_line', '待判断')))}</div>"
            f"<div>{html.escape(trim_text(str(row.get('source_text', '')), 120))}</div>"
            "</div>"
        )
    return "".join(items)
