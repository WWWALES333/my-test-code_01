from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence


def write_jsonl(path: Path, rows: Iterable[Dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        for row in rows:
            file.write(json.dumps(row, ensure_ascii=False) + "\n")


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_markdown(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_csv_template(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "sample_id,task_id,report_id,segment_id,salesperson_id,review_reason_code,current_labels,reviewed_fields,final_labels,is_pass,review_comment,reviewer,reviewed_at,error_reason_primary,review_necessity,actionability,action_bucket,learning_note,need_rule_update,need_prompt_update,need_annotation_update\n",
        encoding="utf-8",
    )


def write_web_pages(path: Path, pages: Dict[str, str]) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for name, content in pages.items():
        (path / name).write_text(content, encoding="utf-8")


def build_summary_markdown(
    snapshot: Dict[str, object],
    trend_explanations: Sequence[Dict[str, object]],
    salesperson_profiles: Sequence[Dict[str, object]],
    insight_cards: Sequence[Dict[str, object]],
    review_tasks: Sequence[Dict[str, object]],
    region_rollups: Sequence[Dict[str, object]],
) -> str:
    lines: List[str] = ["# AI 一线情报工作台摘要", ""]
    lines.extend(["## 本期核心判断"])
    for item in snapshot.get("headline_judgements", [])[:5]:
        lines.append(f"- {item}")
    if not snapshot.get("headline_judgements"):
        lines.append("- 当前还缺少足够稳定的趋势判断。")

    lines.extend(
        [
            "",
            "## 关键指标",
            f"- 最新统计周期：{snapshot.get('latest_year_month', '')}",
            f"- 时间范围：{snapshot.get('data_range', {}).get('start_month', '')} 至 {snapshot.get('data_range', {}).get('end_month', '')}",
            f"- 默认周度窗口：{snapshot.get('latest_year_week', '')}",
            f"- AI 证据：{snapshot.get('total_ai_mentions', 0)} 条",
            f"- 活跃销售对象：{snapshot.get('active_sales_count', 0)} 个",
            f"- 当前在岗销售：{snapshot.get('roster_active_count', 0)} 个",
            f"- 待复核任务：{snapshot.get('open_review_tasks', 0)} 条",
            f"- 已沉淀洞察卡：{snapshot.get('insight_card_count', 0)} 张",
        ]
    )

    lines.extend(["", "## 趋势解释"])
    for item in trend_explanations[:4]:
        lines.append(f"- {item.get('explanation', '')}（可信度：{item.get('confidence_level', '')}）")

    lines.extend(["", "## 销售画像观察"])
    high = [row for row in salesperson_profiles if str(row.get("segment", "")) == "高频使用者"][:5]
    silent = [row for row in salesperson_profiles if str(row.get("segment", "")) == "长期未提及者"][:5]
    if high:
        lines.append("- 高频使用者：")
        for row in high:
            lines.append(
                f"  - {row.get('display_name', '')}｜{row.get('battle_zone_name', '')}/{row.get('region_name', '')}｜提及 {row.get('total_mentions', 0)} 次｜医生反馈 {row.get('doctor_feedback_mentions', 0)} 次"
            )
    if silent:
        lines.append("- 长期未提及者样例：")
        for row in silent:
            lines.append(f"  - {row.get('display_name', '')}｜{row.get('battle_zone_name', '')}/{row.get('region_name', '')}")

    lines.extend(["", "## 区域构成"])
    for row in region_rollups[:5]:
        top = row.get("top_salespeople", [])
        top_names = "、".join(str(item.get("display_name", "")) for item in top[:3])
        lines.append(
            f"- {row.get('battle_zone_name', '')}/{row.get('region_name', '')}：活跃销售 {row.get('active_sales_count', 0)} 个，沉默销售 {row.get('silent_sales_count', 0)} 个，判断：{row.get('maturity_judgement', '')}。核心支撑：{top_names}"
        )

    lines.extend(["", "## 重点结论"])
    for card in insight_cards[:8]:
        lines.append(
            f"- {card.get('title', '')}：{card.get('judgement', '')} {card.get('why_it_matters', '')} 建议：{card.get('action_recommendation', '')}"
        )

    lines.extend(["", "## 复核优先级"])
    open_tasks = [row for row in review_tasks if str(row.get("task_status", "")) == "open"][:8]
    if open_tasks:
        for row in open_tasks:
            lines.append(
                f"- {row.get('salesperson_name', '未识别对象')}｜{row.get('review_reason_code', '')}｜{trim_text(str(row.get('source_text', '')), 100)}"
            )
    else:
        lines.append("- 当前没有待复核任务。")

    lines.extend(
        [
            "",
            "## 验收入口",
            "- 总览：web/overview.html",
            "- 趋势：web/trends.html",
            "- 销售画像：web/sales.html",
            "- 洞察树：web/insights.html",
            "- 复核：web/review.html（交互写回请使用 `python3 -m src.analysis_v15.webapp --data ...`）",
            "- 证据：web/evidence.html",
        ]
    )
    return "\n".join(lines)


def build_workbench_pages(
    snapshot: Dict[str, object],
    trend_cube: Sequence[Dict[str, object]],
    trend_explanations: Sequence[Dict[str, object]],
    salesperson_profiles: Sequence[Dict[str, object]],
    region_rollups: Sequence[Dict[str, object]],
    insight_tree: Dict[str, object],
    review_tasks: Sequence[Dict[str, object]],
    evidence_index: Sequence[Dict[str, object]],
    review_learning_summary: Dict[str, object] | None = None,
    review_candidates: Sequence[Dict[str, object]] | None = None,
    review_batch_summaries: Sequence[Dict[str, object]] | None = None,
    interactive_review: bool = False,
) -> Dict[str, str]:
    overview = _render_page(
        "overview",
        "AI 一线情报工作台",
        _overview_body(snapshot),
    )
    trends = _render_page(
        "trends",
        "趋势中心",
        _trends_body(snapshot),
        extra_scripts=_trends_scripts(snapshot, trend_cube, trend_explanations, evidence_index),
    )
    sales = _render_page(
        "sales",
        "销售分析中心",
        _sales_body(snapshot),
        extra_scripts=_sales_scripts(snapshot, salesperson_profiles, region_rollups),
    )
    insights = _render_page(
        "insights",
        "结论中心",
        _insights_body(insight_tree),
    )
    review = _render_page(
        "review",
        "复核工作台",
        _review_body(
            review_tasks,
            interactive_review=interactive_review,
            snapshot=snapshot,
            review_learning_summary=review_learning_summary or {},
            review_candidates=review_candidates or [],
            review_batch_summaries=review_batch_summaries or [],
        ),
        extra_scripts=_review_scripts(review_tasks, review_learning_summary or {}, review_candidates or [], review_batch_summaries or []) if interactive_review else "",
    )
    evidence = _render_page(
        "evidence",
        "证据下钻",
        _evidence_body(snapshot),
        extra_scripts=_evidence_scripts(snapshot, evidence_index),
    )
    pages = {
        "overview.html": overview,
        "AI情报工作台.html": overview,
        "trends.html": trends,
        "sales.html": sales,
        "insights.html": insights,
        "review.html": review,
        "evidence.html": evidence,
    }
    return pages


def trim_text(text: str, max_len: int) -> str:
    return text if len(text) <= max_len else text[: max_len - 1] + "…"


def _render_page(current: str, title: str, body: str, extra_scripts: str = "") -> str:
    nav_items = [
        ("overview", "总览", "overview.html"),
        ("trends", "趋势", "trends.html"),
        ("sales", "销售", "sales.html"),
        ("insights", "结论", "insights.html"),
        ("review", "复核", "review.html"),
        ("evidence", "证据", "evidence.html"),
    ]
    nav = "".join(
        f"<a class='nav-item {'active' if key == current else ''}' href='{href}'>{label}</a>"
        for key, label, href in nav_items
    )
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{html.escape(title)}</title>
  <style>
    :root {{
      --bg:#f6f4ee;
      --panel:#fffdfa;
      --line:#ddd4c8;
      --ink:#16212b;
      --muted:#5d6772;
      --accent:#a64e2c;
      --accent-2:#275574;
      --warn:#8c5a14;
      --good:#236c46;
    }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; font-family:"PingFang SC","Helvetica Neue",sans-serif; background:var(--bg); color:var(--ink); }}
    .shell {{ max-width:1400px; margin:0 auto; padding:20px; }}
    .topbar {{ display:flex; align-items:center; justify-content:space-between; gap:16px; margin-bottom:16px; }}
    .title h1 {{ margin:0; font-size:30px; }}
    .title p {{ margin:6px 0 0; color:var(--muted); }}
    .nav {{ display:flex; flex-wrap:wrap; gap:8px; }}
    .nav-item {{ text-decoration:none; color:var(--ink); padding:8px 12px; border:1px solid var(--line); border-radius:999px; background:#fff; }}
    .nav-item.active {{ background:var(--accent); color:#fff; border-color:var(--accent); }}
    .hero {{ background:linear-gradient(125deg,#23160f,#6a2a18,#a26a25); color:#fff; border-radius:18px; padding:22px; margin-bottom:16px; }}
    .hero p {{ margin:8px 0 0; color:#f4dfc3; }}
    .grid {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; }}
    .grid.two {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
    .grid.three {{ grid-template-columns:repeat(3,minmax(0,1fr)); }}
    .panel {{ background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:16px; box-shadow:0 4px 12px rgba(0,0,0,.04); }}
    .panel h2, .panel h3 {{ margin:0 0 10px; }}
    .metric-label {{ font-size:13px; color:var(--muted); }}
    .metric-value {{ font-size:30px; font-weight:700; margin-top:6px; }}
    .list {{ display:grid; gap:10px; }}
    .item {{ border:1px solid var(--line); border-radius:12px; padding:12px; background:#fff; }}
    .item-title {{ font-weight:700; margin-bottom:6px; }}
    .item-meta {{ font-size:12px; color:var(--muted); margin-top:6px; }}
    .pill {{ display:inline-block; padding:2px 8px; border-radius:999px; font-size:12px; background:#f2e8dc; color:#7a351f; margin-right:6px; }}
    .pill.good {{ background:#e3f0e8; color:var(--good); }}
    .pill.warn {{ background:#f5ead7; color:var(--warn); }}
    table {{ width:100%; border-collapse:collapse; font-size:13px; }}
    th, td {{ text-align:left; padding:10px; border-bottom:1px solid var(--line); vertical-align:top; }}
    th {{ background:#f8f2ea; }}
    .table-wrap {{ overflow:auto; border:1px solid var(--line); border-radius:12px; background:#fff; }}
    .muted {{ color:var(--muted); }}
    .code {{ font-family:ui-monospace,SFMono-Regular,Menlo,monospace; }}
    textarea, select, input {{ width:100%; padding:8px 10px; border-radius:10px; border:1px solid var(--line); font:inherit; }}
    input[type="checkbox"] {{ width:auto; transform:translateY(1px); }}
    button {{ border:0; border-radius:10px; padding:10px 14px; background:var(--accent-2); color:#fff; font:inherit; cursor:pointer; }}
    button.secondary {{ background:#fff; color:var(--ink); border:1px solid var(--line); }}
    .toolbar {{ display:flex; flex-wrap:wrap; gap:8px; align-items:center; }}
    .field-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }}
    .checkbox-grid {{ display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:8px; }}
    .checkbox-item {{ display:flex; gap:8px; align-items:center; padding:8px 10px; border:1px solid var(--line); border-radius:10px; background:#fff; }}
    .review-shell {{ display:grid; grid-template-columns:1.15fr .95fr; gap:12px; }}
    .review-card-head {{ display:flex; gap:10px; align-items:flex-start; justify-content:space-between; }}
    .status-chip {{ display:inline-block; padding:2px 8px; border-radius:999px; font-size:12px; background:#ece4d7; color:#694127; margin-right:6px; }}
    .status-chip.good {{ background:#e3f0e8; color:var(--good); }}
    .status-chip.warn {{ background:#f5ead7; color:var(--warn); }}
    .status-chip.muted {{ background:#eef1f5; color:#51606e; }}
    details.panel summary {{ cursor:pointer; font-weight:700; list-style:none; }}
    details.panel summary::-webkit-details-marker {{ display:none; }}
    .stepbar {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; margin-top:12px; }}
    .step {{ border:1px solid var(--line); border-radius:12px; padding:10px; background:#fff; }}
    .step strong {{ display:block; margin-bottom:4px; }}
    .hint {{ font-size:12px; color:var(--muted); }}
    .kv-list {{ display:grid; gap:8px; }}
    .kv-item {{ display:flex; justify-content:space-between; gap:12px; border-bottom:1px dashed var(--line); padding-bottom:6px; }}
    .kv-item:last-child {{ border-bottom:none; padding-bottom:0; }}
    .kv-label {{ color:var(--muted); font-size:12px; }}
    .kv-value {{ text-align:right; }}
    @media (max-width: 1100px) {{
      .grid, .grid.two, .grid.three {{ grid-template-columns:1fr; }}
      .field-grid, .checkbox-grid, .review-shell {{ grid-template-columns:1fr; }}
      .stepbar {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <div class="topbar">
      <div class="title">
        <h1>{html.escape(title)}</h1>
        <p>从 AI 看板切换到 AI 一线情报工作台，优先承载判断、下钻和复核闭环。</p>
      </div>
      <div class="nav">{nav}</div>
    </div>
    <div class="hero">
      <div class="item-title">统一入口</div>
      <p>先看判断，再看趋势、销售、结论和复核。所有结论都必须能回到证据和原文。</p>
    </div>
    {body}
  </div>
  {extra_scripts}
</body>
</html>"""


def _overview_body(snapshot: Dict[str, object]) -> str:
    judgements = list(snapshot.get("headline_judgements", []))[:5]
    stable_insights = list(snapshot.get("stable_insights", []))[:5]
    risk_insights = list(snapshot.get("risk_insights", []))[:5]
    opportunity_items = [row for row in stable_insights if bool(row.get("is_actionable", False))][:5]
    return f"""
    <div class="panel" style="margin-bottom:16px;">
      <h2>统计口径</h2>
      <div class="item">
        <div class="item-title">时间范围</div>
        <div>{html.escape(str(snapshot.get('data_range', {}).get('start_month', '')))} 至 {html.escape(str(snapshot.get('data_range', {}).get('end_month', '')))}</div>
        <div class="item-meta">{html.escape(str(snapshot.get('time_scope_note', '')))}</div>
      </div>
    </div>
    <div class="grid">
      {_metric_card("AI 证据", snapshot.get("total_ai_mentions", 0), "当前可追溯 AI 相关证据总量")}
      {_metric_card("活跃销售", snapshot.get("active_sales_count", 0), "当前窗口内出现 AI 信号的销售对象")}
      {_metric_card("在岗销售", snapshot.get("roster_active_count", 0), "来自花名册的当前有效销售主名单")}
      {_metric_card("待复核任务", snapshot.get("open_review_tasks", 0), "仍需人工确认的任务")}
    </div>
    <div class="grid two" style="margin-top:16px;">
      <div class="panel">
        <h2>本期核心判断</h2>
        <div class="list">{''.join(_bullet_item(text) for text in judgements) or _empty_item("当前还缺少足够稳定的判断。")}</div>
      </div>
      <div class="panel">
        <h2>趋势一句话</h2>
        <div class="list">{_bullet_item(judgements[0] if judgements else "当前还没有稳定趋势结论。")}</div>
      </div>
    </div>
    <div class="grid two" style="margin-top:16px;">
      <div class="panel">
        <h2>重点机会</h2>
        <div class="list">{''.join(_insight_brief_item(row) for row in opportunity_items) or _empty_item("当前还没有足够稳定的可执行机会。")}</div>
      </div>
      <div class="panel">
        <h2>重点风险</h2>
        <div class="list">{''.join(_insight_brief_item(row) for row in risk_insights) or _empty_item("当前没有高风险结论卡。")}</div>
      </div>
    </div>
    """


def _trends_body(snapshot: Dict[str, object]) -> str:
    return f"""
    <div class="panel">
      <h2>统计周期与对比口径</h2>
      <div class="grid three">
        <div><label>主周期（月）</label><select id="trend-primary-month"></select></div>
        <div><label>对比周期（月）</label><select id="trend-compare-month"></select></div>
        <div><label>周度下钻</label><select id="trend-week"></select></div>
      </div>
      <div class="item" style="margin-top:12px;">
        <div class="item-title">当前口径</div>
        <div id="trend-scope-note">{html.escape(str(snapshot.get('time_scope_note', '')))}</div>
      </div>
    </div>
    <div class="grid two" style="margin-top:16px;">
      <div class="panel">
        <h2>趋势解释</h2>
        <div id="trend-explanations" class="list"></div>
      </div>
      <div class="panel">
        <h2>关键指标</h2>
        <div id="trend-metrics" class="grid"></div>
      </div>
    </div>
    <div class="grid two" style="margin-top:16px;">
      <div class="panel">
        <h2>月度总趋势</h2>
        <div class="table-wrap"><table><thead><tr><th>周期</th><th>AI 证据</th><th>活跃销售</th><th>渗透率</th><th>区域覆盖率</th><th>待复核率</th></tr></thead><tbody id="trend-monthly-table"></tbody></table></div>
      </div>
      <div class="panel">
        <h2>选中月份的周度明细</h2>
        <div class="table-wrap"><table><thead><tr><th>周次</th><th>AI 证据</th><th>活跃销售</th><th>待复核率</th></tr></thead><tbody id="trend-weekly-table"></tbody></table></div>
      </div>
    </div>
    <div class="grid two" style="margin-top:16px;">
      <div class="panel">
        <h2>结构趋势</h2>
        <div id="trend-structure"></div>
      </div>
      <div class="panel">
        <h2>周度证据预览</h2>
        <div id="trend-week-evidence" class="list"></div>
      </div>
    </div>
    """


def _sales_body(snapshot: Dict[str, object]) -> str:
    return f"""
    <div class="panel">
      <h2>筛选条件</h2>
      <div class="grid three">
        <div><label>销售分层</label><select id="sales-segment-filter"></select></div>
        <div><label>主体标签</label><select id="sales-actor-filter"></select></div>
        <div><label>业务线标签</label><select id="sales-line-filter"></select></div>
      </div>
      <div class="grid two" style="margin-top:12px;">
        <div><label>时间窗口（月）</label><select id="sales-month-filter"></select></div>
        <div><label>销售搜索</label><input id="sales-search" placeholder="输入花名搜索" /></div>
      </div>
    </div>
    <div id="sales-summary-cards" class="grid" style="margin-top:16px;"></div>
    <div class="grid two">
      <div class="panel">
        <h2>销售分层</h2>
        <div id="sales-segment-panels" class="grid two"></div>
      </div>
      <div class="panel">
        <h2>区域背后的销售构成</h2>
        <div class="table-wrap">
          <table>
            <thead><tr><th>战区 / 区域</th><th>活跃销售</th><th>沉默销售</th><th>主导度</th><th>判断</th><th>核心支撑</th></tr></thead>
            <tbody id="sales-region-table"></tbody>
          </table>
        </div>
      </div>
    </div>
    <div class="panel" style="margin-top:16px;">
      <h2>销售画像</h2>
      <div class="item" style="margin-bottom:12px;">
        <div class="item-title">统计口径</div>
        <div>{html.escape(str(snapshot.get('time_scope_note', '')))}</div>
      </div>
      <div id="sales-profile-grid" class="grid three"></div>
    </div>
    """


def _insights_body(insight_tree: Dict[str, object]) -> str:
    groups = []
    for line_group in insight_tree.get("business_lines", []):
        topic_html = []
        for topic_group in line_group.get("topics", []):
            cards = topic_group.get("cards", [])
            topic_html.append(
                "<div class='panel'>"
                f"<h3>{html.escape(str(topic_group.get('topic', '')))}</h3>"
                f"<div class='list'>{''.join(_insight_card_item(card) for card in cards) or _empty_item('暂无结论卡')}</div>"
                "</div>"
            )
        groups.append(
            "<div class='panel'>"
            f"<h2>{html.escape(str(line_group.get('business_line', '')))}</h2>"
            f"<div class='grid two'>{''.join(topic_html)}</div>"
            "</div>"
        )
    return "".join(groups) if groups else _empty_item("暂无洞察树。")


def _review_body(
    review_tasks: Sequence[Dict[str, object]],
    interactive_review: bool,
    snapshot: Dict[str, object],
    review_learning_summary: Dict[str, object],
    review_candidates: Sequence[Dict[str, object]],
    review_batch_summaries: Sequence[Dict[str, object]],
) -> str:
    open_tasks = [row for row in review_tasks if str(row.get("task_status", "")) == "open"]
    reviewed_tasks = [row for row in review_tasks if str(row.get("task_status", "")) == "reviewed"]
    note = (
        "<div class='item'>当前页面默认是基础批阅模式。先看原文，再改核心标签，再补一个错因和行动价值，最后提交下一条。</div>"
        if interactive_review
        else "<div class='item'>静态页面只展示任务。若要直接写回，请运行 <span class='code'>python3 -m src.analysis_v15.webapp --data data/output/insights/v1.5</span>。</div>"
    )
    editor = _review_editor(review_tasks) if interactive_review else ""
    candidate_items = "".join(_candidate_item(row) for row in list(review_candidates)[:8]) or _empty_item("当前还没有学习候选。")
    return f"""
    <div class="grid">
      {_metric_card("待复核任务", len(open_tasks), "当前仍需逐条批阅的任务")}
      {_metric_card("已复核任务", len(reviewed_tasks), "当前已写回的人工复核记录")}
      {_metric_card("默认批次", "20 条", "复核按固定小批次推进，做完一批就可以停下来优化")}
      {_metric_card("学习候选", review_learning_summary.get("candidate_count", 0), "已从复核记录中汇总出的规则 / Prompt / 标注候选")}
      {_metric_card("高优先候选", review_learning_summary.get("high_priority_candidate_count", 0), "累计达到阈值的候选模式")}
    </div>
    <div class="grid two" style="margin-top:16px;">
      <div class="panel">
        <h2>你现在要做什么</h2>
        <div class="list">{note}</div>
        <div class="stepbar">
          <div class="step"><strong>1. 看原文</strong><span class="hint">先判断系统是不是理解对了这段话。</span></div>
          <div class="step"><strong>2. 改核心标签</strong><span class="hint">只改你确定错的字段，不必每项都动。</span></div>
          <div class="step"><strong>3. 看有没有行动价值</strong><span class="hint">只标这条值不值得进入动作池。</span></div>
          <div class="step"><strong>4. 提交下一条</strong><span class="hint">系统会根据你的修改自动归因并进入下一张。</span></div>
        </div>
      </div>
      <div class="panel">
        <h2>字段释义</h2>
        <div class="list">
          <div class="item"><span class="code">decision_status</span>：<strong>confirmed</strong> 已确认；<strong>uncertain</strong> 仍不确定；<strong>pending_human_review</strong> 继续挂起复核。</div>
          <div class="item"><span class="code">ai_scope</span>：<strong>product_ai</strong> 我方产品 AI；<strong>market_trend</strong> 市场趋势；<strong>competitor_ai</strong> 竞品 AI；<strong>general_ai</strong> 泛 AI 话题。</div>
          <div class="item"><span class="code">actionability</span>：这条值不值得形成动作；<span class="code">error_reason_primary</span> 由系统根据你的复核自动归因。</div>
        </div>
      </div>
    </div>
    <div class="panel" style="margin-top:16px;">
      <h2>批次策略</h2>
      <div class="list">
        <div class="item">系统会把任务按优先级自动分成 <strong>每批 20 条</strong>。你可以只做当前批次，做完就停，不需要先把全部任务复核完。</div>
        <div class="item">做完一批后，优先看本批次的错因分布、候选池和高频修改模式，再决定是否先优化规则或 Prompt。</div>
      </div>
    </div>
    <details class="panel" style="margin-top:16px;">
      <summary>高级筛选与任务状态</summary>
      <div class="hint" style="margin:8px 0 12px;">首轮批阅时可以先忽略这里，默认直接从第一条 open 任务开始。</div>
      <div class="field-grid">
        <div><label>复核批次</label><select id="filter-batch-id"></select></div>
        <div><label>任务状态</label><select id="filter-task-status"></select></div>
        <div><label>原因码</label><select id="filter-review-reason"></select></div>
        <div><label>年份</label><select id="filter-year"></select></div>
        <div><label>月份</label><select id="filter-month"></select></div>
        <div><label>周次</label><select id="filter-week"></select></div>
        <div><label>业务线</label><select id="filter-business-line"></select></div>
        <div><label>主体类型</label><select id="filter-actor-primary"></select></div>
        <div><label>销售搜索</label><input id="filter-salesperson" placeholder="按花名或对象搜索" /></div>
      </div>
      <div class="item" style="margin-top:12px;">
        <div class="item-title">当前筛选摘要</div>
        <div id="review-filter-summary" class="muted">加载中…</div>
      </div>
      <div class="item" style="margin-top:12px;">
        <div class="item-title">原因码分布</div>
        <div id="review-reason-distribution" class="toolbar"></div>
      </div>
    </details>
    {editor}
    <details class="panel" style="margin-top:16px;">
      <summary>学习候选与最近历史</summary>
      <div class="hint" style="margin:8px 0 12px;">这里是系统内部学习视角。首轮批阅时可以先不看。</div>
      <div class="grid two">
        <div class="panel">
          <h2>当前任务最近提交记录</h2>
          <div id="review-current-history" class="list">{_empty_item("当前任务暂无复核记录。")}</div>
        </div>
        <div class="panel">
          <h2>学习候选池</h2>
          <div class="list" id="review-candidate-list">{candidate_items}</div>
        </div>
      </div>
      <div class="panel" style="margin-top:16px;">
        <h2>最近已复核样例</h2>
        <div class="list" id="review-reviewed-list">{''.join(_review_task_item(row, reviewed=True) for row in reviewed_tasks[:8]) or _empty_item("当前还没有已复核记录。")}</div>
      </div>
    </details>
    """


def _review_editor(review_tasks: Sequence[Dict[str, object]]) -> str:
    if not review_tasks:
        return ""
    return f"""
    <div class="panel" style="margin-top:16px;">
      <h2>逐条复核卡片</h2>
      <div class="review-shell">
        <div>
          <div class="review-card-head">
            <div>
              <div class="item-title" id="review-card-title">待复核卡片</div>
              <div class="item-meta" id="review-card-meta"></div>
            </div>
            <div class="toolbar">
              <button id="review-prev" type="button" class="secondary">上一张</button>
              <button id="review-next" type="button" class="secondary">下一张</button>
              <button id="review-skip" type="button" class="secondary">跳过，保持 open</button>
              <button id="review-open-evidence" type="button" class="secondary">查看原文证据页</button>
            </div>
          </div>
          <div class="item" style="margin-top:12px;">
            <div class="item-title">原文片段</div>
            <div id="task-source" class="muted"></div>
            <div class="item-meta" id="task-context"></div>
            <div class="item-meta" id="task-file"></div>
          </div>
          <div class="item" style="margin-top:12px;">
            <div class="item-title">当前系统判断</div>
            <div id="task-current-fields" class="code muted"></div>
          </div>
          <div class="item" style="margin-top:12px;">
            <div class="item-title">基本信息</div>
            <div id="task-basic-info" class="muted"></div>
          </div>
        </div>
        <div>
          <div class="item" style="margin-bottom:12px;">
            <div class="item-title">基础复核</div>
            <div class="hint">先只改这些核心项：是否命中 AI、业务线、主体、复核状态、行动价值。错因由系统自动归因。</div>
          </div>
          <div class="field-grid">
            <div><label>是否命中 AI</label><select id="field-is-ai-hit"><option value="true">是</option><option value="false">否</option></select></div>
            <div><label>业务线</label><select id="field-business-line"><option>云诊室</option><option>云管家</option><option>混合</option><option>待判断</option></select></div>
            <div><label>主体类型</label><select id="field-actor-primary"><option value="">现有标签不适用 / 待扩增</option><option>销售自用</option><option>销售对外介绍</option><option>医生反馈</option><option>潜在 AI 机会</option></select></div>
            <div><label>复核状态</label><select id="field-decision-status"><option value="confirmed">confirmed（已确认）</option><option value="uncertain">uncertain（仍不确定）</option><option value="pending_human_review">pending_human_review（继续复核）</option></select></div>
            <div><label>行动价值</label><select id="field-actionability"><option value="">请选择</option><option value="actionable">actionable（可行动）</option><option value="observe">observe（先观察）</option><option value="no_action">no_action（无动作）</option></select></div>
            <div><label>复核人</label><input id="field-reviewer" value="wales" /></div>
          </div>
          <label style="margin-top:12px;display:block;">复核备注</label>
          <textarea id="field-review-comment" rows="4"></textarea>
          <div class="toolbar" style="margin-top:12px;">
            <button id="submit-review-next" type="button">提交并进入下一张</button>
            <button id="submit-review-stay" type="button">仅提交</button>
            <button id="rebuild-workbench" type="button" class="secondary">按当前复核结果重建工作台</button>
            <span class="muted" id="submit-status"></span>
          </div>
          <details class="panel" style="margin-top:12px;">
            <summary>高级设置（可选）</summary>
            <div class="hint" style="margin:8px 0 12px;">这里只保留补充信息。至于这是规则问题、Prompt 问题还是标注口径问题，系统会根据你的复核结果自动归因。</div>
            <div class="field-grid">
              <div><label>范围类型</label><select id="field-ai-scope"><option value="product_ai">product_ai（我方产品 AI）</option><option value="market_trend">market_trend（市场趋势）</option><option value="competitor_ai">competitor_ai（竞品 AI）</option><option value="general_ai">general_ai（泛 AI 话题）</option></select></div>
              <div><label>是否值得进复核</label><select id="field-review-necessity"><option value="should_review">should_review（确实该复核）</option><option value="could_auto_confirm">could_auto_confirm（本可自动确认）</option><option value="could_auto_reject">could_auto_reject（本可自动排除）</option><option value="low_value_noise">low_value_noise（低价值噪音）</option></select></div>
              <div><label>行动归档</label><select id="field-action-bucket"><option value="">请选择</option><option value="product_pool">product_pool（产品池）</option><option value="sales_enablement_pool">sales_enablement_pool（销售赋能池）</option><option value="watchlist">watchlist（观察池）</option><option value="none">none（不归档）</option></select></div>
            </div>
            <label style="margin-top:12px;display:block;">学习备注</label>
            <textarea id="field-learning-note" rows="4"></textarea>
          </details>
        </div>
      </div>
    </div>
    """


def _review_scripts(
    review_tasks: Sequence[Dict[str, object]],
    review_learning_summary: Dict[str, object],
    review_candidates: Sequence[Dict[str, object]],
    review_batch_summaries: Sequence[Dict[str, object]],
) -> str:
    tasks_json = json.dumps(list(review_tasks), ensure_ascii=False)
    learning_summary_json = json.dumps(dict(review_learning_summary), ensure_ascii=False)
    candidates_json = json.dumps(list(review_candidates), ensure_ascii=False)
    batch_summaries_json = json.dumps(list(review_batch_summaries), ensure_ascii=False)
    return f"""
<script>
let reviewTasks = {tasks_json};
let reviewLearningSummary = {learning_summary_json};
let reviewCandidates = {candidates_json};
let reviewBatchSummaries = {batch_summaries_json};
let currentIndex = 0;
const sourceNode = document.getElementById("task-source");
const contextNode = document.getElementById("task-context");
const fileNode = document.getElementById("task-file");
const titleNode = document.getElementById("review-card-title");
const metaNode = document.getElementById("review-card-meta");
const currentFieldsNode = document.getElementById("task-current-fields");
const basicInfoNode = document.getElementById("task-basic-info");
const currentHistoryNode = document.getElementById("review-current-history");
const reviewedListNode = document.getElementById("review-reviewed-list");
const filterSummaryNode = document.getElementById("review-filter-summary");
const reasonDistributionNode = document.getElementById("review-reason-distribution");
const submitStatus = document.getElementById("submit-status");
const filterBatchNode = document.getElementById("filter-batch-id");
const filterStatusNode = document.getElementById("filter-task-status");
const filterReasonNode = document.getElementById("filter-review-reason");
const filterYearNode = document.getElementById("filter-year");
const filterMonthNode = document.getElementById("filter-month");
const filterWeekNode = document.getElementById("filter-week");
const filterBusinessLineNode = document.getElementById("filter-business-line");
const filterActorPrimaryNode = document.getElementById("filter-actor-primary");
const filterSalespersonNode = document.getElementById("filter-salesperson");
function esc(text) {{
  return String(text ?? "").replace(/[&<>"]/g, (c) => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}})[c]);
}}
function labelFor(key, value) {{
  const maps = {{
    is_ai_hit: {{ true: "是", false: "否" }},
    decision_status: {{
      confirmed: "confirmed（已确认）",
      uncertain: "uncertain（仍不确定）",
      pending_human_review: "pending_human_review（继续复核）",
    }},
    ai_scope: {{
      product_ai: "product_ai（我方产品 AI）",
      market_trend: "market_trend（市场趋势）",
      competitor_ai: "competitor_ai（竞品 AI）",
      general_ai: "general_ai（泛 AI 话题）",
    }},
    error_reason_primary: {{
      label_gap: "label_gap（现有标签不适用）",
      actor_boundary: "actor_boundary（主体边界错）",
      business_line_boundary: "business_line_boundary（业务线边界错）",
      ai_scope_boundary: "ai_scope_boundary（范围边界错）",
      low_signal_noise: "low_signal_noise（低信号噪音）",
      context_loss: "context_loss（上下文丢失）",
      parser_or_segmentation_error: "parser_or_segmentation_error（抽取/切分错）",
      model_misread: "model_misread（模型误读）",
      rule_threshold_issue: "rule_threshold_issue（规则阈值问题）",
      other: "other（其他）",
    }},
    review_necessity: {{
      should_review: "should_review（确实该复核）",
      could_auto_confirm: "could_auto_confirm（本可自动确认）",
      could_auto_reject: "could_auto_reject（本可自动排除）",
      low_value_noise: "low_value_noise（低价值噪音）",
    }},
    actionability: {{
      actionable: "actionable（可行动）",
      observe: "observe（先观察）",
      no_action: "no_action（无动作）",
    }},
    action_bucket: {{
      product_pool: "product_pool（产品池）",
      sales_enablement_pool: "sales_enablement_pool（销售赋能池）",
      watchlist: "watchlist（观察池）",
      none: "none（不归档）",
    }},
  }};
  const map = maps[key] || {{}};
  const normalized = key === "is_ai_hit" ? String(Boolean(value)) : String(value ?? "");
  return map[normalized] || String(value ?? "");
}}
function fill(node, values, selected="all") {{
  node.innerHTML = ['<option value="all">全部</option>', ...values.map(value => `<option value="${{esc(value)}}" ${{String(value) === String(selected) ? "selected" : ""}}>${{esc(value)}}</option>`)].join("");
}}
function effectiveFields(task) {{
  return Object.assign({{}}, task.current_fields || {{}}, task.edited_fields || {{}});
}}
function filteredTasks() {{
  const batchId = filterBatchNode.value || "all";
  const status = filterStatusNode.value || "open";
  const reason = filterReasonNode.value || "all";
  const year = filterYearNode.value || "all";
  const month = filterMonthNode.value || "all";
  const week = filterWeekNode.value || "all";
  const businessLine = filterBusinessLineNode.value || "all";
  const actorPrimary = filterActorPrimaryNode.value || "all";
  const salesperson = (filterSalespersonNode.value || "").trim();
  return reviewTasks.filter(item => {{
    const fields = effectiveFields(item);
    if (batchId !== "all" && String(item.batch_id || "") !== batchId) return false;
    if (status !== "all" && item.task_status !== status) return false;
    if (reason !== "all" && String(item.review_reason_code || "") !== reason) return false;
    if (year !== "all" && String(item.year || "") !== year) return false;
    if (month !== "all" && String(item.month || "") !== month) return false;
    if (week !== "all" && String(item.week_of_month || "") !== week) return false;
    if (businessLine !== "all" && String(fields.business_line || "") !== businessLine) return false;
    if (actorPrimary !== "all" && String(fields.actor_primary || "") !== actorPrimary) return false;
    if (salesperson && !String(item.salesperson_name || "").includes(salesperson) && !String(item.owner_hint || "").includes(salesperson)) return false;
    return true;
  }});
}}
function preferredOpenBatchId() {{
  const firstOpen = reviewTasks.find(item => item.task_status === "open");
  return firstOpen ? String(firstOpen.batch_id || "all") : "all";
}}
function taskById(taskId) {{
  return reviewTasks.find(item => item.task_id === taskId) || null;
}}
function currentTask() {{
  const tasks = filteredTasks();
  if (!tasks.length) return null;
  if (currentIndex >= tasks.length) currentIndex = tasks.length - 1;
  if (currentIndex < 0) currentIndex = 0;
  return tasks[currentIndex];
}}
function setDisabled(disabled) {{
  [
    "field-is-ai-hit",
    "field-business-line",
    "field-actor-primary",
    "field-ai-scope",
    "field-decision-status",
    "field-review-necessity",
    "field-actionability",
    "field-action-bucket",
    "field-review-comment",
    "field-learning-note",
    "submit-review-next",
    "submit-review-stay",
  ].forEach(id => {{
    const node = document.getElementById(id);
    if (node) node.disabled = disabled;
  }});
}}
function collectReviewedList(tasks) {{
  const reviewed = tasks.filter(item => item.task_status === "reviewed").slice(0, 8);
  reviewedListNode.innerHTML = reviewed.map(item => {{
    const diff = JSON.stringify(item.change_diff || {{}}, null, 0);
    return `<div class="item"><div class="item-title">${{esc(item.salesperson_name || "未识别对象")}}｜${{esc(item.review_reason_code || "")}}</div><div>${{esc(String(item.source_text || "").slice(0, 120))}}</div><div class="item-meta">修改差异：${{esc(diff)}}</div></div>`;
  }}).join("") || `<div class="item muted">当前筛选条件下没有已复核样例。</div>`;
}}
function renderCandidateList() {{
  const topCandidates = reviewCandidates.slice(0, 8);
  document.getElementById("review-candidate-list").innerHTML = topCandidates.map(item => {{
    const sampleTasks = Array.isArray(item.sample_task_ids) ? item.sample_task_ids.slice(0, 3).join("、") : "";
    const priority = item.priority_level === "high" ? "高优先" : "观察中";
    return `<div class="item"><div class="item-title">${{esc(item.update_type || "")}}｜${{esc(item.error_reason_primary || "未标记")}}｜${{esc(priority)}}</div><div>累计 ${{esc(item.count || 0)}} 条，原因码 ${{esc(item.review_reason_code || "未标记")}}</div><div class="item-meta">样例任务：${{esc(sampleTasks || "暂无")}}</div></div>`;
  }}).join("") || `<div class="item muted">当前还没有学习候选。</div>`;
}}
function renderReasonDistribution(tasks) {{
  const counter = new Map();
  for (const item of tasks) {{
    const key = String(item.review_reason_code || "未标记");
    counter.set(key, (counter.get(key) || 0) + 1);
  }}
  const chips = Array.from(counter.entries()).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([key, count]) => `<span class="status-chip muted">${{esc(key)}}：${{count}}</span>`);
  reasonDistributionNode.innerHTML = chips.join("") || `<span class="status-chip muted">当前没有任务</span>`;
}}
function hydrateFilterOptions() {{
  if (filterStatusNode.innerHTML) return;
  fill(filterBatchNode, reviewBatchSummaries.map(item => String(item.batch_id || "")), preferredOpenBatchId());
  fill(filterStatusNode, ["open", "reviewed"], "open");
  fill(filterReasonNode, [...new Set(reviewTasks.map(item => String(item.review_reason_code || "")).filter(Boolean))]);
  fill(filterYearNode, [...new Set(reviewTasks.map(item => String(item.year || "")).filter(Boolean))]);
  fill(filterMonthNode, [...new Set(reviewTasks.map(item => String(item.month || "")).filter(Boolean))]);
  fill(filterWeekNode, [...new Set(reviewTasks.map(item => String(item.week_of_month || "")).filter(Boolean))]);
  fill(filterBusinessLineNode, [...new Set(reviewTasks.map(item => String((effectiveFields(item).business_line) || "")).filter(Boolean))]);
  fill(filterActorPrimaryNode, [...new Set(reviewTasks.map(item => String((effectiveFields(item).actor_primary) || "")).filter(Boolean))]);
}}
function syncTask() {{
  hydrateFilterOptions();
  const tasks = filteredTasks();
  const batchId = filterBatchNode.value || "all";
  const currentBatch = reviewBatchSummaries.find(item => String(item.batch_id || "") === batchId);
  const batchText = currentBatch
    ? `当前批次 ${{currentBatch.batch_number}}：已完成 ${{currentBatch.reviewed_count}} / ${{currentBatch.task_count}} 条`
    : "当前为全部任务视角";
  filterSummaryNode.textContent = `筛选结果 ${{tasks.length}} 条｜open ${{tasks.filter(item => item.task_status === "open").length}} 条｜reviewed ${{tasks.filter(item => item.task_status === "reviewed").length}} 条｜${{batchText}}。系统学习候选 ${{reviewLearningSummary.candidate_count || 0}} 条，高优先 ${{reviewLearningSummary.high_priority_candidate_count || 0}} 条。`;
  renderReasonDistribution(tasks);
  collectReviewedList(tasks);
  renderCandidateList();
  const task = currentTask();
  if (!task) {{
    titleNode.textContent = "当前筛选已清空";
    metaNode.textContent = "";
    sourceNode.textContent = "";
    contextNode.textContent = "";
    fileNode.textContent = "";
    currentFieldsNode.textContent = "";
    basicInfoNode.textContent = "";
    currentHistoryNode.innerHTML = `<div class="item muted">当前筛选条件下没有可展示任务。</div>`;
    setDisabled(true);
    return;
  }}
  const fields = effectiveFields(task);
  const learning = Object.assign({{}}, {{
    error_reason_primary: "",
    review_necessity: "should_review",
    actionability: "",
    action_bucket: "",
    need_rule_update: false,
    need_prompt_update: false,
    need_annotation_update: false,
    learning_note: "",
  }}, task.learning_fields || {{}});
  titleNode.textContent = (task.salesperson_name || "未识别对象") + "｜" + (task.review_reason_code || "");
  metaNode.innerHTML = `<span class="status-chip ${{task.task_status === "reviewed" ? "good" : "warn"}}">${{esc(task.task_status)}}</span><span class="status-chip muted">批次 ${{esc(String(task.batch_number || ""))}} ｜ 第 ${{currentIndex + 1}} / ${{tasks.length}} 张</span><span class="status-chip muted">${{esc(String(task.year || ""))}}-${{esc(String(task.month || "").padStart(2, "0"))}}${{task.week_of_month ? "-W" + esc(String(task.week_of_month)) : ""}}</span>`;
  sourceNode.textContent = task.source_text || "";
  contextNode.textContent = task.source_context || "";
  fileNode.textContent = task.file_path || "";
  currentFieldsNode.innerHTML = [
    ["是否命中 AI", labelFor("is_ai_hit", fields.is_ai_hit !== false)],
    ["业务线", fields.business_line || "待判断"],
    ["主体类型", fields.actor_primary || "未标注"],
    ["范围类型", labelFor("ai_scope", fields.ai_scope || "product_ai")],
    ["复核状态", labelFor("decision_status", fields.decision_status || "pending_human_review")],
    ["原因码", task.review_reason_code || "未标记"],
  ].map(([label, value]) => `<div class="kv-item"><div class="kv-label">${{esc(label)}}</div><div class="kv-value">${{esc(value)}}</div></div>`).join("");
  basicInfoNode.innerHTML = `report_id=${{esc(task.report_id || "")}} ｜ segment_id=${{esc(task.segment_id || "")}} ｜ 销售=${{esc(task.salesperson_name || "")}} ｜ 战区=${{esc(task.battle_zone_name || "")}} ｜ 区域=${{esc(task.region_name || "")}}`;
  document.getElementById("field-is-ai-hit").value = String(fields.is_ai_hit !== false);
  document.getElementById("field-business-line").value = fields.business_line || "待判断";
  document.getElementById("field-actor-primary").value = fields.actor_primary || "";
  document.getElementById("field-ai-scope").value = fields.ai_scope || "product_ai";
  document.getElementById("field-decision-status").value = fields.decision_status || "pending_human_review";
  document.getElementById("field-review-comment").value = task.review_comment || "";
  document.getElementById("field-review-necessity").value = learning.review_necessity || "should_review";
  document.getElementById("field-actionability").value = learning.actionability || "";
  document.getElementById("field-action-bucket").value = learning.action_bucket || "";
  document.getElementById("field-learning-note").value = learning.learning_note || "";
  currentHistoryNode.innerHTML = task.task_status === "reviewed"
    ? `<div class="item"><div class="item-title">最近一次提交</div><div>复核人：${{esc(task.reviewer || "")}} ｜ 时间：${{esc(task.reviewed_at || "")}}</div><div class="item-meta">备注：${{esc(task.review_comment || "无")}}</div><div class="item-meta">系统归因：${{esc(labelFor("error_reason_primary", (task.learning_fields || {{}}).error_reason_primary || ""))}} ｜ 行动价值 ${{esc(labelFor("actionability", (task.learning_fields || {{}}).actionability || ""))}}</div><div class="item-meta">修改差异：${{esc(JSON.stringify(task.change_diff || {{}}, null, 0))}}</div></div>`
    : `<div class="item muted">当前任务暂无复核记录。</div>`;
  setDisabled(task.task_status === "reviewed");
}}
function validatePayload(task, reviewed_fields, learning_fields) {{
  if (!learning_fields.review_necessity) return "提交复核时必须填写是否值得进复核。";
  if (reviewed_fields.is_ai_hit && !learning_fields.actionability) return "AI 命中条目必须填写行动价值。";
  if (learning_fields.actionability === "actionable" && !learning_fields.action_bucket) return "行动价值为 actionable 时必须填写行动归档。";
  return "";
}}
document.getElementById("review-prev")?.addEventListener("click", () => {{
  currentIndex -= 1;
  syncTask();
}});
document.getElementById("review-next")?.addEventListener("click", () => {{
  currentIndex += 1;
  syncTask();
}});
document.getElementById("review-skip")?.addEventListener("click", () => {{
  currentIndex += 1;
  syncTask();
}});
document.getElementById("review-open-evidence")?.addEventListener("click", () => {{
  window.open("/evidence", "_blank");
}});
async function submitReview(moveNext) {{
  const task = currentTask();
  if (!task) return;
  const reviewed_fields = {{
    is_ai_hit: document.getElementById("field-is-ai-hit").value === "true",
    business_line: document.getElementById("field-business-line").value,
    actor_primary: document.getElementById("field-actor-primary").value,
    ai_scope: document.getElementById("field-ai-scope").value,
    decision_status: document.getElementById("field-decision-status").value,
  }};
  const learning_fields = {{
    error_reason_primary: "",
    review_necessity: document.getElementById("field-review-necessity").value,
    actionability: document.getElementById("field-actionability").value,
    action_bucket: document.getElementById("field-action-bucket").value,
    need_rule_update: false,
    need_prompt_update: false,
    need_annotation_update: false,
    learning_note: document.getElementById("field-learning-note").value || "",
  }};
  const validationError = validatePayload(task, reviewed_fields, learning_fields);
  if (validationError) {{
    submitStatus.textContent = validationError;
    return;
  }}
  submitStatus.textContent = "提交中...";
  const payload = {{
    task_id: task.task_id,
    report_id: task.report_id,
    segment_id: task.segment_id,
    reviewed_fields,
    reviewer: document.getElementById("field-reviewer").value || "wales",
    review_comment: document.getElementById("field-review-comment").value || "",
    learning_fields,
  }};
  const res = await fetch("/api/review-decisions", {{
    method: "POST",
    headers: {{ "Content-Type": "application/json" }},
    body: JSON.stringify(payload),
  }});
  if (!res.ok) {{
    const payload = await res.json().catch(() => ({{}}));
    submitStatus.textContent = payload.error || "提交失败";
    return;
  }}
  const response = await res.json();
  const decision = response.decision || {{}};
  reviewLearningSummary = response.review_learning_summary || reviewLearningSummary;
  reviewCandidates = response.review_candidates || reviewCandidates;
  reviewBatchSummaries = response.review_batch_summaries || reviewBatchSummaries;
  task.task_status = "reviewed";
  task.edited_fields = reviewed_fields;
  task.learning_fields = learning_fields;
  task.review_comment = payload.review_comment;
  task.reviewer = payload.reviewer;
  task.reviewed_at = decision.reviewed_at || "";
  task.change_diff = decision.change_diff || {{}};
  if (moveNext) {{
    filterStatusNode.value = "open";
    const sameBatchOpen = reviewTasks.filter(item => item.task_status === "open" && item.batch_id === task.batch_id);
    submitStatus.textContent = sameBatchOpen.length
      ? "已写回，已自动切到下一条待复核卡片。"
      : "本批次已全部复核完成，可以先停下来做一轮规则优化。";
  }} else {{
    filterStatusNode.value = "all";
    const allTasks = filteredTasks();
    currentIndex = Math.max(0, allTasks.findIndex(item => item.task_id === task.task_id));
    submitStatus.textContent = "已写回，当前卡片已切到只读状态。";
  }}
  syncTask();
}}
document.getElementById("submit-review-next")?.addEventListener("click", () => submitReview(true));
document.getElementById("submit-review-stay")?.addEventListener("click", () => submitReview(false));
document.getElementById("rebuild-workbench")?.addEventListener("click", async () => {{
  submitStatus.textContent = "正在重建工作台...";
  const res = await fetch("/api/rebuild", {{
    method: "POST",
    headers: {{ "Content-Type": "application/json" }},
    body: JSON.stringify({{ source: "review-page" }}),
  }});
  if (!res.ok) {{
    submitStatus.textContent = "重建失败";
    return;
  }}
  const payload = await res.json();
  const mentionCount = payload?.result?.total_ai_mentions ?? "-";
  const reviewCount = payload?.result?.open_review_tasks ?? "-";
  submitStatus.textContent = "已重建：AI证据 " + mentionCount + " 条，待复核 " + reviewCount + " 条。页面即将刷新。";
  setTimeout(() => window.location.reload(), 800);
}});
[
  filterBatchNode,
  filterStatusNode,
  filterReasonNode,
  filterYearNode,
  filterMonthNode,
  filterWeekNode,
  filterBusinessLineNode,
  filterActorPrimaryNode,
].forEach((node) => node?.addEventListener("change", () => {{
  currentIndex = 0;
  syncTask();
}}));
filterSalespersonNode?.addEventListener("input", () => {{
  currentIndex = 0;
  syncTask();
}});
syncTask();
</script>"""


def _trends_scripts(
    snapshot: Dict[str, object],
    trend_cube: Sequence[Dict[str, object]],
    trend_explanations: Sequence[Dict[str, object]],
    evidence_index: Sequence[Dict[str, object]],
) -> str:
    monthly = [row for row in trend_cube if str(row.get("grain", "")) == "month"]
    weekly = [row for row in trend_cube if str(row.get("grain", "")) == "week"]
    payload = {
        "monthly": monthly,
        "weekly": weekly,
        "explanations": list(trend_explanations),
        "evidence": list(evidence_index),
        "latestMonth": str(snapshot.get("latest_year_month", "")),
        "latestWeek": str(snapshot.get("latest_year_week", "")),
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return f"""
<script>
const trendPayload = {data_json};
const primaryMonthNode = document.getElementById("trend-primary-month");
const compareMonthNode = document.getElementById("trend-compare-month");
const weekNode = document.getElementById("trend-week");
const metricNode = document.getElementById("trend-metrics");
const explanationsNode = document.getElementById("trend-explanations");
const monthlyTableNode = document.getElementById("trend-monthly-table");
const weeklyTableNode = document.getElementById("trend-weekly-table");
const structureNode = document.getElementById("trend-structure");
const weekEvidenceNode = document.getElementById("trend-week-evidence");

function esc(text) {{
  return String(text ?? "").replace(/[&<>"]/g, (c) => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}})[c]);
}}
function pct(value) {{
  return (Number(value || 0) * 100).toFixed(1) + "%";
}}
function monthLabel(row) {{
  return String(row.year).padStart(4, "0") + "-" + String(row.month).padStart(2, "0");
}}
function weekLabel(row) {{
  return monthLabel(row) + "-W" + String(row.week_of_month);
}}
function deltaText(a, b, isRate=false) {{
  const before = Number(a || 0);
  const after = Number(b || 0);
  const delta = after - before;
  if (isRate) return pct(before) + " → " + pct(after) + "（" + (delta >= 0 ? "+" : "") + pct(Math.abs(delta)).replace("%","pct") + "）";
  return before + " → " + after + "（" + (delta >= 0 ? "+" : "") + delta + "）";
}}
function fillSelect(node, values, selected, includeAll=false) {{
  const opts = [];
  if (includeAll) opts.push(`<option value="all">全部</option>`);
  for (const value of values) {{
    opts.push(`<option value="${{esc(value)}}" ${{value === selected ? "selected" : ""}}>${{esc(value)}}</option>`);
  }}
  node.innerHTML = opts.join("");
}}
function monthRows() {{
  return trendPayload.monthly.slice();
}}
function weekRowsForMonth(month) {{
  return trendPayload.weekly.filter(row => monthLabel(row) === month);
}}
function evidenceForWeek(month, week) {{
  return trendPayload.evidence.filter(row => {{
    const rowMonth = String(row.year).padStart(4, "0") + "-" + String(row.month).padStart(2, "0");
    const rowWeek = rowMonth + "-W" + String(row.week_of_month);
    if (week && week !== "all") return rowWeek === week;
    return rowMonth === month;
  }});
}}
function renderTrendPage() {{
  const monthly = monthRows();
  const monthLabels = monthly.map(monthLabel);
  if (!primaryMonthNode.innerHTML) {{
    const latestMonth = trendPayload.latestMonth || monthLabels[monthLabels.length - 1] || "";
    let compareDefault = monthLabels.find(label => label.startsWith(String(Number(latestMonth.slice(0,4)) - 1)) && label.slice(5) === latestMonth.slice(5));
    if (!compareDefault) compareDefault = monthLabels[Math.max(0, monthLabels.length - 2)] || latestMonth;
    fillSelect(primaryMonthNode, monthLabels, latestMonth);
    fillSelect(compareMonthNode, monthLabels, compareDefault);
  }}
  const selectedMonth = primaryMonthNode.value;
  const compareMonth = compareMonthNode.value;
  const selected = monthly.find(row => monthLabel(row) === selectedMonth) || monthly[monthly.length - 1] || {{}};
  const compare = monthly.find(row => monthLabel(row) === compareMonth) || {{}};
  const weeks = weekRowsForMonth(selectedMonth);
  if (!weekNode.dataset.bound || weekNode.dataset.month !== selectedMonth) {{
    fillSelect(weekNode, weeks.map(weekLabel), trendPayload.latestWeek && trendPayload.latestWeek.startsWith(selectedMonth) ? trendPayload.latestWeek : (weeks.length ? weekLabel(weeks[weeks.length - 1]) : "all"), true);
    weekNode.dataset.bound = "1";
    weekNode.dataset.month = selectedMonth;
  }}
  const selectedWeek = weekNode.value;
  metricNode.innerHTML = [
    ["AI 证据", deltaText(compare.ai_mentions, selected.ai_mentions), "选中月份与对比月份的命中量"],
    ["活跃销售", deltaText(compare.active_sales_count, selected.active_sales_count), "看增长来自更多销售还是更高频"],
    ["销售渗透率", pct(selected.sales_penetration_rate), "当前月活跃销售占在岗销售比例"],
    ["区域覆盖率", pct(selected.region_coverage_rate), "当前月覆盖到的区域比例"],
    ["待复核率", pct(selected.pending_review_rate), "当前月结论可信度的主要噪声来源"]
  ].map(item => `<div class="panel"><div class="metric-label">${{esc(item[0])}}</div><div class="metric-value" style="font-size:22px;">${{esc(item[1])}}</div><div class="muted">${{esc(item[2])}}</div></div>`).join("");
  const derivedExplanation = (() => {{
    const compareAvg = Number(compare.active_sales_count || 0) ? Number(compare.ai_mentions || 0) / Number(compare.active_sales_count || 1) : 0;
    const selectedAvg = Number(selected.active_sales_count || 0) ? Number(selected.ai_mentions || 0) / Number(selected.active_sales_count || 1) : 0;
    let judgement = "变化有限";
    if (Number(selected.active_sales_count || 0) > Number(compare.active_sales_count || 0) && selectedAvg > compareAvg) judgement = "覆盖和单人强度都在上升";
    else if (Number(selected.active_sales_count || 0) > Number(compare.active_sales_count || 0)) judgement = "更多销售开始加入";
    else if (selectedAvg > compareAvg) judgement = "少数销售提得更频繁";
    return `${{selectedMonth}} 对比 ${{compareMonth}}：AI 证据 ${{compare.ai_mentions || 0}}→${{selected.ai_mentions || 0}}，活跃销售 ${{compare.active_sales_count || 0}}→${{selected.active_sales_count || 0}}。当前判断：${{judgement}}。`;
  }})();
  const staticExplanations = trendPayload.explanations.map(item => `<div class="item"><div class="item-title">${{esc(item.metric_name)}}｜${{esc(item.period)}}</div><div>${{esc(item.explanation)}}</div><div class="item-meta">可信度：${{esc(item.confidence_level)}}</div></div>`);
  explanationsNode.innerHTML = [`<div class="item"><div class="item-title">当前选中对比</div><div>${{esc(derivedExplanation)}}</div><div class="item-meta">主周期：${{esc(selectedMonth)}} ｜ 对比周期：${{esc(compareMonth)}}</div></div>`, ...staticExplanations].join("");
  monthlyTableNode.innerHTML = monthly.map(row => `<tr><td>${{esc(monthLabel(row))}}</td><td>${{esc(row.ai_mentions)}}</td><td>${{esc(row.active_sales_count)}}</td><td>${{esc(pct(row.sales_penetration_rate))}}</td><td>${{esc(pct(row.region_coverage_rate))}}</td><td>${{esc(pct(row.pending_review_rate))}}</td></tr>`).join("") || `<tr><td colspan="6">暂无数据</td></tr>`;
  weeklyTableNode.innerHTML = weeks.map(row => `<tr><td>${{esc(weekLabel(row))}}</td><td>${{esc(row.ai_mentions)}}</td><td>${{esc(row.active_sales_count)}}</td><td>${{esc(pct(row.pending_review_rate))}}</td></tr>`).join("") || `<tr><td colspan="4">暂无周度数据</td></tr>`;
  const actorRows = Object.entries(selected.actor_breakdown || {{}}).map(([k, v]) => `<div class="item"><div class="item-title">${{esc(k)}}</div><div>${{esc(v)}} 条</div></div>`).join("") || `<div class="item muted">暂无主体结构</div>`;
  const lineRows = Object.entries(selected.business_line_breakdown || {{}}).map(([k, v]) => `<div class="item"><div class="item-title">${{esc(k)}}</div><div>${{esc(v)}} 条</div></div>`).join("") || `<div class="item muted">暂无业务线结构</div>`;
  structureNode.innerHTML = `<div class="grid two"><div><h3>主体结构</h3><div class="list">${{actorRows}}</div></div><div><h3>业务线结构</h3><div class="list">${{lineRows}}</div></div></div>`;
  const evidenceRows = evidenceForWeek(selectedMonth, selectedWeek).slice(0, 20);
  weekEvidenceNode.innerHTML = evidenceRows.map(row => `<div class="item"><div class="item-title">${{esc(String(row.year).padStart(4,'0') + '-' + String(row.month).padStart(2,'0') + '-W' + String(row.week_of_month))}}｜${{esc(row.salesperson_name || '未识别对象')}}｜${{esc(row.actor_primary || '未标注')}}</div><div>${{esc(row.source_text || '')}}</div><div class="item-meta">${{esc(row.business_line || '待判断')}} ｜ ${{esc(row.decision_status || '')}}</div></div>`).join("") || `<div class="item muted">当前周次暂无可展示证据。</div>`;
}}
primaryMonthNode?.addEventListener("change", renderTrendPage);
compareMonthNode?.addEventListener("change", renderTrendPage);
weekNode?.addEventListener("change", renderTrendPage);
renderTrendPage();
</script>"""


def _sales_scripts(
    snapshot: Dict[str, object],
    salesperson_profiles: Sequence[Dict[str, object]],
    region_rollups: Sequence[Dict[str, object]],
) -> str:
    payload = {
        "profiles": list(salesperson_profiles),
        "regions": list(region_rollups),
        "months": list(snapshot.get("available_months", [])),
        "latestMonth": str(snapshot.get("latest_year_month", "")),
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return f"""
<script>
const salesPayload = {data_json};
const segmentNode = document.getElementById("sales-segment-filter");
const actorNode = document.getElementById("sales-actor-filter");
const lineNode = document.getElementById("sales-line-filter");
const monthNode = document.getElementById("sales-month-filter");
const searchNode = document.getElementById("sales-search");
const summaryNode = document.getElementById("sales-summary-cards");
const segmentPanelsNode = document.getElementById("sales-segment-panels");
const regionTableNode = document.getElementById("sales-region-table");
const profileGridNode = document.getElementById("sales-profile-grid");
function esc(text) {{
  return String(text ?? "").replace(/[&<>"]/g, (c) => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}})[c]);
}}
function fill(node, values, selected="all") {{
  node.innerHTML = ['<option value="all">全部</option>', ...values.map(value => `<option value="${{esc(value)}}" ${{value === selected ? "selected" : ""}}>${{esc(value)}}</option>`)].join("");
}}
function historyForMonth(profile, month) {{
  return (profile.history || []).find(item => item.period === month);
}}
function miniHistory(history) {{
  const values = history.map(item => Number(item.mentions || 0));
  const maxValue = Math.max(...values, 1);
  return `<div style="display:flex;gap:4px;align-items:flex-end;height:40px;margin-top:8px;">${{history.map(item => `<div title="${{esc(item.period)}}:${{esc(item.mentions)}}" style="flex:1;background:#d5b189;border-radius:4px 4px 0 0;height:${{Math.max(10, Math.round((Number(item.mentions || 0) / maxValue) * 40))}}px;"></div>`).join("")}}</div>`;
}}
function buildRegionRows(profiles) {{
  const grouped = new Map();
  for (const row of profiles) {{
    const key = `${{row.battle_zone_name || "未识别战区"}}||${{row.region_name || "未识别区域"}}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(row);
  }}
  return Array.from(grouped.entries()).map(([key, items]) => {{
    const [battle, region] = key.split("||");
    const active = items.filter(item => Number(item.total_mentions || 0) > 0);
    const silent = items.filter(item => Number(item.total_mentions || 0) === 0 && item.employment_status === "active");
    const totalMentions = active.reduce((sum, item) => sum + Number(item.total_mentions || 0), 0);
    const top = active.slice().sort((a, b) => Number(b.total_mentions || 0) - Number(a.total_mentions || 0));
    const dominance = totalMentions ? Number(top[0]?.total_mentions || 0) / totalMentions : 0;
    const maturity = !active.length ? "未启动" : dominance >= 0.6 ? "依赖少数销售" : active.length >= 3 ? "相对均衡" : "起步中";
    return {{ battle, region, activeCount: active.length, silentCount: silent.length, dominance, maturity, top }};
  }}).sort((a, b) => (b.activeCount - a.activeCount) || (b.top[0]?.total_mentions || 0) - (a.top[0]?.total_mentions || 0));
}}
function renderSalesPage() {{
  const profiles = salesPayload.profiles;
  const segmentValues = [...new Set(profiles.map(item => item.segment).filter(Boolean))];
  const actorValues = [...new Set(profiles.flatMap(item => Object.keys(item.actor_breakdown || {{}})).filter(Boolean))];
  const lineValues = [...new Set(profiles.flatMap(item => Object.keys(item.business_line_breakdown || {{}})).filter(Boolean))];
  if (!segmentNode.innerHTML) {{
    fill(segmentNode, segmentValues);
    fill(actorNode, actorValues);
    fill(lineNode, lineValues);
    fill(monthNode, salesPayload.months || [], salesPayload.latestMonth || "all");
  }}
  const segment = segmentNode.value;
  const actor = actorNode.value;
  const line = lineNode.value;
  const month = monthNode.value;
  const search = (searchNode.value || "").trim();
  const filtered = profiles.filter(profile => {{
    if (segment !== "all" && profile.segment !== segment) return false;
    if (actor !== "all" && !Number((profile.actor_breakdown || {{}})[actor] || 0)) return false;
    if (line !== "all" && !Number((profile.business_line_breakdown || {{}})[line] || 0)) return false;
    if (month !== "all") {{
      const hit = historyForMonth(profile, month);
      if (!hit || !Number(hit.mentions || 0)) return false;
    }}
    if (search && !(profile.display_name || "").includes(search)) return false;
    return true;
  }});
  const totalMentions = filtered.reduce((sum, item) => sum + Number(item.total_mentions || 0), 0);
  const doctorFeedback = filtered.reduce((sum, item) => sum + Number(item.doctor_feedback_mentions || 0), 0);
  const opportunities = filtered.reduce((sum, item) => sum + Number(item.opportunity_mentions || 0), 0);
  const reviewOpen = filtered.reduce((sum, item) => sum + Number(item.review_open_count || 0), 0);
  summaryNode.innerHTML = [
    ["筛选后销售", filtered.length, "当前筛选条件下的销售对象数量"],
    ["AI 总提及", totalMentions, "这些销售累计提到 AI 的次数"],
    ["医生反馈", doctorFeedback, "这些销售带回的医生反馈数量"],
    ["潜在机会", opportunities, "这些销售带回的 AI 机会数量"],
    ["待复核", reviewOpen, "这些销售相关的人工复核任务量"]
  ].map(item => `<div class="panel"><div class="metric-label">${{esc(item[0])}}</div><div class="metric-value">${{esc(item[1])}}</div><div class="muted">${{esc(item[2])}}</div></div>`).join("");
  const segmentOrder = ["高频使用者", "中频使用者", "偶发使用者", "长期未提及者"];
  segmentPanelsNode.innerHTML = segmentOrder.map(name => {{
    const rows = filtered.filter(item => item.segment === name);
    return `<div class="item"><div class="item-title">${{esc(name)}}</div><div class="item-meta">人数：${{rows.length}}</div><div>${{rows.slice(0, 6).map(item => esc(item.display_name || "")).join("、") || "暂无"}}</div></div>`;
  }}).join("");
  const regionRows = buildRegionRows(filtered);
  regionTableNode.innerHTML = regionRows.slice(0, 20).map(row => `<tr><td>${{esc(row.battle)}} / ${{esc(row.region)}}</td><td>${{esc(row.activeCount)}}</td><td>${{esc(row.silentCount)}}</td><td>${{esc((row.dominance * 100).toFixed(1) + "%")}}</td><td>${{esc(row.maturity)}}</td><td>${{esc(row.top.slice(0, 3).map(item => item.display_name || "").join("、"))}}</td></tr>`).join("") || `<tr><td colspan="6">暂无区域数据</td></tr>`;
  profileGridNode.innerHTML = filtered.slice(0, 24).map(row => {{
    const actorSummary = Object.entries(row.actor_breakdown || {{}}).map(([k, v]) => `${{k}}:${{v}}`).join("，") || "暂无";
    const lineSummary = Object.entries(row.business_line_breakdown || {{}}).map(([k, v]) => `${{k}}:${{v}}`).join("，") || "暂无";
    const history = row.history || [];
    const sample = (row.high_quality_evidence || [])[0];
    return `<div class="panel"><h3>${{esc(row.display_name || "")}}</h3><div class="item-meta">${{esc(row.battle_zone_name || "")}} / ${{esc(row.region_name || "")}}</div><div class="item-meta"><span class="pill">${{esc(row.segment || "")}}</span><span class="pill ${{row.recommended_case ? "good" : "warn"}}">${{row.recommended_case ? "建议沉淀案例" : "继续观察"}}</span></div><div>首次提及：${{esc(row.first_seen_period || "")}} ｜ 最近提及：${{esc(row.last_seen_period || "")}}</div><div>总提及：${{esc(row.total_mentions || 0)}} ｜ 医生反馈：${{esc(row.doctor_feedback_mentions || 0)}} ｜ 潜在机会：${{esc(row.opportunity_mentions || 0)}} ｜ 待复核：${{esc(row.review_open_count || 0)}}</div><div class="item-meta">主体分布：${{esc(actorSummary)}}</div><div class="item-meta">业务线分布：${{esc(lineSummary)}}</div><div class="item-meta">月度趋势</div>${{miniHistory(history)}}<div class="item-meta">代表证据：${{esc(sample?.source_text || "暂无代表证据")}}</div></div>`;
  }}).join("") || `<div class="item muted">当前筛选条件下没有销售画像。</div>`;
}}
segmentNode?.addEventListener("change", renderSalesPage);
actorNode?.addEventListener("change", renderSalesPage);
lineNode?.addEventListener("change", renderSalesPage);
monthNode?.addEventListener("change", renderSalesPage);
searchNode?.addEventListener("input", renderSalesPage);
renderSalesPage();
</script>"""


def _evidence_scripts(snapshot: Dict[str, object], evidence_index: Sequence[Dict[str, object]]) -> str:
    payload = {
        "evidence": list(evidence_index),
        "months": list(snapshot.get("available_months", [])),
        "latestMonth": str(snapshot.get("latest_year_month", "")),
        "latestWeek": str(snapshot.get("latest_year_week", "")),
    }
    data_json = json.dumps(payload, ensure_ascii=False)
    return f"""
<script>
const evidencePayload = {data_json};
const evidenceMonthNode = document.getElementById("evidence-month-filter");
const evidenceWeekNode = document.getElementById("evidence-week-filter");
const evidenceSearchNode = document.getElementById("evidence-search");
const evidenceActorNode = document.getElementById("evidence-actor-filter");
const evidenceLineNode = document.getElementById("evidence-line-filter");
const evidenceTableNode = document.getElementById("evidence-table");
function esc(text) {{
  return String(text ?? "").replace(/[&<>"]/g, (c) => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}})[c]);
}}
function monthOf(row) {{
  return String(row.year).padStart(4, "0") + "-" + String(row.month).padStart(2, "0");
}}
function weekOf(row) {{
  return monthOf(row) + "-W" + String(row.week_of_month);
}}
function fill(node, values, selected="all") {{
  node.innerHTML = ['<option value="all">全部</option>', ...values.map(value => `<option value="${{esc(value)}}" ${{value === selected ? "selected" : ""}}>${{esc(value)}}</option>`)].join("");
}}
function renderEvidencePage() {{
  const rows = evidencePayload.evidence;
  const actorValues = [...new Set(rows.map(item => item.actor_primary).filter(Boolean))];
  const lineValues = [...new Set(rows.map(item => item.business_line).filter(Boolean))];
  if (!evidenceMonthNode.innerHTML) {{
    fill(evidenceMonthNode, evidencePayload.months || [], evidencePayload.latestMonth || "all");
    fill(evidenceActorNode, actorValues);
    fill(evidenceLineNode, lineValues);
  }}
  const selectedMonth = evidenceMonthNode.value;
  const weekValues = rows.filter(item => selectedMonth === "all" || monthOf(item) === selectedMonth).map(weekOf).filter((value, index, arr) => arr.indexOf(value) === index);
  const currentWeek = evidenceWeekNode.value;
  const nextWeek = weekValues.includes(currentWeek) ? currentWeek : (evidencePayload.latestWeek && weekValues.includes(evidencePayload.latestWeek) ? evidencePayload.latestWeek : "all");
  fill(evidenceWeekNode, weekValues, nextWeek);
  const selectedWeek = evidenceWeekNode.value;
  const actor = evidenceActorNode.value;
  const line = evidenceLineNode.value;
  const search = (evidenceSearchNode.value || "").trim();
  const filtered = rows.filter(row => {{
    if (selectedMonth !== "all" && monthOf(row) !== selectedMonth) return false;
    if (selectedWeek !== "all" && weekOf(row) !== selectedWeek) return false;
    if (actor !== "all" && row.actor_primary !== actor) return false;
    if (line !== "all" && row.business_line !== line) return false;
    if (search && !(row.salesperson_name || "").includes(search) && !(row.source_text || "").includes(search)) return false;
    return true;
  }});
  evidenceTableNode.innerHTML = filtered.slice(0, 200).map(row => `<tr><td>${{esc(weekOf(row))}}</td><td>${{esc(row.salesperson_name || "")}}</td><td>${{esc(row.battle_zone_name || "")}} / ${{esc(row.region_name || "")}}</td><td>${{esc(row.business_line || "")}}</td><td>${{esc(row.actor_primary || "")}}</td><td>${{esc(row.decision_status || "")}}</td><td style="min-width:380px;">${{esc(row.source_text || "")}}</td><td class="code">${{esc(row.file_path || "")}}</td></tr>`).join("") || `<tr><td colspan="8">当前筛选条件下没有证据。</td></tr>`;
}}
evidenceMonthNode?.addEventListener("change", renderEvidencePage);
evidenceWeekNode?.addEventListener("change", renderEvidencePage);
evidenceActorNode?.addEventListener("change", renderEvidencePage);
evidenceLineNode?.addEventListener("change", renderEvidencePage);
evidenceSearchNode?.addEventListener("input", renderEvidencePage);
renderEvidencePage();
</script>"""


def _evidence_body(snapshot: Dict[str, object]) -> str:
    return f"""
    <div class="panel">
      <h2>证据与原文下钻</h2>
      <div class="grid three">
        <div><label>月份</label><select id="evidence-month-filter"></select></div>
        <div><label>周次</label><select id="evidence-week-filter"></select></div>
        <div><label>销售搜索</label><input id="evidence-search" placeholder="按花名搜索" /></div>
      </div>
      <div class="grid two" style="margin-top:12px;">
        <div><label>主体标签</label><select id="evidence-actor-filter"></select></div>
        <div><label>业务线标签</label><select id="evidence-line-filter"></select></div>
      </div>
      <div class="item" style="margin-top:12px;">
        <div class="item-title">默认口径</div>
        <div>{html.escape(str(snapshot.get('time_scope_note', '')))}</div>
      </div>
    </div>
    <div class="panel" style="margin-top:16px;">
      <h2>命中明细</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>时间</th><th>销售</th><th>战区 / 区域</th><th>业务线</th><th>主体</th><th>状态</th><th>原文</th><th>路径</th></tr></thead>
          <tbody id="evidence-table"></tbody>
        </table>
      </div>
    </div>
    """


def _metric_card(title: str, value: object, desc: str) -> str:
    return (
        "<div class='panel'>"
        f"<div class='metric-label'>{html.escape(str(title))}</div>"
        f"<div class='metric-value'>{html.escape(str(value))}</div>"
        f"<div class='muted'>{html.escape(desc)}</div>"
        "</div>"
    )


def _bullet_item(text: str) -> str:
    return f"<div class='item'>{html.escape(str(text))}</div>"


def _empty_item(text: str) -> str:
    return f"<div class='item muted'>{html.escape(text)}</div>"


def _insight_brief_item(row: Dict[str, object]) -> str:
    return (
        "<div class='item'>"
        f"<div class='item-title'>{html.escape(str(row.get('title', '')))}</div>"
        f"<div>{html.escape(str(row.get('judgement', '')))}</div>"
        f"<div class='item-meta'>可信度：{html.escape(str(row.get('confidence_level', '')))} ｜ 待复核：{html.escape(str(row.get('pending_review_count', 0)))} 条</div>"
        "</div>"
    )


def _trend_item(row: Dict[str, object]) -> str:
    return (
        "<div class='item'>"
        f"<div class='item-title'>{html.escape(str(row.get('metric_name', '')))}｜{html.escape(str(row.get('period', '')))}</div>"
        f"<div>{html.escape(str(row.get('explanation', '')))}</div>"
        f"<div class='item-meta'>变化类型：{html.escape(str(row.get('change_type', '')))} ｜ 可信度：{html.escape(str(row.get('confidence_level', '')))}</div>"
        "</div>"
    )


def _trend_row(row: Dict[str, object]) -> str:
    return (
        "<tr>"
        f"<td>{html.escape(f'{int(row.get('year', 0)):04d}-{int(row.get('month', 0)):02d}')}</td>"
        f"<td>{html.escape(str(row.get('ai_mentions', 0)))}</td>"
        f"<td>{html.escape(str(row.get('active_sales_count', 0)))}</td>"
        f"<td>{html.escape(f'{float(row.get('sales_penetration_rate', 0.0)):.1%}')}</td>"
        f"<td>{html.escape(f'{float(row.get('region_coverage_rate', 0.0)):.1%}')}</td>"
        f"<td>{html.escape(f'{float(row.get('pending_review_rate', 0.0)):.1%}')}</td>"
        "</tr>"
    )


def _structure_panels(row: Dict[str, object]) -> str:
    actor_rows = row.get("actor_breakdown", {})
    line_rows = row.get("business_line_breakdown", {})
    actor_html = "".join(f"<div class='item'><div class='item-title'>{html.escape(str(k))}</div><div>{v} 条</div></div>" for k, v in actor_rows.items()) or _empty_item("暂无主体结构")
    line_html = "".join(f"<div class='item'><div class='item-title'>{html.escape(str(k))}</div><div>{v} 条</div></div>" for k, v in line_rows.items()) or _empty_item("暂无业务线结构")
    return f"<div class='grid two'><div><h3>主体结构</h3><div class='list'>{actor_html}</div></div><div><h3>业务线结构</h3><div class='list'>{line_html}</div></div></div>"


def _segment_panel(name: str, rows: Sequence[Dict[str, object]]) -> str:
    return (
        "<div class='item'>"
        f"<div class='item-title'>{html.escape(name)}</div>"
        f"<div class='item-meta'>人数：{len(rows)}</div>"
        f"<div>{'、'.join(html.escape(str(row.get('display_name', ''))) for row in rows[:6]) or '暂无'}</div>"
        "</div>"
    )


def _profile_card(row: Dict[str, object]) -> str:
    actors = row.get("actor_breakdown", {})
    lines = row.get("business_line_breakdown", {})
    history = " / ".join(f"{item.get('period', '')}:{item.get('mentions', 0)}" for item in row.get("history", []) if int(item.get("mentions", 0)) > 0)
    evidence = row.get("high_quality_evidence", [])
    evidence_text = trim_text(str(evidence[0].get("source_text", "")), 80) if evidence else "暂无代表证据"
    return (
        "<div class='panel'>"
        f"<h3>{html.escape(str(row.get('display_name', '')))}</h3>"
        f"<div class='item-meta'>{html.escape(str(row.get('battle_zone_name', '')))} / {html.escape(str(row.get('region_name', '')))}</div>"
        f"<div class='item-meta'><span class='pill'>{html.escape(str(row.get('segment', '')))}</span><span class='pill {'good' if bool(row.get('recommended_case', False)) else 'warn'}'>{'建议沉淀案例' if bool(row.get('recommended_case', False)) else '继续观察'}</span></div>"
        f"<div>首次提及：{html.escape(str(row.get('first_seen_period', '')))} ｜ 最近提及：{html.escape(str(row.get('last_seen_period', '')))}</div>"
        f"<div>总提及：{html.escape(str(row.get('total_mentions', 0)))} ｜ 医生反馈：{html.escape(str(row.get('doctor_feedback_mentions', 0)))} ｜ 潜在机会：{html.escape(str(row.get('opportunity_mentions', 0)))} ｜ 待复核：{html.escape(str(row.get('review_open_count', 0)))}</div>"
        f"<div class='item-meta'>主体：{html.escape(', '.join(f'{k}:{v}' for k, v in actors.items()) or '暂无')}</div>"
        f"<div class='item-meta'>业务线：{html.escape(', '.join(f'{k}:{v}' for k, v in lines.items()) or '暂无')}</div>"
        f"<div class='item-meta'>历史趋势：{html.escape(history or '暂无')}</div>"
        f"<div class='item-meta'>代表证据：{html.escape(evidence_text)}</div>"
        "</div>"
    )


def _region_row(row: Dict[str, object]) -> str:
    top_names = "、".join(str(item.get("display_name", "")) for item in row.get("top_salespeople", [])[:3])
    return (
        "<tr>"
        f"<td>{html.escape(str(row.get('battle_zone_name', '')))} / {html.escape(str(row.get('region_name', '')))}</td>"
        f"<td>{html.escape(str(row.get('active_sales_count', 0)))}</td>"
        f"<td>{html.escape(str(row.get('silent_sales_count', 0)))}</td>"
        f"<td>{html.escape(f'{float(row.get('dominance_ratio', 0.0)):.1%}')}</td>"
        f"<td>{html.escape(str(row.get('maturity_judgement', '')))}</td>"
        f"<td>{html.escape(top_names)}</td>"
        "</tr>"
    )


def _insight_card_item(card: Dict[str, object]) -> str:
    owner_refs = "、".join(str(item) for item in card.get("owner_refs", []))
    evidence_refs = card.get("representative_evidence_refs", [])
    preview = trim_text(str(evidence_refs[0].get("source_text", "")), 100) if evidence_refs else ""
    return (
        "<div class='item'>"
        f"<div class='item-title'>{html.escape(str(card.get('title', '')))}</div>"
        f"<div>{html.escape(str(card.get('judgement', '')))}</div>"
        f"<div class='item-meta'>{html.escape(str(card.get('why_it_matters', '')))}</div>"
        f"<div class='item-meta'>证据 {html.escape(str(card.get('evidence_count', 0)))} 条｜已确认 {html.escape(str(card.get('confirmed_evidence_count', 0)))} 条｜待复核 {html.escape(str(card.get('pending_review_count', 0)))} 条｜可信度 {html.escape(str(card.get('confidence_level', '')))}</div>"
        f"<div class='item-meta'>涉及销售：{html.escape(owner_refs or '未识别')}</div>"
        f"<div class='item-meta'>建议动作：{html.escape(str(card.get('action_recommendation', '')))}</div>"
        f"<div class='item-meta'>代表证据：{html.escape(preview)}</div>"
        "</div>"
    )


def _candidate_item(row: Dict[str, object]) -> str:
    return (
        "<div class='item'>"
        f"<div class='item-title'>{html.escape(str(row.get('update_type', '')))}｜{html.escape(str(row.get('error_reason_primary', '')))}</div>"
        f"<div>累计 {html.escape(str(row.get('count', 0)))} 条｜优先级 {html.escape(str(row.get('priority_level', '')))}</div>"
        f"<div class='item-meta'>原因码：{html.escape(str(row.get('review_reason_code', '')))}</div>"
        f"<div class='item-meta'>目标标签：{html.escape(json.dumps(row.get('final_labels', {}), ensure_ascii=False))}</div>"
        "</div>"
    )


def _review_task_item(row: Dict[str, object], reviewed: bool = False) -> str:
    current = row.get("current_fields", {})
    diff = row.get("change_diff", {})
    return (
        "<div class='item'>"
        f"<div class='item-title'>{html.escape(str(row.get('salesperson_name', '未识别对象')))}｜{html.escape(str(row.get('review_reason_code', '')))}</div>"
        f"<div>{html.escape(trim_text(str(row.get('source_text', '')), 120))}</div>"
        f"<div class='item-meta'>当前字段：{html.escape(json.dumps(current, ensure_ascii=False))}</div>"
        + (
            f"<div class='item-meta'>修改差异：{html.escape(json.dumps(diff, ensure_ascii=False))}</div>"
            if reviewed and diff
            else ""
        )
        + f"<div class='item-meta'>路径：{html.escape(str(row.get('file_path', '')))}</div>"
        + "</div>"
    )


def _evidence_row(row: Dict[str, object]) -> str:
    return (
        "<tr>"
        f"<td>{html.escape(str(row.get('salesperson_name', '')))}</td>"
        f"<td>{html.escape(str(row.get('battle_zone_name', '')))} / {html.escape(str(row.get('region_name', '')))}</td>"
        f"<td>{html.escape(str(row.get('business_line', '')))}</td>"
        f"<td>{html.escape(str(row.get('actor_primary', '')))}</td>"
        f"<td>{html.escape(str(row.get('decision_status', '')))}</td>"
        f"<td>{html.escape(trim_text(str(row.get('source_text', '')), 80))}</td>"
        f"<td class='code'>{html.escape(str(row.get('file_path', '')))}</td>"
        "</tr>"
    )
