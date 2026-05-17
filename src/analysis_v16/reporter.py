from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence

from .schema import (
    ACTIONABILITY_LABELS,
    BUSINESS_QUESTION_LABELS,
    BUSINESS_QUESTION_VALUES,
    COMPETITOR_SIGNAL_LABELS,
    DOCTOR_ACCEPTANCE_LABELS,
    DOCTOR_NEED_LABELS,
    SALES_AI_USAGE_LABELS,
)


def write_jsonl(path: Path, rows: Iterable[Dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        for row in rows:
            handle.write(json.dumps(row, ensure_ascii=False) + "\n")


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_markdown(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_web_pages(path: Path, pages: Dict[str, str]) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for name, content in pages.items():
        (path / name).write_text(content, encoding="utf-8")


def build_weekly_brief(
    business_summary: Dict[str, object],
    business_insights: Sequence[Dict[str, object]],
    review_batch: Sequence[Dict[str, object]],
) -> str:
    lines = [
        "# AI 一线情报周度摘要",
        "",
        "## 本期判断",
        f"- 业务证据：{business_summary.get('total_business_evidence', 0)} 条",
        f"- 需复核证据：{business_summary.get('review_needed_count', 0)} 条",
        f"- 可行动证据：{business_summary.get('actionable_count', 0)} 条",
        "",
        "## 业务问题分布",
    ]
    for key, count in dict(business_summary.get("business_question_breakdown", {})).items():
        lines.append(f"- {BUSINESS_QUESTION_LABELS.get(key, key)}：{count}")
    lines.extend(["", "## 重点洞察"])
    for insight in business_insights[:8]:
        lines.append(
            f"- {insight.get('title', '')}：{insight.get('judgement', '')} "
            f"{insight.get('why_it_matters', '')} 建议：{insight.get('action_recommendation', '')}"
        )
    lines.extend(["", "## 本轮复核"])
    open_tasks = [row for row in review_batch if str(row.get("task_status", "")) == "open"]
    if open_tasks:
        lines.append(f"- 本轮待复核：{len(open_tasks[:20])} 条，建议按工作台卡片顺序处理。")
    else:
        lines.append("- 当前没有打开状态的复核任务。")
    return "\n".join(lines)


def build_workbench_pages(
    dashboard_snapshot: Dict[str, object],
    business_summary: Dict[str, object],
    business_facts: Sequence[Dict[str, object]],
    business_insights: Sequence[Dict[str, object]],
    review_batch: Sequence[Dict[str, object]],
    trend_cube: Sequence[Dict[str, object]],
    salesperson_profiles: Sequence[Dict[str, object]],
) -> Dict[str, str]:
    pages = {
        "overview.html": _page("overview", "AI 一线情报工作台 V1.6", _overview(dashboard_snapshot, business_summary, business_insights)),
        "trends.html": _page("trends", "趋势中心", _trends(dashboard_snapshot, trend_cube, business_summary)),
        "sales.html": _page("sales", "销售画像", _sales(salesperson_profiles, business_facts)),
        "insights.html": _page("insights", "洞察中心", _insights(business_insights)),
        "review.html": _page("review", "复核学习工作台", _review(review_batch), extra_script=_review_script()),
        "evidence.html": _page("evidence", "证据下钻", _evidence(business_facts)),
    }
    pages["AI一线情报工作台.html"] = pages["overview.html"]
    return pages


def _page(current: str, title: str, body: str, *, extra_script: str = "") -> str:
    nav = "".join(
        f"<a class='nav {'active' if key == current else ''}' href='{key}.html'>{label}</a>"
        for key, label in [
            ("overview", "总览"),
            ("trends", "趋势"),
            ("sales", "销售"),
            ("insights", "洞察"),
            ("review", "复核"),
            ("evidence", "证据"),
        ]
    )
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{html.escape(title)}</title>
  <style>
    :root {{
      --ink:#16201d;
      --muted:#697570;
      --paper:#faf7ef;
      --panel:#fffdf8;
      --line:#e4dac8;
      --accent:#176b5b;
      --accent-2:#b45f2a;
      --soft:#edf4ef;
      --risk:#8c2d18;
      --shadow:0 18px 45px rgba(30,40,30,.08);
    }}
    * {{ box-sizing:border-box; }}
    body {{
      margin:0; color:var(--ink); background:
        radial-gradient(circle at 20% 0%, #e9f4ed 0, transparent 32%),
        linear-gradient(135deg,#fbf7ed 0%,#f4efe5 100%);
      font-family:"Avenir Next","PingFang SC","Hiragino Sans GB",sans-serif;
    }}
    .shell {{ display:grid; grid-template-columns:230px minmax(0,1fr); min-height:100vh; }}
    aside {{ padding:28px 20px; border-right:1px solid var(--line); background:rgba(255,253,248,.72); backdrop-filter:blur(16px); position:sticky; top:0; height:100vh; }}
    .brand {{ font-size:24px; line-height:1.08; font-weight:800; letter-spacing:-.04em; margin-bottom:10px; }}
    .version {{ color:var(--muted); font-size:13px; margin-bottom:30px; }}
    .nav {{ display:block; padding:12px 14px; color:var(--muted); text-decoration:none; border-radius:14px; margin:5px 0; font-weight:700; }}
    .nav.active,.nav:hover {{ color:var(--ink); background:var(--soft); }}
    main {{ padding:34px 42px 56px; max-width:1440px; }}
    h1 {{ font-size:42px; letter-spacing:-.05em; margin:0 0 10px; }}
    h2 {{ font-size:22px; margin:0 0 16px; letter-spacing:-.02em; }}
    .lead {{ font-size:17px; color:var(--muted); margin:0 0 28px; max-width:900px; }}
    .grid {{ display:grid; gap:18px; }}
    .cols-4 {{ grid-template-columns:repeat(4,minmax(0,1fr)); }}
    .cols-3 {{ grid-template-columns:repeat(3,minmax(0,1fr)); }}
    .cols-2 {{ grid-template-columns:repeat(2,minmax(0,1fr)); }}
    .panel {{ background:rgba(255,253,248,.9); border:1px solid var(--line); border-radius:24px; box-shadow:var(--shadow); padding:22px; }}
    .metric {{ font-size:34px; font-weight:850; letter-spacing:-.04em; }}
    .label {{ color:var(--muted); font-size:13px; font-weight:700; margin-top:4px; }}
    .tag {{ display:inline-flex; align-items:center; gap:6px; border-radius:999px; padding:6px 10px; background:var(--soft); color:var(--accent); font-size:12px; font-weight:800; margin:2px 6px 2px 0; }}
    .tag.warn {{ background:#f9ebe5; color:var(--risk); }}
    .card-title {{ font-size:18px; font-weight:850; margin-bottom:10px; }}
    .quote {{ border-left:3px solid var(--accent); padding-left:14px; color:#39433f; line-height:1.72; }}
    .muted {{ color:var(--muted); }}
    .table {{ width:100%; border-collapse:collapse; font-size:14px; }}
    .table th,.table td {{ text-align:left; border-bottom:1px solid var(--line); padding:11px 8px; vertical-align:top; }}
    .review-card {{ display:grid; grid-template-columns:minmax(0,1.1fr) minmax(360px,.9fr); gap:18px; align-items:start; }}
    .field {{ padding:12px; border:1px solid var(--line); border-radius:16px; background:#fffaf1; margin-bottom:10px; }}
    .btn {{ display:inline-block; border:0; border-radius:14px; background:var(--accent); color:white; padding:11px 15px; font-weight:800; text-decoration:none; }}
    .btn.secondary {{ background:#efe5d4; color:var(--ink); }}
    .toolbar {{ display:flex; gap:10px; flex-wrap:wrap; margin:16px 0 22px; }}
    select,input,textarea {{ border:1px solid var(--line); border-radius:12px; padding:10px 12px; background:white; font:inherit; width:100%; }}
    textarea {{ min-height:92px; resize:vertical; }}
    .form-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }}
    .form-field label {{ display:block; font-size:13px; color:var(--muted); font-weight:800; margin:0 0 7px; }}
    .status-line {{ color:var(--muted); font-weight:700; margin-top:10px; }}
    .task-hidden {{ display:none; }}
    @media (max-width: 920px) {{
      .shell {{ grid-template-columns:1fr; }}
      aside {{ position:static; height:auto; }}
      main {{ padding:24px 18px; }}
      .cols-4,.cols-3,.cols-2,.review-card {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <aside>
      <div class="brand">AI 一线<br/>情报工作台</div>
      <div class="version">V1.6 quality build</div>
      {nav}
    </aside>
    <main>{body}</main>
  </div>
  {extra_script}
</body>
</html>"""


def _overview(snapshot: Dict[str, object], summary: Dict[str, object], insights: Sequence[Dict[str, object]]) -> str:
    return f"""
    <h1>本期判断总览</h1>
    <p class="lead">按业务问题组织，而不是按标签堆叠。优先回答医生接纳度、医生诉求、销售使用 AI、竞品动作和需要复核的边界样本。</p>
    <section class="grid cols-4">
      {_metric("业务证据", summary.get("total_business_evidence", 0), "进入 v1.6 业务问题层的证据")}
      {_metric("待复核", summary.get("review_needed_count", 0), "需要进入 20 条一轮复核")}
      {_metric("可行动", summary.get("actionable_count", 0), "可进报告或行动池")}
      {_metric("活跃销售", snapshot.get("active_sales_count", 0), "归一后的销售对象")}
    </section>
    <section class="panel" style="margin-top:20px">
      <h2>核心业务问题</h2>
      <div>{_breakdown(summary.get("business_question_breakdown", {}), BUSINESS_QUESTION_LABELS)}</div>
    </section>
    <section class="grid cols-2" style="margin-top:20px">
      {''.join(_insight_card(item) for item in insights[:4])}
    </section>
    """


def _trends(snapshot: Dict[str, object], trend_cube: Sequence[Dict[str, object]], summary: Dict[str, object]) -> str:
    monthly = [row for row in trend_cube if str(row.get("grain", "")) == "month"][-12:]
    weekly = [row for row in trend_cube if str(row.get("grain", "")) == "week"][-12:]
    return f"""
    <h1>趋势中心</h1>
    <p class="lead">当前版本先把周/月时间维度显性化，后续继续增强筛选和同比环比交互。</p>
    <section class="grid cols-3">
      {_metric("最新月", snapshot.get("latest_month", ""), "默认月度窗口")}
      {_metric("最新周", snapshot.get("latest_week", ""), "默认周度窗口")}
      {_metric("证据总量", summary.get("total_business_evidence", 0), "业务问题层证据")}
    </section>
    <section class="grid cols-2" style="margin-top:20px">
      <div class="panel"><h2>近 12 个月</h2>{_period_table(monthly)}</div>
      <div class="panel"><h2>近 12 周</h2>{_period_table(weekly)}</div>
    </section>
    <section class="panel" style="margin-top:20px">
      <h2>医生接纳度结构</h2>
      {_breakdown(summary.get("doctor_acceptance_breakdown", {}), DOCTOR_ACCEPTANCE_LABELS)}
    </section>
    """


def _sales(profiles: Sequence[Dict[str, object]], facts: Sequence[Dict[str, object]]) -> str:
    top = sorted(profiles, key=lambda item: -int(item.get("total_mentions", 0)))[:20]
    return f"""
    <h1>销售画像</h1>
    <p class="lead">销售个人是最小动作单元。这里优先展示提及趋势和业务问题贡献，避免只看区域汇总。</p>
    <section class="panel">
      <h2>高频销售 / 医助样例</h2>
      <table class="table"><thead><tr><th>姓名</th><th>战区/区域</th><th>AI证据</th><th>医生反馈</th><th>分层</th></tr></thead><tbody>
      {''.join(f"<tr><td>{_e(row.get('display_name',''))}</td><td>{_e(row.get('battle_zone_name',''))}/{_e(row.get('region_name',''))}</td><td>{row.get('total_mentions',0)}</td><td>{row.get('doctor_feedback_mentions',0)}</td><td>{_e(row.get('segment',''))}</td></tr>" for row in top)}
      </tbody></table>
    </section>
    """


def _insights(insights: Sequence[Dict[str, object]]) -> str:
    return f"""
    <h1>洞察中心</h1>
    <p class="lead">洞察卡必须说明发生了什么、为什么重要、建议怎么处理，并能回到代表证据。</p>
    <section class="grid cols-2">
      {''.join(_insight_card(item) for item in insights)}
    </section>
    """


def _review(tasks: Sequence[Dict[str, object]]) -> str:
    open_tasks = [row for row in tasks if str(row.get("task_status", "")) == "open"][:20]
    if not open_tasks:
        cards = "<div class='panel'><h2>当前没有打开状态的复核任务</h2><p class='muted'>完成新一轮分析后会自动生成 20 条复核卡片。</p></div>"
    else:
        cards = "".join(_review_card(row, idx, len(open_tasks)) for idx, row in enumerate(open_tasks, 1))
    return f"""
    <h1>复核学习工作台</h1>
    <p class="lead">每轮只处理 20 条。你只判断业务价值和结论是否正确；错因、规则候选、Prompt 候选由系统自动生成。</p>
    {cards}
    """


def _evidence(facts: Sequence[Dict[str, object]]) -> str:
    rows = list(facts)[:200]
    return f"""
    <h1>证据下钻</h1>
    <p class="lead">所有趋势和洞察都必须能回到原文证据和文件路径。</p>
    <section class="panel">
      <table class="table"><thead><tr><th>业务问题</th><th>销售/区域</th><th>结论字段</th><th>原文</th></tr></thead><tbody>
      {''.join(_evidence_row(row) for row in rows)}
      </tbody></table>
    </section>
    """


def _metric(label: str, value: object, note: str) -> str:
    return f"<div class='panel'><div class='metric'>{_e(value)}</div><div class='label'>{_e(label)}</div><p class='muted'>{_e(note)}</p></div>"


def _breakdown(payload: object, labels: Dict[str, str]) -> str:
    if not isinstance(payload, dict) or not payload:
        return "<p class='muted'>暂无数据</p>"
    return "".join(f"<span class='tag'>{_e(labels.get(str(key), str(key)))} · {_e(value)}</span>" for key, value in payload.items() if int(value or 0) > 0)


def _period_table(rows: Sequence[Dict[str, object]]) -> str:
    if not rows:
        return "<p class='muted'>暂无时间序列数据</p>"
    body = "".join(
        f"<tr><td>{row.get('year','')}-{int(row.get('month',0)):02d}{'-W'+str(row.get('week_of_month')) if row.get('grain') == 'week' else ''}</td><td>{row.get('ai_mentions',0)}</td><td>{row.get('active_sales_count',0)}</td><td>{round(float(row.get('sales_penetration_rate',0))*100,1)}%</td></tr>"
        for row in rows
    )
    return f"<table class='table'><thead><tr><th>周期</th><th>证据</th><th>活跃销售</th><th>渗透率</th></tr></thead><tbody>{body}</tbody></table>"


def _insight_card(item: Dict[str, object]) -> str:
    return f"""
    <article class="panel">
      <div class="tag">{_e(item.get('confidence_level',''))}</div>
      <div class="card-title">{_e(item.get('title',''))}</div>
      <p>{_e(item.get('judgement',''))}</p>
      <p class="muted">{_e(item.get('why_it_matters',''))}</p>
      <p><strong>建议：</strong>{_e(item.get('action_recommendation',''))}</p>
      <div class="label">证据 {item.get('evidence_count',0)} 条｜待复核 {item.get('review_needed_count',0)} 条</div>
    </article>
    """


def _review_card(row: Dict[str, object], idx: int, total: int) -> str:
    fields = dict(row.get("current_fields", {}))
    task_json = html.escape(json.dumps({"task_id": row.get("task_id", "")}, ensure_ascii=False), quote=True)
    return f"""
    <section class="panel review-card" style="margin-bottom:20px" data-review-card data-task='{task_json}'>
      <div>
        <div class="tag">第 {idx} / {total} 条</div>
        <div class="tag warn">{_e(row.get('selection_reason',''))}</div>
        <h2>{_e(row.get('salesperson_name','未识别'))} · {_e(row.get('battle_zone_name',''))}/{_e(row.get('region_name',''))}</h2>
        <div class="quote">{_e(row.get('source_text',''))}</div>
        <p class="muted">{_e(row.get('file_path',''))}</p>
      </div>
      <div>
        <div class="field"><strong>业务问题：</strong>{_e(row.get('business_question_label',''))}</div>
        <div class="field"><strong>医生接纳度：</strong>{_e(DOCTOR_ACCEPTANCE_LABELS.get(str(fields.get('doctor_acceptance_level','')), str(fields.get('doctor_acceptance_level',''))))}</div>
        <div class="field"><strong>医生诉求：</strong>{_e(DOCTOR_NEED_LABELS.get(str(fields.get('doctor_need_type','')), str(fields.get('doctor_need_type',''))))}</div>
        <div class="field"><strong>行动价值：</strong>{_e(ACTIONABILITY_LABELS.get(str(fields.get('actionability','')), str(fields.get('actionability',''))))}</div>
        <p class="muted">{_e(row.get('review_guidance',''))}</p>
        <div class="form-grid">
          <div class="form-field">
            <label>这条内容有没有业务价值</label>
            {_select("business_value", [("high", "高：值得重点关注"), ("medium", "中：可作为参考"), ("low", "低：噪声或价值很低")])}
          </div>
          <div class="form-field">
            <label>它主要属于哪个业务问题</label>
            {_select("final_business_question", [(key, BUSINESS_QUESTION_LABELS.get(key, key)) for key in BUSINESS_QUESTION_VALUES], str(fields.get("business_question", "")))}
          </div>
          <div class="form-field">
            <label>是否值得进入报告 / 后续行动</label>
            {_select("is_report_worthy", [("yes", "是：可进入报告或行动"), ("observe", "先观察"), ("no", "否：不进入报告")])}
          </div>
          <div class="form-field">
            <label>复核备注</label>
            <textarea name="review_comment" placeholder="只写业务判断即可，例如：这是医生真实顾虑；这是销售内部学习；这是低价值泛 AI。"></textarea>
          </div>
        </div>
        <div class="toolbar">
          <button class="btn" type="button" data-submit-review>提交并进入下一张</button>
          <a class="btn secondary" href="evidence.html">查看证据页</a>
        </div>
        <div class="status-line" data-review-status></div>
      </div>
    </section>
    """


def _select(name: str, options: Sequence[tuple[str, str]], selected: str = "") -> str:
    option_html = ["<option value=''>请选择</option>"]
    for value, label in options:
        selected_attr = " selected" if value == selected else ""
        option_html.append(f"<option value='{_e(value)}'{selected_attr}>{_e(label)}</option>")
    return f"<select name='{_e(name)}'>{''.join(option_html)}</select>"


def _review_script() -> str:
    return """
  <script>
    function nextOpenCard(current) {
      const cards = Array.from(document.querySelectorAll('[data-review-card]')).filter(card => !card.classList.contains('task-hidden'));
      const idx = cards.indexOf(current);
      return cards[idx + 1] || null;
    }

    function setStatus(card, text, isError) {
      const el = card.querySelector('[data-review-status]');
      if (!el) return;
      el.textContent = text;
      el.style.color = isError ? 'var(--risk)' : 'var(--accent)';
    }

    document.querySelectorAll('[data-submit-review]').forEach(button => {
      button.addEventListener('click', async () => {
        const card = button.closest('[data-review-card]');
        const task = JSON.parse(card.dataset.task || '{}');
        const fields = {};
        card.querySelectorAll('select').forEach(select => {
          fields[select.name] = select.value;
        });
        const comment = (card.querySelector('textarea[name="review_comment"]') || {}).value || '';

        if (!fields.business_value || !fields.final_business_question || !fields.is_report_worthy) {
          setStatus(card, '还有必填项未选择：业务价值、业务问题、是否进报告。', true);
          return;
        }
        if (window.location.protocol === 'file:') {
          setStatus(card, '当前是静态文件预览。提交复核请用：python -m src.analysis_v16.webapp --data data/output/insights/v1.6，然后打开 /review。', true);
          return;
        }

        button.disabled = true;
        setStatus(card, '提交中...', false);
        try {
          const response = await fetch('/api/v16-review-decisions', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
              task_id: task.task_id,
              reviewed_fields: fields,
              review_comment: comment
            })
          });
          const payload = await response.json();
          if (!response.ok || !payload.ok) {
            throw new Error(payload.error || '提交失败');
          }
          setStatus(card, '已提交，已写回复核结果。', false);
          card.classList.add('task-hidden');
          const next = nextOpenCard(card);
          if (next) next.scrollIntoView({behavior: 'smooth', block: 'start'});
        } catch (err) {
          button.disabled = false;
          setStatus(card, err.message || String(err), true);
        }
      });
    });
  </script>
    """


def _evidence_row(row: Dict[str, object]) -> str:
    return (
        "<tr>"
        f"<td>{_e(BUSINESS_QUESTION_LABELS.get(str(row.get('business_question','')), str(row.get('business_question',''))))}</td>"
        f"<td>{_e(row.get('salesperson_name',''))}<br><span class='muted'>{_e(row.get('region_name',''))}</span></td>"
        f"<td>{_e(DOCTOR_ACCEPTANCE_LABELS.get(str(row.get('doctor_acceptance_level','')), str(row.get('doctor_acceptance_level',''))))}<br>{_e(ACTIONABILITY_LABELS.get(str(row.get('actionability','')), str(row.get('actionability',''))))}</td>"
        f"<td>{_e(str(row.get('source_text',''))[:260])}</td>"
        "</tr>"
    )


def _e(value: object) -> str:
    return html.escape(str(value))
