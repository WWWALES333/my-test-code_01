from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from src.analysis_v15.run import run_pipeline as run_v15_pipeline

from .business_questions import (
    BusinessQuestionAnalyzer,
    build_business_insights,
    build_executive_brief,
    summarize_business_questions,
)
from .prompt_context import PROMPT_CONTEXT_VERSION, build_prompt_context, render_prompt_reference_markdown
from .reporter import build_weekly_brief, build_workbench_pages, write_json, write_jsonl, write_markdown, write_web_pages
from .review_learning import (
    apply_review_decisions_to_facts,
    build_learning_outputs,
    build_review_batch,
    load_review_decisions,
)
from .time_windows import compute_time_context


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="v1.6 AI 一线情报工作台质量提升版入口")
    parser.add_argument("--samples", required=True, help="样本目录路径，通常为销售周报归档目录")
    parser.add_argument("--annotations", required=True, help="标注 / 复核输入目录")
    parser.add_argument("--out", required=True, help="输出目录")
    parser.add_argument("--model-mode", choices=["mock", "real"], default="mock", help="模型模式")
    parser.add_argument("--llm-concurrency", type=int, default=4, help="v1.5 基础识别 real 模式并发度")
    parser.add_argument("--llm-chunk-size", type=int, default=50, help="v1.5 基础识别 real 模式分块大小")
    parser.add_argument("--insight-mode", choices=["mock", "real"], default="mock", help="洞察归纳模式；real 只对证据簇调用模型")
    parser.add_argument("--business-llm-batch-size", type=int, default=4, help="v1.6 业务问题边界判定每批样本数")
    parser.add_argument("--business-llm-concurrency", type=int, default=2, help="v1.6 业务问题边界判定并发度")
    parser.add_argument("--roster", default="", help="花名册 Excel 路径；为空时沿用 v1.5 自动发现")
    parser.add_argument("--review-batch-size", type=int, default=20, help="每轮主动复核样本数量")
    parser.add_argument("--skip-base", action="store_true", help="跳过 v1.5 基础链路，直接复用当前 out 下 normalized 产物")
    return parser.parse_args()


def run_pipeline(
    samples_dir: Path,
    annotations_dir: Path,
    out_dir: Path,
    model_mode: str,
    llm_concurrency: int = 4,
    llm_chunk_size: int = 50,
    insight_mode: str = "mock",
    business_llm_batch_size: int = 4,
    business_llm_concurrency: int = 2,
    roster_path: Path | None = None,
    review_batch_size: int = 20,
    skip_base: bool = False,
) -> Dict[str, object]:
    annotations_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not skip_base:
        run_v15_pipeline(
            samples_dir=samples_dir,
            annotations_dir=annotations_dir,
            out_dir=out_dir,
            model_mode=model_mode,
            llm_concurrency=llm_concurrency,
            llm_chunk_size=llm_chunk_size,
            roster_path=roster_path,
        )

    normalized_dir = out_dir / "normalized"
    review_dir = out_dir / "review"
    reports_dir = out_dir / "reports"
    web_dir = out_dir / "web"
    evaluation_dir = out_dir / "evaluation"
    for path in (normalized_dir, review_dir, reports_dir, web_dir, evaluation_dir):
        path.mkdir(parents=True, exist_ok=True)

    evidence_facts = _read_jsonl(normalized_dir / "evidence_facts.jsonl")
    trend_cube = _read_json(normalized_dir / "trend_cube.json", [])
    salesperson_profiles = _read_jsonl(normalized_dir / "salesperson_profile.jsonl")
    dashboard_snapshot = _read_json(normalized_dir / "dashboard_snapshot.json", {})
    time_context = compute_time_context(trend_cube)
    dashboard_snapshot["time_context"] = time_context
    dashboard_snapshot["target_year_month"] = time_context.get("target_month", "")
    dashboard_snapshot["target_year_week"] = time_context.get("target_week", "")

    analyzer = BusinessQuestionAnalyzer(
        mode=model_mode,
        llm_batch_size=business_llm_batch_size,
        llm_concurrency=business_llm_concurrency,
    )
    raw_business_facts = analyzer.analyze_batch(evidence_facts)
    review_decisions = load_review_decisions(review_dir / "review_decisions.jsonl")
    business_facts = apply_review_decisions_to_facts(raw_business_facts, review_decisions)
    business_summary = summarize_business_questions(business_facts)
    business_insights = build_business_insights(business_facts, mode=insight_mode)
    executive_brief = build_executive_brief(business_insights, business_summary)

    review_batch = build_review_batch(business_facts, review_decisions, batch_size=review_batch_size)
    learning_summary, rule_candidates, prompt_candidates, label_candidates, golden_set = build_learning_outputs(review_decisions, review_batch)

    write_jsonl(normalized_dir / "business_question_facts.jsonl", business_facts)
    write_json(normalized_dir / "business_question_summary.json", business_summary)
    write_jsonl(normalized_dir / "business_insights.jsonl", business_insights)
    write_json(normalized_dir / "executive_brief.json", executive_brief)
    write_json(normalized_dir / "time_context.json", time_context)
    write_json(normalized_dir / "prompt_context.json", build_prompt_context())
    write_json(normalized_dir / "dashboard_snapshot.json", dashboard_snapshot)

    write_jsonl(review_dir / "review_batch.jsonl", review_batch)
    write_markdown(review_dir / "learning_summary.md", learning_summary)
    write_jsonl(review_dir / "rule_candidates.jsonl", rule_candidates)
    write_jsonl(review_dir / "prompt_candidates.jsonl", prompt_candidates)
    write_jsonl(review_dir / "label_gap_candidates.jsonl", label_candidates)
    write_jsonl(review_dir / "golden_set.jsonl", golden_set)

    weekly_brief = build_weekly_brief(executive_brief, business_summary, business_insights, review_batch, dashboard_snapshot)
    write_markdown(reports_dir / "AI一线情报周报.md", weekly_brief)
    write_markdown(reports_dir / "AI一线情报月报.md", weekly_brief.replace("周度摘要", "月度复盘"))
    write_markdown(reports_dir / "当前使用Prompt说明.md", render_prompt_reference_markdown())

    write_web_pages(
        web_dir,
        build_workbench_pages(
            dashboard_snapshot=dashboard_snapshot,
            business_summary=business_summary,
            executive_brief=executive_brief,
            business_facts=business_facts,
            business_insights=business_insights,
            review_batch=review_batch,
            trend_cube=trend_cube,
            salesperson_profiles=salesperson_profiles,
        ),
    )

    manifest = {
        "version": "v1.6",
        "samples": str(samples_dir.resolve()),
        "annotations": str(annotations_dir.resolve()),
        "out": str(out_dir.resolve()),
        "model_mode": model_mode,
        "llm_concurrency": llm_concurrency,
        "llm_chunk_size": llm_chunk_size,
        "insight_mode": insight_mode,
        "business_llm_batch_size": business_llm_batch_size,
        "business_llm_concurrency": business_llm_concurrency,
        "review_batch_size": review_batch_size,
        "roster": str(roster_path.resolve()) if roster_path else "",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "minimax_base_url_default": "https://api.minimaxi.com/v1",
        "minimax_model_default": "MiniMax-M2.7",
        "prompt_context_version": PROMPT_CONTEXT_VERSION,
        "target_month": time_context.get("target_month", ""),
        "target_week": time_context.get("target_week", ""),
    }
    write_json(out_dir / "run_manifest.json", manifest)

    return {
        "business_facts": len(business_facts),
        "business_insights": len(business_insights),
        "review_batch_open": sum(1 for row in review_batch if str(row.get("task_status", "")) == "open"),
        "rule_candidates": len(rule_candidates),
        "prompt_candidates": len(prompt_candidates),
        "label_candidates": len(label_candidates),
        "golden_set": len(golden_set),
        "out": str(out_dir),
    }


def _read_jsonl(path: Path) -> List[Dict[str, object]]:
    if not path.exists():
        return []
    rows: List[Dict[str, object]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            rows.append(json.loads(line))
    return rows


def _read_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> None:
    args = parse_args()
    result = run_pipeline(
        samples_dir=Path(args.samples),
        annotations_dir=Path(args.annotations),
        out_dir=Path(args.out),
        model_mode=args.model_mode,
        llm_concurrency=args.llm_concurrency,
        llm_chunk_size=args.llm_chunk_size,
        insight_mode=args.insight_mode,
        business_llm_batch_size=args.business_llm_batch_size,
        business_llm_concurrency=args.business_llm_concurrency,
        roster_path=Path(args.roster) if args.roster else None,
        review_batch_size=args.review_batch_size,
        skip_base=args.skip_base,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
