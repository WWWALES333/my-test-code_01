from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

from src.analysis_v14.loader import build_report_record
from src.analysis_v14.tagger import Tagger
from src.analysis_v15.review_state import (
    apply_review_decision,
    build_review_learning_candidates,
    load_review_decisions,
    validate_review_submission,
)
from src.analysis_v15.run import run_pipeline


class TestAnalysisV15(unittest.TestCase):
    def test_pipeline_generates_multiview_outputs_with_roster(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)
            roster_path = base / "roster.xlsx"
            _write_roster_xlsx(
                roster_path,
                [
                    ["白兰花", "将军汤一战区", "销售部（将军汤）/销售一战区（粤闽）/广东区域", "广东区域", "销售代表"],
                    ["紫萍", "将军汤四战区", "销售部（将军汤）/销售四战区（湘川贵）/湖南区域", "湖南区域", "销售代表"],
                ],
            )

            (samples / "2026年3月第2周-粤闽区域.txt").write_text(
                "\n".join(
                    [
                        "白兰花：我向医生介绍了AI辅助诊疗，医生觉得很有帮助，后续愿意继续体验。",
                        "海金沙：客户担心AI准确度和替代风险，后续需要继续解释。",
                    ]
                ),
                encoding="utf-8",
            )

            result = run_pipeline(samples, annotations, out, model_mode="mock", llm_chunk_size=7, roster_path=roster_path)
            self.assertEqual(result["reports"], 1)
            self.assertEqual(result["roster_sales"], 2)
            self.assertGreaterEqual(result["profiles"], 3)

            normalized = out / "normalized"
            web = out / "web"
            reports = out / "reports"

            expected_files = [
                normalized / "sales_roster.jsonl",
                normalized / "owner_registry.jsonl",
                normalized / "trend_cube.json",
                normalized / "trend_explanations.json",
                normalized / "salesperson_profile.jsonl",
                normalized / "region_sales_rollup.jsonl",
                normalized / "insight_tree.json",
                normalized / "review_tasks.jsonl",
                normalized / "system_review_tasks.jsonl",
                normalized / "review_batch_summaries.json",
                normalized / "review_audit_log.jsonl",
                normalized / "review_candidates.jsonl",
                normalized / "review_learning_summary.json",
                normalized / "evidence_index.jsonl",
                normalized / "dashboard_snapshot.json",
                web / "overview.html",
                web / "trends.html",
                web / "sales.html",
                web / "insights.html",
                web / "review.html",
                web / "evidence.html",
                web / "AI情报工作台.html",
                reports / "AI情报工作台摘要.md",
                out / "run_manifest.json",
                out / "review" / "review_batch_summaries.json",
                out / "review" / "system_review_tasks.jsonl",
            ]
            for path in expected_files:
                self.assertTrue(path.exists(), path)

            snapshot = json.loads((normalized / "dashboard_snapshot.json").read_text(encoding="utf-8"))
            self.assertEqual(snapshot["roster_active_count"], 2)
            self.assertIn("headline_judgements", snapshot)
            self.assertEqual(snapshot["latest_month"], snapshot["latest_year_month"])
            self.assertEqual(snapshot["latest_week"], snapshot["latest_year_week"])
            self.assertEqual(snapshot["active_salespeople"], snapshot["active_sales_count"])
            self.assertEqual(snapshot["active_individuals"], snapshot["active_person_count"])
            self.assertEqual(snapshot["review_open_count"], snapshot["open_review_tasks"])
            self.assertEqual(snapshot["review_reviewed_count"], snapshot["reviewed_tasks"])
            manifest = json.loads((out / "run_manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["llm_chunk_size"], 7)

            profiles = _read_jsonl(normalized / "salesperson_profile.jsonl")
            segments = {row["display_name"]: row["segment"] for row in profiles}
            self.assertEqual(segments["紫萍"], "长期未提及者")
            self.assertIn("白兰花", segments)

            owner_registry = _read_jsonl(normalized / "owner_registry.jsonl")
            unmatched = {row["salesperson_name"]: row["employment_status"] for row in owner_registry}
            self.assertEqual(unmatched["海金沙"], "historical_unmatched")

            sales_html = (web / "sales.html").read_text(encoding="utf-8")
            trends_html = (web / "trends.html").read_text(encoding="utf-8")
            review_html = (web / "review.html").read_text(encoding="utf-8")
            summary_text = (reports / "AI情报工作台摘要.md").read_text(encoding="utf-8")
            self.assertIn("销售分析中心", sales_html)
            self.assertIn("长期未提及者", sales_html)
            self.assertIn("sales-month-filter", sales_html)
            self.assertIn("trend-primary-month", trends_html)
            self.assertIn("你现在要做什么", review_html)
            self.assertIn("默认批次", review_html)
            self.assertIn("filter-batch-id", review_html)
            self.assertIn("学习候选池", review_html)
            self.assertIn("filter-task-status", review_html)
            self.assertIn("本期核心判断", summary_text)
            self.assertIn("时间范围", summary_text)

            evidence_rows = _read_jsonl(normalized / "evidence_index.jsonl")
            self.assertIn("year", evidence_rows[0])
            self.assertIn("month", evidence_rows[0])
            self.assertIn("week_of_month", evidence_rows[0])

            evidence_facts = _read_jsonl(normalized / "evidence_facts.jsonl")
            self.assertIn("triage_status", evidence_facts[0])
            self.assertIn("used_label_gap", evidence_facts[0])
            self.assertIn("llm_invoked", evidence_facts[0])
            self.assertIn("llm_failed", evidence_facts[0])

            batch_summaries = json.loads((normalized / "review_batch_summaries.json").read_text(encoding="utf-8"))
            review_rows = _read_jsonl(normalized / "review_tasks.jsonl")
            if review_rows:
                self.assertIn("batch_id", review_rows[0])
                self.assertEqual(batch_summaries[0]["batch_number"], 1)

    def test_system_failure_review_tasks_are_split_from_business_queue(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)
            roster_path = base / "roster.xlsx"
            _write_roster_xlsx(
                roster_path,
                [["白兰花", "将军汤一战区", "销售部（将军汤）/销售一战区（粤闽）/广东区域", "广东区域", "销售代表"]],
            )

            (samples / "2026年1月第3周-白兰花.txt").write_text(
                "听了龙胆的分享，收获很多，如今AI高速发展，医生通过平台的AI诊疗可以改变很多。",
                encoding="utf-8",
            )

            with patch.dict("os.environ", {"OPENAI_API_KEY": "", "OPENAI_MODEL": ""}, clear=False):
                run_pipeline(samples, annotations, out, model_mode="real", roster_path=roster_path)

            review_rows = _read_jsonl(out / "normalized" / "review_tasks.jsonl")
            system_rows = _read_jsonl(out / "normalized" / "system_review_tasks.jsonl")
            self.assertFalse(review_rows)
            self.assertEqual(len(system_rows), 1)
            self.assertEqual(system_rows[0]["queue_type"], "system")
            self.assertIn("MODEL_CALL_FAILED", system_rows[0]["review_reason_code"])

    def test_review_decision_is_consumed_and_audit_logged(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)
            roster_path = base / "roster.xlsx"
            _write_roster_xlsx(
                roster_path,
                [["白兰花", "将军汤一战区", "销售部（将军汤）/销售一战区（粤闽）/广东区域", "广东区域", "销售代表"]],
            )

            sample = samples / "2026年3月第2周-白兰花.txt"
            sample.write_text("白兰花：客户提到 AI 很火，但暂时还没有明确使用场景，需要继续观察。", encoding="utf-8")

            report_id = str(build_report_record(sample)["report_id"])
            review_csv = annotations / "review_result.csv"
            with review_csv.open("w", encoding="utf-8", newline="") as file:
                writer = csv.DictWriter(
                    file,
                    fieldnames=[
                        "sample_id",
                        "task_id",
                        "report_id",
                        "segment_id",
                        "salesperson_id",
                        "review_reason_code",
                        "current_labels",
                        "reviewed_fields",
                        "final_labels",
                        "is_pass",
                        "review_comment",
                        "reviewer",
                        "reviewed_at",
                        "error_reason_primary",
                        "review_necessity",
                        "actionability",
                        "action_bucket",
                        "learning_note",
                        "need_rule_update",
                        "need_prompt_update",
                        "need_annotation_update",
                    ],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "sample_id": "",
                        "task_id": "",
                        "report_id": report_id,
                        "segment_id": "S001",
                        "salesperson_id": "",
                        "review_reason_code": "ACTOR_OVERLAP",
                        "current_labels": "",
                        "reviewed_fields": "",
                        "final_labels": json.dumps(
                            {
                                "decision_status": "confirmed",
                                "business_line": "云管家",
                                "actor_primary": "销售对外介绍",
                                "ai_scope": "product_ai",
                            },
                            ensure_ascii=False,
                        ),
                        "is_pass": "1",
                        "review_comment": "已人工确认",
                        "reviewer": "tester",
                        "reviewed_at": "2026-03-31T09:00:00",
                        "error_reason_primary": "actor_boundary",
                        "review_necessity": "should_review",
                        "actionability": "actionable",
                        "action_bucket": "sales_enablement_pool",
                        "learning_note": "该条更适合落到销售对外介绍，不应继续模糊处理。",
                        "need_rule_update": "1",
                        "need_prompt_update": "",
                        "need_annotation_update": "",
                    }
                )

            run_pipeline(samples, annotations, out, model_mode="mock", roster_path=roster_path)

            evidence_rows = _read_jsonl(out / "normalized" / "evidence_facts.jsonl")
            first = evidence_rows[0]
            self.assertEqual(first["business_line"], "云管家")
            self.assertEqual(first["actor_primary"], "销售对外介绍")
            self.assertEqual(first["review_status"], "reviewed")

            review_rows = _read_jsonl(out / "normalized" / "review_tasks.jsonl")
            current_fields = review_rows[0]["current_fields"]
            self.assertIn("business_line", current_fields)
            self.assertIn("actor_primary", current_fields)
            self.assertIn("ai_scope", current_fields)
            self.assertIn("is_ai_hit", current_fields)

            audit_rows = _read_jsonl(out / "normalized" / "review_audit_log.jsonl")
            self.assertEqual(audit_rows[0]["reviewer"], "tester")
            self.assertEqual(audit_rows[0]["learning_fields"]["error_reason_primary"], "actor_boundary")

            candidate_rows = _read_jsonl(out / "review" / "review_candidates.jsonl")
            self.assertEqual(candidate_rows[0]["update_type"], "rule")
            learning_summary = json.loads((out / "review" / "review_learning_summary.json").read_text(encoding="utf-8"))
            self.assertEqual(learning_summary["candidate_count"], 1)

    def test_real_mode_cache_supports_resume(self) -> None:
        class _FakeResponse:
            def raise_for_status(self) -> None:
                return None

            def json(self) -> dict:
                return {
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps(
                                    {
                                        "is_ai_hit": True,
                                        "business_line": "云诊室",
                                        "actor_primary": "label_gap",
                                        "ai_scope": "product_ai",
                                        "decision_status": "confirmed",
                                        "confidence": 0.84,
                                        "reason": "边界样本，现有主体不适用。",
                                        "used_label_gap": True,
                                        "should_review": False,
                                    },
                                    ensure_ascii=False,
                                )
                            }
                        }
                    ]
                }

        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)
            roster_path = base / "roster.xlsx"
            _write_roster_xlsx(
                roster_path,
                [["白兰花", "将军汤一战区", "销售部（将军汤）/销售一战区（粤闽）/广东区域", "广东区域", "销售代表"]],
            )

            (samples / "2026年1月第3周-白兰花.txt").write_text(
                "听了龙胆的分享，收获很多，如今AI高速发展，医生通过平台的AI诊疗可以改变很多。",
                encoding="utf-8",
            )

            with patch.dict("os.environ", {"OPENAI_API_KEY": "k", "OPENAI_MODEL": "m"}, clear=False):
                with patch("src.analysis_v14.tagger.requests.post", return_value=_FakeResponse()) as mock_post:
                    run_pipeline(samples, annotations, out, model_mode="real", llm_concurrency=2, llm_chunk_size=1, roster_path=roster_path)
            self.assertEqual(mock_post.call_count, 1)

            cache_path = out / "runtime" / "real_llm_cache.jsonl"
            self.assertTrue(cache_path.exists())
            progress = json.loads((out / "runtime" / "real_llm_progress.json").read_text(encoding="utf-8"))
            self.assertEqual(progress["pending"], 0)
            self.assertEqual(progress["llm_chunk_size"], 1)

            with patch.dict("os.environ", {"OPENAI_API_KEY": "k", "OPENAI_MODEL": "m"}, clear=False):
                with patch("src.analysis_v14.tagger.requests.post", side_effect=AssertionError("should not be called")) as mock_post_again:
                    run_pipeline(samples, annotations, out, model_mode="real", llm_concurrency=2, llm_chunk_size=1, roster_path=roster_path)
            self.assertEqual(mock_post_again.call_count, 0)

    def test_apply_review_decision_writes_jsonl(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            review_dir = Path(tmp) / "review"
            task = {
                "task_id": "task-1",
                "report_id": "r1",
                "segment_id": "S001",
            }
            apply_review_decision(
                review_dir=review_dir,
                task=task,
                reviewed_fields={
                    "is_ai_hit": True,
                    "business_line": "云诊室",
                    "actor_primary": "医生反馈",
                    "ai_scope": "product_ai",
                    "decision_status": "confirmed",
                },
                reviewer="wales",
                reviewed_at="2026-03-31T10:00:00",
                review_comment="确认",
                learning_fields={
                    "error_reason_primary": "actor_boundary",
                    "review_necessity": "should_review",
                    "actionability": "actionable",
                    "action_bucket": "product_pool",
                    "need_rule_update": True,
                    "need_prompt_update": False,
                    "need_annotation_update": True,
                    "learning_note": "需要沉淀到产品机会池。",
                },
            )
            decisions = load_review_decisions(Path(tmp) / "annotations", review_dir=review_dir)
            decision = decisions[("r1", "S001")]
            self.assertEqual(decision["reviewer"], "wales")
            self.assertEqual(decision["final_labels"]["business_line"], "云诊室")
            self.assertTrue(decision["learning_fields"]["need_rule_update"])
            self.assertEqual(decision["learning_fields"]["action_bucket"], "product_pool")
            self.assertIn("final_labels", decision["change_diff"])
            self.assertIn("learning_fields", decision["change_diff"])

    def test_candidates_can_be_inferred_without_manual_update_flags(self) -> None:
        decisions = {
            ("r1", "S001"): {
                "task_id": "task-1",
                "report_id": "r1",
                "segment_id": "S001",
                "review_reason_code": "ACTOR_OVERLAP",
                "current_labels": {
                    "business_line": "待判断",
                    "actor_primary": "销售对外介绍",
                    "ai_scope": "product_ai",
                },
                "final_labels": {
                    "business_line": "云诊室",
                    "actor_primary": "医生反馈",
                    "ai_scope": "product_ai",
                },
                "learning_fields": {
                    "error_reason_primary": "actor_boundary",
                    "review_necessity": "should_review",
                    "actionability": "actionable",
                    "action_bucket": "product_pool",
                    "need_rule_update": False,
                    "need_prompt_update": False,
                    "need_annotation_update": False,
                    "learning_note": "",
                },
                "reviewed_at": "2026-04-01T10:00:00",
                "source_text": "医生更关心 AI 诊疗结果是否可靠。",
            }
        }
        candidates = build_review_learning_candidates(decisions)
        self.assertEqual(candidates[0]["update_type"], "annotation")

    def test_validate_review_submission_requires_learning_fields(self) -> None:
        task = {
            "current_fields": {
                "is_ai_hit": True,
                "business_line": "待判断",
                "actor_primary": "销售对外介绍",
                "ai_scope": "product_ai",
                "decision_status": "pending_human_review",
            }
        }
        errors = validate_review_submission(
            task,
            reviewed_fields={
                "is_ai_hit": True,
                "business_line": "云诊室",
                "actor_primary": "医生反馈",
                "ai_scope": "product_ai",
                "decision_status": "confirmed",
            },
            learning_fields={
                "review_necessity": "",
                "actionability": "",
            },
        )
        self.assertTrue(errors)
        self.assertIn("review_necessity", errors[0])

    def test_blank_actor_can_be_used_as_label_gap(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            review_dir = Path(tmp) / "review"
            task = {
                "task_id": "task-gap",
                "report_id": "r2",
                "segment_id": "S002",
                "current_fields": {
                    "is_ai_hit": True,
                    "business_line": "云诊室",
                    "actor_primary": "销售对外介绍",
                    "ai_scope": "product_ai",
                    "decision_status": "pending_human_review",
                },
            }
            apply_review_decision(
                review_dir=review_dir,
                task=task,
                reviewed_fields={
                    "is_ai_hit": True,
                    "business_line": "云诊室",
                    "actor_primary": "",
                    "ai_scope": "product_ai",
                    "decision_status": "confirmed",
                },
                reviewer="wales",
                reviewed_at="2026-04-01T12:00:00",
                review_comment="现有主体标签都不合适，先留空。",
                learning_fields={
                    "review_necessity": "should_review",
                    "actionability": "observe",
                    "action_bucket": "",
                    "learning_note": "需要后续扩增主体标签。",
                },
            )
            decisions = load_review_decisions(Path(tmp) / "annotations", review_dir=review_dir)
            decision = decisions[("r2", "S002")]
            self.assertEqual(decision["learning_fields"]["error_reason_primary"], "label_gap")

    def test_noise_patterns_are_not_forced_into_pending_review(self) -> None:
        tagger = Tagger(mode="mock")

        intro = tagger.classify("本周线上微信宣传ai功能为主。")
        self.assertEqual(intro["actor_primary"], "销售对外介绍")
        self.assertEqual(intro["decision_status"], "confirmed")

        trend = tagger.classify("人工智能系统越来越强大了，deep seek 一夜爆火，成了公众的热点话题。")
        self.assertIn(trend["ai_scope"], {"market_trend", "general_ai"})
        self.assertEqual(trend["decision_status"], "confirmed")

        qr = tagger.classify("推荐了我们的ai二维码做个了解。")
        self.assertEqual(qr["actor_primary"], "销售对外介绍")
        self.assertEqual(qr["decision_status"], "confirmed")


def _read_jsonl(path: Path) -> list[dict]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def _write_roster_xlsx(path: Path, rows: list[list[str]]) -> None:
    header = ["花名", "汤名", "组织全称", "部门", "职务"]
    all_rows = [header, *rows]
    shared_strings: list[str] = []
    string_index: dict[str, int] = {}

    def shared_id(value: str) -> int:
        if value not in string_index:
            string_index[value] = len(shared_strings)
            shared_strings.append(value)
        return string_index[value]

    row_xml: list[str] = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            cell_ref = f"{_col_name(c_idx)}{r_idx}"
            cells.append(f'<c r="{cell_ref}" t="s"><v>{shared_id(value)}</v></c>')
        row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
        + "".join(f"<si><t>{escape(value)}</t></si>" for value in shared_strings)
        + "</sst>"
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(row_xml)}</sheetData>"
        "</worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Roster" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    )

    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/sharedStrings.xml", shared_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def _col_name(index: int) -> str:
    result = ""
    value = index
    while value:
        value, remain = divmod(value - 1, 26)
        result = chr(65 + remain) + result
    return result


if __name__ == "__main__":
    unittest.main()
