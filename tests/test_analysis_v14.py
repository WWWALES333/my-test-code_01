from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src.analysis_v14.run import run_pipeline


class TestAnalysisV14(unittest.TestCase):
    def test_pipeline_generates_required_outputs_and_contract_fields(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "S001_2026年3月第2周周报.txt").write_text(
                "本周我使用 DeepSeek 整理回访记录，并向医生介绍了 AI 辅助方案。",
                encoding="utf-8",
            )
            (samples / "S002_2026年3月第2周周报.txt").write_text(
                "本周常规拜访，无 AI 相关表达。",
                encoding="utf-8",
            )

            result = run_pipeline(samples, annotations, out, model_mode="mock")
            self.assertEqual(result["reports"], 2)
            self.assertTrue(result["run_id"])

            extracted = out / "extracted"
            reports = out / "reports"
            review = out / "review"
            self.assertTrue((extracted / "report_index.jsonl").exists())
            self.assertTrue((extracted / "tag_result.jsonl").exists())
            self.assertTrue((extracted / "evidence_span.jsonl").exists())
            self.assertTrue((extracted / "review_queue.jsonl").exists())
            self.assertTrue((reports / "review_queue.csv").exists())
            self.assertTrue((reports / "AI专题摘要.md").exists())
            self.assertTrue((reports / "dashboard_weekly.csv").exists())
            self.assertTrue((reports / "dashboard_monthly.csv").exists())
            self.assertTrue((reports / "opportunity_backlog.csv").exists())
            self.assertTrue((reports / "evidence_trace.csv").exists())
            self.assertTrue((reports / "AI专题看板.html").exists())
            self.assertTrue((review / "review_queue.csv").exists())
            self.assertTrue((review / "review_result_template.csv").exists())
            summary_text = (reports / "AI专题摘要.md").read_text(encoding="utf-8")
            dashboard_html = (reports / "AI专题看板.html").read_text(encoding="utf-8")
            self.assertIn("一页结论（先给业务看）", summary_text)
            self.assertIn("现状如何", summary_text)
            self.assertIn("趋势如何", summary_text)
            self.assertIn("可反哺业务的机会点", summary_text)
            self.assertIn("给产品负责人的重点", summary_text)
            self.assertIn("给销售管理者的重点", summary_text)
            self.assertIn("系统自测结果（本轮自动验收）", summary_text)
            self.assertIn("AI专题业务看板", dashboard_html)

            report_lines = (extracted / "report_index.jsonl").read_text(encoding="utf-8").strip().splitlines()
            report_row = json.loads(report_lines[0])
            self.assertIn("run_id", report_row)
            self.assertIn("parse_status", report_row)
            self.assertIn("parse_reason_code", report_row)

            tag_lines = (extracted / "tag_result.jsonl").read_text(encoding="utf-8").strip().splitlines()
            self.assertGreaterEqual(len(tag_lines), 1)
            first_tag = json.loads(tag_lines[0])
            self.assertIn("model_mode", first_tag)
            self.assertIn("model_name", first_tag)
            self.assertIn("run_id", first_tag)
            self.assertIn("parse_status", first_tag)

    def test_doc_parse_failure_is_recorded_with_reason_code(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "S003_2026年3月第2周周报.doc").write_bytes(b"not-a-real-doc")
            run_pipeline(samples, annotations, out, model_mode="mock")

            report_rows = [
                json.loads(line)
                for line in (out / "extracted" / "report_index.jsonl").read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertEqual(len(report_rows), 1)
            self.assertIn(report_rows[0]["parse_status"], {"success", "failed"})
            if report_rows[0]["parse_status"] == "failed":
                self.assertIn(
                    report_rows[0]["parse_reason_code"],
                    {"PARSE_FAILED_DOC", "PARSER_TOOL_MISSING"},
                )

                review_rows = [
                    json.loads(line)
                    for line in (out / "extracted" / "review_queue.jsonl").read_text(encoding="utf-8").splitlines()
                    if line.strip()
                ]
                self.assertGreaterEqual(len(review_rows), 1)
                self.assertIn(
                    review_rows[0]["review_reason_code"],
                    {"PARSE_FAILED_DOC", "PARSER_TOOL_MISSING"},
                )
                self.assertEqual(review_rows[0]["decision_status"], "pending_human_review")

    def test_real_mode_model_failure_goes_pending_human_review(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "S004_2026年3月第2周周报.txt").write_text(
                "我今天使用 ChatGPT 总结问诊沟通重点，并对医生做了演示。",
                encoding="utf-8",
            )

            with patch.dict("os.environ", {"OPENAI_API_KEY": "", "OPENAI_MODEL": ""}, clear=False):
                run_pipeline(samples, annotations, out, model_mode="real")

            tag_rows = [
                json.loads(line)
                for line in (out / "extracted" / "tag_result.jsonl").read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertGreaterEqual(len(tag_rows), 1)
            first = tag_rows[0]
            self.assertEqual(first["model_mode"], "real")
            self.assertEqual(first["decision_status"], "pending_human_review")
            self.assertIn("MODEL_CALL_FAILED", first["review_reason_code"])

    def test_pdf_parse_failure_records_reason(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "S005_2026年3月月报.pdf").write_bytes(b"%PDF-1.4 invalid")
            run_pipeline(samples, annotations, out, model_mode="mock")

            report_rows = [
                json.loads(line)
                for line in (out / "extracted" / "report_index.jsonl").read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertEqual(len(report_rows), 1)
            self.assertEqual(report_rows[0]["parse_status"], "failed")
            self.assertIn(
                report_rows[0]["parse_reason_code"],
                {"PARSE_FAILED_PDF", "PARSER_TOOL_MISSING"},
            )


if __name__ == "__main__":
    unittest.main()
