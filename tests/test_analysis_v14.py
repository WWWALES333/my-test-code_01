from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from src.analysis_v14.run import run_pipeline
from src.analysis_v14.tagger import Tagger


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
            self.assertIn("triage_status", first_tag)
            self.assertIn("used_label_gap", first_tag)
            self.assertIn("llm_invoked", first_tag)
            self.assertIn("llm_failed", first_tag)
            self.assertIn("rule_baseline", first_tag)

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
                "听了龙胆的分享，收获很多，如今AI高速发展，医生通过平台的AI诊疗可以改变很多。",
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
            self.assertEqual(first["triage_status"], "needs_llm")
            self.assertTrue(first["llm_invoked"])
            self.assertTrue(first["llm_failed"])

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

    def test_obvious_samples_do_not_invoke_llm_in_real_mode(self) -> None:
        tagger = Tagger(mode="real")
        with patch.dict("os.environ", {"OPENAI_API_KEY": "k", "OPENAI_MODEL": "m"}, clear=False):
            with patch("src.analysis_v14.tagger.requests.post") as mock_post:
                result = tagger.classify(
                    "回访王老师，演示平台AI诊疗助手，并加微信发送体验链接。",
                    context={"file_path": "2026年3月第2周/周报.txt"},
                )
        self.assertEqual(result["triage_status"], "auto_confirm")
        self.assertFalse(result["llm_invoked"])
        self.assertFalse(result["llm_failed"])
        mock_post.assert_not_called()

    def test_obvious_doctor_feedback_cloud_clinic_sample_is_auto_confirmed(self) -> None:
        tagger = Tagger(mode="mock")
        result = tagger.classify(
            "回访温老师，介绍了平台的AI诊疗功能和诊后随访管理功能，老师觉得很不错，后面体验一下。",
            context={"file_path": "2026年3月第2周/周报.txt"},
        )
        self.assertEqual(result["business_line"], "云诊室")
        self.assertEqual(result["actor_primary"], "医生反馈")
        self.assertEqual(result["decision_status"], "confirmed")
        self.assertEqual(result["triage_status"], "auto_confirm")

    def test_ambiguous_samples_invoke_llm_and_support_label_gap(self) -> None:
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
                                        "reason": "该片段描述的是销售内部学习后的业务理解，不适合现有主体标签。",
                                        "used_label_gap": True,
                                        "should_review": False,
                                    },
                                    ensure_ascii=False,
                                )
                            }
                        }
                    ]
                }

        tagger = Tagger(mode="real")
        with patch.dict("os.environ", {"OPENAI_API_KEY": "k", "OPENAI_MODEL": "m"}, clear=False):
            with patch("src.analysis_v14.tagger.requests.post", return_value=_FakeResponse()) as mock_post:
                result = tagger.classify(
                    "听了龙胆的分享，收获很多，如今AI高速发展，医生通过平台的AI诊疗可以改变很多。",
                    context={"file_path": "2026年1月第3周/周报.txt"},
                )
        self.assertEqual(result["triage_status"], "needs_llm")
        self.assertTrue(result["llm_invoked"])
        self.assertFalse(result["llm_failed"])
        self.assertEqual(result["business_line"], "云诊室")
        self.assertEqual(result["actor_primary"], "label_gap")
        self.assertTrue(result["used_label_gap"])
        mock_post.assert_called_once()

    def test_classify_batch_only_invokes_llm_for_needs_llm(self) -> None:
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
                                        "confidence": 0.8,
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

        tagger = Tagger(mode="real")
        items = [
            ("回访王老师，演示平台AI诊疗助手，并加微信发送体验链接。", {"file_path": "2026年3月第2周/周报.txt"}),
            ("听了龙胆的分享，收获很多，如今AI高速发展，医生通过平台的AI诊疗可以改变很多。", {"file_path": "2026年1月第3周/周报.txt"}),
        ]
        with patch.dict("os.environ", {"OPENAI_API_KEY": "k", "OPENAI_MODEL": "m"}, clear=False):
            with patch("src.analysis_v14.tagger.requests.post", return_value=_FakeResponse()) as mock_post:
                results = tagger.classify_batch(items, llm_concurrency=4)

        self.assertEqual(len(results), 2)
        self.assertEqual(results[0]["triage_status"], "auto_confirm")
        self.assertFalse(results[0]["llm_invoked"])
        self.assertEqual(results[1]["triage_status"], "needs_llm")
        self.assertTrue(results[1]["llm_invoked"])
        self.assertEqual(results[1]["actor_primary"], "label_gap")
        mock_post.assert_called_once()


if __name__ == "__main__":
    unittest.main()
