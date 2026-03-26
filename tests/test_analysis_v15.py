from __future__ import annotations

import csv
import json
import tempfile
import unittest
from pathlib import Path

from src.analysis_v14.loader import build_report_record
from src.analysis_v15.owner import infer_owner_record
from src.analysis_v15.parser import segment_text_with_owner
from src.analysis_v15.run import run_pipeline


class TestAnalysisV15(unittest.TestCase):
    def test_pipeline_generates_normalized_outputs_and_workbench(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "2026年3月第2周-白兰花.docx").write_text(
                "我向医生介绍了AI辅助诊疗，医生觉得很有帮助，后续愿意继续体验。",
                encoding="utf-8",
            )
            (samples / "2026年3月第2周-闽鄂区域.txt").write_text(
                "本周继续拜访客户，介绍AI诊疗助手，并记录了几位老师的反馈。",
                encoding="utf-8",
            )

            result = run_pipeline(samples, annotations, out, model_mode="mock")
            self.assertEqual(result["reports"], 2)
            self.assertGreaterEqual(result["insight_cards"], 1)

            normalized = out / "normalized"
            reports = out / "reports"
            review = out / "review"
            web = out / "web"

            self.assertTrue((normalized / "owner_registry.jsonl").exists())
            self.assertTrue((normalized / "report_facts.jsonl").exists())
            self.assertTrue((normalized / "evidence_facts.jsonl").exists())
            self.assertTrue((normalized / "sales_monthly_rollup.jsonl").exists())
            self.assertTrue((normalized / "insight_cards.jsonl").exists())
            self.assertTrue((normalized / "review_tasks.jsonl").exists())
            self.assertTrue((normalized / "review_decisions.jsonl").exists())
            self.assertTrue((normalized / "dashboard_snapshot.json").exists())
            self.assertTrue((reports / "AI情报工作台摘要.md").exists())
            self.assertTrue((review / "review_result_template.csv").exists())
            self.assertTrue((review / "review_result.csv").exists())
            self.assertTrue((web / "AI情报工作台.html").exists())

            summary_text = (reports / "AI情报工作台摘要.md").read_text(encoding="utf-8")
            workbench_html = (web / "AI情报工作台.html").read_text(encoding="utf-8")
            self.assertIn("一句话判断", summary_text)
            self.assertIn("趋势判断", summary_text)
            self.assertIn("结论中心", workbench_html)
            self.assertIn("销售分析中心", workbench_html)

            snapshot = json.loads((normalized / "dashboard_snapshot.json").read_text(encoding="utf-8"))
            self.assertIn("total_ai_mentions", snapshot)
            self.assertIn("active_sales_count", snapshot)

    def test_review_decision_is_consumed_into_evidence_facts(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            sample = samples / "2026年3月第2周-白兰花.txt"
            sample.write_text("我介绍了AI辅助诊疗。", encoding="utf-8")

            report_id = str(build_report_record(sample)["report_id"])
            segment_id = "S001"
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
                        "need_rule_update",
                        "need_skill_update",
                        "need_annotation_update",
                    ],
                )
                writer.writeheader()
                writer.writerow(
                    {
                        "sample_id": "",
                        "task_id": "",
                        "report_id": report_id,
                        "segment_id": segment_id,
                        "salesperson_id": "",
                        "review_reason_code": "ACTOR_OVERLAP",
                        "current_labels": "",
                        "reviewed_fields": "",
                        "final_labels": json.dumps(
                            {
                                "decision_status": "confirmed",
                                "business_line": "云管家",
                                "actor_primary": "销售对外介绍",
                            },
                            ensure_ascii=False,
                        ),
                        "is_pass": "1",
                        "review_comment": "已人工确认",
                        "reviewer": "tester",
                        "reviewed_at": "2026-03-26",
                        "need_rule_update": "",
                        "need_skill_update": "",
                        "need_annotation_update": "",
                    }
                )

            run_pipeline(samples, annotations, out, model_mode="mock")

            evidence_rows = [
                json.loads(line)
                for line in (out / "normalized" / "evidence_facts.jsonl").read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertGreaterEqual(len(evidence_rows), 1)
            first = evidence_rows[0]
            self.assertEqual(first["decision_status"], "confirmed")
            self.assertEqual(first["business_line"], "云管家")
            self.assertEqual(first["actor_primary"], "销售对外介绍")
            self.assertEqual(first["review_status"], "reviewed")

    def test_parser_and_owner_filter_heading_noise(self) -> None:
        text = "\n".join(
            [
                "本周主要工作：继续跟进老客户，整理资料。",
                "白兰花：介绍AI辅助诊疗，医生觉得很有帮助。",
                "工作内容：继续学习话术。",
            ]
        )
        segments = segment_text_with_owner(text, "粤海区域")
        owner_hints = [item["owner_hint"] for item in segments]
        self.assertIn("白兰花", owner_hints)
        self.assertNotIn("本周主要工作", owner_hints)
        self.assertNotIn("工作内容", owner_hints)
        self.assertEqual(infer_owner_record("白兰花", "")["owner_type"], "person")
        self.assertEqual(infer_owner_record("本周主要工作", "")["owner_type"], "group")


if __name__ == "__main__":
    unittest.main()
