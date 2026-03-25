from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from src.analysis_v13.run import run_pipeline


class TestAnalysisV13(unittest.TestCase):
    def test_pipeline_generates_required_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "S001_周报.txt").write_text(
                "本周我使用 DeepSeek 优化话术，并向医生介绍了 AI 辅助诊疗能力。",
                encoding="utf-8",
            )
            (samples / "S002_周报.txt").write_text(
                "本周正常拜访医生，未提到专题内容。",
                encoding="utf-8",
            )

            result = run_pipeline(samples, annotations, out, model_mode="mock")
            self.assertEqual(result["reports"], 2)
            self.assertGreaterEqual(result["tag_rows"], 1)

            extracted = out / "extracted"
            reports = out / "reports"
            self.assertTrue((extracted / "report_index.jsonl").exists())
            self.assertTrue((extracted / "tag_result.jsonl").exists())
            self.assertTrue((extracted / "evidence_span.jsonl").exists())
            self.assertTrue((extracted / "review_queue.jsonl").exists())
            self.assertTrue((reports / "review_queue.csv").exists())
            self.assertTrue((reports / "AI专题摘要.md").exists())

    def test_unsupported_file_goes_to_review_queue(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            # PDF 在 v1.3 首版中不解析，应进入 review_queue
            (samples / "S003_月报.pdf").write_bytes(b"%PDF-1.4 test")
            run_pipeline(samples, annotations, out, model_mode="mock")

            review_path = out / "extracted" / "review_queue.jsonl"
            lines = review_path.read_text(encoding="utf-8").strip().splitlines()
            self.assertGreaterEqual(len(lines), 1)
            first = json.loads(lines[0])
            self.assertEqual(first["decision_status"], "pending_human_review")
            self.assertTrue(first["review_reason_code"])

    def test_scope_and_actor_upgrades_for_key_cases(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            samples = base / "samples"
            annotations = base / "annotations"
            out = base / "out"
            samples.mkdir(parents=True)
            annotations.mkdir(parents=True)

            (samples / "S001_周报.txt").write_text(
                "回访朱老师，固生堂最近在联系老师，有人给AI投喂黑我们的文章。",
                encoding="utf-8",
            )
            (samples / "S002_周报.txt").write_text(
                "江西客户通过AI搜索，AI推荐的甘草云管家，我们向客户介绍了系统能力。",
                encoding="utf-8",
            )
            (samples / "S003_周报.txt").write_text(
                "本周两会强调AI赋能基层医疗，后续要重视合规和服务闭环。",
                encoding="utf-8",
            )

            run_pipeline(samples, annotations, out, model_mode="mock")
            rows = [json.loads(line) for line in (out / "extracted" / "tag_result.jsonl").read_text(encoding="utf-8").splitlines()]
            self.assertEqual(len(rows), 3)
            by_file = {}
            for row in rows:
                name = Path(row["file_path"]).name
                by_file[name] = row

            # 关键用例 1：竞品语境不应误判为我方 product_ai
            first = by_file["S001_周报.txt"]
            self.assertEqual(first["ai_scope"], "competitor_ai")

            # 关键用例 2：客户 AI 搜索语境归 market_trend，并保留销售介绍信号
            second = by_file["S002_周报.txt"]
            self.assertEqual(second["ai_scope"], "market_trend")
            self.assertEqual(second["actor_primary"], "销售对外介绍")

            # 关键用例 3：政策趋势语境优先归 market_trend
            third = by_file["S003_周报.txt"]
            self.assertEqual(third["ai_scope"], "market_trend")
            self.assertIn(third["business_line"], {"混合", "待判断"})

            for row in rows:
                self.assertIn("actor_primary", row)
                self.assertIn("actor_subtype", row)
                self.assertIn("interaction_outcome", row)
                self.assertIn("certainty_level", row)
                self.assertIn("review_reason_code", row)
                self.assertNotIn("/", row["ai_actor"])


if __name__ == "__main__":
    unittest.main()
