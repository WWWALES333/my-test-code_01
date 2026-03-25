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


if __name__ == "__main__":
    unittest.main()

