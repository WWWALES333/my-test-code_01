from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path

from src.analysis_v16.business_questions import BusinessQuestionAnalyzer, build_business_insights, build_executive_brief
from src.analysis_v16.model_adapter import parse_json_payload, strip_think_blocks
from src.analysis_v16.prompt_context import PROMPT_CONTEXT_VERSION, build_prompt_context, render_prompt_reference_markdown
from src.analysis_v16.review_learning import (
    apply_review_decision,
    apply_review_decisions_to_facts,
    build_review_feedback,
    build_learning_outputs,
    build_review_batch,
    load_review_decisions,
    validate_review_payload,
)
from src.analysis_v16.run import run_pipeline
from src.analysis_v16.schema import (
    BUSINESS_QUESTION_COMPETITOR_AI,
    BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
    BUSINESS_QUESTION_SALES_AI_USAGE,
)
from src.analysis_v16.time_windows import compute_time_context


class TestAnalysisV16(unittest.TestCase):
    def test_model_adapter_strips_think_and_extracts_json(self) -> None:
        payload = '<think>先分析但不应进入结果</think>\n说明文字 {"answer": "ok", "score": 0.9}'
        self.assertEqual(strip_think_blocks(payload).strip(), '说明文字 {"answer": "ok", "score": 0.9}')
        self.assertEqual(parse_json_payload(payload), {"answer": "ok", "score": 0.9})

    def test_prompt_context_contains_business_background(self) -> None:
        context = build_prompt_context()
        serialized = json.dumps(context, ensure_ascii=False)
        self.assertEqual(context["context_version"], PROMPT_CONTEXT_VERSION)
        self.assertIn("云诊室", serialized)
        self.assertIn("云管家", serialized)
        self.assertIn("销售", serialized)
        self.assertIn("医生反馈", serialized)
        prompt_doc = render_prompt_reference_markdown()
        self.assertIn(PROMPT_CONTEXT_VERSION, prompt_doc)

    def test_time_context_uses_system_date_target_not_latest_data(self) -> None:
        trend_cube = [
            {"grain": "month", "year": 2026, "month": 3, "ai_mentions": 10, "active_sales_count": 3},
            {"grain": "week", "year": 2026, "month": 3, "week_of_month": 4, "ai_mentions": 5, "active_sales_count": 2},
        ]
        context = compute_time_context(trend_cube, today=date(2026, 5, 18))
        self.assertEqual(context["target_month"], "2026-04")
        self.assertEqual(context["target_week"], "2026-05-W2")
        self.assertEqual(context["latest_available_month"], "2026-03")
        self.assertFalse(context["month_observation"]["available"])
        self.assertIn("2026-04", context["month_observation"]["status_note"])

    def test_llm_prompt_contains_business_context_and_role_fields(self) -> None:
        client = _FakeModelClient(
            {
                "business_question": BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
                "doctor_acceptance_level": "interest_exploration",
                "doctor_need_type": "unknown",
                "sales_ai_usage_type": "not_applicable",
                "competitor_signal_type": "not_applicable",
                "speaker_role": "doctor_or_clinic_user",
                "business_actor": "doctor",
                "evidence_type": "doctor_feedback",
                "actionability": "report_ready",
                "confidence": 0.77,
                "should_review": False,
                "reason": "医生在反馈云诊室 AI 体验。",
                "reasoning_summary": "医生表达了对平台 AI 的兴趣，属于医生反馈。",
            }
        )
        analyzer = BusinessQuestionAnalyzer(mode="real", model_client=client)
        result = analyzer.analyze(
            {
                "evidence_id": "E101",
                "source_text": "医生想了解平台AI诊疗助手是否能帮助提升问诊效率。",
                "actor_primary": "待判断",
                "ai_scope": "product_ai",
                "business_line": "待判断",
                "decision_status": "pending_human_review",
            }
        )
        prompt_text = json.dumps(client.last_messages, ensure_ascii=False)
        self.assertIn("project_business_context", prompt_text)
        self.assertIn("云诊室", prompt_text)
        self.assertIn("speaker_role", prompt_text)
        self.assertEqual(result["speaker_role"], "doctor_or_clinic_user")
        self.assertEqual(result["evidence_type"], "doctor_feedback")

    def test_business_question_rules_cover_core_scenarios(self) -> None:
        analyzer = BusinessQuestionAnalyzer(mode="mock")

        doctor = analyzer.analyze(
            {
                "evidence_id": "E001",
                "source_text": "医生认可平台AI诊疗助手，觉得对辨证准确度和效率都有帮助。",
                "actor_primary": "医生反馈",
                "ai_scope": "product_ai",
                "business_line": "云诊室",
                "decision_status": "confirmed",
            }
        )
        self.assertEqual(doctor["business_question"], BUSINESS_QUESTION_DOCTOR_ACCEPTANCE)
        self.assertEqual(doctor["doctor_acceptance_level"], "positive_acceptance")

        sales = analyzer.analyze(
            {
                "evidence_id": "E002",
                "source_text": "销售复盘后用AI整理话术，并计划下周对外介绍给医生。",
                "actor_primary": "销售自用",
                "ai_scope": "product_ai",
                "business_line": "云诊室",
                "decision_status": "confirmed",
            }
        )
        self.assertEqual(sales["business_question"], BUSINESS_QUESTION_SALES_AI_USAGE)
        self.assertEqual(sales["sales_ai_usage_type"], "external_pitch")

        competitor = analyzer.analyze(
            {
                "evidence_id": "E003",
                "source_text": "客户提到竞品已经上线AI问诊功能，需要单独关注。",
                "actor_primary": "潜在AI机会",
                "ai_scope": "competitor_ai",
                "business_line": "云诊室",
                "decision_status": "confirmed",
            }
        )
        self.assertEqual(competitor["business_question"], BUSINESS_QUESTION_COMPETITOR_AI)
        self.assertEqual(competitor["competitor_signal_type"], "competitor_product")

    def test_review_batch_is_twenty_item_active_learning_loop(self) -> None:
        facts = [_business_fact(idx) for idx in range(25)]
        batch = build_review_batch(facts, {}, batch_size=20)
        open_items = [row for row in batch if row["task_status"] == "open"]
        self.assertEqual(len(open_items), 20)
        self.assertEqual(open_items[0]["batch_position"], 1)
        self.assertEqual(open_items[-1]["batch_size"], 20)

    def test_insights_are_business_readable_not_enum_templates(self) -> None:
        facts = [
            {
                **_business_fact(1),
                "doctor_acceptance_level": "positive_acceptance",
                "doctor_need_type": "diagnosis_quality",
                "confidence": 0.86,
                "should_review": False,
                "actionability": "report_ready",
                "source_text": "医生认可平台AI诊疗助手，觉得对辨证准确度和效率都有帮助。",
            },
            {
                **_business_fact(2),
                "doctor_acceptance_level": "explicit_concern",
                "doctor_need_type": "trust_and_safety",
                "confidence": 0.82,
                "should_review": False,
                "actionability": "report_ready",
                "source_text": "医生担心AI辨证不准，后续希望看到更多专业依据。",
            },
        ]
        insights = build_business_insights(facts)
        brief = build_executive_brief(insights, {"total_business_evidence": 2, "review_needed_count": 0})

        self.assertEqual(len(insights), 1)
        first = insights[0]
        self.assertIn("医生", first["conclusion"])
        self.assertIn("正向接受", json.dumps(first, ensure_ascii=False))
        self.assertNotIn("positive_acceptance", first["conclusion"])
        self.assertNotIn("unknown", first["conclusion"])
        self.assertTrue(first["representative_quotes"])
        self.assertIn("医生", brief["doctor_acceptance_answer"])

    def test_review_decision_writes_learning_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            review_dir = Path(tmp)
            task = build_review_batch([_business_fact(1)], {}, batch_size=20)[0]
            errors = validate_review_payload(
                task,
                {
                    "business_value": "high",
                    "final_business_question": BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
                    "is_report_worthy": "yes",
                },
            )
            self.assertEqual(errors, [])
            apply_review_decision(
                review_dir,
                task,
                {
                    "business_value": "high",
                    "final_business_question": BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
                    "is_report_worthy": "yes",
                },
                reviewer="tester",
                reviewed_at="2026-05-17T10:00:00",
                review_comment="医生正向反馈，可进报告。",
            )
            decisions = load_review_decisions(review_dir / "review_decisions.jsonl")
            self.assertEqual(len(decisions), 1)
            summary, rule_candidates, prompt_candidates, label_candidates, golden_set = build_learning_outputs(decisions, [task])
            feedback = build_review_feedback(
                next(iter(decisions.values())),
                decisions,
                rule_candidates=rule_candidates,
                prompt_candidates=prompt_candidates,
                label_candidates=label_candidates,
                golden_set=golden_set,
            )
            self.assertIn("已复核样本数：1", summary)
            self.assertIn("复核 1 条即可写回生效", summary)
            self.assertIn("立即生效", feedback["saved_message"])
            self.assertEqual(feedback["same_reason_count"], 1)
            self.assertEqual(rule_candidates, [])
            self.assertEqual(label_candidates, [])
            self.assertEqual(len(golden_set), 1)
            self.assertTrue(prompt_candidates or "other" in next(iter(decisions.values()))["system_inferred_error_reason"])

            updated = apply_review_decisions_to_facts([_business_fact(1)], decisions)
            self.assertEqual(updated[0]["review_status"], "reviewed")
            self.assertFalse(updated[0]["should_review"])
            self.assertEqual(updated[0]["actionability"], "report_ready")

    def test_pipeline_can_reuse_existing_v15_normalized_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            out = base / "out"
            normalized = out / "normalized"
            normalized.mkdir(parents=True)
            (normalized / "evidence_facts.jsonl").write_text(
                json.dumps(_evidence_fact(), ensure_ascii=False) + "\n",
                encoding="utf-8",
            )
            (normalized / "trend_cube.json").write_text(
                json.dumps(
                    [
                        {"grain": "month", "year": 2026, "month": 3, "ai_mentions": 1, "active_sales_count": 1, "sales_penetration_rate": 0.1},
                        {"grain": "week", "year": 2026, "month": 3, "week_of_month": 4, "ai_mentions": 1, "active_sales_count": 1, "sales_penetration_rate": 0.1},
                    ],
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            (normalized / "salesperson_profile.jsonl").write_text(
                json.dumps({"display_name": "白兰花", "battle_zone_name": "一战区", "region_name": "广东区域", "total_mentions": 1, "doctor_feedback_mentions": 1, "segment": "偶发使用者"}, ensure_ascii=False) + "\n",
                encoding="utf-8",
            )
            (normalized / "dashboard_snapshot.json").write_text(
                json.dumps({"active_sales_count": 1, "latest_month": "2026-03", "latest_week": "2026-03-W4"}, ensure_ascii=False),
                encoding="utf-8",
            )

            result = run_pipeline(
                samples_dir=base / "samples",
                annotations_dir=base / "annotations",
                out_dir=out,
                model_mode="mock",
                review_batch_size=20,
                skip_base=True,
            )

            self.assertEqual(result["business_facts"], 1)
            self.assertTrue((normalized / "business_question_facts.jsonl").exists())
            self.assertTrue((out / "review" / "review_batch.jsonl").exists())
            self.assertTrue((out / "reports" / "AI一线情报周报.md").exists())
            overview_html = (out / "web" / "overview.html").read_text(encoding="utf-8")
            insights_html = (out / "web" / "insights.html").read_text(encoding="utf-8")
            self.assertIn("5 个必须回答的问题", overview_html)
            self.assertIn("2026-04", overview_html)
            self.assertIn("<svg", overview_html)
            self.assertIn("为什么重要", insights_html)
            self.assertIn("产品含义", insights_html)
            self.assertIn("销售管理含义", insights_html)
            self.assertIn("代表原文", insights_html)
            self.assertNotIn("external_pitch", overview_html)
            self.assertTrue((normalized / "prompt_context.json").exists())
            self.assertTrue((normalized / "time_context.json").exists())
            self.assertTrue((out / "reports" / "当前使用Prompt说明.md").exists())
            review_html = (out / "web" / "review.html").read_text(encoding="utf-8")
            self.assertIn("/api/v16-review-decisions", review_html)
            self.assertIn("提交并进入下一张", review_html)
            self.assertIn("复核 1 条也会立即写回生效", review_html)


def _business_fact(idx: int) -> dict[str, object]:
    return {
        "business_fact_id": f"BF{idx:03d}",
        "evidence_id": f"E{idx:03d}",
        "report_id": f"R{idx:03d}",
        "segment_id": "S001",
        "business_question": BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
        "business_question_label": "医生 AI 接纳度",
        "doctor_acceptance_level": "unknown",
        "doctor_need_type": "unknown",
        "sales_ai_usage_type": "not_applicable",
        "competitor_signal_type": "not_applicable",
        "actionability": "report_ready" if idx % 2 == 0 else "observe",
        "confidence": 0.35 if idx % 3 == 0 else 0.72,
        "should_review": idx % 3 == 0,
        "speaker_role": "doctor_or_clinic_user",
        "business_actor": "doctor",
        "evidence_type": "doctor_feedback",
        "reasoning_summary": "测试样本。",
        "source_text": f"医生反馈样本 {idx}，提到AI诊疗助手需要进一步确认真实态度。",
        "file_path": f"/tmp/report-{idx}.docx",
        "year": 2026,
        "month": 3,
        "week_of_month": idx % 5 + 1,
        "salesperson_name": "白兰花",
        "battle_zone_name": "一战区",
        "region_name": "广东区域",
    }


class _FakeModelClient:
    def __init__(self, response: dict[str, object]) -> None:
        self.response = response
        self.last_messages = None

    def chat_json(self, messages, max_tokens=900):  # noqa: ANN001, ANN201
        self.last_messages = messages
        return self.response


def _evidence_fact() -> dict[str, object]:
    return {
        "evidence_id": "E001",
        "report_id": "R001",
        "segment_id": "S001",
        "source_text": "白兰花反馈，医生认可平台AI诊疗助手，觉得对辨证准确度和效率都有帮助。",
        "file_path": "/tmp/2026年3月第4周-白兰花.docx",
        "year": 2026,
        "month": 3,
        "week_of_month": 4,
        "salesperson_id": "SP001",
        "salesperson_name": "白兰花",
        "battle_zone_name": "一战区",
        "region_name": "广东区域",
        "business_line": "云诊室",
        "actor_primary": "医生反馈",
        "ai_scope": "product_ai",
        "decision_status": "confirmed",
        "review_status": "",
    }


if __name__ == "__main__":
    unittest.main()
