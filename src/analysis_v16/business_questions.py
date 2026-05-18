from __future__ import annotations

import json
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Iterable, List, Sequence

from src.analysis_v14.schema import DECISION_PENDING_HUMAN, TRIAGE_NEEDS_LLM, stable_hash

from .model_adapter import OpenAICompatibleClient
from .prompt_context import (
    BUSINESS_ACTOR_LABELS,
    BUSINESS_ACTOR_VALUES,
    EVIDENCE_TYPE_LABELS,
    EVIDENCE_TYPE_VALUES,
    PROMPT_CONTEXT_VERSION,
    SPEAKER_ROLE_LABELS,
    SPEAKER_ROLE_VALUES,
    build_prompt_context,
)
from .schema import (
    ACTIONABILITY_VALUES,
    ACTIONABILITY_LABELS,
    BUSINESS_QUESTION_COMPETITOR_AI,
    BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
    BUSINESS_QUESTION_DOCTOR_DIRECT_NEED,
    BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY,
    BUSINESS_QUESTION_LABELS,
    BUSINESS_QUESTION_REGIONAL_SALES_DIFF,
    BUSINESS_QUESTION_SALES_AI_USAGE,
    BUSINESS_QUESTION_VALUES,
    COMPETITOR_SIGNAL_VALUES,
    COMPETITOR_SIGNAL_LABELS,
    DOCTOR_ACCEPTANCE_LABELS,
    DOCTOR_ACCEPTANCE_VALUES,
    DOCTOR_NEED_LABELS,
    DOCTOR_NEED_VALUES,
    SALES_AI_USAGE_LABELS,
    SALES_AI_USAGE_VALUES,
)


class BusinessQuestionAnalyzer:
    """把 v1.5 证据进一步映射为 v1.6 业务问题层。"""

    def __init__(
        self,
        mode: str = "mock",
        model_client: OpenAICompatibleClient | None = None,
        *,
        llm_batch_size: int = 4,
        llm_concurrency: int = 2,
    ) -> None:
        if mode not in {"mock", "real"}:
            raise ValueError(f"不支持的业务问题分析模式: {mode}")
        self.mode = mode
        self.model_client = model_client
        self.llm_batch_size = max(1, llm_batch_size)
        self.llm_concurrency = max(1, llm_concurrency)

    def analyze_batch(self, evidence_facts: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
        baselines = [_rule_business_analysis(row) for row in evidence_facts]
        if self.mode != "real":
            return baselines

        needs_llm = [(idx, row, baselines[idx]) for idx, row in enumerate(evidence_facts) if baselines[idx]["needs_llm"]]
        if not needs_llm:
            return baselines

        batches = [needs_llm[start : start + self.llm_batch_size] for start in range(0, len(needs_llm), self.llm_batch_size)]
        if self.model_client is not None or self.llm_concurrency == 1:
            for batch in batches:
                _apply_batch_result(baselines, batch, self.model_client)
            return baselines

        with ThreadPoolExecutor(max_workers=self.llm_concurrency) as executor:
            future_map = {executor.submit(_classify_batch_with_llm, batch, None): batch for batch in batches}
            for future in as_completed(future_map):
                batch = future_map[future]
                try:
                    results = future.result()
                    _merge_batch_payloads(baselines, batch, results)
                except Exception as exc:  # noqa: BLE001 - 批量失败时必须转入复核，不可静默确认
                    _mark_batch_failed(baselines, batch, exc)
        return baselines

    def analyze(self, evidence: Dict[str, object]) -> Dict[str, object]:
        baseline = _rule_business_analysis(evidence)
        if self.mode != "real" or not baseline["needs_llm"]:
            return baseline
        try:
            client = self.model_client or OpenAICompatibleClient()
            refined = client.chat_json(_build_messages(evidence, baseline), max_tokens=900)
            return _merge_llm_result(evidence, baseline, refined)
        except Exception as exc:  # noqa: BLE001 - real 模式必须把模型失败转成复核任务
            failed = dict(baseline)
            failed["llm_invoked"] = True
            failed["llm_failed"] = True
            failed["should_review"] = True
            failed["confidence"] = min(float(failed.get("confidence", 0.5)), 0.35)
            failed["review_reason"] = f"模型业务问题判定失败: {type(exc).__name__}: {str(exc)[:160]}"
            return failed


def summarize_business_questions(rows: Sequence[Dict[str, object]]) -> Dict[str, object]:
    by_question = Counter(str(row.get("business_question", "unknown")) for row in rows)
    by_acceptance = Counter(str(row.get("doctor_acceptance_level", "not_applicable")) for row in rows)
    by_need = Counter(str(row.get("doctor_need_type", "not_applicable")) for row in rows)
    by_sales_usage = Counter(str(row.get("sales_ai_usage_type", "not_applicable")) for row in rows)
    by_competitor = Counter(str(row.get("competitor_signal_type", "not_applicable")) for row in rows)
    actionable = [row for row in rows if str(row.get("actionability", "")) in {"report_ready", "actionable"}]
    review_needed = [row for row in rows if bool(row.get("should_review", False))]
    return {
        "total_business_evidence": len(rows),
        "business_question_breakdown": dict(by_question),
        "doctor_acceptance_breakdown": dict(by_acceptance),
        "doctor_need_breakdown": dict(by_need),
        "sales_ai_usage_breakdown": dict(by_sales_usage),
        "competitor_signal_breakdown": dict(by_competitor),
        "actionable_count": len(actionable),
        "review_needed_count": len(review_needed),
        "top_actionable_evidence_ids": [str(row.get("evidence_id", "")) for row in actionable[:10]],
    }


def build_business_insights(
    rows: Sequence[Dict[str, object]],
    *,
    mode: str = "mock",
    model_client: OpenAICompatibleClient | None = None,
) -> List[Dict[str, object]]:
    grouped: Dict[str, List[Dict[str, object]]] = defaultdict(list)
    for row in rows:
        grouped[str(row.get("business_question", "unknown"))].append(row)

    insights: List[Dict[str, object]] = []
    for question in BUSINESS_QUESTION_VALUES:
        items = grouped.get(question, [])
        if not items:
            continue
        if question not in BUSINESS_QUESTION_VALUES:
            continue
        insight = _build_evidence_cluster_insight(question, items)
        if mode == "real":
            insight = _refine_cluster_insight_with_llm(insight, items, model_client=model_client)
        insights.append(insight)
    return sorted(insights, key=lambda item: int(item.get("display_order", 99)))


def build_executive_brief(insights: Sequence[Dict[str, object]], summary: Dict[str, object]) -> Dict[str, object]:
    insight_map = {str(item.get("business_question", "")): item for item in insights}
    doctor = insight_map.get(BUSINESS_QUESTION_DOCTOR_ACCEPTANCE, {})
    direct = insight_map.get(BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, {})
    indirect = insight_map.get(BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY, {})
    sales = insight_map.get(BUSINESS_QUESTION_SALES_AI_USAGE, {})
    competitor = insight_map.get(BUSINESS_QUESTION_COMPETITOR_AI, {})
    regional = insight_map.get(BUSINESS_QUESTION_REGIONAL_SALES_DIFF, {})
    return {
        "headline": _headline_from_insights(insight_map, summary),
        "doctor_acceptance_answer": doctor.get("conclusion", "医生接纳度证据不足，需要先完成复核。"),
        "standout_answer": regional.get("conclusion", "区域和销售差异仍需结合销售画像继续下钻。"),
        "direct_need_answer": direct.get("conclusion", "医生直接诉求证据不足，需要继续沉淀代表原文。"),
        "indirect_opportunity_answer": indirect.get("conclusion", "医生间接 AI 机会仍需从证据中继续聚类。"),
        "sales_and_competitor_answer": " ".join(
            part
            for part in [
                f"销售线：{sales.get('conclusion', '')}" if sales.get("conclusion") else "",
                f"市场雷达：{competitor.get('conclusion', '')}" if competitor.get("conclusion") else "",
            ]
            if part
        )
        or "销售使用 AI 和竞品动作暂未形成稳定判断。",
        "top_opportunities": _top_actions(insights),
        "top_risks": _top_risks(insights),
    }


def _rule_business_analysis(evidence: Dict[str, object]) -> Dict[str, object]:
    text = str(evidence.get("source_text", ""))
    lower = text.lower()
    actor = str(evidence.get("actor_primary", ""))
    ai_scope = str(evidence.get("ai_scope", ""))
    business_line = str(evidence.get("business_line", ""))
    decision_status = str(evidence.get("decision_status", ""))
    review_status = str(evidence.get("review_status", ""))
    reason_code = str(evidence.get("review_reason_code", ""))

    question = _infer_business_question(text, lower, actor, ai_scope)
    acceptance = _infer_doctor_acceptance(text, lower, question)
    need_type = _infer_doctor_need(text, lower, question)
    sales_usage = _infer_sales_usage(text, lower, actor, question)
    competitor_signal = _infer_competitor_signal(text, lower, ai_scope, question)
    speaker_role = _infer_speaker_role(text, actor, ai_scope, question)
    business_actor = _infer_business_actor(speaker_role, question)
    evidence_type = _infer_evidence_type(question, speaker_role, action_hint=actor)
    actionability = _infer_actionability(question, acceptance, need_type, sales_usage, competitor_signal, decision_status)
    confidence = _infer_business_confidence(
        text=text,
        question=question,
        acceptance=acceptance,
        need_type=need_type,
        decision_status=decision_status,
        review_status=review_status,
        reason_code=reason_code,
    )
    needs_llm = _needs_llm(evidence, question, acceptance, need_type, sales_usage, competitor_signal, confidence)
    fact = {
        "business_fact_id": stable_hash(str(evidence.get("evidence_id", "")), "v16-business"),
        "evidence_id": evidence.get("evidence_id", ""),
        "report_id": evidence.get("report_id", ""),
        "segment_id": evidence.get("segment_id", ""),
        "business_question": question,
        "business_question_label": BUSINESS_QUESTION_LABELS.get(question, question),
        "doctor_acceptance_level": acceptance,
        "doctor_need_type": need_type,
        "sales_ai_usage_type": sales_usage,
        "competitor_signal_type": competitor_signal,
        "speaker_role": speaker_role,
        "speaker_role_label": SPEAKER_ROLE_LABELS.get(speaker_role, speaker_role),
        "business_actor": business_actor,
        "business_actor_label": BUSINESS_ACTOR_LABELS.get(business_actor, business_actor),
        "evidence_type": evidence_type,
        "evidence_type_label": EVIDENCE_TYPE_LABELS.get(evidence_type, evidence_type),
        "reasoning_summary": _rule_reasoning_summary(question, speaker_role, business_actor, evidence_type),
        "actionability": actionability,
        "confidence": confidence,
        "should_review": bool(needs_llm or confidence < 0.55),
        "review_reason": _review_reason(question, confidence, needs_llm, reason_code),
        "needs_llm": needs_llm,
        "llm_invoked": False,
        "llm_failed": False,
        "rule_baseline": {
            "business_question": question,
            "doctor_acceptance_level": acceptance,
            "doctor_need_type": need_type,
            "sales_ai_usage_type": sales_usage,
            "competitor_signal_type": competitor_signal,
            "speaker_role": speaker_role,
            "business_actor": business_actor,
            "evidence_type": evidence_type,
            "actionability": actionability,
            "confidence": confidence,
        },
        "source_text": text,
        "file_path": evidence.get("file_path", ""),
        "year": evidence.get("year", 0),
        "month": evidence.get("month", 0),
        "week_of_month": evidence.get("week_of_month", 0),
        "salesperson_id": evidence.get("salesperson_id", ""),
        "salesperson_name": evidence.get("salesperson_name", ""),
        "battle_zone_name": evidence.get("battle_zone_name", ""),
        "region_name": evidence.get("region_name", ""),
        "business_line": business_line,
        "actor_primary": actor,
        "ai_scope": ai_scope,
        "decision_status": decision_status,
    }
    return _post_process_business_fact(fact)


def _build_messages(evidence: Dict[str, object], baseline: Dict[str, object]) -> List[Dict[str, str]]:
    context = build_prompt_context()
    return [
        {
            "role": "system",
            "content": (
                "你是将军汤 AI 一线情报工作台的业务语义判定器，服务产品负责人和销售管理者。"
                "你不是关键词分类器，而是要基于项目背景判断：谁在说、说给谁、发生在什么业务动作里、是否和 AI 有真实业务关系。"
                "必须只输出一个 JSON 对象，不要输出 markdown。证据不足时不要强行归类，应标记 should_review=true。"
                "输出字段必须包含 business_question,doctor_acceptance_level,doctor_need_type,"
                "sales_ai_usage_type,competitor_signal_type,speaker_role,business_actor,evidence_type,"
                "actionability,confidence,should_review,reason,reasoning_summary。"
            ),
        },
        {
            "role": "user",
            "content": json.dumps(
                {
                    "project_business_context": context,
                    "source_text": evidence.get("source_text", ""),
                    "context": {
                        "file_path": evidence.get("file_path", ""),
                        "year": evidence.get("year", 0),
                        "month": evidence.get("month", 0),
                        "week_of_month": evidence.get("week_of_month", 0),
                        "salesperson_name": evidence.get("salesperson_name", ""),
                        "battle_zone_name": evidence.get("battle_zone_name", ""),
                        "region_name": evidence.get("region_name", ""),
                        "business_line": evidence.get("business_line", ""),
                        "actor_primary": evidence.get("actor_primary", ""),
                        "ai_scope": evidence.get("ai_scope", ""),
                    },
                    "rule_baseline": baseline.get("rule_baseline", {}),
                    "allowed_values": {
                        "business_question": list(BUSINESS_QUESTION_VALUES),
                        "doctor_acceptance_level": list(DOCTOR_ACCEPTANCE_VALUES),
                        "doctor_need_type": list(DOCTOR_NEED_VALUES),
                        "sales_ai_usage_type": list(SALES_AI_USAGE_VALUES),
                        "competitor_signal_type": list(COMPETITOR_SIGNAL_VALUES),
                        "speaker_role": list(SPEAKER_ROLE_VALUES),
                        "business_actor": list(BUSINESS_ACTOR_VALUES),
                        "evidence_type": list(EVIDENCE_TYPE_VALUES),
                        "actionability": list(ACTIONABILITY_VALUES),
                    },
                    "labels_cn": {
                        "business_question": BUSINESS_QUESTION_LABELS,
                        "doctor_acceptance_level": DOCTOR_ACCEPTANCE_LABELS,
                        "doctor_need_type": DOCTOR_NEED_LABELS,
                        "sales_ai_usage_type": SALES_AI_USAGE_LABELS,
                        "competitor_signal_type": COMPETITOR_SIGNAL_LABELS,
                        "speaker_role": SPEAKER_ROLE_LABELS,
                        "business_actor": BUSINESS_ACTOR_LABELS,
                        "evidence_type": EVIDENCE_TYPE_LABELS,
                        "actionability": ACTIONABILITY_LABELS,
                    },
                    "decision_steps": [
                        "先判断说话主体和行动主体。",
                        "再判断是否是医生真实反馈、销售自述、市场观察、竞品动作或公司机会。",
                        "再判断业务问题和细分类。",
                        "最后给出置信度、是否复核、中文原因和推理摘要。",
                    ],
                },
                ensure_ascii=False,
            ),
        },
    ]


def _apply_batch_result(
    baselines: List[Dict[str, object]],
    batch: Sequence[tuple[int, Dict[str, object], Dict[str, object]]],
    model_client: OpenAICompatibleClient | None,
) -> None:
    try:
        results = _classify_batch_with_llm(batch, model_client)
        _merge_batch_payloads(baselines, batch, results)
    except Exception as exc:  # noqa: BLE001
        _mark_batch_failed(baselines, batch, exc)


def _classify_batch_with_llm(
    batch: Sequence[tuple[int, Dict[str, object], Dict[str, object]]],
    model_client: OpenAICompatibleClient | None,
) -> Dict[str, Dict[str, object]]:
    client = model_client or OpenAICompatibleClient()
    payload = {
        "task": "对规则无法稳定判断的边界样本做业务问题语义判定。",
        "project_business_context": build_prompt_context(),
        "classification_principles": [
            "先判断谁在说、谁在行动、说给谁、发生在什么业务动作里，再判断标签。",
            "只根据原文、上下文、项目背景和规则基线判断，不要为了完整性强行归类。",
            "医生反馈必须是医生/诊所用户表达，不是销售自己的复盘、学习或判断。",
            "医生直接诉求必须是医生明确提出 AI、平台、问诊、内容、体验、可靠性、效率、成本等诉求；药价、经济、旅游、药房、剂型等泛业务内容不能硬算医生 AI 诉求。",
            "医生接纳度必须体现医生对我方云诊室 AI 或平台 AI 的态度，包括正向、兴趣、观望、顾虑、拒绝。",
            "销售日常 AI 使用必须体现销售用 AI 自用提效、对外介绍、学习复盘、话术生成或客户触达。",
            "公司内部可用 AI 降本增效的机会不能误标为医生反馈或销售动作。",
            "竞品/同行 AI 动作必须单独进入市场雷达，不污染我方云诊室 AI 结论。",
            "如果现有标签不适用或证据信号太弱，should_review=true，confidence 不得高于 0.55。",
        ],
        "allowed_values": {
            "business_question": list(BUSINESS_QUESTION_VALUES),
            "doctor_acceptance_level": list(DOCTOR_ACCEPTANCE_VALUES),
            "doctor_need_type": list(DOCTOR_NEED_VALUES),
            "sales_ai_usage_type": list(SALES_AI_USAGE_VALUES),
            "competitor_signal_type": list(COMPETITOR_SIGNAL_VALUES),
            "speaker_role": list(SPEAKER_ROLE_VALUES),
            "business_actor": list(BUSINESS_ACTOR_VALUES),
            "evidence_type": list(EVIDENCE_TYPE_VALUES),
            "actionability": list(ACTIONABILITY_VALUES),
        },
        "labels_cn": {
            "business_question": BUSINESS_QUESTION_LABELS,
            "doctor_acceptance_level": DOCTOR_ACCEPTANCE_LABELS,
            "doctor_need_type": DOCTOR_NEED_LABELS,
            "sales_ai_usage_type": SALES_AI_USAGE_LABELS,
            "competitor_signal_type": COMPETITOR_SIGNAL_LABELS,
            "speaker_role": SPEAKER_ROLE_LABELS,
            "business_actor": BUSINESS_ACTOR_LABELS,
            "evidence_type": EVIDENCE_TYPE_LABELS,
            "actionability": ACTIONABILITY_LABELS,
        },
        "items": [
            {
                "evidence_id": str(evidence.get("evidence_id", "")),
                "source_text": str(evidence.get("source_text", ""))[:900],
                "context": {
                    "file_path": evidence.get("file_path", ""),
                    "year": evidence.get("year", 0),
                    "month": evidence.get("month", 0),
                    "week_of_month": evidence.get("week_of_month", 0),
                    "salesperson_name": evidence.get("salesperson_name", ""),
                    "battle_zone_name": evidence.get("battle_zone_name", ""),
                    "region_name": evidence.get("region_name", ""),
                    "business_line": evidence.get("business_line", ""),
                    "actor_primary": evidence.get("actor_primary", ""),
                    "ai_scope": evidence.get("ai_scope", ""),
                },
                "rule_baseline": baseline.get("rule_baseline", {}),
            }
            for _, evidence, baseline in batch
        ],
        "output_contract": {
            "items": [
                {
                    "evidence_id": "必须与输入一致",
                    "business_question": "allowed_values.business_question 之一",
                    "doctor_acceptance_level": "allowed_values.doctor_acceptance_level 之一",
                    "doctor_need_type": "allowed_values.doctor_need_type 之一",
                    "sales_ai_usage_type": "allowed_values.sales_ai_usage_type 之一",
                    "competitor_signal_type": "allowed_values.competitor_signal_type 之一",
                    "speaker_role": "allowed_values.speaker_role 之一",
                    "business_actor": "allowed_values.business_actor 之一",
                    "evidence_type": "allowed_values.evidence_type 之一",
                    "actionability": "allowed_values.actionability 之一",
                    "confidence": "0-1 数字",
                    "should_review": "布尔值",
                    "reason": "一句中文原因",
                    "reasoning_summary": "一句中文说明：谁在说、什么业务动作、为什么这样判定",
                }
            ]
        },
    }
    raw = client.chat_json(
        [
            {
                "role": "system",
                "content": (
                    "你是将军汤 AI 一线情报系统的业务语义判定器，服务产品负责人和销售管理者。"
                    "你只处理边界样本，目标是减少因为缺少业务上下文导致的低级错判。"
                    "重点判断谁在说、谁在行动、AI 与业务动作是否真实相关。"
                    "必须只输出一个 JSON 对象，格式为 {\"items\":[...]}。"
                    "不得输出 markdown，不得输出英文解释。"
                ),
            },
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        max_tokens=5000,
    )
    raw_items = raw.get("items", [])
    if not isinstance(raw_items, list):
        raise RuntimeError("模型批量分类结果缺少 items 数组")
    results: Dict[str, Dict[str, object]] = {}
    for item in raw_items:
        if isinstance(item, dict):
            evidence_id = str(item.get("evidence_id", "")).strip()
            if evidence_id:
                results[evidence_id] = item
    return results


def _merge_batch_payloads(
    baselines: List[Dict[str, object]],
    batch: Sequence[tuple[int, Dict[str, object], Dict[str, object]]],
    results: Dict[str, Dict[str, object]],
) -> None:
    for idx, evidence, baseline in batch:
        evidence_id = str(evidence.get("evidence_id", ""))
        raw = results.get(evidence_id)
        if raw:
            baselines[idx] = _merge_llm_result(evidence, baseline, raw)
        else:
            failed = dict(baseline)
            failed["llm_invoked"] = True
            failed["llm_failed"] = True
            failed["should_review"] = True
            failed["confidence"] = min(float(failed.get("confidence", 0.5)), 0.35)
            failed["review_reason"] = "模型批量判定未返回该证据，需人工复核"
            baselines[idx] = failed


def _mark_batch_failed(
    baselines: List[Dict[str, object]],
    batch: Sequence[tuple[int, Dict[str, object], Dict[str, object]]],
    exc: Exception,
) -> None:
    for idx, _, baseline in batch:
        failed = dict(baseline)
        failed["llm_invoked"] = True
        failed["llm_failed"] = True
        failed["should_review"] = True
        failed["confidence"] = min(float(failed.get("confidence", 0.5)), 0.35)
        failed["review_reason"] = f"模型批量业务问题判定失败: {type(exc).__name__}: {str(exc)[:160]}"
        baselines[idx] = failed


def _merge_llm_result(evidence: Dict[str, object], baseline: Dict[str, object], raw: Dict[str, object]) -> Dict[str, object]:
    merged = dict(baseline)
    merged["business_question"] = _pick(raw.get("business_question"), BUSINESS_QUESTION_VALUES, str(baseline["business_question"]))
    merged["business_question_label"] = BUSINESS_QUESTION_LABELS.get(str(merged["business_question"]), str(merged["business_question"]))
    merged["doctor_acceptance_level"] = _pick(raw.get("doctor_acceptance_level"), DOCTOR_ACCEPTANCE_VALUES, str(baseline["doctor_acceptance_level"]))
    merged["doctor_need_type"] = _pick(raw.get("doctor_need_type"), DOCTOR_NEED_VALUES, str(baseline["doctor_need_type"]))
    merged["sales_ai_usage_type"] = _pick(raw.get("sales_ai_usage_type"), SALES_AI_USAGE_VALUES, str(baseline["sales_ai_usage_type"]))
    merged["competitor_signal_type"] = _pick(raw.get("competitor_signal_type"), COMPETITOR_SIGNAL_VALUES, str(baseline["competitor_signal_type"]))
    merged["speaker_role"] = _pick(raw.get("speaker_role"), SPEAKER_ROLE_VALUES, str(baseline.get("speaker_role", "unclear")))
    merged["speaker_role_label"] = SPEAKER_ROLE_LABELS.get(str(merged["speaker_role"]), str(merged["speaker_role"]))
    merged["business_actor"] = _pick(raw.get("business_actor"), BUSINESS_ACTOR_VALUES, str(baseline.get("business_actor", "unclear")))
    merged["business_actor_label"] = BUSINESS_ACTOR_LABELS.get(str(merged["business_actor"]), str(merged["business_actor"]))
    merged["evidence_type"] = _pick(raw.get("evidence_type"), EVIDENCE_TYPE_VALUES, str(baseline.get("evidence_type", "low_signal_context")))
    merged["evidence_type_label"] = EVIDENCE_TYPE_LABELS.get(str(merged["evidence_type"]), str(merged["evidence_type"]))
    merged["actionability"] = _pick(raw.get("actionability"), ACTIONABILITY_VALUES, str(baseline["actionability"]))
    merged["confidence"] = max(0.0, min(1.0, float(raw.get("confidence", baseline.get("confidence", 0.5)) or 0.5)))
    merged["should_review"] = bool(raw.get("should_review", merged["confidence"] < 0.65))
    merged["review_reason"] = str(raw.get("reason", "")).strip() or str(baseline.get("review_reason", ""))
    merged["reasoning_summary"] = str(raw.get("reasoning_summary", "")).strip() or str(baseline.get("reasoning_summary", ""))
    merged["llm_invoked"] = True
    merged["llm_failed"] = False
    return _post_process_business_fact(merged)


def _post_process_business_fact(fact: Dict[str, object]) -> Dict[str, object]:
    """对模型/规则结果加业务硬门槛，避免宽泛文本污染关键口径。"""
    text = str(fact.get("source_text", ""))
    question = str(fact.get("business_question", ""))
    if question == BUSINESS_QUESTION_DOCTOR_DIRECT_NEED and not _has_direct_ai_need_signal(text):
        updated = dict(fact)
        updated["business_question"] = BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY if _has_ai_related_context(text) else BUSINESS_QUESTION_DOCTOR_ACCEPTANCE
        updated["business_question_label"] = BUSINESS_QUESTION_LABELS.get(str(updated["business_question"]), str(updated["business_question"]))
        updated["should_review"] = True
        updated["confidence"] = min(float(updated.get("confidence", 0.5) or 0.5), 0.55)
        reason = str(updated.get("review_reason", "")).strip()
        suffix = "医生直接诉求缺少明确 AI/平台/问诊/功能体验语境，降级为待复核观察"
        updated["review_reason"] = f"{reason}；{suffix}" if reason else suffix
        updated["speaker_role_label"] = SPEAKER_ROLE_LABELS.get(str(updated.get("speaker_role", "")), str(updated.get("speaker_role", "")))
        updated["business_actor_label"] = BUSINESS_ACTOR_LABELS.get(str(updated.get("business_actor", "")), str(updated.get("business_actor", "")))
        updated["evidence_type_label"] = EVIDENCE_TYPE_LABELS.get(str(updated.get("evidence_type", "")), str(updated.get("evidence_type", "")))
        return updated
    fact["business_question_label"] = BUSINESS_QUESTION_LABELS.get(question, question)
    fact["speaker_role_label"] = SPEAKER_ROLE_LABELS.get(str(fact.get("speaker_role", "")), str(fact.get("speaker_role", "")))
    fact["business_actor_label"] = BUSINESS_ACTOR_LABELS.get(str(fact.get("business_actor", "")), str(fact.get("business_actor", "")))
    fact["evidence_type_label"] = EVIDENCE_TYPE_LABELS.get(str(fact.get("evidence_type", "")), str(fact.get("evidence_type", "")))
    return fact


def _infer_business_question(text: str, lower: str, actor: str, ai_scope: str) -> str:
    if ai_scope == "competitor_ai" or _has_any(text, ["竞品", "同行", "友商", "deepseek", "DeepSeek", "豆包", "通义"]):
        return BUSINESS_QUESTION_COMPETITOR_AI
    if actor in {"销售自用", "销售对外介绍"}:
        return BUSINESS_QUESTION_SALES_AI_USAGE
    if actor == "医生反馈" or _has_any(text, ["医生", "老师", "医馆", "卫生院", "诊所"]):
        if _has_any(text, ["需求", "希望", "建议", "诉求", "需要", "想要", "能不能", "是否可以"]):
            return BUSINESS_QUESTION_DOCTOR_DIRECT_NEED
        if _has_any(text, ["认可", "接受", "觉得好", "有帮助", "方便", "满意", "担心", "顾虑", "拒绝", "不需要", "没兴趣"]):
            return BUSINESS_QUESTION_DOCTOR_ACCEPTANCE
        if _has_any(text, ["痛点", "效率", "重复", "回访", "随访", "解释", "科普", "记录", "整理"]):
            return BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY
        return BUSINESS_QUESTION_DOCTOR_ACCEPTANCE
    if _has_any(text, ["销售", "自己", "学习", "复盘", "话术", "介绍", "演示", "推荐"]):
        return BUSINESS_QUESTION_SALES_AI_USAGE
    return BUSINESS_QUESTION_REGIONAL_SALES_DIFF


def _infer_doctor_acceptance(text: str, lower: str, question: str) -> str:
    if question not in {BUSINESS_QUESTION_DOCTOR_ACCEPTANCE, BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY}:
        return "not_applicable"
    if _has_any(text, ["拒绝", "不需要", "不愿", "没兴趣", "排斥", "不用"]):
        return "explicit_rejection"
    if _has_any(text, ["担心", "顾虑", "不成熟", "鸡肋", "不好", "不准", "风险", "不专业"]):
        return "explicit_concern"
    if _has_any(text, ["观望", "考虑", "再看", "试试", "体验一下", "了解一下"]):
        return "hesitation_observation"
    if _has_any(text, ["感兴趣", "想了解", "愿意了解", "咨询", "问到"]):
        return "interest_exploration"
    if _has_any(text, ["认可", "接受", "觉得好", "有帮助", "方便", "提升", "愿意", "满意"]):
        return "positive_acceptance"
    return "unknown"


def _infer_doctor_need(text: str, lower: str, question: str) -> str:
    if question not in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY, BUSINESS_QUESTION_DOCTOR_ACCEPTANCE}:
        return "not_applicable"
    if _has_any(text, ["准确", "诊断", "辨证", "疗效", "质量", "专业"]):
        return "diagnosis_quality"
    if _has_any(text, ["效率", "省时", "快速", "提升", "方便", "节省"]):
        return "efficiency"
    if _has_any(text, ["流程", "系统", "操作", "工作流", "录入", "记录"]):
        return "workflow_fit"
    if _has_any(text, ["风险", "安全", "信任", "担心", "不准", "责任"]):
        return "trust_and_safety"
    if _has_any(text, ["价格", "费用", "成本", "价值", "收费"]):
        return "cost_value"
    if _has_any(text, ["患者", "解释", "科普", "教育", "沟通"]):
        return "patient_education"
    if _has_any(text, ["回访", "随访", "复诊", "提醒"]):
        return "follow_up"
    return "unknown"


def _infer_sales_usage(text: str, lower: str, actor: str, question: str) -> str:
    if question != BUSINESS_QUESTION_SALES_AI_USAGE:
        return "not_applicable"
    if _has_any(text, ["介绍", "演示", "推荐", "讲解", "对外"]):
        return "external_pitch"
    if _has_any(text, ["学习", "复盘", "分享", "培训", "总结"]):
        return "learning_review"
    if _has_any(text, ["写", "生成", "话术", "文案", "内容"]):
        return "content_generation"
    if _has_any(text, ["回访", "问候", "维护", "触达", "加微信"]):
        return "customer_followup"
    if actor == "销售自用" or _has_any(text, ["自己", "提效", "整理", "分析"]):
        return "self_efficiency"
    return "unknown"


def _infer_competitor_signal(text: str, lower: str, ai_scope: str, question: str) -> str:
    if question != BUSINESS_QUESTION_COMPETITOR_AI:
        return "not_applicable"
    if _has_any(text, ["竞品", "友商"]):
        return "competitor_product"
    if _has_any(text, ["同行", "其他平台", "别的平台"]):
        return "peer_action"
    if _has_any(text, ["趋势", "市场", "政策", "热度"]):
        return "market_trend"
    if _has_any(text, ["对比", "比较", "问到", "客户说"]):
        return "customer_comparison"
    return "unknown"


def _infer_speaker_role(text: str, actor: str, ai_scope: str, question: str) -> str:
    if ai_scope == "competitor_ai" or question == BUSINESS_QUESTION_COMPETITOR_AI:
        return "competitor_or_market"
    if actor == "医生反馈" or _has_any(text, ["医生反馈", "老师反馈", "医生表示", "老师表示", "医生认为", "医生担心", "医生认可"]):
        return "doctor_or_clinic_user"
    if _has_any(text, ["诊所老板", "馆长", "负责人", "院长", "经营", "门店"]):
        return "clinic_operator"
    if actor in {"销售自用", "销售对外介绍"} or question == BUSINESS_QUESTION_SALES_AI_USAGE:
        return "salesperson_reporter"
    if _has_any(text, ["公司", "总部", "产品部", "内部", "流程"]):
        return "company_internal"
    if _has_any(text, ["市场", "政策", "行业", "同行"]):
        return "competitor_or_market"
    return "unclear"


def _infer_business_actor(speaker_role: str, question: str) -> str:
    if speaker_role == "doctor_or_clinic_user":
        return "doctor"
    if speaker_role == "clinic_operator":
        return "clinic_operator"
    if speaker_role == "salesperson_reporter":
        return "salesperson"
    if speaker_role == "competitor_or_market":
        return "competitor" if question == BUSINESS_QUESTION_COMPETITOR_AI else "market"
    if speaker_role == "company_internal":
        return "company"
    if question in {
        BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
        BUSINESS_QUESTION_DOCTOR_DIRECT_NEED,
        BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY,
    }:
        return "doctor"
    return "unclear"


def _infer_evidence_type(question: str, speaker_role: str, *, action_hint: str) -> str:
    if question == BUSINESS_QUESTION_COMPETITOR_AI:
        return "competitor_signal"
    if speaker_role == "doctor_or_clinic_user":
        return "doctor_feedback"
    if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY}:
        return "product_opportunity"
    if question == BUSINESS_QUESTION_SALES_AI_USAGE:
        return "sales_reflection" if _has_any(action_hint, ["自用"]) else "sales_action"
    if speaker_role == "company_internal":
        return "company_efficiency_opportunity"
    if speaker_role == "competitor_or_market":
        return "market_observation"
    return "low_signal_context"


def _rule_reasoning_summary(question: str, speaker_role: str, business_actor: str, evidence_type: str) -> str:
    return (
        f"规则初判为「{BUSINESS_QUESTION_LABELS.get(question, question)}」，"
        f"说话/行动主体为「{SPEAKER_ROLE_LABELS.get(speaker_role, speaker_role)}」，"
        f"业务对象为「{BUSINESS_ACTOR_LABELS.get(business_actor, business_actor)}」，"
        f"证据类型为「{EVIDENCE_TYPE_LABELS.get(evidence_type, evidence_type)}」。"
    )


def _infer_actionability(question: str, acceptance: str, need_type: str, sales_usage: str, competitor_signal: str, decision_status: str) -> str:
    if decision_status == DECISION_PENDING_HUMAN:
        return "observe"
    if acceptance in {"positive_acceptance", "explicit_concern", "explicit_rejection"}:
        return "report_ready"
    if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY} and need_type not in {"unknown", "not_applicable"}:
        return "actionable"
    if question == BUSINESS_QUESTION_COMPETITOR_AI and competitor_signal not in {"unknown", "not_applicable"}:
        return "report_ready"
    if question == BUSINESS_QUESTION_SALES_AI_USAGE and sales_usage not in {"unknown", "not_applicable"}:
        return "observe"
    return "observe"


def _infer_business_confidence(
    *,
    text: str,
    question: str,
    acceptance: str,
    need_type: str,
    decision_status: str,
    review_status: str,
    reason_code: str,
) -> float:
    score = 0.58
    if decision_status == "confirmed":
        score += 0.12
    if review_status == "reviewed":
        score += 0.18
    if question in BUSINESS_QUESTION_VALUES:
        score += 0.08
    if acceptance not in {"unknown", "not_applicable"}:
        score += 0.08
    if need_type not in {"unknown", "not_applicable"}:
        score += 0.05
    if reason_code:
        score -= 0.12
    if len(text) < 18:
        score -= 0.12
    return round(max(0.05, min(0.95, score)), 2)


def _needs_llm(
    evidence: Dict[str, object],
    question: str,
    acceptance: str,
    need_type: str,
    sales_usage: str,
    competitor_signal: str,
    confidence: float,
) -> bool:
    if str(evidence.get("triage_status", "")) == TRIAGE_NEEDS_LLM:
        return True
    if str(evidence.get("actor_primary", "")) in {"待判断", "label_gap", ""}:
        return True
    if str(evidence.get("business_line", "")) in {"待判断", ""}:
        return True
    if str(evidence.get("review_reason_code", "")):
        return True
    if confidence < 0.68:
        return True
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE and acceptance == "unknown":
        return True
    if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY} and need_type == "unknown":
        return True
    if question == BUSINESS_QUESTION_SALES_AI_USAGE and sales_usage == "unknown":
        return True
    if question == BUSINESS_QUESTION_COMPETITOR_AI and competitor_signal == "unknown":
        return True
    return False


def _review_reason(question: str, confidence: float, needs_llm: bool, reason_code: str) -> str:
    if reason_code:
        return f"沿用识别层复核原因：{reason_code}"
    if needs_llm:
        return "业务问题边界不稳定，需要语义确认"
    if confidence < 0.55:
        return "业务价值或结论置信度偏低"
    return ""


def _confidence_from_items(items: Sequence[Dict[str, object]]) -> str:
    if not items:
        return "low"
    avg = sum(float(item.get("confidence", 0.0)) for item in items) / len(items)
    review_rate = sum(1 for item in items if bool(item.get("should_review", False))) / len(items)
    if avg >= 0.78 and review_rate <= 0.25:
        return "high"
    if avg >= 0.58 and review_rate <= 0.5:
        return "medium"
    return "low"


def _build_why_it_matters(question: str) -> str:
    mapping = {
        BUSINESS_QUESTION_DOCTOR_ACCEPTANCE: "它决定云诊室 AI 是否具备持续推广基础。",
        BUSINESS_QUESTION_REGIONAL_SALES_DIFF: "它能帮助识别哪些区域和个人正在推动变化。",
        BUSINESS_QUESTION_DOCTOR_DIRECT_NEED: "它能直接反哺产品需求池。",
        BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY: "它能发现医生未明确表达但 AI 可解决的机会点。",
        BUSINESS_QUESTION_SALES_AI_USAGE: "它能判断销售是否把 AI 变成常规武器。",
        BUSINESS_QUESTION_COMPETITOR_AI: "它能帮助判断市场教育和竞品压力。",
    }
    return mapping.get(question, "它能帮助理解一线变化。")


def _build_evidence_cluster_insight(question: str, items: Sequence[Dict[str, object]]) -> Dict[str, object]:
    confidence = _confidence_from_items(items)
    evidence_count = len(items)
    review_needed = sum(1 for item in items if bool(item.get("should_review", False)))
    review_rate = review_needed / evidence_count if evidence_count else 0.0
    actionables = [item for item in items if str(item.get("actionability", "")) in {"report_ready", "actionable"}]
    examples = _representative_examples(items, question=question)
    breakdown = _question_breakdown(question, items)
    top_regions = _top_named(items, "region_name")
    top_sales = _top_named(items, "salesperson_name")
    trend = _trend_snapshot(items)
    conclusion = _build_cluster_conclusion(question, items, breakdown, trend, top_regions, top_sales, review_rate)
    action = _build_cluster_action(question, confidence, review_rate, actionables, top_regions, top_sales)
    caveats = _build_caveats(review_rate, breakdown, evidence_count)
    evidence_basis = _evidence_basis_sentence(breakdown, examples, review_needed, evidence_count)
    product_implication, sales_implication = _implications(question, breakdown, top_regions, top_sales)
    return {
        "insight_id": stable_hash("v16-cluster", question, str(evidence_count), str(trend.get("latest_period", ""))),
        "display_order": _question_order(question),
        "business_question": question,
        "business_question_label": BUSINESS_QUESTION_LABELS.get(question, question),
        "title": BUSINESS_QUESTION_LABELS.get(question, question),
        "insight_title": _build_insight_title(question, breakdown, trend),
        "conclusion": conclusion,
        "judgement": conclusion,
        "why_it_matters": _build_why_it_matters(question),
        "evidence_basis": evidence_basis,
        "trend_judgement": _trend_sentence(trend),
        "driving_factors": _driver_sentence(top_regions, top_sales, review_rate),
        "counter_evidence_or_uncertainty": caveats,
        "product_implication": product_implication,
        "sales_management_implication": sales_implication,
        "action_recommendation": action,
        "caveats": caveats,
        "trend_sentence": _trend_sentence(trend),
        "driver_sentence": _driver_sentence(top_regions, top_sales, review_rate),
        "evidence_count": evidence_count,
        "actionable_count": len(actionables),
        "review_needed_count": review_needed,
        "review_rate": round(review_rate, 3),
        "confidence_level": confidence,
        "confidence_label": _confidence_label(confidence),
        "breakdown": breakdown,
        "top_regions": top_regions,
        "top_sales": top_sales,
        "representative_evidence_refs": [str(item.get("evidence_id", "")) for item in examples],
        "representative_quotes": [_quote_payload(item) for item in examples],
    }


def _refine_cluster_insight_with_llm(
    insight: Dict[str, object],
    items: Sequence[Dict[str, object]],
    *,
    model_client: OpenAICompatibleClient | None = None,
) -> Dict[str, object]:
    try:
        client = model_client or OpenAICompatibleClient()
        payload = {
            "project_business_context": build_prompt_context(),
            "business_question": insight.get("title", ""),
            "current_structured_insight": {
                "conclusion": insight.get("conclusion", ""),
                "evidence_basis": insight.get("evidence_basis", ""),
                "trend_sentence": insight.get("trend_sentence", ""),
                "driver_sentence": insight.get("driver_sentence", ""),
                "counter_evidence_or_uncertainty": insight.get("counter_evidence_or_uncertainty", ""),
                "product_implication": insight.get("product_implication", ""),
                "sales_management_implication": insight.get("sales_management_implication", ""),
                "breakdown": insight.get("breakdown", {}),
                "top_regions": insight.get("top_regions", []),
                "top_sales": insight.get("top_sales", []),
                "review_rate": insight.get("review_rate", 0),
            },
            "representative_evidence": [
                _quote_payload_for_llm(item)
                for item in _representative_examples(items, question=str(insight.get("business_question", "")), limit=5)
            ],
            "output_contract": {
                "insight_title": "不超过24字",
                "conclusion": "不超过80字，直接回答业务问题",
                "evidence_basis": "不超过80字，说明证据依据",
                "trend_judgement": "不超过60字，说明趋势或数据不足",
                "driving_factors": "不超过80字，说明区域/销售/场景驱动",
                "counter_evidence_or_uncertainty": "不超过80字，说明反证或不确定性",
                "why_it_matters": "不超过60字，说明为什么重要",
                "product_implication": "不超过80字，说明产品含义",
                "sales_management_implication": "不超过80字，说明销售管理含义",
                "action_recommendation": "不超过80字，给出下一步动作",
                "caveats": "不超过60字，可信度限制",
            },
        }
        refined = client.chat_json(
            [
                {
                    "role": "system",
                "content": (
                    "你是高级 AI 业务洞察员，正在为将军汤 AI 一线情报工作台生成业务判断。"
                    "你必须置身于云诊室/云管家业务、销售周/月报和医生反馈场景中理解证据。"
                    "只能输出 JSON，不要输出 markdown。"
                    "禁止输出英文枚举、技术字段、内部 evidence_id、空泛套话。"
                    "每个字段必须短句输出，不要输出长段落，避免 JSON 被截断。"
                    "每个结论必须包含业务结论、证据依据、趋势判断、驱动因素、反证/不确定性、产品含义、销售管理含义和下一步动作。"
                    "证据不足就明确写可信度限制，不要过度推断。"
                ),
                },
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            max_tokens=2400,
        )
        merged = dict(insight)
        for key in (
            "insight_title",
            "conclusion",
            "evidence_basis",
            "trend_judgement",
            "driving_factors",
            "counter_evidence_or_uncertainty",
            "why_it_matters",
            "product_implication",
            "sales_management_implication",
            "action_recommendation",
            "caveats",
        ):
            value = str(refined.get(key, "")).strip()
            if value:
                merged[key] = value
        merged["llm_refined"] = True
        return merged
    except Exception as exc:  # noqa: BLE001 - 洞察归纳失败必须回退到规则摘要
        merged = dict(insight)
        merged["llm_refined"] = False
        merged["llm_refine_error"] = f"{type(exc).__name__}: {str(exc)[:160]}"
        return merged


def _question_breakdown(question: str, items: Sequence[Dict[str, object]]) -> Dict[str, int]:
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE:
        return _labeled_counter(items, "doctor_acceptance_level", DOCTOR_ACCEPTANCE_LABELS)
    if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY}:
        return _labeled_counter(items, "doctor_need_type", DOCTOR_NEED_LABELS)
    if question == BUSINESS_QUESTION_SALES_AI_USAGE:
        return _labeled_counter(items, "sales_ai_usage_type", SALES_AI_USAGE_LABELS)
    if question == BUSINESS_QUESTION_COMPETITOR_AI:
        return _labeled_counter(items, "competitor_signal_type", COMPETITOR_SIGNAL_LABELS)
    return _labeled_counter(items, "business_line", {})


def _labeled_counter(items: Sequence[Dict[str, object]], field: str, labels: Dict[str, str]) -> Dict[str, int]:
    counts = Counter()
    for item in items:
        raw = str(item.get(field, "")).strip()
        label = labels.get(raw, raw or "待判断")
        if label in {"不适用", "not_applicable"}:
            continue
        counts[label] += 1
    return dict(counts.most_common())


def _representative_examples(items: Sequence[Dict[str, object]], *, question: str = "", limit: int = 4) -> List[Dict[str, object]]:
    return sorted(
        items,
        key=lambda item: (
            _quote_relevance_score(question, item),
            str(item.get("actionability", "")) in {"report_ready", "actionable"},
            not bool(item.get("should_review", False)),
            float(item.get("confidence", 0.0) or 0.0),
            min(len(str(item.get("source_text", ""))), 260),
        ),
        reverse=True,
    )[:limit]


def _quote_relevance_score(question: str, item: Dict[str, object]) -> int:
    text = str(item.get("source_text", ""))
    score = 0
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE:
        if _has_any(text, ["认可", "接受", "感兴趣", "愿意", "顾虑", "担心", "拒绝", "婉拒", "没什么用", "鸡肋"]):
            score += 5
        if _has_any(text, ["医生", "老师", "医院", "医馆"]):
            score += 2
    elif question == BUSINESS_QUESTION_DOCTOR_DIRECT_NEED:
        if _has_any(text, ["AI问诊", "ai问诊", "AI辅助", "ai辅助", "ai诊疗助手", "AI诊疗", "轩岐问对"]):
            score += 6
        if _has_any(text, ["运行", "链接失败", "打不开", "固化", "修复", "优化", "开通", "功能", "体验"]):
            score += 4
        if _has_any(text, ["合作需求", "颗粒剂需求", "人工智能通识课程", "招聘会"]):
            score -= 5
    elif question == BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY:
        if _has_any(text, ["回访", "随访", "患者", "效率", "年轻医生", "经验不足", "信任", "病例"]):
            score += 5
    elif question == BUSINESS_QUESTION_SALES_AI_USAGE:
        if _has_any(text, ["借助AI", "AI工具", "AI系统", "介绍", "演示", "话术", "拜访医生前"]):
            score += 5
    elif question == BUSINESS_QUESTION_COMPETITOR_AI:
        if _has_any(text, ["竞品", "同行", "DeepSeek", "deepseek", "AI开方平台", "卫健委", "固生堂", "黑我们的文章"]):
            score += 5
    elif question == BUSINESS_QUESTION_REGIONAL_SALES_DIFF:
        if str(item.get("salesperson_name", "")).strip():
            score += 2
        if str(item.get("region_name", "")).strip():
            score += 2
    return score


def _quote_payload(item: Dict[str, object]) -> Dict[str, object]:
    period = _period_label(item)
    return {
        "evidence_id": item.get("evidence_id", ""),
        "quote": str(item.get("source_text", ""))[:260],
        "salesperson": item.get("salesperson_name", "") or "未识别销售",
        "region": item.get("region_name", "") or "未识别区域",
        "period": period,
        "file_path": item.get("file_path", ""),
    }


def _quote_payload_for_llm(item: Dict[str, object]) -> Dict[str, object]:
    period = _period_label(item)
    return {
        "quote": str(item.get("source_text", ""))[:160],
        "salesperson": item.get("salesperson_name", "") or "未识别销售",
        "region": item.get("region_name", "") or "未识别区域",
        "period": period,
    }


def _top_named(items: Sequence[Dict[str, object]], field: str, limit: int = 5) -> List[Dict[str, object]]:
    counts = Counter(str(item.get(field, "")).strip() or "未识别" for item in items)
    return [{"name": name, "count": count} for name, count in counts.most_common(limit) if name and name != "未识别"]


def _trend_snapshot(items: Sequence[Dict[str, object]]) -> Dict[str, object]:
    monthly = Counter()
    weekly = Counter()
    sales_by_month: Dict[str, set[str]] = defaultdict(set)
    for item in items:
        year = int(item.get("year", 0) or 0)
        month = int(item.get("month", 0) or 0)
        week = int(item.get("week_of_month", 0) or 0)
        if year and month:
            key = f"{year}-{month:02d}"
            monthly[key] += 1
            salesperson = str(item.get("salesperson_name", "")).strip()
            if salesperson:
                sales_by_month[key].add(salesperson)
        if year and month and week:
            weekly[f"{year}-{month:02d}-W{week}"] += 1
    month_keys = sorted(monthly)
    latest = month_keys[-1] if month_keys else ""
    previous = month_keys[-2] if len(month_keys) >= 2 else ""
    latest_value = monthly.get(latest, 0)
    previous_value = monthly.get(previous, 0)
    latest_sales = len(sales_by_month.get(latest, set())) if latest else 0
    previous_sales = len(sales_by_month.get(previous, set())) if previous else 0
    return {
        "latest_period": latest,
        "previous_period": previous,
        "latest_value": latest_value,
        "previous_value": previous_value,
        "delta": latest_value - previous_value if previous else 0,
        "latest_sales_count": latest_sales,
        "previous_sales_count": previous_sales,
        "sales_delta": latest_sales - previous_sales if previous else 0,
        "monthly_series": [{"period": key, "count": monthly[key], "sales_count": len(sales_by_month.get(key, set()))} for key in month_keys[-12:]],
        "weekly_series": [{"period": key, "count": weekly[key]} for key in sorted(weekly)[-12:]],
    }


def _trend_sentence(trend: Dict[str, object]) -> str:
    latest = str(trend.get("latest_period", ""))
    previous = str(trend.get("previous_period", ""))
    if not latest:
        return "当前证据缺少稳定时间信息，暂不能判断趋势。"
    if not previous:
        return f"{latest} 有 {trend.get('latest_value', 0)} 条相关证据，缺少上一周期对比。"
    delta = int(trend.get("delta", 0) or 0)
    direction = "增加" if delta > 0 else "减少" if delta < 0 else "持平"
    return f"{latest} 相比 {previous} {direction} {abs(delta)} 条，活跃销售人数变化 {int(trend.get('sales_delta', 0) or 0)} 人。"


def _driver_sentence(top_regions: Sequence[Dict[str, object]], top_sales: Sequence[Dict[str, object]], review_rate: float) -> str:
    region_text = "、".join(f"{row['name']}({row['count']})" for row in top_regions[:3]) or "区域未识别"
    sales_text = "、".join(f"{row['name']}({row['count']})" for row in top_sales[:3]) or "销售未识别"
    review_text = "待复核占比较高，判断需谨慎" if review_rate >= 0.4 else "待复核占比可控"
    return f"主要由 {region_text} 和 {sales_text} 贡献；{review_text}。"


def _build_cluster_conclusion(
    question: str,
    items: Sequence[Dict[str, object]],
    breakdown: Dict[str, int],
    trend: Dict[str, object],
    top_regions: Sequence[Dict[str, object]],
    top_sales: Sequence[Dict[str, object]],
    review_rate: float,
) -> str:
    count = len(items)
    dominant_label, dominant_count = _dominant(breakdown)
    trend_text = _trend_sentence(trend)
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE:
        if not breakdown or dominant_label == "待判断":
            return f"医生接纳度已有 {count} 条线索，但可判定态度不足，当前应先复核高价值片段，避免把泛 AI 表达当成真实医生反馈。{trend_text}"
        return f"医生对云诊室 AI 的反馈以“{dominant_label}”最集中（{dominant_count} 条），说明一线已经出现可讨论的接纳/顾虑信号。{trend_text}"
    if question == BUSINESS_QUESTION_DOCTOR_DIRECT_NEED:
        if dominant_label and dominant_label != "待判断":
            return f"医生直接诉求集中在“{dominant_label}”（{dominant_count} 条），这些内容最适合进入产品机会池做优先级评估。{trend_text}"
        return f"医生直接诉求有 {count} 条，但诉求类型尚未稳定，不能直接把数量当成产品需求；应先复核代表原文，确认哪些是真需求、哪些只是泛 AI 讨论。{trend_text}"
    if question == BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY:
        if dominant_label and dominant_label != "待判断":
            return f"间接 AI 机会主要来自“{dominant_label}”，这类线索不是医生直接提 AI，而是暴露了可被 AI 改善的工作流问题。{trend_text}"
        return f"间接 AI 机会有 {count} 条，但场景尚未收敛；需要从原文里确认医生到底卡在效率、随访、患者沟通还是诊疗质量。{trend_text}"
    if question == BUSINESS_QUESTION_SALES_AI_USAGE:
        if dominant_label and dominant_label != "待判断":
            return f"销售 AI 使用主要集中在“{dominant_label}”，需要继续判断这是少数销售高频使用，还是团队使用面扩大。{trend_text}"
        return f"销售 AI 使用证据有 {count} 条，但用途尚未定性；需要区分自用提效、对外介绍、复盘学习和客户触达。{trend_text}"
    if question == BUSINESS_QUESTION_COMPETITOR_AI:
        if dominant_label and dominant_label != "待判断":
            return f"竞品/同行 AI 信号以“{dominant_label}”为主，必须单独进入市场雷达，不能混入我方产品 AI 成效。{trend_text}"
        return f"竞品/同行 AI 信号已有 {count} 条，但很多还停留在市场讨论或客户比较，需要先复核后再进入市场雷达。{trend_text}"
    return f"区域和个人差异线索集中在 {top_regions[0]['name'] if top_regions else '少数区域'} / {top_sales[0]['name'] if top_sales else '少数销售'}，需要下钻判断是区域普遍变化还是个人驱动。"


def _build_cluster_action(
    question: str,
    confidence: str,
    review_rate: float,
    actionables: Sequence[Dict[str, object]],
    top_regions: Sequence[Dict[str, object]],
    top_sales: Sequence[Dict[str, object]],
) -> str:
    if review_rate >= 0.45:
        return "先完成本轮 20 条复核，再决定是否进入正式报告；当前不建议直接下结论。"
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE:
        return "优先抽取正向/顾虑/拒绝各 2-3 条原文，形成医生接纳度证据页。"
    if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY}:
        return "把可行动片段进入产品机会池，并标明对应业务线、医生场景和证据原文。"
    if question == BUSINESS_QUESTION_SALES_AI_USAGE:
        return f"优先找 {top_sales[0]['name'] if top_sales else '高频销售'} 复盘真实用法，判断是否可沉淀话术或案例。"
    if question == BUSINESS_QUESTION_COMPETITOR_AI:
        return "作为市场雷达单独跟踪，只汇报同行动作和客户比较，不作为我方 AI 接纳度依据。"
    return f"优先下钻 {top_regions[0]['name'] if top_regions else '高频区域'}，确认活跃来源是普遍扩散还是个人贡献。"


def _build_caveats(review_rate: float, breakdown: Dict[str, int], evidence_count: int) -> str:
    caveats = []
    if review_rate >= 0.4:
        caveats.append(f"待复核占比 {round(review_rate * 100, 1)}%，当前结论可信度受影响。")
    if breakdown.get("待判断", 0) >= max(3, evidence_count * 0.25):
        caveats.append("待判断样本偏多，说明上下文理解或标签边界仍需修正。")
    if evidence_count < 10:
        caveats.append("样本量偏小，只能作为观察线索。")
    return " ".join(caveats) if caveats else "证据质量相对稳定，但仍需保留原文追溯。"


def _evidence_basis_sentence(
    breakdown: Dict[str, int],
    examples: Sequence[Dict[str, object]],
    review_needed: int,
    evidence_count: int,
) -> str:
    dominant_label, dominant_count = _dominant(breakdown)
    quote_hint = ""
    if examples:
        first = str(examples[0].get("source_text", ""))[:80]
        quote_hint = f"代表原文显示：{first}"
    if dominant_label:
        return f"证据共 {evidence_count} 条，其中「{dominant_label}」最多（{dominant_count} 条），待复核 {review_needed} 条。{quote_hint}"
    return f"证据共 {evidence_count} 条，待复核 {review_needed} 条；当前分类结构尚不稳定。{quote_hint}"


def _implications(
    question: str,
    breakdown: Dict[str, int],
    top_regions: Sequence[Dict[str, object]],
    top_sales: Sequence[Dict[str, object]],
) -> tuple[str, str]:
    dominant_label, _ = _dominant(breakdown)
    region = top_regions[0]["name"] if top_regions else "高频区域"
    salesperson = top_sales[0]["name"] if top_sales else "高频销售"
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE:
        return (
            f"产品侧应优先拆解医生对 AI 的「{dominant_label or '接纳/顾虑'}」来源，判断是能力体验、信任、安全还是流程适配问题。",
            f"销售管理侧应要求 {region} / {salesperson} 补充医生原话和场景，沉淀可复用话术与反对意见处理。",
        )
    if question == BUSINESS_QUESTION_DOCTOR_DIRECT_NEED:
        return (
            f"产品侧应把「{dominant_label or '待判断诉求'}」作为候选需求池，结合原文判断是否是真需求而非泛业务问题。",
            "销售管理侧应规范需求收集格式，要求补充医生角色、使用场景、当前替代方案和影响程度。",
        )
    if question == BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY:
        return (
            f"产品侧应从「{dominant_label or '间接场景'}」里识别 AI 可介入的工作流，而不是只看医生是否直接提 AI。",
            "销售管理侧应引导销售记录具体阻塞点，避免只写泛泛的客户困难。",
        )
    if question == BUSINESS_QUESTION_SALES_AI_USAGE:
        return (
            "产品侧可观察销售真实使用方式，反推一线需要的 AI 资料、话术和培训素材。",
            f"销售管理侧应复盘 {salesperson} 的具体用法，判断是否能形成团队案例或培训材料。",
        )
    if question == BUSINESS_QUESTION_COMPETITOR_AI:
        return (
            "产品侧应把竞品/同行动作作为市场雷达，不直接纳入我方 AI 接纳度。",
            "销售管理侧应沉淀客户比较话术，避免一线对竞品 AI 缺少统一回应。",
        )
    return (
        "产品侧暂不直接行动，先看区域/销售差异是否由真实需求驱动。",
        f"销售管理侧应下钻 {region} 与 {salesperson}，判断是个人高频还是区域普遍扩散。",
    )


def _build_insight_title(question: str, breakdown: Dict[str, int], trend: Dict[str, object]) -> str:
    dominant, _ = _dominant(breakdown)
    prefix = BUSINESS_QUESTION_LABELS.get(question, question)
    if dominant and dominant != "待判断":
        return f"{prefix}：{dominant}最突出"
    return f"{prefix}：需要复核后定性"


def _headline_from_insights(insight_map: Dict[str, Dict[str, object]], summary: Dict[str, object]) -> str:
    doctor = insight_map.get(BUSINESS_QUESTION_DOCTOR_ACCEPTANCE, {})
    sales = insight_map.get(BUSINESS_QUESTION_SALES_AI_USAGE, {})
    review_count = int(summary.get("review_needed_count", 0) or 0)
    total = int(summary.get("total_business_evidence", 0) or 0)
    if review_count >= max(20, total * 0.35):
        return "当前 AI 一线信号已经具备观察价值，但待复核比例偏高，不能直接把数量当成业务结论。"
    if doctor and sales:
        return "当前最值得看的不是 AI 被提了多少次，而是医生反馈与销售使用是否正在形成可复用模式。"
    return "当前证据可以支撑方向性观察，但仍需要围绕医生、销售、竞品三条线分开判断。"


def _top_actions(insights: Sequence[Dict[str, object]]) -> List[str]:
    actions = []
    for item in insights:
        if str(item.get("confidence_level", "")) != "low":
            actions.append(str(item.get("action_recommendation", "")))
    deduped = []
    for action in actions:
        if action and action not in deduped:
            deduped.append(action)
    return deduped[:5] or ["先完成本轮复核，再输出正式行动建议。"]


def _top_risks(insights: Sequence[Dict[str, object]]) -> List[str]:
    risks = []
    for item in insights:
        caveat = str(item.get("caveats", "")).strip()
        if caveat:
            risks.append(f"{item.get('title', '')}：{caveat}")
    return risks[:5] or ["当前主要风险是样本仍需追溯原文，避免过度解读。"]


def _period_label(item: Dict[str, object]) -> str:
    year = int(item.get("year", 0) or 0)
    month = int(item.get("month", 0) or 0)
    week = int(item.get("week_of_month", 0) or 0)
    if year and month and week:
        return f"{year}-{month:02d}-W{week}"
    if year and month:
        return f"{year}-{month:02d}"
    return "未知周期"


def _dominant(payload: Dict[str, int]) -> tuple[str, int]:
    if not payload:
        return "", 0
    return max(payload.items(), key=lambda item: item[1])


def _question_order(question: str) -> int:
    order = {
        BUSINESS_QUESTION_DOCTOR_ACCEPTANCE: 1,
        BUSINESS_QUESTION_REGIONAL_SALES_DIFF: 2,
        BUSINESS_QUESTION_DOCTOR_DIRECT_NEED: 3,
        BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY: 4,
        BUSINESS_QUESTION_SALES_AI_USAGE: 5,
        BUSINESS_QUESTION_COMPETITOR_AI: 6,
    }
    return order.get(question, 99)


def _confidence_label(value: str) -> str:
    return {"high": "高可信", "medium": "中可信", "low": "低可信"}.get(value, "待判断")


def _pick(value: object, allowed: Iterable[str], default: str) -> str:
    text = str(value or "").strip()
    return text if text in set(allowed) else default


def _has_any(text: str, keywords: Sequence[str]) -> bool:
    lower = text.lower()
    return any(keyword.lower() in lower for keyword in keywords)


def _has_ai_related_context(text: str) -> bool:
    return _has_any(
        text,
        [
            "ai",
            "AI",
            "人工智能",
            "平台",
            "问诊",
            "诊疗",
            "诊疗助手",
            "轩岐",
            "智能",
            "工具",
            "系统",
        ],
    )


def _has_direct_ai_need_signal(text: str) -> bool:
    if not _has_ai_related_context(text):
        return False
    return _has_any(
        text,
        [
            "希望",
            "建议",
            "需求",
            "诉求",
            "想要",
            "想让",
            "能不能",
            "是否可以",
            "运行",
            "太慢",
            "慢",
            "不准",
            "准确",
            "固化",
            "优化",
            "改进",
            "体验",
            "效率",
            "辅助",
            "帮助",
            "患者沟通",
            "科普",
            "解释",
            "内容",
            "功能",
        ],
    )
