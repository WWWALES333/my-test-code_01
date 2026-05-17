from __future__ import annotations

import json
from collections import Counter, defaultdict
from typing import Dict, Iterable, List, Sequence

from src.analysis_v14.schema import DECISION_PENDING_HUMAN, TRIAGE_NEEDS_LLM, stable_hash

from .model_adapter import OpenAICompatibleClient
from .schema import (
    ACTIONABILITY_VALUES,
    BUSINESS_QUESTION_COMPETITOR_AI,
    BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
    BUSINESS_QUESTION_DOCTOR_DIRECT_NEED,
    BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY,
    BUSINESS_QUESTION_LABELS,
    BUSINESS_QUESTION_REGIONAL_SALES_DIFF,
    BUSINESS_QUESTION_SALES_AI_USAGE,
    BUSINESS_QUESTION_VALUES,
    COMPETITOR_SIGNAL_VALUES,
    DOCTOR_ACCEPTANCE_VALUES,
    DOCTOR_NEED_VALUES,
    SALES_AI_USAGE_VALUES,
)


class BusinessQuestionAnalyzer:
    """把 v1.5 证据进一步映射为 v1.6 业务问题层。"""

    def __init__(self, mode: str = "mock", model_client: OpenAICompatibleClient | None = None) -> None:
        if mode not in {"mock", "real"}:
            raise ValueError(f"不支持的业务问题分析模式: {mode}")
        self.mode = mode
        self.model_client = model_client

    def analyze_batch(self, evidence_facts: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
        return [self.analyze(row) for row in evidence_facts]

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


def build_business_insights(rows: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    grouped: Dict[str, List[Dict[str, object]]] = defaultdict(list)
    for row in rows:
        grouped[str(row.get("business_question", "unknown"))].append(row)

    insights: List[Dict[str, object]] = []
    for question, items in grouped.items():
        if question not in BUSINESS_QUESTION_VALUES:
            continue
        confidence = _confidence_from_items(items)
        actionables = [item for item in items if str(item.get("actionability", "")) in {"report_ready", "actionable"}]
        examples = sorted(
            items,
            key=lambda item: (-float(item.get("confidence", 0.0)), -len(str(item.get("source_text", "")))),
        )[:3]
        insights.append(
            {
                "insight_id": stable_hash("v16", question, str(len(items))),
                "business_question": question,
                "title": BUSINESS_QUESTION_LABELS.get(question, question),
                "judgement": _build_judgement(question, items),
                "why_it_matters": _build_why_it_matters(question),
                "action_recommendation": _build_action_recommendation(question, actionables, confidence),
                "evidence_count": len(items),
                "actionable_count": len(actionables),
                "review_needed_count": sum(1 for item in items if bool(item.get("should_review", False))),
                "confidence_level": confidence,
                "representative_evidence_refs": [str(item.get("evidence_id", "")) for item in examples],
                "representative_quotes": [str(item.get("source_text", ""))[:180] for item in examples],
            }
        )
    return sorted(insights, key=lambda item: (-int(item.get("evidence_count", 0)), str(item.get("business_question", ""))))


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
    return {
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


def _build_messages(evidence: Dict[str, object], baseline: Dict[str, object]) -> List[Dict[str, str]]:
    return [
        {
            "role": "system",
            "content": (
                "你是将军汤销售周报 AI 一线情报的边界判定器。"
                "你只处理规则无法稳定判断的边界样本。必须只输出一个 JSON 对象。"
                "如果现有分类不适用，不要强行归类，应标记 should_review=true。"
                "字段必须包含 business_question,doctor_acceptance_level,doctor_need_type,"
                "sales_ai_usage_type,competitor_signal_type,actionability,confidence,should_review,reason。"
            ),
        },
        {
            "role": "user",
            "content": json.dumps(
                {
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
                        "actionability": list(ACTIONABILITY_VALUES),
                    },
                    "business_questions_cn": BUSINESS_QUESTION_LABELS,
                },
                ensure_ascii=False,
            ),
        },
    ]


def _merge_llm_result(evidence: Dict[str, object], baseline: Dict[str, object], raw: Dict[str, object]) -> Dict[str, object]:
    merged = dict(baseline)
    merged["business_question"] = _pick(raw.get("business_question"), BUSINESS_QUESTION_VALUES, str(baseline["business_question"]))
    merged["business_question_label"] = BUSINESS_QUESTION_LABELS.get(str(merged["business_question"]), str(merged["business_question"]))
    merged["doctor_acceptance_level"] = _pick(raw.get("doctor_acceptance_level"), DOCTOR_ACCEPTANCE_VALUES, str(baseline["doctor_acceptance_level"]))
    merged["doctor_need_type"] = _pick(raw.get("doctor_need_type"), DOCTOR_NEED_VALUES, str(baseline["doctor_need_type"]))
    merged["sales_ai_usage_type"] = _pick(raw.get("sales_ai_usage_type"), SALES_AI_USAGE_VALUES, str(baseline["sales_ai_usage_type"]))
    merged["competitor_signal_type"] = _pick(raw.get("competitor_signal_type"), COMPETITOR_SIGNAL_VALUES, str(baseline["competitor_signal_type"]))
    merged["actionability"] = _pick(raw.get("actionability"), ACTIONABILITY_VALUES, str(baseline["actionability"]))
    merged["confidence"] = max(0.0, min(1.0, float(raw.get("confidence", baseline.get("confidence", 0.5)) or 0.5)))
    merged["should_review"] = bool(raw.get("should_review", merged["confidence"] < 0.65))
    merged["review_reason"] = str(raw.get("reason", "")).strip() or str(baseline.get("review_reason", ""))
    merged["llm_invoked"] = True
    merged["llm_failed"] = False
    return merged


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


def _build_judgement(question: str, items: Sequence[Dict[str, object]]) -> str:
    count = len(items)
    if question == BUSINESS_QUESTION_DOCTOR_ACCEPTANCE:
        dominant = Counter(str(item.get("doctor_acceptance_level", "unknown")) for item in items).most_common(1)[0][0]
        return f"当前医生 AI 接纳度相关证据 {count} 条，主要集中在 {dominant}。"
    if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY}:
        dominant = Counter(str(item.get("doctor_need_type", "unknown")) for item in items).most_common(1)[0][0]
        return f"当前医生诉求 / 机会证据 {count} 条，最集中方向是 {dominant}。"
    if question == BUSINESS_QUESTION_SALES_AI_USAGE:
        dominant = Counter(str(item.get("sales_ai_usage_type", "unknown")) for item in items).most_common(1)[0][0]
        return f"当前销售 AI 使用证据 {count} 条，主要场景是 {dominant}。"
    if question == BUSINESS_QUESTION_COMPETITOR_AI:
        return f"当前竞品 / 同行 AI 信号 {count} 条，应单独进入市场雷达。"
    return f"当前区域和销售差异相关证据 {count} 条，需要结合销售画像下钻。"


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


def _build_action_recommendation(question: str, actionables: Sequence[Dict[str, object]], confidence: str) -> str:
    if confidence == "low":
        return "先进入复核，不建议直接行动。"
    if actionables:
        if question in {BUSINESS_QUESTION_DOCTOR_DIRECT_NEED, BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY}:
            return "建议进入产品机会池，并补充代表原文。"
        if question == BUSINESS_QUESTION_SALES_AI_USAGE:
            return "建议沉淀优秀销售话术或使用案例。"
        if question == BUSINESS_QUESTION_COMPETITOR_AI:
            return "建议进入市场雷达，和竞品动作分开跟踪。"
        return "建议进入周报/月报摘要。"
    return "继续观察，暂不形成行动项。"


def _pick(value: object, allowed: Iterable[str], default: str) -> str:
    text = str(value or "").strip()
    return text if text in set(allowed) else default


def _has_any(text: str, keywords: Sequence[str]) -> bool:
    lower = text.lower()
    return any(keyword.lower() in lower for keyword in keywords)
