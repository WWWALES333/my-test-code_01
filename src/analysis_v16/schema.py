from __future__ import annotations

from src.analysis_v14.schema import stable_hash

BUSINESS_QUESTION_DOCTOR_ACCEPTANCE = "doctor_ai_acceptance"
BUSINESS_QUESTION_REGIONAL_SALES_DIFF = "regional_sales_difference"
BUSINESS_QUESTION_DOCTOR_DIRECT_NEED = "doctor_direct_need"
BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY = "doctor_indirect_opportunity"
BUSINESS_QUESTION_SALES_AI_USAGE = "sales_ai_usage"
BUSINESS_QUESTION_COMPETITOR_AI = "competitor_ai_signal"

BUSINESS_QUESTION_VALUES = (
    BUSINESS_QUESTION_DOCTOR_ACCEPTANCE,
    BUSINESS_QUESTION_REGIONAL_SALES_DIFF,
    BUSINESS_QUESTION_DOCTOR_DIRECT_NEED,
    BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY,
    BUSINESS_QUESTION_SALES_AI_USAGE,
    BUSINESS_QUESTION_COMPETITOR_AI,
)

BUSINESS_QUESTION_LABELS = {
    BUSINESS_QUESTION_DOCTOR_ACCEPTANCE: "医生 AI 接纳度",
    BUSINESS_QUESTION_REGIONAL_SALES_DIFF: "区域 / 销售个人差异",
    BUSINESS_QUESTION_DOCTOR_DIRECT_NEED: "医生直接诉求",
    BUSINESS_QUESTION_DOCTOR_INDIRECT_OPPORTUNITY: "医生间接 AI 机会",
    BUSINESS_QUESTION_SALES_AI_USAGE: "销售日常 AI 使用",
    BUSINESS_QUESTION_COMPETITOR_AI: "竞品 / 同行 AI 动作",
}

DOCTOR_ACCEPTANCE_VALUES = (
    "positive_acceptance",
    "interest_exploration",
    "hesitation_observation",
    "explicit_concern",
    "explicit_rejection",
    "not_applicable",
    "unknown",
)

DOCTOR_ACCEPTANCE_LABELS = {
    "positive_acceptance": "正向接受",
    "interest_exploration": "兴趣探索",
    "hesitation_observation": "犹豫观望",
    "explicit_concern": "明确顾虑",
    "explicit_rejection": "拒绝排斥",
    "not_applicable": "不适用",
    "unknown": "待判断",
}

DOCTOR_NEED_VALUES = (
    "diagnosis_quality",
    "efficiency",
    "workflow_fit",
    "trust_and_safety",
    "cost_value",
    "patient_education",
    "follow_up",
    "unknown",
    "not_applicable",
)

DOCTOR_NEED_LABELS = {
    "diagnosis_quality": "诊疗质量 / 准确性",
    "efficiency": "效率提升",
    "workflow_fit": "工作流适配",
    "trust_and_safety": "信任与安全",
    "cost_value": "成本与价值",
    "patient_education": "患者沟通 / 科普",
    "follow_up": "回访与随访",
    "unknown": "待判断",
    "not_applicable": "不适用",
}

SALES_AI_USAGE_VALUES = (
    "self_efficiency",
    "external_pitch",
    "learning_review",
    "content_generation",
    "customer_followup",
    "not_applicable",
    "unknown",
)

SALES_AI_USAGE_LABELS = {
    "self_efficiency": "销售自用提效",
    "external_pitch": "对外介绍 / 演示",
    "learning_review": "内部学习 / 复盘",
    "content_generation": "内容 / 话术生成",
    "customer_followup": "客户触达 / 回访",
    "not_applicable": "不适用",
    "unknown": "待判断",
}

COMPETITOR_SIGNAL_VALUES = (
    "competitor_product",
    "peer_action",
    "market_trend",
    "customer_comparison",
    "not_applicable",
    "unknown",
)

COMPETITOR_SIGNAL_LABELS = {
    "competitor_product": "竞品产品动作",
    "peer_action": "同行动作",
    "market_trend": "市场趋势",
    "customer_comparison": "客户比较",
    "not_applicable": "不适用",
    "unknown": "待判断",
}

ACTIONABILITY_VALUES = (
    "report_ready",
    "actionable",
    "observe",
    "no_action",
)

ACTIONABILITY_LABELS = {
    "report_ready": "可直接进报告",
    "actionable": "可形成行动建议",
    "observe": "继续观察",
    "no_action": "无行动价值",
}

REVIEW_ERROR_REASON_VALUES = (
    "rule_issue",
    "prompt_issue",
    "label_gap",
    "context_missing",
    "low_value_noise",
    "parser_segmentation_issue",
    "business_definition_gap",
    "model_output_format",
    "other",
)
