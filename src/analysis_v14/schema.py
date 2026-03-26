from __future__ import annotations

import hashlib
from typing import Dict, List

DECISION_CONFIRMED = "confirmed"
DECISION_UNCERTAIN = "uncertain"
DECISION_PENDING_HUMAN = "pending_human_review"

DECISION_VALUES = {
    DECISION_CONFIRMED,
    DECISION_UNCERTAIN,
    DECISION_PENDING_HUMAN,
}

CERTAINTY_LEVEL_VALUES = {"high", "medium", "low"}

BUSINESS_LINE_VALUES = {"云诊室", "云管家", "混合", "待判断"}

ACTOR_PRIMARY_VALUES = {
    "销售自用",
    "销售对外介绍",
    "医生反馈",
    "潜在 AI 机会",
    "待判断",
}

AI_ACTOR_VALUES = ACTOR_PRIMARY_VALUES

AI_SCOPE_VALUES = {
    "product_ai",
    "market_trend",
    "competitor_ai",
    "general_ai",
}

INTERACTION_OUTCOME_VALUES = {
    "no_feedback",
    "positive_feedback",
    "negative_or_observing",
    "converted",
    "not_applicable",
}

REVIEW_REASON_CODE_VALUES = {
    "ACTOR_OVERLAP",
    "SCOPE_AMBIGUOUS",
    "BROAD_STATEMENT",
    "BUSINESSLINE_LOW_SIGNAL",
    "OUTCOME_UNCLEAR",
    "PARSE_FAILED",
    "PARSE_FAILED_DOC",
    "PARSE_FAILED_PDF",
    "PARSER_TOOL_MISSING",
    "MODEL_CALL_FAILED",
}

PARSE_STATUS_SUCCESS = "success"
PARSE_STATUS_FAILED = "failed"

PARSE_REASON_PDF = "PARSE_FAILED_PDF"
PARSE_REASON_DOC = "PARSE_FAILED_DOC"
PARSE_REASON_TOOL_MISSING = "PARSER_TOOL_MISSING"
PARSE_REASON_GENERIC = "PARSE_FAILED"
MODEL_REASON_FAILED = "MODEL_CALL_FAILED"

AI_EXPLICIT_KEYWORDS = [
    "ai",
    "人工智能",
    "deepseek",
    "chatgpt",
    "智能问诊",
    "智能小结",
    "智能辅助",
    "ai辅助",
    "ai诊疗",
]

AI_OPPORTUNITY_KEYWORDS = [
    "优化话术",
    "资料整理",
    "智能推荐",
    "自动整理",
    "个性化沟通",
    "智能化",
    "效率提升",
    "辅助诊疗",
]

BUSINESS_LINE_KEYWORDS: Dict[str, List[str]] = {
    "云诊室": [
        "医生",
        "问诊",
        "接诊",
        "复诊",
        "开方",
        "处方",
        "药方",
        "代煎",
        "履约",
    ],
    "云管家": [
        "云管家",
        "诊所",
        "门诊",
        "卫生所",
        "健康管理公司",
        "会员",
        "储值",
        "saaS",
        "上线",
        "经营",
        "管理",
    ],
}

CONTEXT_BUSINESS_PRIOR_KEYWORDS: Dict[str, List[str]] = {
    "云诊室": [
        "医生工作周报",
        "战区工作周报",
        "云诊室",
    ],
    "云管家": [
        "云管家",
        "卫生所",
        "健康管理公司",
        "门诊",
        "诊所",
    ],
}

ACTOR_KEYWORDS: Dict[str, List[str]] = {
    "销售自用": [
        "我用",
        "我在用",
        "我也经常使用",
        "自己用",
        "学习",
        "复盘",
        "查资料",
        "写话术",
        "deepseek",
        "chatgpt",
    ],
    "销售对外介绍": [
        "介绍",
        "演示",
        "沟通",
        "同步了一下",
        "给医生讲",
        "向医生介绍",
        "对外介绍",
        "宣讲",
    ],
    "医生反馈": [
        "医生反馈",
        "老师反馈",
        "医生表示",
        "医生觉得",
        "医生提出",
        "老板反馈",
        "诊所反馈",
        "感兴趣",
        "不成熟",
        "观望",
        "抵触",
    ],
    "潜在 AI 机会": AI_OPPORTUNITY_KEYWORDS,
}

SCOPE_KEYWORDS: Dict[str, List[str]] = {
    "competitor_ai": [
        "竞品",
        "固生堂",
        "小鹿",
        "黑我们的文章",
    ],
    "market_trend": [
        "两会",
        "行业方向",
        "政策",
        "趋势",
        "ai搜索",
        "推荐的甘草",
    ],
    "general_ai": [
        "科技发展的潮流",
        "未来社会",
        "有一个更深度的认识",
        "属于销售个人思考",
    ],
}

ACTION_KEYWORDS = [
    "回访",
    "拜访",
    "介绍",
    "沟通",
    "演示",
    "开方",
    "接诊",
    "使用",
    "跟进",
]

POSITIVE_FEEDBACK_KEYWORDS = [
    "感兴趣",
    "认可",
    "提升",
    "方便",
    "有增长",
]

NEGATIVE_FEEDBACK_KEYWORDS = [
    "不成熟",
    "观望",
    "否定",
    "不适应",
    "抵触",
]

CONVERSION_KEYWORDS = [
    "有增长",
    "转化",
    "继续开方",
    "后续合作",
]

CUSTOMER_ENTITY_HINTS = [
    "医生",
    "老师",
    "诊所",
    "卫生所",
    "客户",
]


def stable_hash(*parts: str, length: int = 16) -> str:
    payload = "||".join(parts).encode("utf-8", errors="ignore")
    return hashlib.md5(payload).hexdigest()[:length]
