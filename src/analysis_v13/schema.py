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

BUSINESS_LINE_VALUES = {"云诊室", "云管家", "混合", "待判断"}

AI_ACTOR_VALUES = {
    "销售自用",
    "销售对外介绍",
    "医生反馈",
    "潜在 AI 机会",
    "待判断",
}

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
        "会员",
        "储值",
        "saaS",
        "上线",
        "经营",
        "管理",
    ],
}

ACTOR_KEYWORDS: Dict[str, List[str]] = {
    "销售自用": [
        "我用",
        "我在用",
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
        "给医生讲",
        "向医生介绍",
        "对外介绍",
        "宣讲",
    ],
    "医生反馈": [
        "医生反馈",
        "医生表示",
        "医生觉得",
        "医生提出",
        "老板反馈",
        "诊所反馈",
        "观望",
        "抵触",
    ],
    "潜在 AI 机会": AI_OPPORTUNITY_KEYWORDS,
}


def stable_hash(*parts: str, length: int = 16) -> str:
    payload = "||".join(parts).encode("utf-8", errors="ignore")
    return hashlib.md5(payload).hexdigest()[:length]

