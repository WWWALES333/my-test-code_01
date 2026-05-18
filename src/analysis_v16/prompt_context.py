from __future__ import annotations

import json
from typing import Dict, Tuple

PROMPT_CONTEXT_VERSION = "v1.6-business-context-20260518"

PROMPT_SOURCE_DOCUMENTS: Tuple[str, ...] = (
    "docs/01_business_context.md",
    "docs/02_domain_glossary.md",
    "data/input/v1.6/business_question_taxonomy.md",
    "docs/releases/v1.6/PRD.md",
)

SPEAKER_ROLE_VALUES = (
    "salesperson_reporter",
    "doctor_or_clinic_user",
    "clinic_operator",
    "competitor_or_market",
    "company_internal",
    "unclear",
)

SPEAKER_ROLE_LABELS = {
    "salesperson_reporter": "销售本人在汇报或复盘",
    "doctor_or_clinic_user": "医生或诊所用户在反馈",
    "clinic_operator": "诊所经营者 / 老板视角",
    "competitor_or_market": "竞品 / 同行 / 市场环境",
    "company_internal": "公司内部能力或流程机会",
    "unclear": "说话主体不清",
}

BUSINESS_ACTOR_VALUES = (
    "doctor",
    "salesperson",
    "clinic_operator",
    "patient",
    "competitor",
    "company",
    "market",
    "unclear",
)

BUSINESS_ACTOR_LABELS = {
    "doctor": "医生",
    "salesperson": "销售",
    "clinic_operator": "诊所经营者",
    "patient": "患者",
    "competitor": "竞品 / 同行",
    "company": "公司内部",
    "market": "市场环境",
    "unclear": "待判断",
}

EVIDENCE_TYPE_VALUES = (
    "doctor_feedback",
    "sales_action",
    "sales_reflection",
    "market_observation",
    "competitor_signal",
    "product_opportunity",
    "company_efficiency_opportunity",
    "low_signal_context",
)

EVIDENCE_TYPE_LABELS = {
    "doctor_feedback": "医生真实反馈",
    "sales_action": "销售动作",
    "sales_reflection": "销售复盘 / 学习",
    "market_observation": "市场观察",
    "competitor_signal": "竞品 / 同行动作",
    "product_opportunity": "产品机会",
    "company_efficiency_opportunity": "公司降本增效机会",
    "low_signal_context": "低信号背景",
}


def build_prompt_context() -> Dict[str, object]:
    """返回所有 Minimax 调用共享的压缩业务背景包。

    这里刻意不把上位文档全文塞进 prompt，而是沉淀为稳定、可审阅、可版本化的业务上下文。
    """
    return {
        "context_version": PROMPT_CONTEXT_VERSION,
        "source_documents": list(PROMPT_SOURCE_DOCUMENTS),
        "project_stage": (
            "v1.6 正在把 v1.5 的 AI 结果展示页纠偏为 AI 一线情报工作台。"
            "目标不是统计 AI 被提了多少次，而是帮助产品负责人和销售管理者判断一线趋势、机会和风险。"
        ),
        "business_model": {
            "company": "将军汤 / 甘草医生相关一线销售体系",
            "core_logic": (
                "业务通过销售团队触达医生、医馆、卫生院、诊所等一线对象，推动医生上线、问诊、开方、复诊、患者经营和诊所经营相关业务。"
                "周报和月报是销售记录一线动作、客户反馈、市场竞争、产品问题和机会线索的主要文本来源。"
            ),
            "why_ai_matters": (
                "AI 不是独立话题，而是可能影响医生接纳、销售话术、诊疗效率、患者沟通、诊所经营效率、竞品竞争和公司内部降本增效的业务变量。"
            ),
        },
        "business_lines": {
            "云诊室": (
                "偏医生侧和诊疗服务侧，关注医生上线、线上问诊、辨证开方、处方、复诊、回访、患者沟通和平台 AI 诊疗助手等。"
                "医生对 AI 的接受、顾虑、功能诉求和体验反馈优先归到云诊室语境。"
            ),
            "云管家": (
                "偏诊所经营和管理侧，关注诊所老板/经营者、会员、储值、经营效率、SaaS、门店管理、降本增效等。"
                "如果 AI 主要用于诊所经营管理、客户经营或老板视角，应与云诊室医生诊疗反馈区分。"
            ),
            "混合": "同一片段同时涉及医生诊疗和诊所经营，且二者都对结论有贡献时使用。",
        },
        "roles": {
            "sales": (
                "销售是周/月报作者，日常做拜访、回访、转化、演示、推荐、客情维护、收集反馈、复盘学习和推动医生/诊所使用平台。"
                "销售自己说“我用 AI 学习/写话术/整理材料”属于销售 AI 使用，不等于医生反馈。"
            ),
            "doctor": (
                "医生或医馆老师是云诊室关键用户。只有医生/诊所用户表达的认可、兴趣、观望、担忧、拒绝、功能诉求，才算医生反馈。"
            ),
            "market_or_competitor": "竞品、同行、政策、市场趋势必须进入市场雷达，不要污染我方云诊室 AI 结论。",
            "company": "公司内部用 AI 降本增效或改进运营是公司机会，不应误标为医生反馈或销售动作。",
        },
        "analysis_goals": [
            "判断医生对 AI 的整体接纳程度和趋势。",
            "判断哪些区域、医助或销售个人表现突出，以及是广度增加还是少数人高频。",
            "抽取医生直接诉求：明确提出的功能、体验、可靠性、成本、效率等需求。",
            "抽取医生间接机会：医生没直接说 AI，但暴露出可由 AI 解决的工作流、沟通、随访、诊疗质量问题。",
            "区分销售日常 AI 使用：自用提效、对外介绍、学习复盘、话术生成、客户触达。",
            "识别竞品和同行 AI 动作，作为独立市场雷达。",
        ],
        "judgement_rules": [
            "先判断谁在说、说给谁、发生在什么业务动作里，再判断标签。",
            "不要因为文本出现 AI、平台、医生等关键词就直接判定为医生反馈。",
            "药价、剂型、旅游、药房、普通客情等泛业务内容，除非与 AI 能力或平台 AI 场景有明确关系，否则不能算 AI 诉求。",
            "销售个人复盘和医生真实反馈必须分开。",
            "竞品/同行 AI 动作必须单独归入市场雷达。",
            "证据不足时应标记待复核或低置信，不要为了完整性强行归类。",
            "结论必须能追溯到代表原文，并说明反证、不确定性和下一步动作。",
        ],
        "output_quality_bar": {
            "good": "能回答业务问题，说明为什么重要，指出证据和风险，并给出产品或销售管理动作。",
            "bad": "只复述数量、只翻译标签、输出泛泛而谈的管理套话、暴露英文枚举或忽略证据限制。",
        },
    }


def prompt_context_json() -> str:
    return json.dumps(build_prompt_context(), ensure_ascii=False, indent=2)


def render_prompt_reference_markdown() -> str:
    context = build_prompt_context()
    return "\n".join(
        [
            "# v1.6 当前使用 Prompt 说明",
            "",
            f"- Prompt 背景版本：`{PROMPT_CONTEXT_VERSION}`",
            f"- 来源文档：{', '.join(PROMPT_SOURCE_DOCUMENTS)}",
            "- 使用范围：Minimax 边界样本判定、Minimax 证据簇洞察归纳。",
            "- 不使用范围：文件读取、时间窗口、同比/环比、基础聚合、下载归档主链路。",
            "",
            "## 压缩业务背景包",
            "```json",
            json.dumps(context, ensure_ascii=False, indent=2),
            "```",
            "",
            "## 边界样本判定 Prompt 目标",
            "- 先判断谁在说、谁在行动、说给谁、发生在哪个业务动作里。",
            "- 区分医生真实反馈、销售自述、市场观察、竞品动作、公司内部机会。",
            "- 只在证据足够时输出高置信结论；证据不足必须进入复核。",
            "",
            "## 洞察归纳 Prompt 目标",
            "- 按业务结论、证据依据、趋势判断、驱动因素、反证/不确定性、产品含义、销售管理含义、下一步动作输出。",
            "- 禁止只做标签汇总或泛泛总结。",
            "- 结论必须绑定代表原文，不得包装低置信样本。",
        ]
    )
