from __future__ import annotations

from collections import defaultdict
from typing import Dict, Iterable, List, Sequence, Tuple

from .schema import DECISION_CONFIRMED, TASK_STATUS_OPEN, stable_hash

TOPIC_ORDER = ("接受点", "顾虑点", "产品机会", "销售话术", "待验证假设")
LINE_ORDER = ("云诊室", "云管家", "混合", "待判断")


def build_insight_tree(evidence_facts: Sequence[Dict[str, object]]) -> Dict[str, object]:
    """将证据簇组织成 业务线 × 主题 的洞察树。"""
    grouped: Dict[Tuple[str, str], List[Dict[str, object]]] = defaultdict(list)
    for row in evidence_facts:
        business_line = str(row.get("business_line", "")) or "待判断"
        topic = infer_insight_topic(row)
        grouped[(business_line, topic)].append(row)

    tree: Dict[str, object] = {"business_lines": []}
    for business_line in LINE_ORDER:
        topics = []
        for topic in TOPIC_ORDER:
            rows = grouped.get((business_line, topic), [])
            if not rows:
                continue
            topics.append(
                {
                    "topic": topic,
                    "cards": [_build_card(business_line, topic, rows)],
                }
            )
        if topics:
            tree["business_lines"].append({"business_line": business_line, "topics": topics})
    return tree


def flatten_insight_tree(tree: Dict[str, object]) -> List[Dict[str, object]]:
    cards: List[Dict[str, object]] = []
    for line_group in tree.get("business_lines", []):
        for topic_group in line_group.get("topics", []):
            cards.extend(topic_group.get("cards", []))
    return cards


def infer_insight_topic(row: Dict[str, object]) -> str:
    actor = str(row.get("actor_primary", ""))
    scope = str(row.get("ai_scope", ""))
    text = str(row.get("source_text", ""))
    if actor == "潜在 AI 机会":
        return "产品机会"
    if actor == "销售对外介绍":
        return "销售话术"
    concern_keywords = ("担心", "顾虑", "不能", "限制", "风险", "不准", "复杂", "接受不了", "麻烦", "诈骗")
    accept_keywords = ("认可", "感兴趣", "有帮助", "愿意", "好用", "方便", "接受", "信任", "体验")
    if any(keyword in text for keyword in concern_keywords):
        return "顾虑点"
    if any(keyword in text for keyword in accept_keywords) or actor == "医生反馈":
        return "接受点"
    if scope in {"market_trend", "competitor_ai"}:
        return "待验证假设"
    return "待验证假设"


def _build_card(business_line: str, topic: str, rows: Sequence[Dict[str, object]]) -> Dict[str, object]:
    representative_rows = _select_representative_rows(rows)
    evidence_count = len(rows)
    confirmed_count = sum(1 for row in rows if str(row.get("decision_status", "")) == DECISION_CONFIRMED)
    pending_count = sum(1 for row in rows if str(row.get("review_status", "")) == TASK_STATUS_OPEN)
    confidence_level = _confidence_level(evidence_count, confirmed_count, pending_count)
    owner_refs = sorted({str(row.get("salesperson_name", "")) for row in representative_rows if str(row.get("salesperson_name", ""))})
    return {
        "insight_id": stable_hash(business_line, topic, str(evidence_count)),
        "business_line": business_line,
        "topic": topic,
        "title": _build_title(business_line, topic, rows),
        "judgement": _build_judgement(business_line, topic, rows),
        "why_it_matters": _build_why_it_matters(business_line, topic, rows),
        "action_recommendation": _build_action_recommendation(topic, confidence_level, rows),
        "evidence_count": evidence_count,
        "confirmed_evidence_count": confirmed_count,
        "pending_review_count": pending_count,
        "confidence_level": confidence_level,
        "representative_evidence_refs": [
            {
                "report_id": str(row.get("report_id", "")),
                "segment_id": str(row.get("segment_id", "")),
                "source_text": str(row.get("source_text", "")),
                "file_path": str(row.get("file_path", "")),
            }
            for row in representative_rows
        ],
        "owner_refs": owner_refs[:5],
        "is_actionable": topic in {"产品机会", "销售话术"} and confidence_level in {"high", "medium"},
    }


def _select_representative_rows(rows: Sequence[Dict[str, object]]) -> List[Dict[str, object]]:
    def sort_key(row: Dict[str, object]) -> tuple:
        text = str(row.get("source_text", ""))
        return (
            str(row.get("review_status", "")) == TASK_STATUS_OPEN,
            str(row.get("decision_status", "")) != DECISION_CONFIRMED,
            not _is_high_value_text(text),
            -len(text),
        )

    selected = sorted(rows, key=sort_key)[:3]
    return list(selected)


def _build_title(business_line: str, topic: str, rows: Sequence[Dict[str, object]]) -> str:
    if topic == "接受点":
        return f"{business_line}中已出现可复用的 AI 正向接受信号"
    if topic == "顾虑点":
        return f"{business_line}中的 AI 顾虑开始稳定暴露"
    if topic == "产品机会":
        return f"{business_line}里已有可继续产品化的 AI 机会"
    if topic == "销售话术":
        return f"{business_line}的一线 AI 介绍方式正在形成"
    return f"{business_line}中仍有待验证的 AI 假设"


def _build_judgement(business_line: str, topic: str, rows: Sequence[Dict[str, object]]) -> str:
    owner_names = "、".join(sorted({str(row.get("salesperson_name", "")) for row in rows if str(row.get("salesperson_name", ""))})[:3]) or "多位销售"
    if topic == "接受点":
        return f"{owner_names}在 {business_line} 语境中记录到的 AI 信号里，已经出现真实的接受和兴趣表达。"
    if topic == "顾虑点":
        return f"{owner_names}在 {business_line} 语境中反复提到准确度、复杂度或风险顾虑，说明落地阻力并非个别现象。"
    if topic == "产品机会":
        return f"{owner_names}在 {business_line} 语境中持续暴露出值得产品投入的 AI 问题和场景。"
    if topic == "销售话术":
        return f"{owner_names}在 {business_line} 里已经开始形成较稳定的 AI 介绍动作，而不只是偶发提及。"
    return f"{owner_names}在 {business_line} 中提到的 AI 信号仍偏早期，需要继续验证是否能形成稳定方向。"


def _build_why_it_matters(business_line: str, topic: str, rows: Sequence[Dict[str, object]]) -> str:
    if topic == "接受点":
        return f"这说明 {business_line} 中至少已有一部分客户愿意讨论 AI，适合把正向表达沉淀成更标准的销售打法。"
    if topic == "顾虑点":
        return f"这类顾虑会直接影响 {business_line} 的成交和复访，如果不解释清楚，会削弱 AI 的一线可信度。"
    if topic == "产品机会":
        return f"这些问题已经不只是零散抱怨，而是能反哺产品选题和优先级的真实输入。"
    if topic == "销售话术":
        return f"一线是否已经形成成熟话术，决定了 AI 能否从卖点变成常规拓客工具。"
    return "这类内容当前仍不够稳定，但值得保留为下一轮复核和策略更新的输入。"


def _build_action_recommendation(topic: str, confidence_level: str, rows: Sequence[Dict[str, object]]) -> str:
    if confidence_level == "low":
        return "先进入待验证池，优先补复核和补证据。"
    if topic == "产品机会":
        return "进入产品机会池，结合代表证据做需求拆解。"
    if topic == "销售话术":
        return "沉淀为可复用话术候选，并补充成功/失败场景。"
    if topic == "顾虑点":
        return "进入风险解释清单，补充应对话术和产品说明。"
    if topic == "接受点":
        return "沉淀正向案例，反向验证哪些表达最容易触发接受。"
    return "继续观察，等待更多确认信号。"


def _confidence_level(evidence_count: int, confirmed_count: int, pending_count: int) -> str:
    if evidence_count >= 5 and pending_count <= 1 and confirmed_count >= max(3, evidence_count - 1):
        return "high"
    if evidence_count >= 2 and pending_count <= evidence_count // 2:
        return "medium"
    return "low"


def _is_high_value_text(text: str) -> bool:
    weak_keywords = ("资料整理", "模板", "申报", "很重要", "工作安排", "客户跟进")
    return len(text) >= 20 and not any(keyword in text for keyword in weak_keywords)
