from __future__ import annotations

from typing import Dict, List, Tuple

from .schema import (
    ACTOR_KEYWORDS,
    AI_EXPLICIT_KEYWORDS,
    AI_OPPORTUNITY_KEYWORDS,
    BUSINESS_LINE_KEYWORDS,
    DECISION_CONFIRMED,
    DECISION_PENDING_HUMAN,
    DECISION_UNCERTAIN,
)


class Tagger:
    def __init__(self, mode: str = "mock") -> None:
        if mode not in {"mock", "real"}:
            raise ValueError(f"model mode 不支持: {mode}")
        self.mode = mode

    def classify(self, text: str) -> Dict[str, object]:
        # v1.3 首版默认 mock；real 模式当前复用同口径，后续接入真实模型时替换此分支。
        return self._classify_mock(text)

    def _classify_mock(self, text: str) -> Dict[str, object]:
        lower_text = text.lower()

        explicit_hits = [kw for kw in AI_EXPLICIT_KEYWORDS if kw.lower() in lower_text]
        opportunity_hits = [kw for kw in AI_OPPORTUNITY_KEYWORDS if kw in text]
        is_ai_hit = bool(explicit_hits or opportunity_hits)

        business_line, business_reasons = _detect_business_line(text)
        ai_actor, actor_reasons = _detect_ai_actor(text)

        reason_parts: List[str] = []
        if explicit_hits:
            reason_parts.append(f"显式关键词: {','.join(explicit_hits[:3])}")
        if opportunity_hits:
            reason_parts.append(f"机会关键词: {','.join(opportunity_hits[:3])}")
        reason_parts.extend(business_reasons)
        reason_parts.extend(actor_reasons)

        confidence = 0.5
        if explicit_hits:
            confidence += 0.25
        if business_line != "待判断":
            confidence += 0.15
        if ai_actor != "待判断":
            confidence += 0.10
        confidence = min(confidence, 0.99)

        if not is_ai_hit:
            decision_status = DECISION_CONFIRMED
        elif business_line == "待判断" or ai_actor == "待判断":
            decision_status = DECISION_PENDING_HUMAN
            confidence = min(confidence, 0.69)
        elif confidence < 0.70:
            decision_status = DECISION_UNCERTAIN
        else:
            decision_status = DECISION_CONFIRMED

        return {
            "is_ai_hit": is_ai_hit,
            "business_line": business_line,
            "ai_actor": ai_actor,
            "decision_status": decision_status,
            "confidence": round(confidence, 2),
            "reason": "；".join(reason_parts) if reason_parts else "无明确命中",
        }


def _detect_business_line(text: str) -> Tuple[str, List[str]]:
    hits = {}
    for line, keywords in BUSINESS_LINE_KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in text.lower()]
        if matched:
            hits[line] = matched
    if len(hits) >= 2:
        return "混合", [f"业务线双命中: {','.join(sorted(hits.keys()))}"]
    if len(hits) == 1:
        line = next(iter(hits.keys()))
        return line, [f"{line}关键词命中: {','.join(hits[line][:3])}"]
    return "待判断", ["业务线关键词不足"]


def _detect_ai_actor(text: str) -> Tuple[str, List[str]]:
    hit_map: Dict[str, List[str]] = {}
    for actor, keywords in ACTOR_KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in text.lower()]
        if matched:
            hit_map[actor] = matched

    if not hit_map:
        return "待判断", ["主体关键词不足"]

    # 优先级：医生反馈 > 销售对外介绍 > 销售自用 > 潜在 AI 机会
    for actor in ["医生反馈", "销售对外介绍", "销售自用", "潜在 AI 机会"]:
        if actor in hit_map:
            return actor, [f"{actor}关键词命中: {','.join(hit_map[actor][:3])}"]
    return "待判断", ["主体命中冲突"]

