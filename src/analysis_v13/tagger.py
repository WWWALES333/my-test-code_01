from __future__ import annotations

from typing import Dict, List, Tuple

from .schema import (
    ACTOR_KEYWORDS,
    ACTION_KEYWORDS,
    AI_EXPLICIT_KEYWORDS,
    AI_OPPORTUNITY_KEYWORDS,
    BUSINESS_LINE_KEYWORDS,
    CONTEXT_BUSINESS_PRIOR_KEYWORDS,
    CONVERSION_KEYWORDS,
    CUSTOMER_ENTITY_HINTS,
    DECISION_CONFIRMED,
    DECISION_PENDING_HUMAN,
    NEGATIVE_FEEDBACK_KEYWORDS,
    POSITIVE_FEEDBACK_KEYWORDS,
    SCOPE_KEYWORDS,
)


class Tagger:
    def __init__(self, mode: str = "mock") -> None:
        if mode not in {"mock", "real"}:
            raise ValueError(f"model mode 不支持: {mode}")
        self.mode = mode

    def classify(self, text: str, context: Dict[str, object] | None = None) -> Dict[str, object]:
        # v1.3 首版默认 mock；real 模式当前复用同口径，后续接入真实模型时替换此分支。
        return self._classify_mock(text, context or {})

    def _classify_mock(self, text: str, context: Dict[str, object]) -> Dict[str, object]:
        lower_text = text.lower()
        context_text = str(context.get("file_path", ""))

        explicit_hits = [kw for kw in AI_EXPLICIT_KEYWORDS if kw.lower() in lower_text]
        opportunity_hits = [kw for kw in AI_OPPORTUNITY_KEYWORDS if kw in text]
        is_ai_hit = bool(explicit_hits or opportunity_hits)

        ai_scope, scope_reasons, scope_flags = _detect_ai_scope(text, lower_text, is_ai_hit)
        business_line, business_reasons = _detect_business_line(text, context_text, ai_scope)
        actor_primary, actor_subtype, actor_reasons, actor_flags = _detect_actor(text, lower_text, ai_scope)
        interaction_outcome, outcome_reasons = _detect_interaction_outcome(text, actor_flags)
        review_reason_codes = _infer_review_reason_codes(
            ai_scope=ai_scope,
            business_line=business_line,
            actor_primary=actor_primary,
            interaction_outcome=interaction_outcome,
            scope_flags=scope_flags,
            actor_flags=actor_flags,
        )
        certainty_level, confidence = _infer_certainty(
            is_ai_hit=is_ai_hit,
            ai_scope=ai_scope,
            business_line=business_line,
            actor_primary=actor_primary,
            review_reason_codes=review_reason_codes,
            explicit_hits=explicit_hits,
        )
        decision_status = _infer_decision_status(
            is_ai_hit=is_ai_hit,
            ai_scope=ai_scope,
            business_line=business_line,
            actor_primary=actor_primary,
            interaction_outcome=interaction_outcome,
            certainty_level=certainty_level,
            review_reason_codes=review_reason_codes,
        )

        reason_parts: List[str] = []
        if explicit_hits:
            reason_parts.append(f"显式关键词: {','.join(explicit_hits[:3])}")
        if opportunity_hits:
            reason_parts.append(f"机会关键词: {','.join(opportunity_hits[:3])}")
        reason_parts.extend(scope_reasons)
        reason_parts.extend(business_reasons)
        reason_parts.extend(actor_reasons)
        reason_parts.extend(outcome_reasons)
        if review_reason_codes:
            reason_parts.append(f"复核原因码: {','.join(review_reason_codes)}")

        return {
            "is_ai_hit": is_ai_hit,
            "business_line": business_line,
            "ai_actor": actor_primary,
            "actor_primary": actor_primary,
            "actor_subtype": ";".join(actor_subtype[:2]),
            "ai_scope": ai_scope,
            "interaction_outcome": interaction_outcome,
            "certainty_level": certainty_level,
            "review_reason_code": ";".join(review_reason_codes),
            "decision_status": decision_status,
            "confidence": round(confidence, 2),
            "reason": "；".join(reason_parts) if reason_parts else "无明确命中",
        }


def _detect_ai_scope(text: str, lower_text: str, is_ai_hit: bool) -> Tuple[str, List[str], Dict[str, bool]]:
    reasons: List[str] = []
    flags = {"broad_statement": False, "scope_ambiguous": False}
    if not is_ai_hit:
        return "general_ai", ["未命中 AI 关键词"], flags

    competitor_hits = [kw for kw in SCOPE_KEYWORDS["competitor_ai"] if kw.lower() in lower_text]
    market_hits = [kw for kw in SCOPE_KEYWORDS["market_trend"] if kw.lower() in lower_text]
    broad_hits = [kw for kw in SCOPE_KEYWORDS["general_ai"] if kw.lower() in lower_text]
    has_action = _contains_any(lower_text, ACTION_KEYWORDS)

    if competitor_hits:
        reasons.append(f"范围判定: competitor_ai({','.join(competitor_hits[:2])})")
        return "competitor_ai", reasons, flags

    if market_hits:
        reasons.append(f"范围判定: market_trend({','.join(market_hits[:2])})")
        if not has_action:
            flags["broad_statement"] = True
        return "market_trend", reasons, flags

    if broad_hits and not has_action:
        flags["broad_statement"] = True
        reasons.append("范围判定: general_ai(泛化表达)")
        return "general_ai", reasons, flags

    if broad_hits and has_action:
        # 仅有个人思考/使用表达、缺少客体对象时，仍应归 general_ai。
        has_customer_entity = _contains_any(lower_text, CUSTOMER_ENTITY_HINTS)
        if not has_customer_entity:
            flags["broad_statement"] = True
            reasons.append("范围判定: general_ai(宽泛个人表达)")
            return "general_ai", reasons, flags
        flags["scope_ambiguous"] = True
        reasons.append("范围判定: product_ai(含业务动作)")
        return "product_ai", reasons, flags

    reasons.append("范围判定: product_ai(默认)")
    return "product_ai", reasons, flags


def _detect_business_line(text: str, context_text: str, ai_scope: str) -> Tuple[str, List[str]]:
    lower_text = text.lower()
    lower_ctx = context_text.lower()
    text_hits = {}
    prior_hits = {}

    for line, keywords in BUSINESS_LINE_KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in lower_text]
        if matched:
            text_hits[line] = matched

    for line, keywords in CONTEXT_BUSINESS_PRIOR_KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in lower_ctx]
        if matched:
            prior_hits[line] = matched

    merged_lines = set(text_hits.keys()) | set(prior_hits.keys())
    if len(merged_lines) >= 2:
        return "混合", [f"业务线双命中: {','.join(sorted(merged_lines))}"]

    if len(text_hits) == 1:
        line = next(iter(text_hits.keys()))
        # 政策/趋势语句通常不只作用于单一业务线，优先给混合避免误收窄。
        if ai_scope in {"market_trend", "general_ai"} and line == "云诊室":
            return "混合", ["趋势表达优先按混合业务线处理"]
        return line, [f"{line}关键词命中: {','.join(text_hits[line][:3])}"]

    if ai_scope in {"market_trend", "general_ai"}:
        return "混合", ["范围为趋势/泛AI，业务线按混合兜底"]

    if len(prior_hits) == 1:
        line = next(iter(prior_hits.keys()))
        return line, [f"上下文先验命中: {line}"]

    if ai_scope in {"product_ai", "general_ai"} and not _contains_any(lower_text, CUSTOMER_ENTITY_HINTS):
        return "混合", ["缺少明确客体对象，业务线按混合兜底"]

    return "待判断", ["业务线关键词不足"]


def _detect_actor(
    text: str,
    lower_text: str,
    ai_scope: str,
) -> Tuple[str, List[str], List[str], Dict[str, bool]]:
    reasons: List[str] = []
    subtypes: List[str] = []
    flags = {"actor_overlap": False}

    sales_intro = _contains_any(lower_text, ACTOR_KEYWORDS["销售对外介绍"])
    doctor_feedback = _contains_any(lower_text, ACTOR_KEYWORDS["医生反馈"]) or _looks_like_doctor_feedback(lower_text)
    sales_self_use = _contains_any(lower_text, ACTOR_KEYWORDS["销售自用"])
    opportunity = _contains_any(lower_text, ACTOR_KEYWORDS["潜在 AI 机会"]) or "降本增效" in lower_text
    competitor_product_hint = _contains_any(lower_text, ["ai诊疗", "ai问诊", "ai辅助", "问诊助手", "诊疗助手", "沟通"])

    if "ai搜索" in lower_text:
        subtypes.append("客户AI搜索线索")
    if ai_scope == "market_trend":
        subtypes.append("行业趋势观察")
    if ai_scope == "competitor_ai":
        subtypes.append("竞品AI动态")
    if ai_scope == "general_ai":
        subtypes.append("销售个人思考")
    if opportunity:
        subtypes.append("效率提效机会")

    primary = "待判断"
    if ai_scope == "competitor_ai" and not competitor_product_hint and not sales_intro:
        primary = "待判断"
        subtypes.append("竞品AI动态")
    elif ai_scope == "competitor_ai" and competitor_product_hint:
        primary = "销售对外介绍"
        if doctor_feedback:
            subtypes.append("销售介绍后收到反馈")
    elif sales_intro and doctor_feedback:
        primary = "销售对外介绍"
        subtypes.append("销售介绍后收到反馈")
    elif sales_intro:
        primary = "销售对外介绍"
        subtypes.append("销售介绍_无明确反馈")
    elif doctor_feedback:
        primary = "医生反馈"
    elif ai_scope == "general_ai" and opportunity:
        primary = "潜在 AI 机会"
    elif ai_scope == "general_ai":
        primary = "销售自用"
    elif sales_self_use:
        primary = "销售自用"
    elif opportunity:
        primary = "潜在 AI 机会"
    elif ai_scope == "market_trend":
        primary = "潜在 AI 机会"
    elif ai_scope == "competitor_ai":
        primary = "待判断"

    if primary == "待判断":
        flags["actor_overlap"] = True
        reasons.append("主体关键词不足")
    else:
        reasons.append(f"主体判定: {primary}")

    if sales_intro and doctor_feedback:
        reasons.append("主体复合: 销售介绍 + 医生反馈")

    subtypes = _dedupe_non_empty(subtypes)
    return primary, subtypes, reasons, flags


def _detect_interaction_outcome(text: str, actor_flags: Dict[str, bool]) -> Tuple[str, List[str]]:
    lower_text = text.lower()
    reasons: List[str] = []
    if _contains_any(lower_text, CONVERSION_KEYWORDS):
        reasons.append("结果判定: converted")
        return "converted", reasons
    if _contains_any(lower_text, POSITIVE_FEEDBACK_KEYWORDS):
        reasons.append("结果判定: positive_feedback")
        return "positive_feedback", reasons
    if _contains_any(lower_text, NEGATIVE_FEEDBACK_KEYWORDS):
        reasons.append("结果判定: negative_or_observing")
        return "negative_or_observing", reasons
    if _contains_any(lower_text, ACTOR_KEYWORDS["销售对外介绍"]):
        reasons.append("结果判定: no_feedback")
        return "no_feedback", reasons
    reasons.append("结果判定: not_applicable")
    return "not_applicable", reasons


def _infer_review_reason_codes(
    ai_scope: str,
    business_line: str,
    actor_primary: str,
    interaction_outcome: str,
    scope_flags: Dict[str, bool],
    actor_flags: Dict[str, bool],
) -> List[str]:
    codes: List[str] = []
    if actor_primary == "待判断" or actor_flags.get("actor_overlap"):
        codes.append("ACTOR_OVERLAP")
    if scope_flags.get("scope_ambiguous"):
        codes.append("SCOPE_AMBIGUOUS")
    if scope_flags.get("broad_statement"):
        codes.append("BROAD_STATEMENT")
    if business_line == "待判断":
        codes.append("BUSINESSLINE_LOW_SIGNAL")
    if interaction_outcome == "no_feedback" and ai_scope == "product_ai":
        codes.append("OUTCOME_UNCLEAR")
    return _dedupe_non_empty(codes)


def _infer_certainty(
    is_ai_hit: bool,
    ai_scope: str,
    business_line: str,
    actor_primary: str,
    review_reason_codes: List[str],
    explicit_hits: List[str],
) -> Tuple[str, float]:
    if not is_ai_hit:
        return "high", 0.95

    confidence = 0.60
    if explicit_hits:
        confidence += 0.15
    if ai_scope != "general_ai":
        confidence += 0.10
    if business_line != "待判断":
        confidence += 0.08
    if actor_primary != "待判断":
        confidence += 0.08
    confidence -= 0.07 * len(review_reason_codes)
    confidence = max(0.45, min(confidence, 0.99))

    if confidence >= 0.86:
        return "high", confidence
    if confidence >= 0.70:
        return "medium", confidence
    return "low", confidence


def _infer_decision_status(
    is_ai_hit: bool,
    ai_scope: str,
    business_line: str,
    actor_primary: str,
    interaction_outcome: str,
    certainty_level: str,
    review_reason_codes: List[str],
) -> str:
    if not is_ai_hit:
        return DECISION_CONFIRMED
    if certainty_level == "low":
        return DECISION_PENDING_HUMAN
    if actor_primary == "待判断" and ai_scope == "competitor_ai":
        return DECISION_CONFIRMED
    if actor_primary == "待判断":
        return DECISION_PENDING_HUMAN
    if "SCOPE_AMBIGUOUS" in review_reason_codes or "ACTOR_OVERLAP" in review_reason_codes:
        return DECISION_PENDING_HUMAN
    if (
        "BUSINESSLINE_LOW_SIGNAL" in review_reason_codes
        and ai_scope == "product_ai"
        and interaction_outcome in {"no_feedback", "not_applicable"}
    ):
        return DECISION_PENDING_HUMAN
    return DECISION_CONFIRMED


def _contains_any(text: str, keywords: List[str]) -> bool:
    return any(kw.lower() in text for kw in keywords)


def _looks_like_doctor_feedback(text: str) -> bool:
    if "医生" not in text and "老师" not in text:
        return False
    return any(token in text for token in ["反馈", "表示", "觉得", "发文章", "观望", "不成熟"])


def _dedupe_non_empty(items: List[str]) -> List[str]:
    seen = set()
    result: List[str] = []
    for item in items:
        token = item.strip()
        if not token or token in seen:
            continue
        seen.add(token)
        result.append(token)
    return result
