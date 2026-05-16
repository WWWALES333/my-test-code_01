from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import json
import os
import re
from typing import Dict, List, Tuple

import requests

from .schema import (
    ACTOR_LABEL_GAP,
    ACTOR_KEYWORDS,
    ACTION_KEYWORDS,
    ACTOR_PRIMARY_VALUES,
    AI_EXPLICIT_KEYWORDS,
    AI_OPPORTUNITY_KEYWORDS,
    AI_SCOPE_VALUES,
    BUSINESS_LINE_KEYWORDS,
    BUSINESS_LINE_VALUES,
    CONTEXT_BUSINESS_PRIOR_KEYWORDS,
    CONVERSION_KEYWORDS,
    CUSTOMER_ENTITY_HINTS,
    DECISION_CONFIRMED,
    DECISION_PENDING_HUMAN,
    DECISION_UNCERTAIN,
    DECISION_VALUES,
    MODEL_REASON_FAILED,
    NEGATIVE_FEEDBACK_KEYWORDS,
    POSITIVE_FEEDBACK_KEYWORDS,
    REVIEW_REASON_CODE_VALUES,
    SCOPE_KEYWORDS,
    TRIAGE_AUTO_CONFIRM,
    TRIAGE_AUTO_REJECT,
    TRIAGE_NEEDS_LLM,
)


class Tagger:
    def __init__(self, mode: str = "mock") -> None:
        if mode not in {"mock", "real"}:
            raise ValueError(f"model mode 不支持: {mode}")
        self.mode = mode
        self.model_name = os.getenv("OPENAI_MODEL", "").strip()

    def classify(self, text: str, context: Dict[str, object] | None = None) -> Dict[str, object]:
        prepared = self._prepare_base(text, context or {})
        if self.mode == "mock":
            return self._finalize_mock(prepared)
        if str(prepared.get("triage_status", "")) != TRIAGE_NEEDS_LLM:
            return self._finalize_real_passthrough(prepared)
        return self._classify_real_safe(text, context or {}, prepared)

    def classify_batch(
        self,
        items: List[Tuple[str, Dict[str, object]]],
        llm_concurrency: int = 1,
    ) -> List[Dict[str, object]]:
        prepared_items = [(text, context, self._prepare_base(text, context)) for text, context in items]
        if self.mode == "mock":
            return [self._finalize_mock(base) for _, _, base in prepared_items]

        results: List[Dict[str, object]] = [dict() for _ in prepared_items]
        llm_indexes: List[int] = []
        for idx, (text, context, base) in enumerate(prepared_items):
            if str(base.get("triage_status", "")) == TRIAGE_NEEDS_LLM:
                llm_indexes.append(idx)
            else:
                results[idx] = self._finalize_real_passthrough(base)

        if not llm_indexes:
            return results

        if llm_concurrency <= 1 or len(llm_indexes) == 1:
            for idx in llm_indexes:
                text, context, base = prepared_items[idx]
                results[idx] = self._classify_real_safe(text, context, base)
            return results

        max_workers = max(1, llm_concurrency)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_map = {
                executor.submit(self._classify_real_safe, prepared_items[idx][0], prepared_items[idx][1], prepared_items[idx][2]): idx
                for idx in llm_indexes
            }
            for future, idx in future_map.items():
                results[idx] = future.result()
        return results

    def _prepare_base(self, text: str, context: Dict[str, object]) -> Dict[str, object]:
        base = self._classify_mock(text, context)
        base["triage_status"] = _infer_triage_status(text, context, base)
        base["used_label_gap"] = str(base.get("actor_primary", "")).strip() == ACTOR_LABEL_GAP
        base["llm_invoked"] = False
        base["llm_failed"] = False
        base["rule_baseline"] = _build_rule_baseline(base)
        return base

    def _finalize_mock(self, base: Dict[str, object]) -> Dict[str, object]:
        result = dict(base)
        result["model_mode"] = "mock"
        result["model_name"] = "mock-rule-engine"
        return result

    def _finalize_real_passthrough(self, base: Dict[str, object]) -> Dict[str, object]:
        result = dict(base)
        result["model_mode"] = "real"
        result["model_name"] = self.model_name or "hybrid-rule-gate"
        return result

    def _classify_real_safe(
        self,
        text: str,
        context: Dict[str, object],
        base: Dict[str, object],
    ) -> Dict[str, object]:
        try:
            refined = self._classify_real(text, context, base)
            merged = self._merge_result(base, refined)
            merged["model_mode"] = "real"
            merged["model_name"] = self.model_name or "unknown"
            merged["llm_invoked"] = True
            merged["llm_failed"] = False
            return merged
        except Exception as exc:
            fallback = dict(base)
            fallback["decision_status"] = DECISION_PENDING_HUMAN
            fallback["certainty_level"] = "low"
            fallback["review_reason_code"] = _merge_reason_codes(
                str(base.get("review_reason_code", "")),
                [MODEL_REASON_FAILED],
            )
            fallback["reason"] = f"{base.get('reason', '无明确命中')}；real模型失败:{str(exc)[:120]}"
            fallback["model_mode"] = "real"
            fallback["model_name"] = self.model_name or "unknown"
            fallback["llm_invoked"] = True
            fallback["llm_failed"] = True
            return fallback

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

    def _classify_real(
        self,
        text: str,
        context: Dict[str, object],
        base: Dict[str, object],
    ) -> Dict[str, object]:
        api_key = os.getenv("OPENAI_API_KEY", "").strip()
        model = os.getenv("OPENAI_MODEL", "").strip()
        base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").strip().rstrip("/")
        if not api_key or not model:
            raise RuntimeError("OPENAI_API_KEY 或 OPENAI_MODEL 未配置")

        url = f"{base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": model,
            "temperature": 0,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "你是销售周报AI专题边界判定器。"
                        "输出必须是单个 JSON 对象，不要输出 markdown，不要输出解释性前缀。"
                        "你只处理规则无法稳定确认的模糊样本，不要复述规则已经能确定的简单事实。"
                        "字段: is_ai_hit,business_line,actor_primary,ai_scope,decision_status,confidence,reason,used_label_gap,should_review。"
                        "business_line 只能是 云诊室/云管家/混合/待判断。"
                        "actor_primary 只能是 销售自用/销售对外介绍/医生反馈/潜在 AI 机会/label_gap/待判断。"
                        "ai_scope 只能是 product_ai/market_trend/competitor_ai/general_ai。"
                        "如果文本与我方业务 AI 无关，is_ai_hit=false。"
                        "如果现有主体标签不适用，用 actor_primary=label_gap 且 used_label_gap=true。"
                        "不要为了完整性强行归类。"
                    ),
                },
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "segment_text": text,
                            "report_context": {
                                "file_path": context.get("file_path", ""),
                                "report_type": context.get("report_type", ""),
                                "year": context.get("year", 0),
                                "month": context.get("month", 0),
                                "week_of_month": context.get("week_of_month", 0),
                                "owner_hint": context.get("report_owner_name", "") or context.get("owner_hint", ""),
                                "battle_zone_name": context.get("battle_zone_name", ""),
                                "region_name": context.get("region_name", ""),
                            },
                            "rule_baseline": base.get("rule_baseline", _build_rule_baseline(base)),
                            "constraints": {
                                "decision_status_values": sorted(list(DECISION_VALUES)),
                                "business_line_values": sorted(list(BUSINESS_LINE_VALUES)),
                                "actor_primary_values": sorted(list(ACTOR_PRIMARY_VALUES)),
                                "ai_scope_values": sorted(list(AI_SCOPE_VALUES)),
                                "do_not_force_classify": True,
                            },
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
        }

        resp = requests.post(url, headers=headers, json=payload, timeout=30)
        resp.raise_for_status()
        body = resp.json()
        content = (
            body.get("choices", [{}])[0]
            .get("message", {})
            .get("content", "")
        )
        if not content:
            raise RuntimeError("模型返回空内容")
        return _normalize_model_result(_parse_json_payload(content))

    def _merge_result(self, base: Dict[str, object], refined: Dict[str, object]) -> Dict[str, object]:
        merged = dict(base)
        override_keys = [
            "is_ai_hit",
            "business_line",
            "ai_actor",
            "actor_primary",
            "ai_scope",
            "review_reason_code",
            "decision_status",
            "confidence",
            "reason",
            "used_label_gap",
        ]
        for key in override_keys:
            value = refined.get(key)
            if value in (None, ""):
                continue
            merged[key] = value

        merged["actor_subtype"] = str(merged.get("actor_subtype", "")).strip()
        merged["interaction_outcome"] = str(merged.get("interaction_outcome", "")).strip() or "not_applicable"
        merged["certainty_level"] = _confidence_to_certainty(float(merged.get("confidence", base.get("confidence", 0.6))))
        merged["ai_actor"] = merged.get("actor_primary", merged.get("ai_actor", ""))
        merged["used_label_gap"] = bool(merged.get("used_label_gap", False) or str(merged.get("actor_primary", "")).strip() == ACTOR_LABEL_GAP)
        if merged.get("decision_status") not in DECISION_VALUES:
            merged["decision_status"] = DECISION_UNCERTAIN
            merged["review_reason_code"] = _merge_reason_codes(
                str(merged.get("review_reason_code", "")),
                [MODEL_REASON_FAILED],
            )
        return merged


def _parse_json_payload(content: str) -> Dict[str, object]:
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError as exc:
        extracted = _extract_json_object(content)
        if extracted is None:
            raise RuntimeError(f"模型返回非JSON: {exc}") from exc
        try:
            parsed = json.loads(extracted)
        except json.JSONDecodeError as inner_exc:
            raise RuntimeError(f"模型返回非JSON: {inner_exc}") from inner_exc
    if not isinstance(parsed, dict):
        raise RuntimeError("模型返回结构不是JSON对象")
    return parsed


def _normalize_model_result(raw: Dict[str, object]) -> Dict[str, object]:
    is_ai_hit = bool(raw.get("is_ai_hit", False))
    business_line = str(raw.get("business_line", "")).strip()
    if business_line not in BUSINESS_LINE_VALUES:
        business_line = "待判断"

    actor_primary = str(raw.get("actor_primary", "")).strip()
    used_label_gap = bool(raw.get("used_label_gap", False))
    if actor_primary in {"", "现有标签不适用", "待扩增", "不适用"} and used_label_gap:
        actor_primary = ACTOR_LABEL_GAP
    if actor_primary == ACTOR_LABEL_GAP:
        used_label_gap = True
    if actor_primary not in ACTOR_PRIMARY_VALUES:
        actor_primary = "待判断"

    decision_status = str(raw.get("decision_status", "")).strip() or DECISION_UNCERTAIN
    should_review = bool(raw.get("should_review", False))
    if should_review:
        decision_status = DECISION_PENDING_HUMAN
    elif decision_status not in DECISION_VALUES:
        decision_status = DECISION_UNCERTAIN

    try:
        confidence = float(raw.get("confidence", 0.6))
    except Exception:
        confidence = 0.6
    confidence = max(0.0, min(confidence, 1.0))

    ai_scope = str(raw.get("ai_scope", "")).strip()
    if ai_scope not in AI_SCOPE_VALUES:
        ai_scope = "product_ai"
    reason = str(raw.get("reason", "")).strip() or "模型返回未提供原因"
    review_reason_code = ""
    if decision_status == DECISION_PENDING_HUMAN:
        review_reason_code = "SCOPE_AMBIGUOUS"

    return {
        "is_ai_hit": is_ai_hit,
        "business_line": business_line,
        "ai_actor": actor_primary,
        "actor_primary": actor_primary,
        "actor_subtype": "",
        "ai_scope": ai_scope,
        "interaction_outcome": "not_applicable",
        "certainty_level": _confidence_to_certainty(confidence),
        "review_reason_code": review_reason_code,
        "decision_status": decision_status,
        "confidence": round(confidence, 2),
        "reason": reason,
        "used_label_gap": used_label_gap,
    }


def _merge_reason_codes(existing: str, incoming: List[str]) -> str:
    merged = [code for code in existing.split(";") if code]
    merged.extend(incoming)
    merged = _dedupe_non_empty(merged)
    return ";".join([code for code in merged if code in REVIEW_REASON_CODE_VALUES])


def _detect_ai_scope(text: str, lower_text: str, is_ai_hit: bool) -> Tuple[str, List[str], Dict[str, bool]]:
    reasons: List[str] = []
    flags = {"broad_statement": False, "scope_ambiguous": False}
    if not is_ai_hit:
        return "general_ai", ["未命中 AI 关键词"], flags

    competitor_hits = [kw for kw in SCOPE_KEYWORDS["competitor_ai"] if kw.lower() in lower_text]
    market_hits = [kw for kw in SCOPE_KEYWORDS["market_trend"] if kw.lower() in lower_text]
    broad_hits = [kw for kw in SCOPE_KEYWORDS["general_ai"] if kw.lower() in lower_text]
    trend_hits = _collect_extra_hits(
        lower_text,
        [
            "大模型",
            "爆火",
            "热点话题",
            "大爆炸",
            "科技",
            "未来",
            "创业",
            "机器人",
            "行业",
            "潮流",
            "替代",
            "通识课",
            "人工智能大数据",
            "赋能各个行业",
            "走到前沿",
            "只是一个工具",
        ],
    )
    product_hits = _collect_extra_hits(
        lower_text,
        [
            "ai问诊",
            "ai辅助",
            "ai诊疗",
            "ai二维码",
            "ai病历",
            "病历提取",
            "辅助开方",
            "辅助诊疗",
            "问诊助手",
            "诊疗助手",
        ],
    )
    has_action = _contains_any(lower_text, ACTION_KEYWORDS) or _contains_any(
        lower_text,
        [
            "宣传",
            "推荐",
            "推广",
            "聊了",
            "聊下",
            "报名",
            "引导",
            "审核",
            "认证",
            "发朋友圈",
            "让老师",
            "同步",
            "提到",
            "体验",
            "分享",
        ],
    )
    has_customer_entity = _contains_any(lower_text, CUSTOMER_ENTITY_HINTS)

    if competitor_hits:
        reasons.append(f"范围判定: competitor_ai({','.join(competitor_hits[:2])})")
        return "competitor_ai", reasons, flags

    if market_hits or trend_hits:
        merged_hits = market_hits + trend_hits
        reasons.append(f"范围判定: market_trend({','.join(merged_hits[:2])})")
        if not has_action and not has_customer_entity:
            flags["broad_statement"] = True
        return "market_trend", reasons, flags

    if broad_hits and not has_action and not has_customer_entity:
        flags["broad_statement"] = True
        reasons.append("范围判定: general_ai(泛化表达)")
        return "general_ai", reasons, flags

    if not has_action and not product_hits and _contains_any(lower_text, ["普通人", "工具", "智商外挂", "走到前沿", "我们身边"]):
        flags["broad_statement"] = True
        reasons.append("范围判定: general_ai(泛化观点)")
        return "general_ai", reasons, flags

    if "智商外挂" in lower_text and not product_hits:
        flags["broad_statement"] = True
        reasons.append("范围判定: general_ai(泛化比喻)")
        return "general_ai", reasons, flags

    if not has_action and not has_customer_entity and not product_hits:
        flags["broad_statement"] = True
        reasons.append("范围判定: general_ai(宽泛非业务表达)")
        return "general_ai", reasons, flags

    if len(text.strip()) <= 18 and not has_action and not has_customer_entity and not product_hits:
        flags["broad_statement"] = True
        reasons.append("范围判定: general_ai(短句提示)")
        return "general_ai", reasons, flags

    if broad_hits and has_action:
        # 仅有个人思考/使用表达、缺少客体对象时，仍应归 general_ai。
        if not has_customer_entity and not product_hits:
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
    clinic_feature_hits = _collect_extra_hits(
        lower_text,
        [
            "ai诊疗",
            "辅助诊疗",
            "智能问诊",
            "问诊单",
            "病历提取",
            "病历整理",
            "ai病历",
            "诊后随访",
            "诊疗功能",
            "问诊助手",
            "诊疗助手",
        ],
    )
    steward_feature_hits = _collect_extra_hits(
        lower_text,
        [
            "云管家",
            "经营",
            "会员",
            "储值",
            "随访管理",
            "门诊管理",
            "诊所管理",
        ],
    )

    for line, keywords in BUSINESS_LINE_KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in lower_text]
        if matched:
            text_hits[line] = matched

    if "云诊室" in text_hits and "云管家" in text_hits:
        steward_hits = set(text_hits.get("云管家", []))
        if steward_hits and steward_hits <= {"管理"} and (
            clinic_feature_hits or _contains_any(lower_text, ["老师", "医生", "医馆", "卫生院", "问诊", "处方"])
        ):
            text_hits.pop("云管家", None)

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

    if ai_scope == "product_ai":
        if clinic_feature_hits and not steward_feature_hits:
            return "云诊室", [f"产品特征命中: {','.join(clinic_feature_hits[:3])}"]
        if steward_feature_hits and not clinic_feature_hits:
            return "云管家", [f"产品特征命中: {','.join(steward_feature_hits[:3])}"]

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

    customer_entity = _contains_any(lower_text, CUSTOMER_ENTITY_HINTS)
    sales_intro = _contains_any(lower_text, ACTOR_KEYWORDS["销售对外介绍"]) or _contains_any(
        lower_text,
        [
            "宣传",
            "推荐",
            "推广",
            "聊了",
            "聊下",
            "报名",
            "引导",
            "审核",
            "认证",
            "发朋友圈",
            "让老师",
            "扫码",
            "体验",
            "介绍了ai",
            "ai二维码",
        ],
    )
    doctor_feedback = _contains_any(lower_text, ACTOR_KEYWORDS["医生反馈"]) or _looks_like_doctor_feedback(lower_text) or (
        customer_entity
        and _contains_any(lower_text, ["喜欢", "建议", "关注", "反馈", "说", "表示", "认可", "担心", "感兴趣"])
    )
    explicit_feedback = _looks_like_doctor_feedback(lower_text) or _contains_any(
        lower_text,
        ["感兴趣", "很不错", "还不错", "认可", "喜欢", "体验一下", "想试试", "担心", "婉拒", "不好用", "不靠谱"],
    )
    sales_self_use = _contains_any(lower_text, ACTOR_KEYWORDS["销售自用"]) or (
        not customer_entity
        and _contains_any(lower_text, ["我个人觉得", "我觉得", "我认为", "关注", "研究", "学习", "梳理", "复盘", "需求重新梳理", "很受启发", "期待"])
    )
    opportunity = _contains_any(lower_text, ACTOR_KEYWORDS["潜在 AI 机会"]) or "降本增效" in lower_text or _contains_any(
        lower_text,
        ["需求", "机会", "适合", "痛点", "开发课程", "产品经理", "大杀器", "转化", "引流", "赋能", "信心", "前沿"],
    )
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
    elif sales_intro and doctor_feedback and explicit_feedback:
        primary = "医生反馈"
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
    elif sales_intro:
        primary = "销售对外介绍"
        subtypes.append("销售动作_弱反馈")

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
    if ai_scope in {"general_ai", "market_trend"}:
        codes = [code for code in codes if code not in {"ACTOR_OVERLAP", "BUSINESSLINE_LOW_SIGNAL"}]
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
    if actor_primary == "待判断" and ai_scope in {"competitor_ai", "market_trend", "general_ai"}:
        return DECISION_CONFIRMED
    if ai_scope in {"general_ai", "market_trend"} and actor_primary in {"销售自用", "潜在 AI 机会"}:
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
    return any(
        token in text
        for token in [
            "反馈",
            "表示",
            "觉得",
            "发文章",
            "观望",
            "不成熟",
            "感兴趣",
            "很不错",
            "还不错",
            "认可",
            "喜欢",
            "体验一下",
            "想试试",
            "不好用",
            "不靠谱",
            "担心",
            "婉拒",
        ]
    )


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


def _collect_extra_hits(text: str, keywords: List[str]) -> List[str]:
    return [kw for kw in keywords if kw.lower() in text]


def _extract_json_object(content: str) -> str | None:
    start = content.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escaped = False
    for idx in range(start, len(content)):
        ch = content[idx]
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return content[start : idx + 1]
    return None


def _build_rule_baseline(base: Dict[str, object]) -> Dict[str, object]:
    return {
        "is_ai_hit": bool(base.get("is_ai_hit", False)),
        "business_line": str(base.get("business_line", "待判断")),
        "actor_primary": str(base.get("actor_primary", "待判断")),
        "ai_scope": str(base.get("ai_scope", "product_ai")),
        "decision_status": str(base.get("decision_status", DECISION_PENDING_HUMAN)),
        "review_reason_code": str(base.get("review_reason_code", "")),
        "confidence": float(base.get("confidence", 0.0) or 0.0),
        "reason": str(base.get("reason", "")),
    }


def _infer_triage_status(text: str, context: Dict[str, object], base: Dict[str, object]) -> str:
    if not bool(base.get("is_ai_hit", False)) and str(base.get("decision_status", "")) == DECISION_CONFIRMED:
        return TRIAGE_AUTO_REJECT

    reason_codes = {
        code.strip()
        for code in str(base.get("review_reason_code", "")).split(";")
        if code.strip()
    }
    actor_primary = str(base.get("actor_primary", "")).strip()
    business_line = str(base.get("business_line", "")).strip()
    ai_scope = str(base.get("ai_scope", "")).strip()
    decision_status = str(base.get("decision_status", "")).strip()
    lower_text = text.lower()
    lower_context = json.dumps(context, ensure_ascii=False).lower()

    if _is_obvious_product_ai_sample(lower_text, lower_context, base):
        return TRIAGE_AUTO_CONFIRM

    if decision_status == DECISION_PENDING_HUMAN:
        return TRIAGE_NEEDS_LLM
    if reason_codes & {"ACTOR_OVERLAP", "BUSINESSLINE_LOW_SIGNAL", "SCOPE_AMBIGUOUS", "BROAD_STATEMENT"}:
        return TRIAGE_NEEDS_LLM
    if actor_primary in {"", "待判断", ACTOR_LABEL_GAP}:
        return TRIAGE_NEEDS_LLM
    if business_line == "待判断":
        return TRIAGE_NEEDS_LLM
    if _count_actor_signals(lower_text) >= 2:
        return TRIAGE_NEEDS_LLM

    has_business_action = _contains_any(
        lower_text,
        ACTION_KEYWORDS
        + ["推荐", "问候维护", "加微信", "体验", "宣传", "推广", "聊下", "同步", "拜访", "回访", "演示"],
    )
    has_customer_context = _contains_any(lower_text, CUSTOMER_ENTITY_HINTS + ["平台", "老师", "医馆", "卫生院"])
    if has_business_action and (actor_primary in {"", "待判断"} or business_line == "待判断"):
        return TRIAGE_NEEDS_LLM
    if ai_scope in {"general_ai", "market_trend"} and has_customer_context and has_business_action:
        return TRIAGE_NEEDS_LLM
    if ai_scope in {"general_ai", "market_trend"} and _contains_any(lower_context, ["老师", "医生", "医馆", "平台"]):
        return TRIAGE_NEEDS_LLM

    non_blocking_reason_codes = reason_codes - {"OUTCOME_UNCLEAR"}
    if (
        decision_status == DECISION_CONFIRMED
        and actor_primary not in {"", "待判断"}
        and business_line != "待判断"
        and not non_blocking_reason_codes
    ):
        return TRIAGE_AUTO_CONFIRM
    return TRIAGE_NEEDS_LLM


def _is_obvious_product_ai_sample(lower_text: str, lower_context: str, base: Dict[str, object]) -> bool:
    if not bool(base.get("is_ai_hit", False)):
        return False
    if str(base.get("ai_scope", "")).strip() != "product_ai":
        return False

    actor_primary = str(base.get("actor_primary", "")).strip()
    business_line = str(base.get("business_line", "")).strip()
    if actor_primary not in {"医生反馈", "销售对外介绍"}:
        return False
    if business_line not in {"云诊室", "云管家"}:
        return False

    combined_text = f"{lower_text} {lower_context}"
    has_customer_context = _contains_any(
        combined_text,
        CUSTOMER_ENTITY_HINTS + ["医馆", "卫生院", "医院", "平台", "门诊"],
    )
    has_product_feature = _contains_any(
        combined_text,
        [
            "ai诊疗",
            "辅助诊疗",
            "智能问诊",
            "问诊单",
            "病历提取",
            "ai病历",
            "病历整理",
            "诊后随访",
            "诊疗功能",
            "问诊助手",
            "诊疗助手",
            "云管家ai",
        ],
    )
    has_intro_signal = _contains_any(
        combined_text,
        ["介绍", "演示", "推荐", "回访", "体验", "跟老师提了", "给医生介绍", "向医生介绍"],
    )
    has_feedback_signal = _looks_like_doctor_feedback(combined_text) or _contains_any(
        combined_text,
        ["感兴趣", "很不错", "还不错", "认可", "喜欢", "体验一下", "想试试", "担心", "婉拒", "不好用", "不靠谱"],
    )

    if actor_primary == "医生反馈" and has_customer_context and has_product_feature and has_feedback_signal:
        return True
    if actor_primary == "销售对外介绍" and has_customer_context and has_product_feature and has_intro_signal:
        return True
    if has_customer_context and has_product_feature and has_intro_signal and has_feedback_signal:
        return True
    return False


def _count_actor_signals(lower_text: str) -> int:
    checks = [
        _contains_any(lower_text, ACTOR_KEYWORDS["销售对外介绍"]) or _contains_any(lower_text, ["推荐", "回访", "演示", "体验"]),
        _contains_any(lower_text, ACTOR_KEYWORDS["医生反馈"]) or _looks_like_doctor_feedback(lower_text),
        _contains_any(lower_text, ACTOR_KEYWORDS["销售自用"]),
        _contains_any(lower_text, ACTOR_KEYWORDS["潜在 AI 机会"]),
    ]
    return sum(1 for matched in checks if matched)


def _confidence_to_certainty(confidence: float) -> str:
    if confidence >= 0.86:
        return "high"
    if confidence >= 0.70:
        return "medium"
    return "low"
