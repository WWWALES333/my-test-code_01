from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

from .schema import OWNER_TYPE_GROUP, OWNER_TYPE_PERSON, OWNER_TYPE_UNKNOWN, stable_hash

GROUP_KEYWORDS = (
    "战区",
    "区域",
    "部门",
    "周报",
    "月报",
    "模板",
    "工作周报",
    "工作月报",
    "云管家",
    "云诊室",
    "SaaS",
    "saas",
    "工作",
    "计划",
    "总结",
    "思考",
    "业绩",
    "签约",
    "新增",
    "审核",
    "情况",
    "内容",
    "客户",
    "医生",
    "老师",
    "诊所",
    "医院",
    "医馆",
    "卫生所",
    "门诊",
    "平台",
    "人员",
    "意向",
    "副本",
    "售后",
    "处理",
    "完成",
    "心得",
    "中医",
    "北京",
)
PERSON_STOPWORDS = {
    "整理",
    "提交",
    "模板",
    "周报",
    "月报",
    "区域",
    "战区",
    "部门",
    "将军汤",
    "审核",
    "新增",
}


def extract_owner_hint(file_path: str) -> str:
    """从文件路径中提取最稳定的所有者提示信息。"""
    stem = Path(file_path).stem
    zone_patterns = [
        r"[一二三四五六七八九十0-9]+战区[（(][^）)]+[）)]",
        r"[一二三四五六七八九十0-9]+战区[\u4e00-\u9fa5]{0,12}",
        r"线上战区",
        r"[\u4e00-\u9fa5]{2,10}区域",
    ]
    for pattern in zone_patterns:
        matched = re.search(pattern, stem)
        if matched:
            return matched.group(0).strip()

    trailing_name = re.search(r"[-_—]\s*([\u4e00-\u9fa5]{2,6})\s*$", stem)
    if trailing_name:
        return trailing_name.group(1).strip()

    parts = re.split(r"[-_—]", stem)
    for part in reversed(parts):
        token = _clean_owner_token(part)
        if token:
            return token
    token = _clean_owner_token(stem)
    return token if token else stem[:30]


def build_owner_registry(report_rows: Iterable[Dict[str, object]], extra_owner_hints: Sequence[str] | None = None) -> List[Dict[str, object]]:
    """基于报告级信息构建销售/组织归一注册表。"""
    registry: Dict[str, Dict[str, object]] = {}
    for row in report_rows:
        file_path = str(row.get("file_path", ""))
        owner_hint = extract_owner_hint(file_path)
        owner_record = infer_owner_record(owner_hint, file_path)
        existing = registry.get(owner_record["salesperson_id"])
        if not existing:
            registry[owner_record["salesperson_id"]] = owner_record
            continue
        aliases = set(existing.get("aliases", []))
        aliases.update(owner_record.get("aliases", []))
        existing["aliases"] = sorted(aliases)
        source_paths = int(existing.get("source_paths", 0))
        existing["source_paths"] = source_paths + 1

    for owner_hint in extra_owner_hints or []:
        candidate = canonicalize_owner_name(owner_hint)
        if not candidate:
            continue
        owner_record = infer_owner_record(candidate, "")
        existing = registry.get(owner_record["salesperson_id"])
        if not existing:
            registry[owner_record["salesperson_id"]] = owner_record
            continue
        aliases = set(existing.get("aliases", []))
        aliases.update(owner_record.get("aliases", []))
        existing["aliases"] = sorted(aliases)
    return sorted(registry.values(), key=lambda item: (str(item["owner_type"]), str(item["salesperson_name"])))


def infer_owner_record(owner_hint: str, file_path: str) -> Dict[str, object]:
    """把原始 owner_hint 归一为销售对象或组织对象。"""
    candidate = canonicalize_owner_name(owner_hint)
    owner_type = _infer_owner_type(candidate)
    salesperson_name = candidate
    if owner_type == OWNER_TYPE_PERSON:
        salesperson_name = candidate
    elif owner_type == OWNER_TYPE_GROUP:
        salesperson_name = candidate
    else:
        salesperson_name = candidate or "未识别对象"

    salesperson_id = stable_hash(owner_type, salesperson_name, length=12)
    return {
        "salesperson_id": salesperson_id,
        "salesperson_name": salesperson_name,
        "owner_hint": candidate,
        "owner_type": owner_type,
        "aliases": [candidate] if candidate else [],
        "source_paths": 1,
        "example_file_path": file_path,
    }


def canonicalize_owner_name(owner_hint: str) -> str:
    """尽量把各种 owner_hint 清洗成可聚合的名称。"""
    token = owner_hint.strip()
    token = token.strip("【】[]（）()")
    token = token.replace("  ", " ")
    token = re.sub(r"^\d{4}年?\d{0,2}月?.*", "", token) if "模板" in token else token
    token = re.sub(r"^202\d年?", "", token)
    token = re.sub(r"^20\d{2}", "", token)
    token = token.replace("工作周报", "").replace("工作月报", "")
    token = token.replace("周报", "").replace("月报", "")
    token = token.replace("将军汤", "").strip()
    token = token.replace("（", "").replace("）", "")
    token = token.replace("(", "").replace(")", "")
    token = re.sub(r"\s+", " ", token)
    token = token.strip("-_— ")
    if token.endswith("整理") and len(token) <= 5:
        token = token[:-2]
    return token.strip()


def _infer_owner_type(candidate: str) -> str:
    token = candidate.strip()
    if not token:
        return OWNER_TYPE_UNKNOWN
    if any(keyword in token for keyword in GROUP_KEYWORDS):
        return OWNER_TYPE_GROUP
    if re.search(r"[A-Za-z]", token):
        return OWNER_TYPE_GROUP
    if token.endswith("老师"):
        return OWNER_TYPE_GROUP
    if re.fullmatch(r"[\u4e00-\u9fa5]{2,4}", token) and token not in PERSON_STOPWORDS:
        return OWNER_TYPE_PERSON
    return OWNER_TYPE_GROUP if len(token) > 4 else OWNER_TYPE_UNKNOWN


def _clean_owner_token(raw: str) -> str:
    token = raw.strip()
    token = re.sub(r"20\d{2}", "", token)
    token = re.sub(r"\d{1,2}月", "", token)
    token = re.sub(r"第?\d{1,2}周", "", token)
    token = re.sub(r"[（(].*?[）)]", "", token)
    token = re.sub(r"[0-9. ]+", "", token)
    token = token.strip(" -_—")
    if not token:
        return ""
    return token[:24]
