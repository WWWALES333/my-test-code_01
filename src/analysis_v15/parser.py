from __future__ import annotations

import re
from typing import Dict, List

SPEAKER_STOPWORDS = {
    "周一",
    "周二",
    "周三",
    "周四",
    "周五",
    "周六",
    "周日",
    "星期一",
    "星期二",
    "星期三",
    "星期四",
    "星期五",
    "星期六",
    "星期日",
    "上午",
    "下午",
    "关键词",
    "其他",
    "方向",
    "建议人",
    "存在问题",
    "建设性意见与措施",
    "工作复盘与规划",
    "本周工作回顾",
    "本周工作情况",
    "本周工作",
    "本周思考",
    "重点客户",
    "潜力客户情况",
    "开发审核",
    "维护拔高",
    "流失激活",
    "竞品跟进",
    "云管家相关",
    "工作情况",
    "实际业绩",
    "实际审核",
    "实际新增",
    "云诊室",
    "云管家",
    "下降原因",
    "增长原因",
    "下降区域",
    "增长区域",
    "优秀人员及事件",
    "内训提升",
    "资源异动",
    "组员异动",
    "竞品异动",
    "本月战区会议",
    "本月陪访次数",
    "行动计划",
    "完成进度",
    "业务",
    "内务",
    "具体安排",
    "具体计划",
    "云管家审核",
    "云诊室审核",
    "下滑原因",
    "下滑思考",
    "反思",
    "总结",
    "填表人",
    "工作内容",
    "工作总结",
    "工作计划",
    "工作目标",
    "本周主要工作",
    "本周工作相关情况",
    "本月saas业绩",
    "本周saas业绩",
    "本月saas签约",
    "本周云管家签约",
    "新增",
    "周会启示",
    "本周周会心得",
    "其他收获和思考",
    "其他收获",
    "本周收获",
    "本周客户情况",
    "本周数据",
}

SPEAKER_INVALID_KEYWORDS = (
    "工作",
    "业绩",
    "签约",
    "新增",
    "审核",
    "计划",
    "总结",
    "思考",
    "情况",
    "内容",
    "目标",
    "回顾",
    "周会",
    "启示",
    "客户",
    "医生",
    "老师",
    "诊所",
    "医院",
    "医馆",
    "卫生所",
    "门诊",
    "云管家",
    "云诊室",
    "SaaS",
    "saas",
    "AI",
    "ai",
)


def segment_text_with_owner(text: str, fallback_owner: str) -> List[Dict[str, str]]:
    """按正文中的人员小节切分片段，并尽量给每段分配负责人。"""
    if not text.strip():
        return []

    normalized = re.sub(r"\r\n?", "\n", text)
    normalized = re.sub(r"[ \t]+", " ", normalized)
    lines = [line.strip() for line in re.split(r"\n+", normalized) if line.strip()]

    segments: List[Dict[str, str]] = []
    current_owner = fallback_owner
    buffer: List[str] = []

    for line in lines:
        speaker_match = _extract_speaker(line)
        if speaker_match:
            if buffer:
                segments.extend(_buffer_to_segments(buffer, current_owner))
                buffer = []
            current_owner, remainder = speaker_match
            if remainder:
                buffer.append(remainder)
            continue

        if _is_heading_noise(line):
            if buffer:
                segments.extend(_buffer_to_segments(buffer, current_owner))
                buffer = []
            continue
        buffer.append(line)

    if buffer:
        segments.extend(_buffer_to_segments(buffer, current_owner))

    if not segments:
        return [{"owner_hint": fallback_owner, "source_text": normalized.strip()}]
    return segments


def _extract_speaker(line: str) -> tuple[str, str] | None:
    cleaned = line.strip().lstrip("-•·1234567890.、）) ")
    matched = re.match(r"^([A-Za-z\u4e00-\u9fa5]{2,8})[：:]\s*(.*)$", cleaned)
    if not matched:
        return None
    candidate = matched.group(1).strip()
    remainder = matched.group(2).strip()
    if not _is_valid_speaker(candidate):
        return None
    return candidate, remainder


def _is_valid_speaker(candidate: str) -> bool:
    if candidate in SPEAKER_STOPWORDS:
        return False
    if any(keyword in candidate for keyword in SPEAKER_INVALID_KEYWORDS):
        return False
    if "老师" in candidate:
        return False
    if any(char.isdigit() for char in candidate):
        return False
    if re.search(r"[A-Za-z]", candidate):
        return False
    if len(candidate) <= 1:
        return False
    if len(candidate) > 4:
        return False
    return True


def _is_heading_noise(line: str) -> bool:
    token = line.strip()
    if len(token) <= 14 and token in SPEAKER_STOPWORDS:
        return True
    if re.fullmatch(r"[一二三四五六七八九十0-9]+[、.)）].*", token):
        return False
    return token in {"协同", "专业", "突破", "降本增效", "协同攻坚", "高质量发展"}


def _buffer_to_segments(lines: List[str], owner_hint: str) -> List[Dict[str, str]]:
    raw = "\n".join(lines).strip()
    if not raw:
        return []
    pieces: List[str] = []
    for line in lines:
        sentences = [item.strip() for item in re.split(r"(?<=[。！？；;.!?])\s*", line) if item.strip()]
        if sentences:
            pieces.extend(sentences)
        else:
            pieces.append(line)

    segments = [item for item in pieces if len(item) >= 12]
    if not segments:
        segments = [raw]
    return [{"owner_hint": owner_hint, "source_text": item} for item in segments]
