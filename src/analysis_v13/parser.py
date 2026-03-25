from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import List, Tuple


def extract_text(path: Path) -> Tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _extract_docx(path)
    if suffix in {".txt", ".md"}:
        try:
            return path.read_text(encoding="utf-8", errors="ignore"), ""
        except Exception as exc:
            return "", f"文本读取失败: {exc}"
    if suffix == ".pdf":
        return "", "PDF 解析未启用（v1.3 首版）"
    if suffix == ".doc":
        return "", "DOC 解析未启用（v1.3 首版）"
    return "", f"不支持的文件类型: {suffix}"


def _extract_docx(path: Path) -> Tuple[str, str]:
    try:
        with zipfile.ZipFile(path) as zf:
            xml_bytes = zf.read("word/document.xml")
        root = ET.fromstring(xml_bytes)
        chunks: List[str] = []
        for elem in root.iter():
            tag = elem.tag
            if tag.endswith("}t"):
                chunks.append(elem.text or "")
            elif tag.endswith("}tab"):
                chunks.append("\t")
            elif tag.endswith("}br") or tag.endswith("}cr"):
                chunks.append("\n")
            elif tag.endswith("}p"):
                chunks.append("\n")
        text = "".join(chunks)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip(), ""
    except Exception as exc:
        return "", f"DOCX 解析失败: {exc}"


def segment_text(text: str) -> List[str]:
    if not text.strip():
        return []
    raw_parts = [p.strip() for p in re.split(r"\n+", text) if p.strip()]
    segments = [p for p in raw_parts if len(p) >= 8]
    if not segments:
        segments = raw_parts[:]
    return segments

