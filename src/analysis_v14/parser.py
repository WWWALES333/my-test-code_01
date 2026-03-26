from __future__ import annotations

import re
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import List, Tuple

from .schema import (
    PARSE_REASON_DOC,
    PARSE_REASON_GENERIC,
    PARSE_REASON_PDF,
    PARSE_REASON_TOOL_MISSING,
)

SUBPROCESS_TIMEOUT_SECONDS = 20


def extract_text(path: Path) -> Tuple[str, str, str]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        try:
            return _extract_docx(path), "", ""
        except Exception as exc:
            return "", str(exc), PARSE_REASON_GENERIC
    if suffix in {".txt", ".md"}:
        try:
            return path.read_text(encoding="utf-8", errors="ignore"), "", ""
        except Exception as exc:
            return "", f"文本读取失败: {exc}", PARSE_REASON_GENERIC
    if suffix == ".pdf":
        return _extract_pdf(path)
    if suffix == ".doc":
        return _extract_doc(path)
    return "", f"不支持的文件类型: {suffix}", PARSE_REASON_GENERIC


def _extract_docx(path: Path) -> str:
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
        return text.strip()
    except Exception as exc:
        raise RuntimeError(f"DOCX 解析失败: {exc}") from exc


def _extract_pdf(path: Path) -> Tuple[str, str, str]:
    # 优先使用 pypdf（若环境安装），其次使用系统 pdftotext。
    try:
        from pypdf import PdfReader  # type: ignore

        reader = PdfReader(str(path))
        text_parts = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            if page_text:
                text_parts.append(page_text)
        text = "\n".join(text_parts).strip()
        if text:
            return text, "", ""
    except ImportError:
        pass
    except Exception as exc:
        # 继续尝试工具链
        pypdf_error = str(exc)
    else:
        pypdf_error = ""

    tool = shutil.which("pdftotext")
    if tool:
        try:
            result = subprocess.run(
                [tool, "-layout", str(path), "-"],
                check=True,
                capture_output=True,
                text=True,
                timeout=SUBPROCESS_TIMEOUT_SECONDS,
            )
            text = result.stdout.strip()
            if text:
                return text, "", ""
            return "", "PDF 解析结果为空", PARSE_REASON_PDF
        except subprocess.TimeoutExpired:
            return "", f"PDF 解析超时（>{SUBPROCESS_TIMEOUT_SECONDS}s）", PARSE_REASON_PDF
        except Exception as exc:
            return "", f"PDF 解析失败: {exc}", PARSE_REASON_PDF

    if "pypdf_error" in locals() and pypdf_error:
        return "", f"PDF 解析失败: {pypdf_error}", PARSE_REASON_PDF
    return "", "PDF 解析工具缺失（需安装 pypdf 或 pdftotext）", PARSE_REASON_TOOL_MISSING


def _extract_doc(path: Path) -> Tuple[str, str, str]:
    # DOC 使用系统工具链稳定提取：antiword 优先，catdoc 次选，macOS textutil 兜底。
    antiword = shutil.which("antiword")
    if antiword:
        try:
            result = subprocess.run(
                [antiword, str(path)],
                check=True,
                capture_output=True,
                text=True,
                timeout=SUBPROCESS_TIMEOUT_SECONDS,
            )
            text = result.stdout.strip()
            if text:
                return text, "", ""
            return "", "DOC 解析结果为空", PARSE_REASON_DOC
        except subprocess.TimeoutExpired:
            return "", f"DOC 解析超时（antiword >{SUBPROCESS_TIMEOUT_SECONDS}s）", PARSE_REASON_DOC
        except Exception as exc:
            return "", f"DOC 解析失败: {exc}", PARSE_REASON_DOC

    catdoc = shutil.which("catdoc")
    if catdoc:
        try:
            result = subprocess.run(
                [catdoc, str(path)],
                check=True,
                capture_output=True,
                text=True,
                timeout=SUBPROCESS_TIMEOUT_SECONDS,
            )
            text = result.stdout.strip()
            if text:
                return text, "", ""
            return "", "DOC 解析结果为空", PARSE_REASON_DOC
        except subprocess.TimeoutExpired:
            return "", f"DOC 解析超时（catdoc >{SUBPROCESS_TIMEOUT_SECONDS}s）", PARSE_REASON_DOC
        except Exception as exc:
            return "", f"DOC 解析失败: {exc}", PARSE_REASON_DOC

    textutil = shutil.which("textutil")
    if textutil:
        try:
            result = subprocess.run(
                [textutil, "-convert", "txt", "-stdout", str(path)],
                check=True,
                capture_output=True,
                text=True,
                timeout=SUBPROCESS_TIMEOUT_SECONDS,
            )
            text = result.stdout.strip()
            if text:
                return text, "", ""
            return "", "DOC 解析结果为空", PARSE_REASON_DOC
        except subprocess.TimeoutExpired:
            return "", f"DOC 解析超时（textutil >{SUBPROCESS_TIMEOUT_SECONDS}s）", PARSE_REASON_DOC
        except Exception as exc:
            return "", f"DOC 解析失败: {exc}", PARSE_REASON_DOC

    return "", "DOC 解析工具缺失（需安装 antiword/catdoc，或使用 macOS textutil）", PARSE_REASON_TOOL_MISSING


def segment_text(text: str) -> List[str]:
    if not text.strip():
        return []
    normalized = re.sub(r"[ \t]+", " ", text.strip())
    line_parts = [p.strip() for p in re.split(r"\n+", normalized) if p.strip()]

    raw_parts: List[str] = []
    for line in line_parts:
        sentence_parts = [item.strip() for item in re.split(r"(?<=[。！？；;.!?])\s*", line) if item.strip()]
        if sentence_parts:
            raw_parts.extend(sentence_parts)
        else:
            raw_parts.append(line)

    segments = [p for p in raw_parts if len(p) >= 12]
    if not segments:
        segments = raw_parts[:]
    return segments
