#!/usr/bin/env python3
"""发布前敏感信息检查脚本。"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path


BLOCKED_PATHS = {
    "config.json",
    "data/input/config.json",
    ".claude/settings.local.json",
}

SKIP_PREFIXES = (
    ".git/",
    ".venv/",
    "venv/",
    "env/",
    "__pycache__/",
    "data/output/",
    "downloaded_reports/",
    ".backup/",
)

SKIP_SUFFIXES = (
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".webp",
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".zip",
    ".tar",
    ".gz",
    ".7z",
    ".pyc",
    ".pyo",
    ".so",
)

PATTERNS = [
    ("飞书 Webhook", re.compile(r"https://open\.feishu\.cn/open-apis/bot/v2/hook/[A-Za-z0-9\-]{20,}")),
    ("OpenAI Key", re.compile(r"\bsk-[A-Za-z0-9]{20,}\b")),
    ("常见私钥头", re.compile(r"-----BEGIN (RSA|OPENSSH|EC|DSA) PRIVATE KEY-----")),
]

PASSWORD_LINE = re.compile(r"(?i)(\"?password\"?\s*[:=]\s*[\"']?)([^\"'\s,}]+)")
WEBHOOK_LINE = re.compile(r"(?i)(\"?webhook_url\"?\s*[:=]\s*[\"']?)([^\"'\s,}]+)")
TOKEN_LINE = re.compile(r"(?i)(\"?(api_key|token|secret)\"?\s*[:=]\s*[\"']?)([^\"'\s,}]+)")

PLACEHOLDER_HINTS = ("replace", "example", "placeholder", "your_", "<", ">")
CONFIG_EXTENSIONS = {".json", ".yaml", ".yml", ".toml", ".ini", ".env", ".cfg", ".conf"}


def _git_file_list() -> list[str]:
    result = subprocess.run(
        ["git", "ls-files", "-co", "--exclude-standard"],
        capture_output=True,
        text=True,
        check=True,
    )
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def _is_placeholder(value: str) -> bool:
    value = value.strip().lower()
    return any(hint in value for hint in PLACEHOLDER_HINTS)


def _should_skip(path: str) -> bool:
    if any(path.startswith(prefix) for prefix in SKIP_PREFIXES):
        return True
    if path.endswith(SKIP_SUFFIXES):
        return True
    return False


def _scan_text(path: Path, rel: str) -> list[str]:
    findings: list[str] = []
    try:
        content = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return findings
    except OSError as exc:
        findings.append(f"{rel}: 文件读取失败: {exc}")
        return findings

    for label, pattern in PATTERNS:
        for hit in pattern.finditer(content):
            findings.append(f"{rel}: 命中 {label} -> {hit.group(0)[:80]}")

    if path.suffix.lower() in CONFIG_EXTENSIONS:
        for idx, line in enumerate(content.splitlines(), start=1):
            pw = PASSWORD_LINE.search(line)
            if pw:
                value = pw.group(2).strip()
                if value and not _is_placeholder(value):
                    findings.append(f"{rel}:{idx}: 命中 password 实值")

            wh = WEBHOOK_LINE.search(line)
            if wh:
                value = wh.group(2).strip()
                if value and not _is_placeholder(value):
                    findings.append(f"{rel}:{idx}: 命中 webhook_url 实值")

            tk = TOKEN_LINE.search(line)
            if tk:
                value = tk.group(3).strip()
                if value and len(value) >= 12 and not _is_placeholder(value):
                    findings.append(f"{rel}:{idx}: 命中 token/api_key/secret 实值")

    return findings


def main() -> int:
    root = Path.cwd()
    findings: list[str] = []

    try:
        paths = _git_file_list()
    except subprocess.CalledProcessError as exc:
        print(f"[ERROR] 无法获取文件列表: {exc}", file=sys.stderr)
        return 2

    for rel in paths:
        if rel in BLOCKED_PATHS:
            findings.append(f"{rel}: 属于本地敏感文件，禁止入库")
            continue
        if _should_skip(rel):
            continue
        path = root / rel
        if not path.exists() or not path.is_file():
            continue
        findings.extend(_scan_text(path, rel))

    if findings:
        print("[FAIL] 检测到敏感信息风险，请先处理后再提交/发布：")
        for item in findings:
            print(f"- {item}")
        return 1

    print("[PASS] 未检测到敏感信息风险。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
