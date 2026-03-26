from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

from .schema import stable_hash

SUPPORTED_EXTENSIONS = {".docx", ".doc", ".pdf", ".txt", ".md"}


def collect_sample_files(samples_dir: Path, year_filter: Optional[int] = None) -> List[Path]:
    if not samples_dir.exists() or not samples_dir.is_dir():
        raise FileNotFoundError(f"样本目录不存在: {samples_dir}")

    files = [p for p in samples_dir.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS]
    if year_filter is not None:
        files = [p for p in files if infer_time_info(p).get("year") == year_filter]
    return sorted(files)


def infer_report_type(path: Path) -> str:
    name = f"{path.parent.name} {path.name}"
    if "月报" in name:
        return "monthly"
    return "weekly"


def infer_time_info(path: Path) -> Dict[str, int]:
    text = f"{path.parent.name} {path.name}"
    year_match = re.search(r"(20\d{2})", text)
    month_match = re.search(r"(\d{1,2})月", text)
    week_match = re.search(r"第(\d{1,2})周", text)
    return {
        "year": int(year_match.group(1)) if year_match else 0,
        "month": int(month_match.group(1)) if month_match else 0,
        "week_of_month": int(week_match.group(1)) if week_match else 0,
    }


def build_report_record(path: Path) -> Dict[str, object]:
    time_info = infer_time_info(path)
    report_type = infer_report_type(path)
    report_id = stable_hash(str(path.resolve()))
    return {
        "report_id": report_id,
        "file_path": str(path.resolve()),
        "report_type": report_type,
        "year": time_info["year"],
        "month": time_info["month"],
        "week_of_month": time_info["week_of_month"],
    }


def build_file_fingerprint(path: Path) -> str:
    stat = path.stat()
    return stable_hash(str(path.resolve()), str(stat.st_size), str(int(stat.st_mtime)))
