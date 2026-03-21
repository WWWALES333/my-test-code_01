#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
销售周报自动下载与分类系统
从阿里邮箱自动下载销售部门的周报邮件附件，并根据区域和时间自动分类存档
"""

import os
import re
import json
import csv
import hashlib
import logging
import email
import schedule
import time
import requests
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import imap_tools
from imap_tools.mailbox import MailBox

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('system.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class WeeklyReportDownloader:
    """销售周报下载器"""

    def __init__(self, config_path: str = "data/input/config.json"):
        """初始化下载器"""
        self.config = self._load_config(config_path)
        self.download_history = self._load_history()
        self.run_log = self._create_run_log()

    def _create_run_log(self) -> dict:
        """创建单次运行日志结构"""
        return {
            "run_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_emails": 0,
            "downloaded": [],
            "skipped": [],
            "failed": [],
            "week_conflicts": [],
            "integrity_summary": {},
            "audit_report": {}
        }

    def _load_config(self, config_path: str) -> dict:
        """加载配置文件"""
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _load_history(self) -> dict:
        """加载下载历史"""
        history_file = self.config.get(
            "download_history_file",
            "data/output/runtime/downloaded_history.json"
        )
        if os.path.exists(history_file):
            with open(history_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {"downloaded": {}}

    def _save_history(self):
        """保存下载历史"""
        history_file = self.config.get(
            "download_history_file",
            "data/output/runtime/downloaded_history.json"
        )
        parent_dir = os.path.dirname(history_file)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(self.download_history, f, ensure_ascii=False, indent=2)

    def _save_run_log(self):
        """保存运行日志"""
        log_file = self.config.get("log_file", "data/output/runtime/run_log.json")
        # 读取已有日志
        logs = []
        if os.path.exists(log_file):
            try:
                with open(log_file, 'r', encoding='utf-8') as f:
                    logs = json.load(f)
            except:
                logs = []
        logs.append(self.run_log)
        # 只保留最近50条日志
        logs = logs[-50:]
        parent_dir = os.path.dirname(log_file)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)
        with open(log_file, 'w', encoding='utf-8') as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)

    def _format_week_folder(self, year: int, month: int, week: int) -> str:
        """格式化周报目录名"""
        return f"{year}年{month:02d}月第{week}周"

    def _resolve_year(self, year_str: str) -> int:
        """将2位或4位年份转为4位年份"""
        return 2000 + int(year_str) if len(year_str) == 2 else int(year_str)

    def _parse_chinese_month(self, month_name: str) -> Optional[int]:
        """将中文月份解析为数字月份"""
        month_names = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
            '七': 7, '八': 8, '九': 9, '十': 10, '十一': 11, '十二': 12,
            '一月': 1, '二月': 2, '三月': 3, '四月': 4, '五月': 5, '六月': 6,
            '七月': 7, '八月': 8, '九月': 9, '十月': 10, '十一月': 11, '十二月': 12
        }
        return month_names.get(month_name)

    def _parse_chinese_number(self, value: str) -> Optional[int]:
        """解析中文数字周次或月份"""
        normalized = value.replace('第', '').replace('周', '').replace('月', '')
        return self._parse_chinese_month(normalized)

    def _is_valid_date_parts(self, year: int, month: int, day: int) -> bool:
        """校验日期片段是否合法"""
        try:
            datetime(year, month, day)
            return True
        except ValueError:
            return False

    def _is_reply_or_notification(self, subject: str) -> bool:
        """过滤回复链、转发和撤回通知"""
        stripped = subject.strip()
        lower = stripped.lower()
        prefixes = ('re:', 'fw:', 'fwd:', '回复：', '回复:', '转发：', '转发:', '通知：', '通知:')
        if lower.startswith(prefixes) or stripped.startswith(prefixes):
            return True
        withdrawal_keywords = ['已被发件人撤回', '邮件已被撤回', '邮件已被发件人撤回']
        return any(keyword in subject for keyword in withdrawal_keywords)

    def _is_target_sales_report(self, subject: str, attachment_names: List[str], report_type: str) -> bool:
        """判断是否属于目标销售周报/月报"""
        searchable_text = " ".join([subject] + attachment_names)
        region_keywords = self.config.get("region_keywords", [])
        if any(keyword in searchable_text for keyword in region_keywords):
            return True

        # 兜底兼容一些常见目标命名
        sales_keywords = ['将军汤', '云管家', '线上战区', '一战区', '二战区', '三战区', '四战区', '五战区', '六战区', '七战区']
        if any(keyword in searchable_text for keyword in sales_keywords):
            return True

        return False

    def _subject_year_matches_scope(self, subject: str, min_year: int) -> bool:
        """按主题中的年份快速过滤旧报告"""
        years = [int(year) for year in re.findall(r'(?<!\d)(20\d{2})(?!\d)', subject)]
        short_years = [2000 + int(year) for year in re.findall(r'(?<!\d)(\d{2})年', subject)]
        year_candidates = years + short_years
        if not year_candidates:
            return True
        return any(year >= min_year for year in year_candidates)

    def _parse_weekly_filename_week(self, text: str) -> Optional[Dict[str, int]]:
        """优先从文件名或主题中的显式周次解析周报时间"""
        explicit_week_patterns = [
            r'(\d{2,4})年\s*(\d{1,2})月第\s*(\d{1,2})周',
            r'(\d{2,4})年(\d{1,2})月(\d{1,2})周\b',
            r'(\d{2,4})年.*?(\d{1,2})月第?\s*(\d{1,2})周',
            r'(\d{4}).*?(\d{1,2})月\s*(\d{1,2})周',
            r'(\d{4}).*?([一二三四五六七八九十]+)月第?\s*([一二三四五六七八九十]+)周',
            r'(\d{2,4})年?\s*(\d{1,2})月第?([一二三四五六七八九十]+)周',
        ]
        date_patterns = [
            r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})',
            r'周报[—\-]+(\d{4})(\d{2})(\d{2})',
            r'工作周报[—\-]?(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})',
            r'(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)',
        ]

        for pattern in explicit_week_patterns:
            match = re.search(pattern, text)
            if not match:
                continue
            year = self._resolve_year(match.group(1))
            month_group = match.group(2)
            week_group = match.group(3)
            month = int(month_group) if month_group.isdigit() else self._parse_chinese_number(month_group)
            week = int(week_group) if week_group.isdigit() else self._parse_chinese_number(week_group)
            if not month or not week:
                continue
            return {
                "year": year,
                "month": month,
                "week": week,
                "source": "filename_explicit_week"
            }

        for pattern in date_patterns:
            match = re.search(pattern, text)
            if not match:
                continue
            if pattern == r'(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)':
                continue
            year = int(match.group(1))
            month = int(match.group(2))
            day = int(match.group(3))
            if not self._is_valid_date_parts(year, month, day):
                continue
            week, week_month, week_year = self._get_work_week(year, month, day)
            return {
                "year": week_year,
                "month": week_month,
                "week": week,
                "day": day,
                "source": "filename_date"
            }

        return None

    def _infer_week_from_email_date(self, email_date) -> Optional[Dict[str, int]]:
        """根据邮件发送日期推导工作周"""
        if not email_date:
            return None

        week, month, week_year = self._get_work_week(email_date.year, email_date.month, email_date.day)
        return {
            "year": week_year,
            "month": month,
            "week": week,
            "day": email_date.day,
            "source": "email_date"
        }

    def _resolve_weekly_time_info(self, text: str, email_date=None, log_failure: bool = True) -> Optional[Dict[str, object]]:
        """混合兼容的周报时间解析：文件名优先，邮件日期兜底，并记录冲突"""
        filename_info = self._parse_weekly_filename_week(text)
        email_info = self._infer_week_from_email_date(email_date)

        if not filename_info and email_date:
            short_date_match = re.search(r'(?<!\d)(\d{1,2})\.(\d{1,2})(?!\d)', text)
            if short_date_match:
                month = int(short_date_match.group(1))
                day = int(short_date_match.group(2))
                if self._is_valid_date_parts(email_date.year, month, day):
                    week, week_month, week_year = self._get_work_week(email_date.year, month, day)
                    filename_info = {
                        "year": week_year,
                        "month": week_month,
                        "week": week,
                        "day": day,
                        "source": "filename_short_date"
                    }

        chosen = filename_info or email_info
        if not chosen:
            if log_failure:
                logger.warning(f"时间解析失败，跳过文件: {text}")
            return None

        explicit_week = None
        if filename_info and filename_info.get("source") == "filename_explicit_week":
            explicit_week = {
                "year": filename_info["year"],
                "month": filename_info["month"],
                "week": filename_info["week"]
            }

        email_week = None
        if email_info:
            email_week = {
                "year": email_info["year"],
                "month": email_info["month"],
                "week": email_info["week"]
            }

        conflict = False
        if explicit_week and email_week:
            conflict = (
                explicit_week["year"] != email_week["year"]
                or explicit_week["month"] != email_week["month"]
                or explicit_week["week"] != email_week["week"]
            )

        return {
            "year_folder": str(chosen["year"]),
            "month_week_folder": self._format_week_folder(chosen["year"], chosen["month"], chosen["week"]),
            "year": chosen["year"],
            "month": chosen["month"],
            "week": chosen["week"],
            "source": chosen["source"],
            "explicit_week": explicit_week,
            "email_week": email_week,
            "conflict": conflict
        }

    def _build_week_conflict_record(
        self,
        subject: str,
        filename: str,
        chosen_info: Dict[str, object],
        current_folder: Optional[str] = None
    ) -> Optional[Dict[str, object]]:
        """构建周次冲突审计记录"""
        explicit_week = chosen_info.get("explicit_week")
        email_week = chosen_info.get("email_week")
        if not explicit_week or not email_week:
            return None
        if not chosen_info.get("conflict"):
            return None

        return {
            "subject": subject,
            "filename": filename,
            "current_folder": current_folder,
            "chosen_folder": chosen_info.get("month_week_folder"),
            "explicit_week": explicit_week,
            "email_week": email_week,
            "reason": "week_conflict"
        }

    def _extract_current_folder_info(self, folder_name: str) -> Dict[str, object]:
        """解析当前目录名称中的时间信息"""
        month_week_match = re.search(r'(\d{4})年(\d{2})月第(\d{1,2})周', folder_name)
        if month_week_match:
            year = int(month_week_match.group(1))
            month = int(month_week_match.group(2))
            week = int(month_week_match.group(3))
            return {
                "report_type": "weekly",
                "year": year,
                "month": month,
                "week": week,
                "folder": folder_name
            }

        month_report_match = re.search(r'(\d{2})月报', folder_name)
        if month_report_match:
            month = int(month_report_match.group(1))
            return {
                "report_type": "monthly",
                "month": month,
                "folder": folder_name
            }

        return {"report_type": "unknown", "folder": folder_name}

    def _extract_report_time_for_audit(self, path: Path) -> Dict[str, object]:
        """为历史审计提取建议归档时间信息"""
        current_info = self._extract_current_folder_info(path.parent.name)
        current_year = current_info.get("year")
        filename = path.name

        if "月报" in filename:
            month_info = self._extract_month_info(filename, log_failure=False)
            if month_info[0] is None:
                return {
                    "report_type": "monthly",
                    "current": current_info,
                    "suggested_folder": None,
                    "parse_status": "parse_failed"
                }

            year_folder, _, year, month = month_info
            return {
                "report_type": "monthly",
                "current": current_info,
                "suggested_folder": f"{month:02d}月报",
                "year_folder": year_folder,
                "year": year,
                "month": month,
                "parse_status": "ok"
            }

        time_info = self._resolve_weekly_time_info(filename, log_failure=False)
        if not time_info and current_year:
            synthetic_text = f"{current_year}年 {filename}"
            time_info = self._resolve_weekly_time_info(synthetic_text, log_failure=False)

        if not time_info:
            return {
                "report_type": "weekly",
                "current": current_info,
                "suggested_folder": None,
                "parse_status": "parse_failed"
            }

        return {
            "report_type": "weekly",
            "current": current_info,
            "suggested_folder": time_info["month_week_folder"],
            "year_folder": time_info["year_folder"],
            "year": time_info["year"],
            "month": time_info["month"],
            "week": time_info["week"],
            "parse_status": "ok",
            "explicit_week": time_info.get("explicit_week"),
            "email_week": time_info.get("email_week"),
            "conflict": time_info.get("conflict", False)
        }

    def _build_audit_metadata_index(self) -> Dict[str, Dict[str, object]]:
        """从历史记录和运行日志构建审计补充索引"""
        index: Dict[str, Dict[str, object]] = {}

        def upsert(path_key: str, candidate: Dict[str, object]):
            existing = index.get(path_key, {})
            merged: Dict[str, object] = {}
            for field in ("subject", "sender", "email_date", "time_info", "report_type"):
                new_val = candidate.get(field)
                old_val = existing.get(field)
                if field == "time_info":
                    if isinstance(new_val, dict) and new_val:
                        merged[field] = new_val
                    else:
                        merged[field] = old_val or {}
                else:
                    if new_val not in (None, ""):
                        merged[field] = new_val
                    else:
                        merged[field] = old_val
            index[path_key] = merged

        downloaded = self.download_history.get("downloaded", {})
        for item in downloaded.values():
            output_path = item.get("output_path")
            if not output_path:
                continue
            upsert(str(Path(output_path)), {
                "subject": item.get("subject", ""),
                "sender": item.get("sender", ""),
                "email_date": item.get("email_date"),
                "time_info": item.get("time_info"),
                "report_type": item.get("report_type")
            })

        log_file = self.config.get("log_file", "data/output/runtime/run_log.json")
        if os.path.exists(log_file):
            try:
                with open(log_file, 'r', encoding='utf-8') as f:
                    logs = json.load(f)
                for run in logs:
                    for item in run.get("downloaded", []):
                        output_path = item.get("output_path")
                        if not output_path:
                            continue
                        upsert(str(Path(output_path)), {
                            "subject": item.get("subject", ""),
                            "sender": item.get("sender", ""),
                            "email_date": item.get("email_date"),
                            "time_info": item.get("time_info"),
                            "report_type": item.get("report_type")
                        })
            except Exception:
                pass

        return index

    def _looks_like_template(self, filename: str) -> bool:
        """判断是否为模板或占位文件"""
        template_keywords = ["模板", "XX汤", "示例", "sample"]
        return any(keyword in filename for keyword in template_keywords)

    def _apply_audit_fallbacks(
        self,
        path: Path,
        info: Dict[str, object],
        metadata: Dict[str, object]
    ) -> Dict[str, object]:
        """用历史元数据和当前目录为审计补全缺失时间信息"""
        if info.get("parse_status") == "ok":
            return info

        current = info.get("current", {})
        report_type = info.get("report_type", current.get("report_type", "unknown"))
        time_info = metadata.get("time_info") or {}

        if report_type == "weekly":
            subject = str(metadata.get("subject") or "")
            email_date_raw = metadata.get("email_date")
            email_date_obj = None
            if email_date_raw:
                try:
                    email_date_obj = datetime.fromisoformat(str(email_date_raw))
                except Exception:
                    email_date_obj = None

            if subject or email_date_obj:
                inferred = self._resolve_weekly_time_info(
                    f"{subject} {path.name}".strip(),
                    email_date_obj,
                    log_failure=False
                )
                if inferred:
                    return {
                        "report_type": "weekly",
                        "current": current,
                        "suggested_folder": inferred["month_week_folder"],
                        "year_folder": inferred.get("year_folder"),
                        "year": inferred.get("year"),
                        "month": inferred.get("month"),
                        "week": inferred.get("week"),
                        "parse_status": "metadata_fallback",
                        "explicit_week": inferred.get("explicit_week"),
                        "email_week": inferred.get("email_week"),
                        "conflict": inferred.get("conflict", False)
                    }

            if time_info and time_info.get("month_week_folder"):
                return {
                    "report_type": "weekly",
                    "current": current,
                    "suggested_folder": time_info["month_week_folder"],
                    "year_folder": time_info.get("year_folder", str(time_info.get("year"))),
                    "year": time_info.get("year"),
                    "month": time_info.get("month"),
                    "week": time_info.get("week"),
                    "parse_status": "history_fallback",
                    "explicit_week": time_info.get("explicit_week"),
                    "email_week": time_info.get("email_week"),
                    "conflict": time_info.get("conflict", False)
                }

            if current.get("report_type") == "weekly" and current.get("folder"):
                return {
                    "report_type": "weekly",
                    "current": current,
                    "suggested_folder": current["folder"],
                    "year_folder": str(current.get("year")),
                    "year": current.get("year"),
                    "month": current.get("month"),
                    "week": current.get("week"),
                    "parse_status": "current_folder_fallback",
                    "explicit_week": None,
                    "email_week": None,
                    "conflict": False
                }

        if report_type == "monthly":
            if current.get("report_type") == "monthly" and current.get("folder"):
                return {
                    "report_type": "monthly",
                    "current": current,
                    "suggested_folder": current["folder"],
                    "year_folder": path.parts[-3] if len(path.parts) >= 3 else "",
                    "year": None,
                    "month": current.get("month"),
                    "parse_status": "current_folder_fallback"
                }

        return info

    def _classify_weekly_count_status(self, count: int) -> str:
        """给周报数量打状态标签"""
        if count < 22:
            return "under_expected"
        if count <= 24:
            return "ok"
        return "over_expected"

    def _classify_monthly_count_status(self, count: int) -> str:
        """给月报数量打状态标签"""
        return "ok" if count >= 9 else "under_expected"

    def _write_audit_files(
        self,
        output_dir: Path,
        summary: Dict[str, object],
        file_rows: List[Dict[str, object]]
    ) -> Dict[str, str]:
        """输出 JSON 和 CSV 审计文件"""
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_path = output_dir / f"audit_report_{timestamp}.json"
        csv_path = output_dir / f"audit_report_{timestamp}.csv"

        json_payload = {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": summary,
            "files": file_rows
        }
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_payload, f, ensure_ascii=False, indent=2)

        fieldnames = [
            "report_type",
            "year",
            "current_folder",
            "filename",
            "current_week",
            "filename_week",
            "email_week",
            "conflict",
            "status",
            "suggested_folder",
            "region",
            "sender",
            "parse_status"
        ]
        with open(csv_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in file_rows:
                writer.writerow({key: row.get(key) for key in fieldnames})

        return {
            "json": str(json_path),
            "csv": str(csv_path)
        }

    def generate_audit_report(self, output_dir: Optional[str] = None) -> Dict[str, object]:
        """扫描归档目录，生成历史审计报告"""
        root = Path(self.config.get("output_root", "./06 销售周报"))
        audit_root = Path(output_dir or "data/output/audit/reports")
        metadata_index = self._build_audit_metadata_index()

        weekly_counts = {}
        monthly_counts = {}
        weekly_regions = {}
        weekly_senders = {}
        file_rows = []
        mismatched_files = []
        parse_failed_files = []
        ignored_files = []
        conflict_count = 0

        if not root.exists():
            summary = {
                "weekly_counts": [],
                "monthly_counts": [],
                "mismatched_files": [],
                "parse_failed_files": [],
                "weekly_region_coverage": [],
                "weekly_sender_coverage": [],
                "ignored_files": []
            }
            paths = self._write_audit_files(audit_root, summary, file_rows)
            return {"summary": summary, "paths": paths, "files": file_rows}

        for path in sorted(root.rglob('*')):
            if not path.is_file():
                continue
            if path.name == '.DS_Store':
                continue

            relative_parts = path.relative_to(root).parts
            if len(relative_parts) < 2:
                continue

            if self._looks_like_template(path.name):
                ignored_files.append({
                    "filename": path.name,
                    "current_folder": path.parent.name,
                    "reason": "template_file"
                })
                file_rows.append({
                    "report_type": "unknown",
                    "year": relative_parts[0],
                    "current_folder": path.parent.name,
                    "filename": path.name,
                    "current_week": "",
                    "filename_week": "",
                    "email_week": "",
                    "conflict": "no",
                    "status": "ignored",
                    "suggested_folder": "",
                    "region": "",
                    "sender": "",
                    "parse_status": "ignored_template"
                })
                continue

            metadata = metadata_index.get(str(path), {})
            info = self._extract_report_time_for_audit(path)
            info = self._apply_audit_fallbacks(path, info, metadata)
            current = info.get("current", {})
            report_type = info.get("report_type", "unknown")
            current_year = str(relative_parts[0])
            target_year = str(info.get("year") or current.get("year") or current_year)
            current_folder = current.get("folder")
            suggested_folder = info.get("suggested_folder")
            parse_status = info.get("parse_status", "ok")
            region_folder, matched_regions = self._extract_region(path.name)
            region = matched_regions[0] if matched_regions else ""
            sender = metadata.get("sender", "")

            current_week = ""
            if current.get("report_type") == "weekly":
                current_week = f"{current.get('year', '')}-{current.get('month', 0):02d}-W{current.get('week', '')}"
            filename_week = ""
            explicit_week = info.get("explicit_week")
            if explicit_week:
                filename_week = f"{explicit_week['year']}-{explicit_week['month']:02d}-W{explicit_week['week']}"
            elif report_type == "weekly" and info.get("week") is not None:
                filename_week = f"{info['year']}-{info['month']:02d}-W{info['week']}"
            email_week = ""
            if info.get("email_week"):
                email_week = f"{info['email_week']['year']}-{info['email_week']['month']:02d}-W{info['email_week']['week']}"

            row = {
                "report_type": report_type,
                "year": target_year,
                "current_year": current_year,
                "target_year": target_year,
                "current_folder": current_folder,
                "filename": path.name,
                "current_week": current_week,
                "filename_week": filename_week,
                "email_week": email_week,
                "conflict": "yes" if info.get("conflict") else "no",
                "status": "ok",
                "suggested_folder": suggested_folder or "",
                "region": region,
                "sender": sender,
                "parse_status": parse_status
            }

            if parse_status == "parse_failed":
                row["status"] = "parse_failed"
                parse_failed_files.append({
                    "filename": path.name,
                    "current_year": current_year,
                    "current_folder": current_folder,
                    "report_type": report_type
                })
            elif suggested_folder and current_folder != suggested_folder:
                row["status"] = "folder_mismatch"
                mismatched_files.append({
                    "filename": path.name,
                    "current_year": current_year,
                    "suggested_year": target_year,
                    "current_folder": current_folder,
                    "suggested_folder": suggested_folder,
                    "report_type": report_type
                })

            if info.get("conflict"):
                conflict_count += 1

            if report_type == "weekly" and suggested_folder:
                key = f"{target_year}/{suggested_folder}"
                weekly_counts[key] = weekly_counts.get(key, 0) + 1
                if region:
                    weekly_regions.setdefault(key, set()).add(region)
                if sender:
                    weekly_senders.setdefault(key, set()).add(sender)
            elif report_type == "monthly" and suggested_folder:
                key = f"{target_year}/{suggested_folder}"
                monthly_counts[key] = monthly_counts.get(key, 0) + 1

            file_rows.append(row)

        weekly_summary = []
        for key, count in sorted(weekly_counts.items()):
            weekly_summary.append({
                "period": key,
                "count": count,
                "status": self._classify_weekly_count_status(count)
            })

        monthly_summary = []
        for key, count in sorted(monthly_counts.items()):
            monthly_summary.append({
                "period": key,
                "count": count,
                "status": self._classify_monthly_count_status(count)
            })

        weekly_region_coverage = []
        for key in sorted(weekly_counts.keys()):
            regions = sorted(weekly_regions.get(key, set()))
            weekly_region_coverage.append({
                "period": key,
                "region_count": len(regions),
                "regions": regions
            })

        weekly_sender_coverage = []
        for key in sorted(weekly_counts.keys()):
            senders = sorted(weekly_senders.get(key, set()))
            weekly_sender_coverage.append({
                "period": key,
                "sender_count": len(senders),
                "senders": senders[:30]
            })

        summary = {
            "weekly_counts": weekly_summary,
            "monthly_counts": monthly_summary,
            "mismatched_files": mismatched_files,
            "parse_failed_files": parse_failed_files,
            "weekly_region_coverage": weekly_region_coverage,
            "weekly_sender_coverage": weekly_sender_coverage,
            "ignored_files": ignored_files,
            "week_conflict_count": conflict_count,
            "under_expected_weeks": [item for item in weekly_summary if item["status"] == "under_expected"],
            "under_expected_months": [item for item in monthly_summary if item["status"] == "under_expected"]
        }
        paths = self._write_audit_files(audit_root, summary, file_rows)
        return {"summary": summary, "paths": paths, "files": file_rows}

    def _hash_file(self, path: Path) -> str:
        """计算文件内容 MD5"""
        digest = hashlib.md5()
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                digest.update(chunk)
        return digest.hexdigest()

    def _update_history_paths(self, moved_pairs: List[Tuple[str, str]]):
        """同步更新下载历史中的输出路径"""
        if not moved_pairs:
            return
        path_map = {old: new for old, new in moved_pairs}
        changed = False
        for item in self.download_history.get("downloaded", {}).values():
            old_path = item.get("output_path")
            if old_path in path_map:
                item["output_path"] = path_map[old_path]
                changed = True
        if changed:
            self._save_history()

    def repair_archive(
        self,
        output_dir: Optional[str] = None,
        dry_run: bool = False
    ) -> Dict[str, object]:
        """按审计结果修复历史归档目录"""
        audit_result = self.generate_audit_report(output_dir)
        root = Path(self.config.get("output_root", "./06 销售周报"))
        moved_pairs: List[Tuple[str, str]] = []
        repaired = []
        duplicate_targets = []
        missing_sources = []
        skipped = []

        for row in audit_result["files"]:
            if row.get("status") != "folder_mismatch":
                continue

            current_year = str(row.get("current_year") or row.get("year"))
            target_year = str(row.get("target_year") or row.get("suggested_year") or row.get("year"))
            current_folder = row.get("current_folder")
            suggested_folder = row.get("suggested_folder")
            filename = row.get("filename")
            src = root / current_year / current_folder / filename
            dst = root / target_year / suggested_folder / filename

            if not src.exists():
                missing_sources.append({"source": str(src), "target": str(dst)})
                continue

            if dry_run:
                skipped.append({"source": str(src), "target": str(dst), "reason": "dry_run"})
                continue

            dst.parent.mkdir(parents=True, exist_ok=True)

            if dst.exists():
                src_hash = self._hash_file(src)
                dst_hash = self._hash_file(dst)
                if src_hash == dst_hash:
                    src.unlink()
                    duplicate_targets.append({
                        "source": str(src),
                        "target": str(dst),
                        "reason": "duplicate_target_same_content"
                    })
                else:
                    stem = dst.stem
                    suffix = dst.suffix
                    counter = 1
                    candidate = dst
                    while candidate.exists():
                        candidate = dst.with_name(f"{stem}__repair_{counter}{suffix}")
                        counter += 1
                    src.rename(candidate)
                    repaired.append({"source": str(src), "target": str(candidate)})
                    moved_pairs.append((str(src), str(candidate)))
                continue

            src.rename(dst)
            repaired.append({"source": str(src), "target": str(dst)})
            moved_pairs.append((str(src), str(dst)))

        self._update_history_paths(moved_pairs)

        # 删除修复后产生的空目录
        if not dry_run:
            for year_dir in sorted(root.glob('*')):
                if not year_dir.is_dir():
                    continue
                for subdir in sorted(year_dir.glob('*')):
                    if subdir.is_dir():
                        try:
                            subdir.rmdir()
                        except OSError:
                            pass

        post_audit = self.generate_audit_report(output_dir)
        return {
            "before": audit_result["summary"],
            "after": post_audit["summary"],
            "paths": post_audit["paths"],
            "repaired": repaired,
            "duplicate_targets": duplicate_targets,
            "missing_sources": missing_sources,
            "skipped": skipped
        }

    def _build_integrity_summary(self) -> Dict[str, object]:
        """构建运行后的完整性摘要"""
        audit_result = self.generate_audit_report()
        summary = audit_result["summary"]
        return {
            "weekly_counts": summary["weekly_counts"],
            "monthly_counts": summary["monthly_counts"],
            "week_conflict_count": summary["week_conflict_count"],
            "under_expected_weeks": summary["under_expected_weeks"],
            "under_expected_months": summary["under_expected_months"],
            "weekly_region_coverage": summary["weekly_region_coverage"][:10],
            "weekly_sender_coverage": summary["weekly_sender_coverage"][:10],
            "ignored_files": summary["ignored_files"][:20],
            "audit_paths": audit_result["paths"]
        }

    def connect_mailbox(self) -> MailBox:
        """连接邮箱"""
        logger.info(f"正在连接邮箱: {self.config['email']}")
        mailbox = MailBox(self.config['imap_server'], self.config['imap_port'])
        mailbox.login(self.config['email'], self.config['password'])
        logger.info("邮箱连接成功")
        return mailbox

    def _get_mail_search_start(self) -> datetime:
        """获取邮箱检索起始时间，默认从 2026-01-01 开始"""
        search_config = self.config.get("mail_search", {})
        start_date_str = search_config.get("start_date", "2026-01-01")
        try:
            return datetime.strptime(start_date_str, "%Y-%m-%d")
        except ValueError:
            logger.warning(f"mail_search.start_date 格式错误，回退到 2026-01-01: {start_date_str}")
            return datetime(2026, 1, 1)

    def search_weekly_report_emails(self, mailbox: MailBox) -> List:
        """搜索周报邮件"""
        logger.info("正在搜索周报邮件...")

        # 获取日期过滤配置
        date_filter = self.config.get("date_filter", {})
        date_enabled = date_filter.get("enabled", False)
        filter_year = date_filter.get("year")
        filter_month = date_filter.get("month")

        search_start = self._get_mail_search_start()
        since_token = search_start.strftime("%d-%b-%Y")
        logger.info(f"邮箱检索范围: {search_start.strftime('%Y-%m-%d')} 及之后")

        # 只搜索起始日期之后的邮件，避免全量扫描历史邮箱
        typ, msg_ids = mailbox.client.search(None, 'SINCE', since_token)

        if typ != 'OK':
            logger.warning(f"搜索失败: {typ}")
            return []

        msg_id_list = msg_ids[0].split()
        logger.info(f"邮箱中共有 {len(msg_id_list)} 封邮件")

        emails = []
        # 正序遍历（从最近的邮件开始）
        for msg_id in msg_id_list:
            try:
                # 获取邮件
                typ, msg_data = mailbox.client.fetch(msg_id, '(RFC822)')
                if typ != 'OK':
                    continue

                msg_bytes = msg_data[0][1]
                # 使用标准email库解析
                msg = email.message_from_bytes(msg_bytes)

                # 解码主题
                subject_raw = msg.get('subject', '')
                if subject_raw:
                    subject = email.header.decode_header(subject_raw)
                    if isinstance(subject, list):
                        subject = ''.join([s[0].decode(s[1] or 'utf-8') if s[0] else '' for s in subject])
                    else:
                        subject = str(subject)
                else:
                    subject = ''

                if self._is_reply_or_notification(subject):
                    continue

                # 判断是周报还是月报
                report_type = None

                # 排除非销售报告（使用精确匹配）
                exclude_keywords = ['微信', '甘草医生', '甘草国医', '产品部', '安心汤', '天雄汤', '拨云汤', '修合汤']

                # 单独检查云管家中后台（精确匹配，不是云管家）
                if '云管家中后台' in subject:
                    continue

                # 检查其他关键词
                if any(kw in subject for kw in exclude_keywords):
                    continue

                if '月报' in subject:
                    report_type = 'monthly'
                elif '周报' in subject:
                    report_type = 'weekly'

                if not report_type:
                    continue

                if not self._subject_year_matches_scope(subject, search_start.year):
                    continue

                # 调试：打印找到的报告类型
                logger.info(f"找到报告: {subject[:40]}... 类型: {report_type}")

                # 获取配置的类型过滤
                type_filter = self.config.get('report_type_filter', 'all')
                if type_filter == 'weekly' and report_type != 'weekly':
                    continue
                if type_filter == 'monthly' and report_type != 'monthly':
                    continue

                # 解析日期
                msg_date_str = msg.get('date', '')
                msg_date = None
                if msg_date_str:
                    try:
                        from email.utils import parsedate_to_datetime
                        msg_date = parsedate_to_datetime(msg_date_str)
                    except:
                        pass

                if msg_date:
                    normalized_msg_date = msg_date.replace(tzinfo=None) if msg_date.tzinfo else msg_date
                    if normalized_msg_date < search_start:
                        continue

                # 检查是否有附件 - 使用标准email库
                attachments = []
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        filename = part.get_filename()
                        content_disp = str(part.get('Content-Disposition', ''))

                        # 解码文件名（可能是Base64编码）
                        if filename:
                            try:
                                decoded = email.header.decode_header(filename)
                                filename = decoded[0][0]
                                if isinstance(filename, bytes):
                                    filename = filename.decode('utf-8')
                            except:
                                pass

                        # 方法1：有文件名的附件
                        if filename:
                            # 检查是否为支持的格式：doc/docx/xlsx/pdf/jpg/png
                            if filename.lower().endswith(('.docx', '.doc', '.pdf', '.xlsx', '.xls')):
                                attachments.append({
                                    'filename': filename,
                                    'part': part
                                })
                            # 方法1.5：没有扩展名的情况（可能是从主题推断的doc文件）
                            elif '.' not in filename:
                                # 检查主题是否包含周报/月报关键词
                                if '周报' in subject or '月报' in subject:
                                    # 假设是docx文件
                                    attachments.append({
                                        'filename': filename + '.docx',
                                        'part': part
                                    })
                        # 方法2：检查 application/octet-stream 类型（可能是嵌入文档）
                        elif content_type == 'application/octet-stream':
                            # 尝试从内容中提取文件名
                            payload = part.get_payload(decode=True)
                            if payload:
                                # 从主题中提取文件名（带扩展名）
                                import re
                                match = re.search(r'([^\s]+\.(docx|doc))', subject)
                                if match:
                                    filename = match.group(1)
                                    attachments.append({
                                        'filename': filename,
                                        'part': part
                                    })
                                # 方法2.5：如果没有找到带扩展名的，尝试提取可能的文件名并添加.docx
                                elif '周报' in subject or '月报' in subject:
                                    # 尝试从主题中提取报告名称
                                    report_match = re.search(r'([^\s]+\b月报|[^\s]+\b周报)', subject)
                                    if report_match:
                                        filename = report_match.group(1) + '.docx'
                                        attachments.append({
                                            'filename': filename,
                                            'part': part
                                        })
                                    else:
                                        # 方法2.6：如果主题中有日期，使用日期+类型作为文件名
                                        date_match = re.search(r'(\d{4})年(\d{1,2})月', subject)
                                        if date_match:
                                            year, month = date_match.groups()
                                            filename = f"{year}年{month}月工作报告.docx"
                                            attachments.append({
                                                'filename': filename,
                                                'part': part
                                            })

                if not attachments:
                    # 方法3：尝试从主题直接推断附件名（针对特殊格式邮件）
                    if '周报' in subject or '月报' in subject:
                        # 尝试从主题提取完整文件名
                        import re
                        # 匹配各种格式的报告名
                        patterns = [
                            r'([^<\s]+周报[^>\s]*)',  # xxx周报xxx
                            r'([^<\s]+月报[^>\s]*)',  # xxx月报xxx
                            r'【([^】]+)】',  # 【xxx】
                        ]
                        for pattern in patterns:
                            match = re.search(pattern, subject)
                            if match:
                                potential_name = match.group(1).strip()
                                # 清理一些特殊字符
                                potential_name = re.sub(r'[<>:"/\\|?*]', '', potential_name)
                                if potential_name:
                                    filename = potential_name
                                    if not filename.lower().endswith(('.docx', '.doc', '.pdf', '.xlsx', '.xls')):
                                        filename += '.docx'
                                    # 创建一个虚拟的part来存储这个文件名
                                    attachments.append({
                                        'filename': filename,
                                        'part': None,  # 标记为需要特殊处理
                                        'from_subject': True
                                    })
                                    break

                if not attachments:
                    logger.debug(f"无法提取附件，跳过: {subject[:50]}")
                    continue

                attachment_names = [att.get('filename', '') for att in attachments]
                if any(any(kw in filename for kw in exclude_keywords) for filename in attachment_names):
                    continue

                if not self._is_target_sales_report(subject, attachment_names, report_type):
                    continue

                # 只从文件名解析时间，不使用邮件日期作为 fallback
                filename_for_filter = attachments[0].get('filename', '')
                # 合并主题和附件名用于日期解析
                date_source = subject + ' ' + filename_for_filter

                # 根据报告类型使用不同的日期解析函数
                if report_type == 'monthly':
                    # 月报：只从文件名解析
                    year_folder, month_folder, file_year, file_month = self._extract_month_info(date_source)
                    # 如果解析失败，跳过该邮件
                    if year_folder is None:
                        logger.info(f"月报时间解析失败，跳过: {subject[:40]}")
                        continue
                    file_week = None
                    time_info = {
                        "year_folder": year_folder,
                        "month_folder": month_folder,
                        "year": file_year,
                        "month": file_month,
                        "week": None,
                        "source": "filename_only",
                        "conflict": False
                    }
                else:
                    # 周报：文件名优先，邮件日期兜底，并记录冲突
                    time_info = self._resolve_weekly_time_info(date_source, msg_date)
                    # 如果解析失败，跳过该邮件
                    if time_info is None:
                        logger.info(f"周报时间解析失败，跳过: {subject[:40]}")
                        continue
                    year_folder = time_info["year_folder"]
                    month_week_folder = time_info["month_week_folder"]
                    file_year = time_info["year"]
                    file_month = time_info["month"]
                    file_week = time_info["week"]

                # 按检索起始年份过滤，避免把 2025 等旧周期混入
                if file_year < search_start.year:
                    continue

                # 日期过滤：按文件名中的日期
                if date_enabled and filter_year:
                    # 如果只设置了年份，则只过滤年份
                    if filter_month:
                        if file_year == filter_year and file_month == filter_month:
                            pass  # 符合条件，保留
                        else:
                            continue  # 不符合，跳过
                    else:
                        # 只按年份过滤
                        if file_year == filter_year:
                            pass  # 符合条件，保留
                        else:
                            continue  # 不符合，跳过

                email_obj = {
                    'subject': subject,
                    'date': msg_date,
                    'sender': msg.get('from', ''),
                    'attachments': attachments,
                    'msg_bytes': msg_bytes,
                    'report_type': report_type,
                    'time_info': time_info
                }
                emails.append(email_obj)

            except Exception as e:
                logger.warning(f"处理邮件失败: {str(e)}")
                continue

        logger.info(f"找到 {len(emails)} 封周报邮件(带附件)")
        self.run_log["total_emails"] = len(emails)
        return emails

    def _get_attachment_content_hash(self, content: bytes) -> str:
        """获取附件内容的MD5哈希"""
        return hashlib.md5(content).hexdigest()

    def _extract_region(self, filename: str) -> Tuple[str, List[str]]:
        """
        从文件名中提取战区信息
        返回: (战区文件夹名, 匹配到的关键词列表)
        """
        region_keywords = self.config.get("region_keywords", [])
        matched_regions = []

        for keyword in region_keywords:
            if keyword in filename:
                matched_regions.append(keyword)

        # 额外识别更多战区
        extra_keywords = {
            '一战区': '一战区(上海)',
            '二战区': '二战区',
            '三战区': '三战区',
            '四战区': '四战区(湘黔)',
            '五战区': '五战区(福建)',
            '六战区': '六战区(北京/陕晋)',
            '七战区': '七战区(东北)',
            '线上战区': '线上战区',
            '云管家': '云管家战区',
            '冀蒙': '冀蒙区域',
            '江苏': '江苏区域',
            '浙江': '浙江区域',
            '黑龙江': '黑龙江区域',
            '湖北': '湖北区域',
            '四川': '四川区域',
            '云南': '云南区域',
            '河南': '河南区域',
            '陕西': '陕西区域',
            '辽宁': '辽宁区域',
            '吉林': '吉林区域',
            '福建': '福建区域',
            '上海': '上海区域',
            '粤海': '粤海区域',
            '甘宁': '甘宁区域',
            '晋甘宁': '晋甘宁区域',
        }

        for keyword, full_name in extra_keywords.items():
            if keyword in filename and full_name not in matched_regions:
                matched_regions.append(full_name)

        if matched_regions:
            # 按字母顺序排序以确保一致性
            matched_regions.sort()
            # 生成文件夹名
            folder_name = "_".join(matched_regions)
            return folder_name, matched_regions

        return "未分类", []

    def _get_work_week(self, year: int, month: int, day: int) -> Tuple[int, int, int]:
        """
        计算工作周编号
        工作周定义：
        - 周一到周五：正常工作周
        - 周末（周六、周日）：发送上周周报的时间，归入上一周

        工作周划分：
        - 第1周：每月1-7日
        - 第2周：每月8-14日
        - 第3周：每月15-21日
        - 第4周：每月22-28日
        - 第5周：29-31日

        返回：(工作周编号, 对应月份, 对应年份)
        """
        import datetime
        date = datetime.date(year, month, day)
        weekday = date.weekday()  # 0=周一, 4=周五, 5=周六, 6=周日

        if weekday <= 4:  # 周一到周五
            work_week = (day - 1) // 7 + 1
            return work_week, month, year
        else:  # 周六、周日 - 归入上一周
            # 找到上周的周一
            last_week_monday = date - datetime.timedelta(days=weekday)
            work_week = (last_week_monday.day - 1) // 7 + 1
            return work_week, last_week_monday.month, last_week_monday.year

    def _extract_time_info(self, filename: str, email_date=None) -> Tuple[str, str, int, int, int]:
        """
        从文件名中解析时间信息
        返回: (年份文件夹, 周文件夹, 年, 月, 周)
        使用"年01月第1周"格式
        工作周逻辑：周一到周五算一周，周末+下周一发送上周周报
        """
        info = self._resolve_weekly_time_info(filename, email_date)
        if not info:
            return None, None, 0, 0, 0
        return (
            info["year_folder"],
            info["month_week_folder"],
            info["year"],
            info["month"],
            info["week"]
        )

    def _parse_docx_filename(self, filename: str) -> str:
        """解析并规范化文件名"""
        # 移除各种扩展名
        for ext in ['.docx', '.doc', '.pdf', '.xlsx', '.xls']:
            if filename.lower().endswith(ext):
                filename = filename[:-len(ext)]
                break
        return filename.strip()

    def download_attachment(self, msg, mailbox: MailBox) -> Optional[Tuple[str, bytes]]:
        """
        下载邮件附件
        返回: (文件名, 文件内容) 或 None
        """
        # 支持字典格式（新格式）
        if isinstance(msg, dict):
            attachments = msg.get('attachments', [])
            for att in attachments:
                filename = att.get('filename', '')
                # 使用 part
                part = att.get('part')
                # 如果是直接从主题推断的附件，跳过（无法获取内容）
                if att.get('from_subject', False):
                    logger.warning(f"无法获取附件内容（需手动下载）: {filename}")
                    continue
                if part:
                    try:
                        content = part.get_payload(decode=True)
                        if content:
                            return filename, content
                    except Exception as e:
                        logger.warning(f"解析附件失败 {filename}: {e}")
                        continue
        return None

    def _extract_month_info(self, filename: str, email_date=None, log_failure: bool = True) -> Tuple[str, str, int, int]:
        """
        从文件名中解析月报时间信息
        优先使用文件名，不使用邮件日期作为 fallback
        返回: (年份文件夹, 月份文件夹, 年, 月)
        """
        # 格式1: 2026年1月月报 或 2026年 1月月报
        pattern = r'(\d{2,4})年\s*(\d{1,2})月月报'
        match = re.search(pattern, filename)
        if match:
            year_str = match.group(1)
            if len(year_str) == 2:
                year = 2000 + int(year_str)
            else:
                year = int(year_str)
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式2: 2026年七战区1月月报 (年份和月份之间有其他文字)
        pattern = r'(\d{2,4})年.*?(\d{1,2})月月报'
        match = re.search(pattern, filename)
        if match:
            year_str = match.group(1)
            if len(year_str) == 2:
                year = 2000 + int(year_str)
            else:
                year = int(year_str)
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式2.5: 26年2月产品部月报 (年份 + 产品部 + 月报)
        pattern = r'(\d{2})年\s*(\d{1,2})\s*月?\s*产品部\s*月报'
        match = re.search(pattern, filename)
        if match:
            year = 2000 + int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式2.6: 产品部月报26年2月
        pattern = r'产品部\s*月报\s*(\d{2})年\s*(\d{1,2})'
        match = re.search(pattern, filename)
        if match:
            year = 2000 + int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式2.7: 26年2月月报 (简写年份 + 月报，中间可能有文字)
        pattern = r'(\d{2})年\s*(\d{1,2})\s*月?\s*月报'
        match = re.search(pattern, filename)
        if match:
            year = 2000 + int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式3: 2026年1月 或 2026年 1月
        pattern = r'(\d{2,4})年\s*(\d{1,2})月'
        match = re.search(pattern, filename)
        if match:
            year_str = match.group(1)
            if len(year_str) == 2:
                year = 2000 + int(year_str)
            else:
                year = int(year_str)
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式4: 2026.1月 (用点号分隔)
        pattern = r'(\d{4})\.(\d{1,2})月'
        match = re.search(pattern, filename)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式4.1: 月报-2026.02 (甘草之星月报格式)
        pattern = r'月报[—\-]\.?(\d{4})[.\-](\d{1,2})(?:\b|$)'
        match = re.search(pattern, filename)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式5: 七战区1月月报 (必须有战区名称前缀，如"一战区1月月报"，无年份则跳过)
        # 只匹配明确带有"战区"关键词的月报格式
        pattern = r'(一战区|二战区|三战区|四战区|五战区|六战区|七战区|线上战区|云管家).*?(\d{1,2})月月报'
        match = re.search(pattern, filename)
        if match:
            # 无年份的月报无法确定年份，跳过
            if log_failure:
                logger.warning(f"月报缺少年份信息，跳过: {filename}")
            return None, None, 0, 0

        # 格式6: 26年一月 (简写年份 + 中文月份)
        month_names = {'一月': 1, '二月': 2, '三月': 3, '四月': 4, '五月': 5, '六月': 6,
                       '七月': 7, '八月': 8, '九月': 9, '十月': 10, '十一月': 11, '十二月': 12}

        for month_name, month_num in month_names.items():
            pattern = rf'(\d{{2}})年{month_name}'
            match = re.search(pattern, filename)
            if match:
                year_short = int(match.group(1))
                year = 2000 + year_short
                year_folder = str(year)
                month_folder = f"{year_folder}年{month_num:02d}月"
                return year_folder, month_folder, year, month_num

        # 格式7: 26年1月 (简写年份 + 数字月份)
        pattern = r'(\d{2})年(\d{1,2})月'
        match = re.search(pattern, filename)
        if match:
            year_short = int(match.group(1))
            year = 2000 + year_short
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 无法解析
        if log_failure:
            logger.warning(f"月报时间解析失败，跳过文件: {filename}")
        return None, None, 0, 0

    def _get_output_path(
        self,
        filename: str,
        email_date=None,
        report_type='weekly',
        subject: str = '',
        time_info: Optional[Dict[str, object]] = None
    ) -> str:
        """计算输出路径

        分类结构:
        - 周报: 年/年01月第1周/文件名 (如 2026/2026年01月第1周/xxx.docx)
        - 月报: 年/月报/文件名 (如 2026/01月报/xxx.docx)
        """
        output_root = self.config.get("output_root", "./06 销售周报")

        # 保留原始文件扩展名
        original_ext = '.docx'  # 默认扩展名
        if filename.lower().endswith('.pdf'):
            original_ext = '.pdf'
        elif filename.lower().endswith('.doc'):
            original_ext = '.doc'
        elif filename.lower().endswith('.docx'):
            original_ext = '.docx'
        elif filename.lower().endswith('.xlsx'):
            original_ext = '.xlsx'
        elif filename.lower().endswith('.xls'):
            original_ext = '.xls'

        # 移除扩展名
        clean_filename = self._parse_docx_filename(filename)

        # 只从文件名解析时间
        date_source = subject + ' ' + filename if subject else filename

        if report_type == 'monthly':
            # 月报: 年/月报/文件名
            if time_info:
                year_folder = time_info["year_folder"]
                year = time_info["year"]
                month = time_info["month"]
            else:
                year_folder, month_folder, year, month = self._extract_month_info(date_source)
            # 解析失败则返回None
            if year_folder is None:
                return None
            # 将 "2026年01月" 转换为 "01月报"
            month_report_folder = f"{month:02d}月报"
            full_path = os.path.join(
                output_root,
                year_folder,
                month_report_folder,
                clean_filename + original_ext
            )
        else:
            # 周报: 年/月/第N周/文件名（文件名优先，邮件日期兜底）
            if time_info:
                year_folder = time_info["year_folder"]
                month_week_folder = time_info["month_week_folder"]
            else:
                year_folder, month_week_folder, year, month, week = self._extract_time_info(date_source, email_date)
            # 解析失败则返回None
            if year_folder is None:
                return None
            full_path = os.path.join(
                output_root,
                year_folder,
                month_week_folder,
                clean_filename + original_ext
            )

        return full_path

    def download_and_classify(self):
        """执行下载和分类"""
        try:
            # 连接邮箱
            mailbox = self.connect_mailbox()

            # 搜索周报邮件
            emails = self.search_weekly_report_emails(mailbox)

            # 确保输出目录存在
            output_root = self.config.get("output_root", "./06 销售周报")
            os.makedirs(output_root, exist_ok=True)

            # 处理每封邮件
            for msg in emails:
                try:
                    # 获取主题（兼容两种格式）
                    if isinstance(msg, dict):
                        subject = msg.get('subject', '')
                    else:
                        subject = msg.subject

                    # 获取附件
                    attachment = self.download_attachment(msg, mailbox)
                    if not attachment:
                        logger.warning(f"邮件无docx附件: {subject}")
                        continue

                    filename, content = attachment
                    content_hash = self._get_attachment_content_hash(content)

                    # 检查是否已下载（去重）
                    if content_hash in self.download_history["downloaded"]:
                        logger.info(f"跳过已下载: {filename}")
                        self.run_log["skipped"].append({
                            "filename": filename,
                            "reason": "duplicate",
                            "subject": subject
                        })
                        continue

                    # 获取邮件日期和类型
                    if isinstance(msg, dict):
                        email_date_obj = msg.get('date')
                        report_type = msg.get('report_type', 'weekly')
                        time_info = msg.get('time_info')
                        sender = msg.get('sender', '')
                    else:
                        email_date_obj = None
                        report_type = 'weekly'
                        time_info = None
                        sender = ''

                    # 计算输出路径（传入邮件日期和类型用于分类）
                    output_path = self._get_output_path(filename, email_date_obj, report_type, subject, time_info)

                    # 如果时间解析失败，跳过该文件
                    if output_path is None:
                        logger.warning(f"时间解析失败，跳过: {filename}")
                        self.run_log["skipped"].append({
                            "filename": filename,
                            "reason": "parse_failed",
                            "subject": subject
                        })
                        continue

                    if report_type == 'weekly' and time_info:
                        conflict_record = self._build_week_conflict_record(subject, filename, time_info)
                        if conflict_record:
                            self.run_log["week_conflicts"].append(conflict_record)

                    # 创建目录
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)

                    # 保存文件
                    with open(output_path, 'wb') as f:
                        f.write(content)

                    logger.info(f"下载成功: {output_path}")

                    # 获取日期（兼容两种格式）
                    if isinstance(msg, dict):
                        msg_date = msg.get('date')
                        email_date = str(msg_date) if msg_date else None
                    else:
                        email_date = str(msg.date) if msg.date else None

                    # 记录下载历史
                    self.download_history["downloaded"][content_hash] = {
                        "filename": filename,
                        "output_path": output_path,
                        "download_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "subject": subject,
                        "email_date": email_date,
                        "sender": sender,
                        "report_type": report_type,
                        "time_info": time_info
                    }

                    self.run_log["downloaded"].append({
                        "filename": filename,
                        "output_path": output_path,
                        "subject": subject,
                        "sender": sender,
                        "report_type": report_type,
                        "time_info": time_info
                    })

                except Exception as e:
                    logger.error(f"处理邮件失败: {subject}, 错误: {str(e)}")
                    self.run_log["failed"].append({
                        "subject": subject,
                        "error": str(e)
                    })

            # 断开邮箱连接（捕获异常避免中断保存）
            try:
                mailbox.logout()
            except Exception as e:
                logger.warning(f"邮箱登出失败: {str(e)}")

            # 保存历史记录
            self._save_history()

            # 生成完整性摘要和审计报告
            self.run_log["integrity_summary"] = self._build_integrity_summary()
            self.run_log["audit_report"] = self.run_log["integrity_summary"].get("audit_paths", {})

            # 保存运行日志
            self._save_run_log()

            # 打印总结
            self._print_summary()

        except Exception as e:
            logger.error(f"程序执行失败: {str(e)}")
            raise

    def _print_summary(self):
        """打印运行总结"""
        print("\n" + "="*50)
        print("运行总结")
        print("="*50)
        print(f"总邮件数: {self.run_log['total_emails']}")
        print(f"成功下载: {len(self.run_log['downloaded'])}")
        print(f"跳过(重复): {len(self.run_log['skipped'])}")
        print(f"失败: {len(self.run_log['failed'])}")
        print(f"周次冲突: {len(self.run_log['week_conflicts'])}")
        print("="*50)

        if self.run_log['downloaded']:
            print("\n成功下载的文件:")
            for item in self.run_log['downloaded']:
                print(f"  - {item['filename']}")
                print(f"    -> {item['output_path']}")

        if self.run_log['skipped']:
            print("\n跳过的文件(重复):")
            for item in self.run_log['skipped']:
                print(f"  - {item['filename']}")

        if self.run_log['failed']:
            print("\n失败的文件:")
            for item in self.run_log['failed']:
                print(f"  - {item['subject']}: {item['error']}")

        integrity_summary = self.run_log.get("integrity_summary", {})
        if integrity_summary.get("under_expected_weeks"):
            print("\n低于预期的周报周次:")
            for item in integrity_summary["under_expected_weeks"][:10]:
                print(f"  - {item['period']}: {item['count']} 份")

        if integrity_summary.get("under_expected_months"):
            print("\n低于预期的月报月份:")
            for item in integrity_summary["under_expected_months"][:10]:
                print(f"  - {item['period']}: {item['count']} 份")

        print()

    def send_notification(self):
        """发送飞书通知"""
        notify_config = self.config.get("notify", {})
        if not notify_config.get("enabled", False):
            return

        webhook_url = notify_config.get("webhook_url", "")
        if not webhook_url:
            logger.warning("飞书Webhook URL未配置，跳过通知")
            return

        # 构建消息内容
        run_time = self.run_log.get("run_time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        total = self.run_log.get("total_emails", 0)
        downloaded_count = len(self.run_log.get("downloaded", []))
        skipped_count = len(self.run_log.get("skipped", []))
        failed_count = len(self.run_log.get("failed", []))
        conflict_count = len(self.run_log.get("week_conflicts", []))
        integrity_summary = self.run_log.get("integrity_summary", {})
        under_expected_weeks = integrity_summary.get("under_expected_weeks", [])
        under_expected_months = integrity_summary.get("under_expected_months", [])

        # 构建文件列表
        downloaded_files = self.run_log.get("downloaded", [])
        file_list_text = ""
        if downloaded_files:
            file_list = [f["filename"][:30] for f in downloaded_files[:5]]
            file_list_text = "\n".join([f"• {f}" for f in file_list])
            if len(downloaded_files) > 5:
                file_list_text += f"\n• ... 还有 {len(downloaded_files) - 5} 个文件"

        # 飞书消息卡片格式
        msg_content = {
            "msg_type": "interactive",
            "card": {
                "header": {
                    "title": {
                        "tag": "plain_text",
                        "content": "📥 销售周报下载完成"
                    },
                    "template": "blue"
                },
                "elements": [
                    {
                        "tag": "div",
                        "fields": [
                            {
                                "is_short": True,
                                "text": {
                                    "tag": "lark_md",
                                    "content": f"**运行时间**\n{run_time}"
                                }
                            },
                            {
                                "is_short": True,
                                "text": {
                                    "tag": "lark_md",
                                    "content": f"**扫描邮件**\n{total} 封"
                                }
                            }
                        ]
                    },
                    {
                        "tag": "div",
                        "fields": [
                            {
                                "is_short": True,
                                "text": {
                                    "tag": "lark_md",
                                    "content": f"**成功下载**\n✅ {downloaded_count} 个"
                                }
                            },
                            {
                                "is_short": True,
                                "text": {
                                    "tag": "lark_md",
                                    "content": f"**跳过(重复)**\n⏭️ {skipped_count} 个"
                                }
                            }
                        ]
                    },
                    {
                        "tag": "div",
                        "fields": [
                            {
                                "is_short": True,
                                "text": {
                                    "tag": "lark_md",
                                    "content": f"**失败**\n❌ {failed_count} 个"
                                }
                            }
                        ]
                    },
                    {
                        "tag": "div",
                        "fields": [
                            {
                                "is_short": True,
                                "text": {
                                    "tag": "lark_md",
                                    "content": f"**周次冲突**\n⚠️ {conflict_count} 个"
                                }
                            }
                        ]
                    }
                ]
            }
        }

        if file_list_text:
            msg_content["card"]["elements"].append({
                "tag": "div",
                "text": {
                    "tag": "lark_md",
                    "content": f"**新增文件**\n{file_list_text}"
                }
            })

        if under_expected_weeks or under_expected_months:
            alert_lines = []
            for item in under_expected_weeks[:5]:
                alert_lines.append(f"- 周报不足: {item['period']} ({item['count']} 份)")
            for item in under_expected_months[:5]:
                alert_lines.append(f"- 月报不足: {item['period']} ({item['count']} 份)")
            msg_content["card"]["elements"].append({
                "tag": "div",
                "text": {
                    "tag": "lark_md",
                    "content": "**完整性提醒**\n" + "\n".join(alert_lines)
                }
            })

        audit_paths = self.run_log.get("audit_report", {})
        if audit_paths:
            msg_content["card"]["elements"].append({
                "tag": "div",
                "text": {
                    "tag": "lark_md",
                    "content": (
                        f"**审计文件**\nJSON: `{audit_paths.get('json', '')}`\n"
                        f"CSV: `{audit_paths.get('csv', '')}`"
                    )
                }
            })

        # 添加底部提示
        msg_content["card"]["elements"].append({
            "tag": "div",
            "text": {
                "tag": "lark_md",
                "content": "请及时检查下载结果，确认文件分类是否正确。"
            }
        })

        try:
            response = requests.post(webhook_url, json=msg_content, timeout=10)
            if response.status_code == 200:
                result = response.json()
                if result.get("code") == 0:
                    logger.info("飞书通知发送成功")
                else:
                    logger.warning(f"飞书通知发送失败: {result.get('msg')}")
            else:
                logger.warning(f"飞书通知HTTP错误: {response.status_code}")
        except Exception as e:
            logger.warning(f"飞书通知发送失败: {str(e)}")

    def run_download(self, report_type_filter: str = "all"):
        """执行下载任务（供定时调用）"""
        # 临时修改报告类型过滤
        original_filter = self.config.get("report_type_filter", "all")
        self.config["report_type_filter"] = report_type_filter

        # 重置运行日志
        self.run_log = self._create_run_log()

        try:
            self.download_and_classify()
        except Exception as e:
            # 定时任务场景下，单次失败不应导致守护进程退出
            logger.error(f"定时任务执行失败: {str(e)}")
            self.run_log["failed"].append({
                "filename": "",
                "error": str(e)
            })
        finally:
            # 恢复原始过滤设置
            self.config["report_type_filter"] = original_filter

            # 发送通知
            self.send_notification()

    def setup_scheduler(self):
        """设置定时任务"""
        scheduler_config = self.config.get("scheduler", {})
        if not scheduler_config.get("enabled", False):
            logger.info("定时任务未启用")
            return

        weekly_day = scheduler_config.get("weekly_day", 1)  # 默认周一
        monthly_day = scheduler_config.get("monthly_day", 1)  # 默认1号
        hour = scheduler_config.get("hour", 9)
        minute = scheduler_config.get("minute", 0)

        # 每周定时任务（周报）
        if weekly_day == 1:
            schedule.every().monday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")
        elif weekly_day == 2:
            schedule.every().tuesday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")
        elif weekly_day == 3:
            schedule.every().wednesday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")
        elif weekly_day == 4:
            schedule.every().thursday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")
        elif weekly_day == 5:
            schedule.every().friday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")
        elif weekly_day == 6:
            schedule.every().saturday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")
        elif weekly_day == 7:
            schedule.every().sunday.at(f"{hour:02d}:{minute:02d}").do(self.run_download, report_type_filter="weekly")

        logger.info(f"已设置每周 {['一','二','三','四','五','六','日'][weekly_day-1]} {hour:02d}:{minute:02d} 执行周报下载")

        # 每月定时任务（月报）
        def run_monthly():
            """每月定时任务"""
            self.run_download(report_type_filter="monthly")

        # 检查是否每月N号（需要额外逻辑判断）
        # 由于schedule库原生不支持"每月N号"，我们使用每日检查的方式
        def check_monthly():
            if datetime.now().day == monthly_day:
                run_monthly()

        schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(check_monthly)
        logger.info(f"已设置每月 {monthly_day} 日 {hour:02d}:{minute:02d} 执行月报下载")

        logger.info("定时任务已启动，程序将持续运行...")

    def run_scheduler(self):
        """运行定时任务模式"""
        self.setup_scheduler()

        # 立即执行一次
        logger.info("立即执行一次下载任务...")
        self.run_download(report_type_filter="all")

        # 进入定时循环
        while True:
            try:
                schedule.run_pending()
            except Exception as e:
                logger.error(f"定时任务调度异常: {str(e)}")
            time.sleep(60)


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description="销售周报自动下载与分类系统")
    parser.add_argument("-c", "--config", default="data/input/config.json", help="配置文件路径")
    parser.add_argument("--daemon", "-d", action="store_true", help="以定时任务模式运行（后台持续运行）")
    parser.add_argument("--once", action="store_true", help="只运行一次，不进入定时循环")
    parser.add_argument("--audit", action="store_true", help="扫描现有归档目录并输出审计报告")
    parser.add_argument("--audit-output", default="data/output/audit/reports", help="审计报告输出目录")
    parser.add_argument("--repair", action="store_true", help="按审计结果修复历史归档目录")
    parser.add_argument("--dry-run", action="store_true", help="配合 --repair 使用，仅预览不落盘")
    args = parser.parse_args()

    downloader = WeeklyReportDownloader(args.config)

    if args.repair:
        result = downloader.repair_archive(args.audit_output, dry_run=args.dry_run)
        print("\n" + "="*50)
        print("修复结果")
        print("="*50)
        print(f"修复前目录不一致文件: {len(result['before']['mismatched_files'])}")
        print(f"修复后目录不一致文件: {len(result['after']['mismatched_files'])}")
        print(f"移动文件数: {len(result['repaired'])}")
        print(f"清理重复目标数: {len(result['duplicate_targets'])}")
        print(f"缺失源文件数: {len(result['missing_sources'])}")
        print(f"JSON: {result['paths']['json']}")
        print(f"CSV: {result['paths']['csv']}")
        print("="*50)
        return

    if args.audit:
        result = downloader.generate_audit_report(args.audit_output)
        summary = result["summary"]
        print("\n" + "="*50)
        print("审计报告")
        print("="*50)
        print(f"周次冲突: {summary['week_conflict_count']}")
        print(f"目录不一致文件: {len(summary['mismatched_files'])}")
        print(f"解析失败文件: {len(summary['parse_failed_files'])}")
        print(f"低于预期的周报周次: {len(summary['under_expected_weeks'])}")
        print(f"低于预期的月报月份: {len(summary['under_expected_months'])}")
        print(f"JSON: {result['paths']['json']}")
        print(f"CSV: {result['paths']['csv']}")
        print("="*50)
        return

    # 检查是否启用定时任务
    scheduler_config = downloader.config.get("scheduler", {})
    if scheduler_config.get("enabled", False) and args.daemon:
        # 定时任务模式
        downloader.run_scheduler()
    else:
        # 单次运行模式
        downloader.download_and_classify()
        # 发送通知（如果启用）
        downloader.send_notification()


if __name__ == "__main__":
    main()
