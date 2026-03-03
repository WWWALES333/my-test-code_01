#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
销售周报自动下载与分类系统
从阿里邮箱自动下载销售部门的周报邮件附件，并根据区域和时间自动分类存档
"""

import os
import re
import json
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
from imap_tools import AND, OR
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

    def __init__(self, config_path: str = "config.json"):
        """初始化下载器"""
        self.config = self._load_config(config_path)
        self.download_history = self._load_history()
        self.run_log = {
            "run_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_emails": 0,
            "downloaded": [],
            "skipped": [],
            "failed": []
        }

    def _load_config(self, config_path: str) -> dict:
        """加载配置文件"""
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _load_history(self) -> dict:
        """加载下载历史"""
        history_file = self.config.get("download_history_file", "downloaded_history.json")
        if os.path.exists(history_file):
            with open(history_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {"downloaded": {}}

    def _save_history(self):
        """保存下载历史"""
        history_file = self.config.get("download_history_file", "downloaded_history.json")
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(self.download_history, f, ensure_ascii=False, indent=2)

    def _save_run_log(self):
        """保存运行日志"""
        log_file = self.config.get("log_file", "run_log.json")
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
        with open(log_file, 'w', encoding='utf-8') as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)

    def connect_mailbox(self) -> MailBox:
        """连接邮箱"""
        logger.info(f"正在连接邮箱: {self.config['email']}")
        mailbox = MailBox(self.config['imap_server'], self.config['imap_port'])
        mailbox.login(self.config['email'], self.config['password'])
        logger.info("邮箱连接成功")
        return mailbox

    def search_weekly_report_emails(self, mailbox: MailBox) -> List:
        """搜索周报邮件"""
        logger.info("正在搜索周报邮件...")

        # 获取日期过滤配置
        date_filter = self.config.get("date_filter", {})
        date_enabled = date_filter.get("enabled", False)
        filter_year = date_filter.get("year")
        filter_month = date_filter.get("month")

        # 先搜索所有邮件（不使用中文，避免编码问题）
        typ, msg_ids = mailbox.client.search(None, 'ALL')

        if typ != 'OK':
            logger.warning(f"搜索失败: {typ}")
            return []

        msg_id_list = msg_ids[0].split()
        logger.info(f"邮箱中共有 {len(msg_id_list)} 封邮件")

        emails = []
        # 倒序遍历（从最近的邮件开始）
        for msg_id in reversed(msg_id_list):
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

                # 判断是周报还是月报
                report_type = None
                if '月报' in subject:
                    report_type = 'monthly'
                elif '周报' in subject:
                    report_type = 'weekly'

                if not report_type:
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
                            # 检查是否为doc/docx
                            if filename.lower().endswith(('.docx', '.doc')):
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
                                    if not filename.lower().endswith(('.docx', '.doc')):
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

                # 优先使用邮件日期解析时间（更准确），其次用文件名
                filename_for_filter = attachments[0].get('filename', '')
                # 合并主题和附件名用于日期解析
                date_source = subject + ' ' + filename_for_filter

                # 根据报告类型使用不同的日期解析函数
                if report_type == 'monthly':
                    # 月报：优先使用邮件日期，其次用文件名解析
                    year_folder, month_folder, file_year, file_month = self._extract_month_info(date_source)
                    # 如果文件名无法解析出有效日期，使用邮件日期
                    if file_year == datetime.now().year and file_month == datetime.now().month and msg_date:
                        file_year = msg_date.year
                        file_month = msg_date.month
                        year_folder = str(file_year)
                        month_folder = f"{year_folder}年{file_month:02d}月"
                    file_week = None
                else:
                    # 周报：优先使用邮件日期
                    year_folder, month_week_folder, file_year, file_month, file_week = self._extract_time_info(date_source, msg_date)

                # 日期过滤：按文件名中的日期
                if date_enabled and filter_year and filter_month:
                    if file_year == filter_year and file_month == filter_month:
                        pass  # 符合条件，保留
                    else:
                        continue  # 不符合，跳过

                email_obj = {
                    'subject': subject,
                    'date': msg_date,
                    'attachments': attachments,
                    'msg_bytes': msg_bytes,
                    'report_type': report_type
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

    def _get_week_number(self, year: int, month: int, day: int) -> int:
        """
        根据日期计算周次
        规则：每月1-7日为第1周，8-14日为第2周，15-21日为第3周，22-28日为第4周，29+为第5周
        """
        import calendar
        # 计算这是该月的第几周
        _, num_days = calendar.monthrange(year, month)
        if day <= 7:
            return 1
        elif day <= 14:
            return 2
        elif day <= 21:
            return 3
        elif day <= 28:
            return 4
        else:
            return 5

    def _extract_time_info(self, filename: str, email_date=None) -> Tuple[str, str, int, int, int]:
        """
        从文件名或邮件日期中解析时间信息
        优先使用邮件发件时间
        返回: (年份文件夹, 月份周次文件夹, 年, 月, 周)
        """
        # 首先尝试从文件名解析
        # 匹配格式：2026年 2月第 2周 或 2026年2月第2周
        pattern = r'(\d{4})年\s*(\d{1,2})月第\s*(\d{1,2})周'
        match = re.search(pattern, filename)

        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            week = int(match.group(3))

            year_folder = str(year)
            month_week_folder = f"{month:02d}月第{week}周"

            return year_folder, month_week_folder, year, month, week

        # 尝试匹配简单日期格式：2026.2.1, 2026.02.01, 20260201
        simple_pattern = r'(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})'
        simple_match = re.search(simple_pattern, filename)
        if simple_match:
            year = int(simple_match.group(1))
            month = int(simple_match.group(2))
            day = int(simple_match.group(3))
            week = self._get_week_number(year, month, day)
            year_folder = str(year)
            month_week_folder = f"{month:02d}月第{week}周"
            return year_folder, month_week_folder, year, month, week

        # 使用邮件发件时间
        if email_date:
            year = email_date.year
            month = email_date.month
            day = email_date.day
            week = self._get_week_number(year, month, day)
            year_folder = str(year)
            month_week_folder = f"{month:02d}月第{week}周"
            return year_folder, month_week_folder, year, month, week

        # 如果都无法解析，返回当前时间
        current_year = datetime.now().year
        current_month = datetime.now().month
        current_day = datetime.now().day
        current_week = self._get_week_number(current_year, current_month, current_day)
        return str(current_year), f"{current_month:02d}月第{current_week}周", current_year, current_month, current_week

    def _parse_docx_filename(self, filename: str) -> str:
        """解析并规范化docx文件名"""
        # 移除 .docx 后缀
        if filename.lower().endswith('.docx'):
            filename = filename[:-5]
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

    def _extract_month_info(self, filename: str) -> Tuple[str, str, int, int]:
        """
        从文件名中解析月报时间信息
        格式:
        - 2026年1月 将军汤 广西区域部门工作月报.docx
        - 三战区2026年1月月报
        - 线上战区2026.1月月报.docx
        - 将军汤四战区26年一月月报
        - 五战区26年1月月报.docx
        返回: (年份文件夹, 月份文件夹, 年, 月)
        """
        # 格式1: 2026年1月 或 2026年 1月
        pattern = r'(\d{4})年\s*(\d{1,2})月'
        match = re.search(pattern, filename)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式2: 2026.1月 (用点号分隔)
        pattern = r'(\d{4})\.(\d{1,2})月'
        match = re.search(pattern, filename)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式2.5: 2026年七战区1月月报, 2026年二战区1月月报
        # 年份和月份之间有其他文字的情况
        pattern = r'(\d{4})年.*?(\d{1,2})月(?!/)'
        match = re.search(pattern, filename)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 格式3: 26年1月 或 26年一月 (简写年份)
        # 需要判断是20xx年
        month_names = {'一月': 1, '二月': 2, '三月': 3, '四月': 4, '五月': 5, '六月': 6,
                       '七月': 7, '八月': 8, '九月': 9, '十月': 10, '十一月': 11, '十二月': 12}

        for month_name, month_num in month_names.items():
            pattern = rf'(\d{{2}})年{month_name}'
            match = re.search(pattern, filename)
            if match:
                year_short = int(match.group(1))
                year = 2000 + year_short  # 假设20xx年
                year_folder = str(year)
                month_folder = f"{year_folder}年{month_num:02d}月"
                return year_folder, month_folder, year, month_num

        # 格式4: 26年1月 (简写年份 + 数字月份)
        pattern = r'(\d{2})年(\d{1,2})月'
        match = re.search(pattern, filename)
        if match:
            year_short = int(match.group(1))
            year = 2000 + year_short
            month = int(match.group(2))
            year_folder = str(year)
            month_folder = f"{year_folder}年{month:02d}月"
            return year_folder, month_folder, year, month

        # 如果无法解析，返回当前时间
        current_year = datetime.now().year
        current_month = datetime.now().month
        return str(current_year), f"{current_year}年{current_month:02d}月", current_year, current_month

    def _get_output_path(self, filename: str, email_date=None, report_type='weekly', subject: str = '') -> str:
        """计算输出路径

        分类结构:
        - 周报: 年/月/第N周/文件名 (如 2026/02月第2周/xxx.docx)
        - 月报: 年/月报/文件名 (如 2026/01月报/xxx.docx)
        """
        output_root = self.config.get("output_root", "./06 销售周报")
        clean_filename = self._parse_docx_filename(filename)

        # 优先使用主题解析日期，其次用附件名
        date_source = subject + ' ' + filename if subject else filename

        if report_type == 'monthly':
            # 月报: 年/月报/文件名 (如 2026/01月报/xxx.docx)
            year_folder, month_folder, year, month = self._extract_month_info(date_source)
            # 将 "2026年01月" 转换为 "01月报"
            month_report_folder = f"{month:02d}月报"
            full_path = os.path.join(
                output_root,
                year_folder,
                month_report_folder,
                clean_filename + ".docx"
            )
        else:
            # 周报: 年/月/第N周/文件名
            year_folder, month_week_folder, year, month, week = self._extract_time_info(date_source, email_date)
            full_path = os.path.join(
                output_root,
                year_folder,
                month_week_folder,
                clean_filename + ".docx"
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
                    else:
                        email_date_obj = None
                        report_type = 'weekly'

                    # 计算输出路径（传入邮件日期和类型用于分类）
                    output_path = self._get_output_path(filename, email_date_obj, report_type, subject)

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
                        "email_date": email_date
                    }

                    self.run_log["downloaded"].append({
                        "filename": filename,
                        "output_path": output_path,
                        "subject": subject
                    })

                except Exception as e:
                    logger.error(f"处理邮件失败: {subject}, 错误: {str(e)}")
                    self.run_log["failed"].append({
                        "subject": subject,
                        "error": str(e)
                    })

            # 断开邮箱连接
            mailbox.logout()

            # 保存历史记录
            self._save_history()

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
        self.run_log = {
            "run_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_emails": 0,
            "downloaded": [],
            "skipped": [],
            "failed": []
        }

        try:
            self.download_and_classify()
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

        # 每月1号执行
        schedule.every().day.at(f"{hour:02d}:{minute:02d}").do(run_monthly)

        # 检查是否每月1号（需要额外逻辑判断）
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
            schedule.run_pending()
            time.sleep(60)


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description="销售周报自动下载与分类系统")
    parser.add_argument("-c", "--config", default="config.json", help="配置文件路径")
    parser.add_argument("--daemon", "-d", action="store_true", help="以定时任务模式运行（后台持续运行）")
    parser.add_argument("--once", action="store_true", help="只运行一次，不进入定时循环")
    args = parser.parse_args()

    downloader = WeeklyReportDownloader(args.config)

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
