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

                if not attachments:
                    continue

                # 优先使用邮件主题解析日期（更准确），其次用附件名
                filename_for_filter = attachments[0].get('filename', '')
                # 合并主题和附件名用于日期解析
                date_source = subject + ' ' + filename_for_filter

                # 根据报告类型使用不同的日期解析函数
                if report_type == 'monthly':
                    # 月报：使用月报日期解析
                    year_folder, month_folder, file_year, file_month = self._extract_month_info(date_source)
                    file_week = None
                else:
                    # 周报：使用周报日期解析
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


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description="销售周报自动下载与分类系统")
    parser.add_argument("-c", "--config", default="config.json", help="配置文件路径")
    args = parser.parse_args()

    downloader = WeeklyReportDownloader(args.config)
    downloader.download_and_classify()


if __name__ == "__main__":
    main()
