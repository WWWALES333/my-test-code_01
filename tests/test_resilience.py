import json
import tempfile
import unittest
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import Mock, patch

from src.main import WeeklyReportDownloader


class TestSchedulerResilience(unittest.TestCase):
    """调度与通知容错回归测试"""

    def _build_config(self, base_dir: Path) -> Path:
        config = {
            "imap_server": "imap.example.com",
            "imap_port": 993,
            "email": "test@example.com",
            "password": "password",
            "output_root": str(base_dir / "output"),
            "region_keywords": [],
            "download_history_file": str(base_dir / "runtime" / "downloaded_history.json"),
            "log_file": str(base_dir / "runtime" / "run_log.json"),
            "mail_search": {
                "start_date": "2026-01-01"
            },
            "date_filter": {
                "enabled": False
            },
            "report_type_filter": "all",
            "scheduler": {
                "enabled": True,
                "weekly_day": 1,
                "monthly_day": 1,
                "hour": 9,
                "minute": 0
            },
            "notify": {
                "enabled": True,
                "webhook_url": "https://open.feishu.cn/open-apis/bot/v2/hook/test",
                "type": "feishu"
            },
            "retry": {
                "imap": {
                    "max_attempts": 3,
                    "base_delay_seconds": 1,
                    "backoff_factor": 1
                },
                "notify": {
                    "max_attempts": 3,
                    "base_delay_seconds": 1,
                    "backoff_factor": 1
                }
            }
        }

        config_path = base_dir / "config.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        return config_path

    def test_connect_mailbox_retries_until_success(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = self._build_config(Path(temp_dir))
            downloader = WeeklyReportDownloader(str(config_path))

            mailbox = Mock()
            mailbox.login.side_effect = [
                RuntimeError("dns failed"),
                RuntimeError("dns failed"),
                None
            ]

            with patch("src.main.MailBox", return_value=mailbox) as mailbox_cls:
                with patch("src.main.time.sleep"):
                    result = downloader.connect_mailbox()

            self.assertIs(result, mailbox)
            self.assertEqual(mailbox_cls.call_count, 3)
            self.assertEqual(mailbox.login.call_count, 3)

    def test_send_notification_retries_until_success(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = self._build_config(Path(temp_dir))
            downloader = WeeklyReportDownloader(str(config_path))

            success_response = Mock()
            success_response.status_code = 200
            success_response.json.return_value = {"code": 0}

            with patch(
                "src.main.requests.post",
                side_effect=[
                    RuntimeError("network error"),
                    RuntimeError("network error"),
                    success_response
                ]
            ) as post_mock:
                with patch("src.main.time.sleep"):
                    downloader.send_notification()

            self.assertEqual(post_mock.call_count, 3)

    def test_run_download_failure_still_persists_run_log(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            config_path = self._build_config(base_dir)
            downloader = WeeklyReportDownloader(str(config_path))

            with patch.object(downloader, "download_and_classify", side_effect=RuntimeError("imap unavailable")):
                with patch.object(downloader, "send_notification"):
                    downloader.run_download(report_type_filter="weekly")

            run_log_path = base_dir / "runtime" / "run_log.json"
            self.assertTrue(run_log_path.exists())
            records = json.loads(run_log_path.read_text(encoding="utf-8"))
            self.assertEqual(len(records), 1)
            self.assertEqual(len(records[0]["failed"]), 1)
            self.assertIn("imap unavailable", records[0]["failed"][0]["error"])
            self.assertEqual(downloader.config["report_type_filter"], "all")

    def test_run_download_no_duplicate_log_when_already_saved(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            base_dir = Path(temp_dir)
            config_path = self._build_config(base_dir)
            downloader = WeeklyReportDownloader(str(config_path))

            def fake_download():
                downloader.run_log["total_emails"] = 1
                downloader._save_run_log()

            with patch.object(downloader, "download_and_classify", side_effect=fake_download):
                with patch.object(downloader, "send_notification"):
                    downloader.run_download(report_type_filter="weekly")

            run_log_path = base_dir / "runtime" / "run_log.json"
            records = json.loads(run_log_path.read_text(encoding="utf-8"))
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["total_emails"], 1)

    def _build_mail_bytes(self, subject: str, filename: str, date_header: str) -> bytes:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = "sender@example.com"
        msg["Date"] = date_header
        msg.set_content("test")
        msg.add_attachment(
            b"demo",
            maintype="application",
            subtype="octet-stream",
            filename=filename
        )
        return msg.as_bytes()

    def test_search_processes_latest_message_first(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = self._build_config(Path(temp_dir))
            downloader = WeeklyReportDownloader(str(config_path))

            mailbox = Mock()
            mailbox.client = Mock()
            mailbox.client.search.return_value = ("OK", [b"1 2"])

            old_msg = self._build_mail_bytes(
                subject="2026年1月第1周 将军汤 周报",
                filename="2026年1月第1周将军汤工作周报.docx",
                date_header="Sun, 04 Jan 2026 10:00:00 +0800"
            )
            new_msg = self._build_mail_bytes(
                subject="2026年3月第2周 将军汤 周报",
                filename="2026年3月第2周将军汤工作周报.docx",
                date_header="Sun, 15 Mar 2026 10:00:00 +0800"
            )

            fetch_order = []

            def fetch_side_effect(msg_id, _spec):
                fetch_order.append(msg_id)
                if msg_id == b"2":
                    return ("OK", [(b"2", new_msg)])
                return ("OK", [(b"1", old_msg)])

            mailbox.client.fetch.side_effect = fetch_side_effect

            emails = downloader.search_weekly_report_emails(mailbox)

            self.assertEqual(fetch_order[:2], [b"2", b"1"])
            self.assertGreaterEqual(len(emails), 2)

    def test_search_reconnects_and_retries_on_fetch_eof(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = self._build_config(Path(temp_dir))
            downloader = WeeklyReportDownloader(str(config_path))

            mailbox = Mock()
            mailbox.client = Mock()
            mailbox.client.search.return_value = ("OK", [b"9"])

            msg_bytes = self._build_mail_bytes(
                subject="2026年3月第2周 将军汤 周报",
                filename="2026年3月第2周将军汤工作周报.docx",
                date_header="Sun, 15 Mar 2026 10:00:00 +0800"
            )
            mailbox.client.fetch.side_effect = [
                RuntimeError("socket error: EOF occurred in violation of protocol"),
                ("OK", [(b"9", msg_bytes)])
            ]

            with patch.object(downloader, "_reconnect_mailbox_for_fetch", return_value=True) as reconnect_mock:
                with patch("src.main.time.sleep"):
                    emails = downloader.search_weekly_report_emails(mailbox)

            self.assertEqual(reconnect_mock.call_count, 1)
            self.assertEqual(mailbox.client.fetch.call_count, 2)
            self.assertEqual(len(emails), 1)


if __name__ == "__main__":
    unittest.main()
