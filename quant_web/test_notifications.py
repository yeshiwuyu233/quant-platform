"""Notification behavior tests."""
import importlib
import os
import sys
import unittest
from unittest import mock


WEB_DIR = os.path.dirname(os.path.abspath(__file__))
if WEB_DIR not in sys.path:
    sys.path.insert(0, WEB_DIR)


class TestPipelineNotification(unittest.TestCase):
    def setUp(self):
        self.app = importlib.import_module("app")

    def test_pipeline_notification_skips_when_smtp_host_missing(self):
        self.app._SMTP_USER = "sender@example.com"
        self.app._SMTP_PASS = "secret"
        self.app._SMTP_HOST = ""

        with mock.patch.object(self.app.smtplib, "SMTP_SSL") as smtp_ssl:
            self.app._send_pipeline_notification(True, "ok")

        smtp_ssl.assert_not_called()


if __name__ == "__main__":
    unittest.main()
