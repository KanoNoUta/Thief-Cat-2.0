import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import Mock

from outlook_graph import (
    OutlookCredentials,
    OutlookGraphClient,
    move_outlook_credential,
)


CLIENT_ID = "00000000-0000-0000-0000-000000000001"


class OutlookCredentialTests(unittest.TestCase):
    def test_parse_card_key_format(self):
        record = OutlookCredentials.parse(
            f"User@outlook.com----pw----refresh-token----{CLIENT_ID}"
        )
        self.assertEqual(record.email, "user@outlook.com")
        self.assertEqual(record.password, "pw")
        self.assertEqual(record.refresh_token, "refresh-token")
        self.assertEqual(record.client_id, CLIENT_ID)

    def test_parse_client_id_before_refresh_token(self):
        record = OutlookCredentials.parse(
            f"User@outlook.com----pw----{CLIENT_ID}----M.C_test-refresh-token"
        )
        self.assertEqual(record.email, "user@outlook.com")
        self.assertEqual(record.client_id, CLIENT_ID)
        self.assertEqual(record.refresh_token, "M.C_test-refresh-token")

    def test_rejects_bad_format(self):
        with self.assertRaises(ValueError):
            OutlookCredentials.parse("only----three----parts")

    def test_selects_email_from_multi_record_file(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "outlook.txt"
            path.write_text(
                f"one@outlook.com----a----token-1----{CLIENT_ID}\n"
                f"two@outlook.com----b----token-2----{CLIENT_ID}\n",
                encoding="utf-8",
            )
            record = OutlookCredentials.from_file(path, email="TWO@outlook.com")
        self.assertEqual(record.email, "two@outlook.com")
        self.assertEqual(record.refresh_token, "token-2")

    def test_uses_first_pending_record_by_default(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "outlook_pending.txt"
            path.write_text(
                f"one@outlook.com----a----token-1----{CLIENT_ID}\n"
                f"two@outlook.com----b----token-2----{CLIENT_ID}\n",
                encoding="utf-8",
            )
            record = OutlookCredentials.from_file(path)
        self.assertEqual(record.email, "one@outlook.com")

    def test_moves_successful_record_and_is_idempotent(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pending = Path(temp_dir) / "outlook_pending.txt"
            success = Path(temp_dir) / "outlook_success.txt"
            first = f"one@outlook.com----a----token-1----{CLIENT_ID}"
            second = f"two@outlook.com----b----token-2----{CLIENT_ID}"
            pending.write_text(f"{first}\n{second}\n", encoding="utf-8")

            self.assertTrue(
                move_outlook_credential(pending, success, "TWO@OUTLOOK.COM")
            )
            self.assertEqual(pending.read_text(encoding="utf-8").strip(), first)
            self.assertEqual(success.read_text(encoding="utf-8").strip(), second)
            self.assertTrue(
                move_outlook_credential(pending, success, "two@outlook.com")
            )
            self.assertEqual(success.read_text(encoding="utf-8").splitlines(), [second])


class OutlookGraphTests(unittest.TestCase):
    def setUp(self):
        credentials = OutlookCredentials(
            "child@outlook.com",
            "pw",
            "refresh-token",
            CLIENT_ID,
        )
        self.client = OutlookGraphClient(credentials)

    def tearDown(self):
        self.client.close()

    def test_recipient_filter_is_case_insensitive(self):
        message = {
            "toRecipients": [
                {"emailAddress": {"address": "CHILD@OUTLOOK.COM"}}
            ]
        }
        self.assertTrue(self.client._is_for_address(message))

    def test_recipient_filter_rejects_other_shared_inbox_alias(self):
        message = {
            "toRecipients": [
                {"emailAddress": {"address": "other@outlook.com"}}
            ]
        }
        self.assertFalse(self.client._is_for_address(message))

    def test_wait_for_databricks_code(self):
        now = datetime.now(timezone.utc)
        recipient = [{"emailAddress": {"address": "child@outlook.com"}}]
        self.client.list_messages = Mock(
            return_value=[
                {
                    "id": "message-id",
                    "subject": "Databricks verification code",
                    "from": {
                        "emailAddress": {"address": "noreply@databricks.com"}
                    },
                    "toRecipients": recipient,
                    "receivedDateTime": now.isoformat(),
                    "bodyPreview": "Your code is ready",
                }
            ]
        )
        self.client.get_message = Mock(
            return_value={
                "id": "message-id",
                "subject": "Databricks verification code",
                "toRecipients": recipient,
                "body": {"content": "Your verification code is ABC-123."},
            }
        )

        code = self.client.wait_for_databricks_code(
            not_before=now,
            timeout=1,
            poll_interval=0,
        )

        self.assertEqual(code, "ABC-123")
        self.client.get_message.assert_called_once_with("message-id")


if __name__ == "__main__":
    unittest.main()
