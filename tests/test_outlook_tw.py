import unittest
from datetime import datetime, timezone
from unittest.mock import Mock, patch

from outlook_tw import OutlookTwClient, OutlookTwError


class OutlookTwTests(unittest.TestCase):
    def test_create_account_uses_anonymous_generate_api(self):
        response = Mock()
        response.status_code = 200
        response.json.return_value = {
            "email": "TempUser@outlook.tw",
            "expires": 1_800_000_000_000,
            "anonymous": True,
        }
        session = Mock()
        session.get.return_value = response

        with patch.object(OutlookTwClient, "_build_session", return_value=session):
            client = OutlookTwClient.create_account(length=12)

        self.assertEqual(client.address, "tempuser@outlook.tw")
        self.assertEqual(client.expires_at.tzinfo, timezone.utc)
        _, kwargs = session.get.call_args
        self.assertEqual(kwargs["params"], {"length": 12, "domainIndex": 0})

    def test_wait_for_databricks_code_reads_message_detail(self):
        client = OutlookTwClient("tempuser@outlook.tw", session=Mock())
        now = datetime.now(timezone.utc)
        client.list_messages = Mock(
            return_value=[
                {
                    "id": 42,
                    "subject": "Databricks verification code",
                    "from_address": "noreply@databricks.com",
                    "received_at": now.isoformat(),
                }
            ]
        )
        client.get_message = Mock(
            return_value={
                "id": 42,
                "subject": "Databricks verification code",
                "html_content": "<p>Your verification code is ABC-123.</p>",
            }
        )

        code = client.wait_for_databricks_code(
            not_before=now,
            timeout=1,
            poll_interval=0,
        )

        self.assertEqual(code, "ABC-123")
        client.get_message.assert_called_once_with("42")

    def test_rejects_invalid_message_list(self):
        client = OutlookTwClient("tempuser@outlook.tw", session=Mock())
        client._get_json = Mock(return_value={"messages": []})
        with self.assertRaises(OutlookTwError):
            client.list_messages()

    def test_message_text_includes_backend_verification_code(self):
        text = OutlookTwClient._message_text(
            {
                "sender": "Databricks",
                "verification_code": "ABC-123",
                "text_content": "Sign-in request",
            }
        )
        self.assertIn("ABC-123", text)
        self.assertIn("Sign-in request", text)


if __name__ == "__main__":
    unittest.main()
