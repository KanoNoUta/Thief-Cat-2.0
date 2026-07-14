import csv
import inspect
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import auto_register
from mail_common import extract_verification_code


class WorkspaceTests(unittest.TestCase):
    def test_clicks_existing_work_account_card(self):
        page = Mock()
        account = Mock()
        account.count.return_value = 1
        account.is_visible.return_value = True
        page.get_by_text.return_value = account

        clicked = auto_register.click_existing_work_account(
            page,
            "existingaccount",
        )

        self.assertTrue(clicked)
        page.get_by_text.assert_called_once_with(
            "existingaccount",
            exact=True,
        )
        account.click.assert_called_once_with()

    def test_detects_databricks_account_setup_urls(self):
        self.assertTrue(
            auto_register.is_onboarding_url(
                "https://login.databricks.com/login/accounts/abc?next_url=%2Fsetup"
            )
        )
        self.assertTrue(
            auto_register.is_onboarding_url(
                "https://accounts.cloud.databricks.com/setup"
            )
        )
        self.assertFalse(
            auto_register.is_onboarding_url(
                "https://accounts.cloud.databricks.com/workspaces"
            )
        )

    def test_resilient_goto_retries_interrupted_navigation(self):
        page = Mock()
        page.url = "https://login.databricks.com/setup"
        page.goto.side_effect = [
            auto_register.PlaywrightError("interrupted by another navigation"),
            "response",
        ]
        with patch.object(
            auto_register,
            "wait_for_navigation_stable",
            return_value=page.url,
        ):
            result = auto_register.resilient_goto(
                page,
                "https://accounts.cloud.databricks.com/workspaces",
            )
        self.assertEqual(result, "response")
        self.assertEqual(page.goto.call_count, 2)

    def test_workspace_domain_accepts_supported_hosts(self):
        self.assertEqual(
            auto_register.workspace_domain(
                "https://dbc-1234.cloud.databricks.com/explore/model-services"
            ),
            "https://dbc-1234.cloud.databricks.com",
        )
        self.assertEqual(
            auto_register.workspace_domain("https://demo.azuredatabricks.net/path"),
            "https://demo.azuredatabricks.net",
        )

    def test_workspace_domain_rejects_accounts_and_suffix_tricks(self):
        self.assertEqual(
            auto_register.workspace_domain(
                "https://accounts.cloud.databricks.com/workspaces"
            ),
            "",
        )
        self.assertEqual(
            auto_register.workspace_domain(
                "https://demo.cloud.databricks.com.example.test/path"
            ),
            "",
        )

    def test_normalize_workspace_rejects_invalid_input(self):
        with self.assertRaises(ValueError):
            auto_register.normalize_workspace("https://example.com")


class TokenAndMailTests(unittest.TestCase):
    def test_detects_qr_from_frame_probe(self):
        frame = Mock()
        frame.evaluate.return_value = True
        page = Mock()
        page.frames = [frame]
        self.assertTrue(auto_register.qr_challenge_visible(page))

    def test_detects_loading_indicator_from_frame_probe(self):
        frame = Mock()
        frame.evaluate.return_value = True
        page = Mock()
        page.frames = [frame]
        self.assertTrue(auto_register.loading_indicator_visible(page))

    def test_detects_blank_databricks_onboarding_page(self):
        page = Mock()
        page.url = "https://login.databricks.com/login/accounts/account-id"
        page.evaluate.return_value = True

        self.assertTrue(auto_register.onboarding_page_is_blank(page))
        page.evaluate.assert_called_once_with(auto_register.ONBOARDING_BLANK_SCRIPT)

    def test_does_not_treat_non_databricks_blank_page_as_onboarding(self):
        page = Mock()
        page.url = "https://example.test/blank"

        self.assertFalse(auto_register.onboarding_page_is_blank(page))
        page.evaluate.assert_not_called()

    def test_open_known_workspace_requires_authenticated_workspace_url(self):
        page = Mock()
        page.url = "https://login.databricks.com/setup"

        with patch.object(auto_register, "resilient_goto"), patch.object(
            auto_register,
            "wait_for_navigation_stable",
            return_value=page.url,
        ):
            self.assertEqual(
                auto_register.open_known_workspace(
                    page,
                    "https://dbc-test.cloud.databricks.com",
                ),
                "",
            )

    def test_complete_onboarding_can_bypass_blank_page_with_known_workspace(self):
        page = Mock()
        page.url = "https://login.databricks.com/login/accounts/account-id"
        target = "https://dbc-test.cloud.databricks.com"

        def open_workspace(current_page, value, **_kwargs):
            self.assertIs(current_page, page)
            self.assertEqual(value, target)
            page.url = target
            return target

        with patch.object(auto_register, "workspace_links", return_value=[]), patch.object(
            auto_register, "accept_cookies", return_value=False
        ), patch.object(
            auto_register, "click_existing_work_account", return_value=False
        ), patch.object(
            auto_register, "fill_optional_profile", return_value=False
        ), patch.object(
            auto_register, "select_first_available_option", return_value=False
        ), patch.object(
            auto_register, "click_named_button", return_value=False
        ), patch.object(
            auto_register, "wait_for_navigation_stable", return_value=page.url
        ), patch.object(
            auto_register, "wait_for_qr_resolution", return_value=False
        ), patch.object(
            auto_register, "loading_indicator_visible", return_value=False
        ), patch.object(
            auto_register, "onboarding_page_is_blank", return_value=True
        ), patch.object(
            auto_register, "open_known_workspace", side_effect=open_workspace
        ) as opener, patch.object(auto_register, "screenshot") as screenshot:
            auto_register.complete_onboarding(
                page,
                "existingaccount",
                "",
                preferred_workspace=target,
            )

        opener.assert_called_once()
        screenshot.assert_not_called()

    def test_qr_wait_continues_without_enter_prompt(self):
        page = Mock()
        page.is_closed.return_value = False
        with patch.object(
            auto_register,
            "qr_challenge_visible",
            side_effect=[True, True, False],
        ), patch.object(auto_register, "screenshot"):
            self.assertTrue(
                auto_register.wait_for_qr_resolution(
                    page,
                    allow_manual=True,
                    timeout=5,
                )
            )
        page.wait_for_timeout.assert_called_once_with(500)

    def test_named_action_uses_first_visible_match(self):
        hidden = Mock()
        hidden.is_visible.return_value = False
        visible = Mock()
        visible.is_visible.return_value = True
        matches = Mock()
        matches.count.return_value = 2
        matches.nth.side_effect = [hidden, visible]
        page = Mock()
        page.get_by_role.return_value = matches

        self.assertIs(
            auto_register.named_action(page, ("Manage",), timeout=100),
            visible,
        )

    def test_named_action_falls_back_to_exact_visible_text(self):
        roles = Mock()
        roles.count.return_value = 0
        text_match = Mock()
        text_match.count.return_value = 1
        text_match.nth.return_value.is_visible.return_value = True
        page = Mock()
        page.get_by_role.return_value = roles
        page.get_by_text.return_value = text_match

        result = auto_register.named_action(page, ("Manage",), timeout=100)

        self.assertIs(result, text_match.nth.return_value)
        page.get_by_text.assert_called_once_with("Manage", exact=True)

    def test_generate_token_uses_ai_gateway_one_click_flow(self):
        page = Mock()
        with patch.object(
            auto_register,
            "generate_glm_access_token",
            return_value="dapiGatewayGenerated123",
        ) as gateway:
            token = auto_register.generate_token(
                page,
                "https://dbc-test.cloud.databricks.com",
            )
        self.assertEqual(token, "dapiGatewayGenerated123")
        gateway.assert_called_once_with(
            page,
            "https://dbc-test.cloud.databricks.com",
            allow_manual=True,
        )

    def test_non_qr_paths_do_not_prompt_for_input(self):
        functions = (
            auto_register.choose_workspace,
            auto_register.complete_onboarding,
            auto_register.submit_email,
            auto_register.verify_email,
            auto_register.generate_token,
        )
        for function in functions:
            self.assertNotIn("input(", inspect.getsource(function), function.__name__)

    def test_token_pattern(self):
        self.assertIsNotNone(auto_register.TOKEN_PATTERN.fullmatch("dapiAbc_123.X-y"))
        self.assertIsNone(auto_register.TOKEN_PATTERN.fullmatch("prefix-dapiAbc123"))

    def test_extracts_segmented_verification_code(self):
        text = "Your verification code is ABC-123. It expires soon."
        self.assertEqual(extract_verification_code(text), "ABC-123")

    def test_extracts_numeric_verification_code(self):
        text = "Databricks login code: 482901"
        self.assertEqual(extract_verification_code(text), "482901")

    def test_does_not_extract_unrelated_short_number(self):
        self.assertEqual(extract_verification_code("Invoice 12345"), "")


class CsvAndCliTests(unittest.TestCase):
    def test_save_record_appends_without_overwriting(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "accounts.csv"
            with patch.object(auto_register, "OUTPUT_CSV", output):
                auto_register.save_record(
                    "one@example.test", "pw1", "https://a.cloud.databricks.com", "dapiOne", "OK"
                )
                auto_register.save_record(
                    "two@example.test", "pw2", "", "", "FAIL", "test"
                )

            with output.open("r", newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.reader(handle))
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[0][0], "Time")
            self.assertEqual(rows[1][1], "one@example.test")
            self.assertEqual(rows[2][5], "FAIL")

    def test_domain_token_output_migrates_legacy_schema_and_upserts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "glm_keys.csv"
            output.write_text(
                "时间,邮箱,工作区域名,Token,备注\n"
                "2026-01-01,a@example.test,https://old.cloud.databricks.com,dapiOld123,\n",
                encoding="utf-8",
            )
            with patch.object(auto_register, "KEYS_CSV", output):
                auto_register.save_domain_token(
                    "https://new.cloud.databricks.com/path",
                    "dapiNew456",
                )

            with output.open("r", newline="", encoding="utf-8") as handle:
                rows = list(csv.DictReader(handle))
            self.assertEqual(list(rows[0]), ["Domain", "Token"])
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[-1]["Domain"], "https://new.cloud.databricks.com")
            self.assertEqual(rows[-1]["Token"], "dapiNew456")

    def test_parser_defaults(self):
        args = auto_register.build_parser().parse_args([])
        self.assertEqual(args.browser_channel, "chrome")
        self.assertTrue(args.outlook_success_file.endswith("outlook_success.txt"))
        self.assertFalse(args.headless)
        self.assertFalse(args.resume)

    def test_outlook_mail_test_does_not_start_browser(self):
        client = Mock()
        client.address = "child@outlook.com"
        client.test_connection.return_value = 1
        with patch.object(auto_register, "create_mail_client", return_value=client):
            result = auto_register.main(["--mail-test"])
        self.assertEqual(result, 0)
        client.test_connection.assert_called_once_with()
        client.close.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
