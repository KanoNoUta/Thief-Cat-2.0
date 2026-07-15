"""outlook.tw anonymous temporary mailbox API client."""

from __future__ import annotations

import html
import re
import time
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from mail_common import extract_verification_code


class OutlookTwError(RuntimeError):
    """outlook.tw request or response error."""


class OutlookTwClient:
    BASE_URL = "https://outlook.tw"

    def __init__(
        self,
        address: str,
        *,
        expires_at: datetime | None = None,
        timeout: int = 20,
        session: requests.Session | None = None,
    ):
        address = address.strip().lower()
        if not re.fullmatch(r"[a-z0-9._-]{1,64}@outlook\.tw", address):
            raise ValueError(f"Invalid outlook.tw address: {address!r}")
        self.address = address
        self.password = ""
        self.expires_at = expires_at
        self.timeout = timeout
        self.session = session or self._build_session()

    @staticmethod
    def _build_session() -> requests.Session:
        session = requests.Session()
        session.headers.update(
            {
                "Accept": "application/json",
                "Cache-Control": "no-cache",
                "Pragma": "no-cache",
                "User-Agent": "Thief-Cat/2.0",
            }
        )
        retry = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=0.5,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset({"GET", "HEAD", "OPTIONS"}),
            respect_retry_after_header=True,
        )
        session.mount("https://", HTTPAdapter(max_retries=retry))
        return session

    @classmethod
    def create_account(
        cls,
        *,
        length: int = 12,
        domain_index: int = 0,
        timeout: int = 20,
    ) -> "OutlookTwClient":
        length = max(8, min(30, int(length)))
        session = cls._build_session()
        try:
            response = session.get(
                f"{cls.BASE_URL}/api/generate",
                params={"length": length, "domainIndex": max(0, int(domain_index))},
                timeout=timeout,
            )
            data = cls._json_response(response, "generate mailbox")
            address = str(data.get("email") or "").strip().lower()
            expires_ms = data.get("expires")
            expires_at = None
            if isinstance(expires_ms, (int, float)):
                expires_at = datetime.fromtimestamp(
                    float(expires_ms) / 1000,
                    tz=timezone.utc,
                )
            return cls(
                address,
                expires_at=expires_at,
                timeout=timeout,
                session=session,
            )
        except Exception:
            session.close()
            raise

    @staticmethod
    def _json_response(response: requests.Response, action: str):
        if response.status_code >= 400:
            body = (response.text or "").strip()[:300]
            raise OutlookTwError(
                f"outlook.tw {action} failed: HTTP {response.status_code}: {body}"
            )
        try:
            return response.json()
        except ValueError as exc:
            raise OutlookTwError(f"outlook.tw {action} returned invalid JSON") from exc

    def _get_json(self, path: str, params: dict | None = None):
        try:
            response = self.session.get(
                f"{self.BASE_URL}{path}",
                params=params,
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise OutlookTwError(f"outlook.tw network error: {exc}") from exc
        return self._json_response(response, path)

    def list_messages(self) -> list[dict]:
        data = self._get_json("/api/emails", params={"mailbox": self.address})
        if not isinstance(data, list):
            raise OutlookTwError("outlook.tw message list has an invalid shape")
        return [item for item in data if isinstance(item, dict)]

    def get_message(self, message_id: str | int) -> dict:
        data = self._get_json(f"/api/email/{quote(str(message_id), safe='')}")
        if not isinstance(data, dict):
            raise OutlookTwError("outlook.tw message detail has an invalid shape")
        return data

    @staticmethod
    def _parse_time(value) -> datetime | None:
        if not value:
            return None
        try:
            parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        except ValueError:
            return None
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)

    @staticmethod
    def _message_text(message: dict) -> str:
        raw_html = str(
            message.get("html_content")
            or message.get("html")
            or message.get("content")
            or message.get("text_content")
            or ""
        )
        plain_html = html.unescape(re.sub(r"<[^>]+>", " ", raw_html))
        fields = (
            message.get("subject"),
            message.get("from_address"),
            message.get("from_name"),
            message.get("sender"),
            message.get("preview"),
            message.get("bodyPreview"),
            message.get("verification_code"),
            plain_html,
        )
        return "\n".join(str(value or "") for value in fields)

    def wait_for_databricks_code(
        self,
        not_before: datetime | None = None,
        timeout: int = 180,
        poll_interval: float = 5,
    ) -> str:
        if not_before is None:
            not_before = datetime.now(timezone.utc)
        elif not_before.tzinfo is None:
            not_before = not_before.replace(tzinfo=timezone.utc)
        threshold = not_before.astimezone(timezone.utc) - timedelta(seconds=30)

        deadline = time.monotonic() + timeout
        ignored_ids: set[str] = set()
        last_error = None
        while time.monotonic() < deadline:
            try:
                for summary in self.list_messages():
                    message_id = str(summary.get("id") or "")
                    if not message_id or message_id in ignored_ids:
                        continue
                    received_at = self._parse_time(
                        summary.get("received_at")
                        or summary.get("created_at")
                        or summary.get("receivedDateTime")
                    )
                    if received_at and received_at < threshold:
                        ignored_ids.add(message_id)
                        continue
                    summary_text = self._message_text(summary)
                    if "databricks" not in summary_text.lower():
                        ignored_ids.add(message_id)
                        continue

                    detail = self.get_message(message_id)
                    code = extract_verification_code(self._message_text(detail))
                    if code:
                        return code
            except OutlookTwError as exc:
                last_error = exc
            time.sleep(max(1, poll_interval))

        message = f"等待 outlook.tw 验证码超时（{timeout} 秒）"
        if last_error:
            message += f": {last_error}"
        raise TimeoutError(message)

    def test_connection(self) -> int:
        return len(self.list_messages())

    def close(self) -> None:
        self.session.close()
