"""使用 Microsoft Graph OAuth refresh token 读取 Outlook 验证码。"""

from __future__ import annotations

import os
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from mail_common import extract_verification_code


class OutlookGraphError(RuntimeError):
    """Microsoft OAuth 或 Graph 请求失败。"""


def _is_uuid(value: str) -> bool:
    try:
        uuid.UUID(value)
        return True
    except (ValueError, AttributeError):
        return False


def _write_lines_atomic(path: Path, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f"{path.name}.tmp")
    text = "".join(f"{line.rstrip()}\n" for line in lines if line.strip())
    temporary.write_text(text, encoding="utf-8")
    os.replace(temporary, path)


@dataclass(frozen=True)
class OutlookCredentials:
    email: str
    password: str
    refresh_token: str
    client_id: str

    @classmethod
    def parse(cls, value: str) -> "OutlookCredentials":
        try:
            email, password, remainder = value.strip().split("----", 2)
        except ValueError:
            raise ValueError(
                "Outlook credential must be email----password----refresh_token----client_id"
            )

        # 同时兼容两种常见导出顺序：
        # email----password----refresh_token----client_id
        # email----password----client_id----refresh_token
        first, separator, tail = remainder.partition("----")
        if separator and _is_uuid(first.strip()):
            client_id, refresh_token = first, tail
        else:
            try:
                refresh_token, client_id = remainder.rsplit("----", 1)
            except ValueError:
                raise ValueError(
                    "Outlook credential must contain refresh_token and client_id"
                )
        email, password, refresh_token, client_id = (
            part.strip()
            for part in (email, password, refresh_token, client_id)
        )
        if not email or "@" not in email:
            raise ValueError("Outlook credential contains an invalid email")
        if not refresh_token:
            raise ValueError("Outlook credential has no refresh token")
        try:
            uuid.UUID(client_id)
        except (ValueError, AttributeError) as exc:
            raise ValueError("Outlook credential contains an invalid client_id") from exc
        return cls(email.lower(), password, refresh_token, client_id)

    @classmethod
    def from_file(
        cls,
        path: str | Path,
        email: str | None = None,
    ) -> "OutlookCredentials":
        records = []
        with Path(path).expanduser().open("r", encoding="utf-8-sig") as handle:
            for raw_line in handle:
                line = raw_line.strip()
                if line and not line.startswith("#"):
                    records.append(cls.parse(line))
        if not records:
            raise ValueError(f"No Outlook credentials found in {path}")
        if email:
            email = email.strip().lower()
            for record in records:
                if record.email == email:
                    return record
            raise ValueError(f"Outlook credential file has no record for {email}")
        # Pending queues are processed from top to bottom when no address is selected.
        return records[0]

    @classmethod
    def from_environment(
        cls,
        email: str | None = None,
        password: str | None = None,
    ) -> "OutlookCredentials":
        combined = os.environ.get("OUTLOOK_CREDENTIAL", "").strip()
        if combined:
            record = cls.parse(combined)
            if email and record.email != email.strip().lower():
                raise ValueError("--email does not match OUTLOOK_CREDENTIAL")
            return record

        address = (email or os.environ.get("OUTLOOK_EMAIL", "")).strip().lower()
        secret = password if password is not None else os.environ.get("OUTLOOK_PASSWORD", "")
        refresh_token = os.environ.get("OUTLOOK_REFRESH_TOKEN", "").strip()
        client_id = os.environ.get("OUTLOOK_CLIENT_ID", "").strip()
        return cls.parse(f"{address}----{secret}----{refresh_token}----{client_id}")


def move_outlook_credential(
    source: str | Path,
    destination: str | Path,
    email: str,
) -> bool:
    """Move one credential record after a successful registration."""
    source_path = Path(source).expanduser().resolve()
    destination_path = Path(destination).expanduser().resolve()
    if source_path == destination_path:
        return False

    target_email = email.strip().lower()
    source_lines = (
        source_path.read_text(encoding="utf-8-sig").splitlines()
        if source_path.exists()
        else []
    )
    destination_lines = (
        destination_path.read_text(encoding="utf-8-sig").splitlines()
        if destination_path.exists()
        else []
    )

    destination_emails: set[str] = set()
    for line in destination_lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        try:
            destination_emails.add(OutlookCredentials.parse(stripped).email)
        except ValueError:
            continue

    moved_lines: list[str] = []
    remaining_lines: list[str] = []
    for line in source_lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            remaining_lines.append(line)
            continue
        try:
            record = OutlookCredentials.parse(stripped)
        except ValueError:
            remaining_lines.append(line)
            continue
        if record.email == target_email:
            moved_lines.append(stripped)
        else:
            remaining_lines.append(line)

    if not moved_lines:
        return target_email in destination_emails

    if target_email not in destination_emails:
        destination_lines.append(moved_lines[0])
        _write_lines_atomic(destination_path, destination_lines)
    _write_lines_atomic(source_path, remaining_lines)
    return True


class OutlookGraphClient:
    TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
    GRAPH_BASE = "https://graph.microsoft.com/v1.0"
    GRAPH_SCOPE = "https://graph.microsoft.com/.default"

    def __init__(
        self,
        credentials: OutlookCredentials,
        timeout: int = 20,
        session: requests.Session | None = None,
    ):
        self.address = credentials.email
        self.password = credentials.password
        self.refresh_token = credentials.refresh_token
        self.client_id = credentials.client_id
        self.timeout = timeout
        self.session = session or self._build_session()
        self.access_token = ""
        self.access_token_deadline = 0.0

    @staticmethod
    def _build_session() -> requests.Session:
        session = requests.Session()
        session.headers.update(
            {
                "Accept": "application/json",
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
    def from_sources(
        cls,
        *,
        credential_file: str | Path | None = None,
        email: str | None = None,
        password: str | None = None,
        timeout: int = 20,
    ) -> "OutlookGraphClient":
        if credential_file:
            credentials = OutlookCredentials.from_file(credential_file, email=email)
        else:
            credentials = OutlookCredentials.from_environment(
                email=email,
                password=password,
            )
        return cls(credentials, timeout=timeout)

    @staticmethod
    def _error_code(response: requests.Response) -> str:
        try:
            data = response.json()
        except ValueError:
            return response.reason or f"HTTP {response.status_code}"
        error = data.get("error")
        if isinstance(error, dict):
            return str(error.get("code") or error.get("message") or "Graph error")
        return str(error or data.get("error_codes") or f"HTTP {response.status_code}")

    def _refresh_access_token(self, force: bool = False) -> str:
        if (
            not force
            and self.access_token
            and time.monotonic() < self.access_token_deadline
        ):
            return self.access_token

        try:
            response = self.session.post(
                self.TOKEN_URL,
                data={
                    "client_id": self.client_id,
                    "grant_type": "refresh_token",
                    "refresh_token": self.refresh_token,
                    "scope": self.GRAPH_SCOPE,
                },
                timeout=self.timeout,
            )
        except requests.RequestException as exc:
            raise OutlookGraphError(f"Microsoft OAuth 网络请求失败: {exc}") from exc
        if response.status_code >= 400:
            raise OutlookGraphError(
                f"Microsoft OAuth {response.status_code}: {self._error_code(response)}"
            )

        try:
            data = response.json()
        except ValueError as exc:
            raise OutlookGraphError("Microsoft OAuth 返回了无效 JSON") from exc
        access_token = str(data.get("access_token") or "")
        if not access_token:
            raise OutlookGraphError("Microsoft OAuth 响应中没有 access_token")

        rotated = str(data.get("refresh_token") or "")
        if rotated:
            self.refresh_token = rotated
        try:
            expires_in = max(60, int(data.get("expires_in") or 3600))
        except (TypeError, ValueError):
            expires_in = 3600
        self.access_token = access_token
        self.access_token_deadline = time.monotonic() + expires_in - 60
        return access_token

    def _graph_get(self, path: str, params: dict | None = None) -> dict:
        for attempt in range(2):
            token = self._refresh_access_token(force=attempt > 0)
            try:
                response = self.session.get(
                    f"{self.GRAPH_BASE}{path}",
                    headers={"Authorization": f"Bearer {token}"},
                    params=params,
                    timeout=self.timeout,
                )
            except requests.RequestException as exc:
                raise OutlookGraphError(f"Microsoft Graph 网络请求失败: {exc}") from exc
            if response.status_code == 401 and attempt == 0:
                continue
            if response.status_code >= 400:
                raise OutlookGraphError(
                    f"Microsoft Graph {response.status_code}: {self._error_code(response)}"
                )
            try:
                data = response.json()
            except ValueError as exc:
                raise OutlookGraphError("Microsoft Graph 返回了无效 JSON") from exc
            if not isinstance(data, dict):
                raise OutlookGraphError("Microsoft Graph 返回格式异常")
            return data
        raise OutlookGraphError("Microsoft Graph authentication failed")

    def list_messages(self, top: int = 50) -> list[dict]:
        data = self._graph_get(
            "/me/mailFolders/inbox/messages",
            params={
                "$top": str(max(1, min(int(top), 100))),
                "$orderby": "receivedDateTime desc",
                "$select": (
                    "id,subject,from,toRecipients,receivedDateTime,bodyPreview"
                ),
            },
        )
        messages = data.get("value") or []
        if not isinstance(messages, list):
            raise OutlookGraphError("Microsoft Graph 邮件列表格式异常")
        return [message for message in messages if isinstance(message, dict)]

    def get_message(self, message_id: str) -> dict:
        encoded_id = quote(message_id, safe="")
        return self._graph_get(
            f"/me/messages/{encoded_id}",
            params={
                "$select": (
                    "id,subject,from,toRecipients,receivedDateTime,bodyPreview,body"
                )
            },
        )

    @staticmethod
    def _parse_time(value: str | None) -> datetime | None:
        if not value:
            return None
        try:
            parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return None
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)

    @staticmethod
    def _address_from(value) -> str:
        if not isinstance(value, dict):
            return ""
        email_address = value.get("emailAddress")
        if not isinstance(email_address, dict):
            return ""
        return str(email_address.get("address") or "").strip().lower()

    def _is_for_address(self, message: dict) -> bool:
        recipients = message.get("toRecipients") or []
        addresses = {
            self._address_from(recipient)
            for recipient in recipients
            if isinstance(recipient, dict)
        }
        addresses.discard("")
        # Graph 某些旧邮件可能省略收件人；有数据时必须精确匹配子邮箱。
        return not addresses or self.address in addresses

    @classmethod
    def _message_text(cls, message: dict) -> str:
        sender = cls._address_from(message.get("from"))
        body = message.get("body") or {}
        body_content = body.get("content", "") if isinstance(body, dict) else ""
        return "\n".join(
            (
                str(message.get("subject") or ""),
                sender,
                str(message.get("bodyPreview") or ""),
                str(body_content or ""),
            )
        )

    def wait_for_databricks_code(
        self,
        not_before: datetime | None = None,
        timeout: int = 180,
        poll_interval: float = 3,
    ) -> str:
        if not_before is None:
            not_before = datetime.now(timezone.utc)
        elif not_before.tzinfo is None:
            not_before = not_before.replace(tzinfo=timezone.utc)
        threshold = not_before.astimezone(timezone.utc) - timedelta(seconds=30)

        deadline = time.monotonic() + timeout
        checked_ids: set[str] = set()
        last_error = None
        while time.monotonic() < deadline:
            try:
                for summary in self.list_messages():
                    message_id = str(summary.get("id") or "")
                    if not message_id or message_id in checked_ids:
                        continue
                    received_at = self._parse_time(summary.get("receivedDateTime"))
                    if received_at and received_at < threshold:
                        checked_ids.add(message_id)
                        continue
                    if not self._is_for_address(summary):
                        checked_ids.add(message_id)
                        continue
                    if "databricks" not in self._message_text(summary).lower():
                        checked_ids.add(message_id)
                        continue

                    detail = self.get_message(message_id)
                    if not self._is_for_address(detail):
                        checked_ids.add(message_id)
                        continue
                    code = extract_verification_code(self._message_text(detail))
                    if code:
                        return code
            except OutlookGraphError as exc:
                last_error = exc

            time.sleep(poll_interval)

        message = f"等待 Outlook 验证码超时（{timeout} 秒）"
        if last_error:
            message += f": {last_error}"
        raise TimeoutError(message)

    def test_connection(self) -> int:
        """刷新 OAuth token 并返回 Inbox 当前可见消息数（最多 1）。"""
        self._refresh_access_token()
        return len(self.list_messages(top=1))

    def close(self) -> None:
        self.access_token = ""
        self.session.close()
