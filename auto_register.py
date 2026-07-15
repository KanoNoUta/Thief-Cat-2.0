"""注册或登录 Databricks，进入 AI Gateway 并生成 GLM Access Token。"""

from __future__ import annotations

import argparse
import csv
import random
import re
import subprocess
import sys
import time
import traceback
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

from playwright.sync_api import (
    Browser,
    BrowserContext,
    Error as PlaywrightError,
    Page,
    Playwright,
    sync_playwright,
)

from outlook_graph import OutlookGraphClient, move_outlook_credential
from outlook_tw import OutlookTwClient


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_CSV = BASE_DIR / "registered_accounts.csv"
KEYS_CSV = BASE_DIR / "glm_keys.csv"
OUTLOOK_PENDING_FILE = BASE_DIR / "outlook_pending.txt"
OUTLOOK_SUCCESS_FILE = BASE_DIR / "outlook_success.txt"
SCREENSHOT_DIR = BASE_DIR / "screenshots"
SIGNUP_URL = "https://login.databricks.com/signup?provider=DB"
WORKSPACES_URL = "https://accounts.cloud.databricks.com/workspaces"
TOKEN_PATTERN = re.compile(r"dapi[A-Za-z0-9._-]+")
AI_GATEWAY_LABEL = "AI Gateway"
GLM_MODEL_LABEL = "GLM 5.2"
GLM_TOKEN_ACTION = "Generate Access Token"

OTP_SELECTORS = (
    'input[autocomplete="one-time-code"]',
    'input[name="code"]',
    'input[name="otp"]',
    'input[inputmode="numeric"]',
    'input[placeholder*="code" i]',
)

FIRST_NAMES = (
    "James",
    "John",
    "Robert",
    "Michael",
    "William",
    "David",
    "Richard",
    "Joseph",
    "Thomas",
    "Charles",
)
LAST_NAMES = (
    "Smith",
    "Johnson",
    "Williams",
    "Brown",
    "Jones",
    "Garcia",
    "Miller",
    "Davis",
    "Rodriguez",
    "Martinez",
)


def create_mail_client(
    provider: str = "outlook",
    email: str | None = None,
    password: str | None = None,
    outlook_credential_file: str | None = None,
) -> OutlookGraphClient | OutlookTwClient:
    if provider == "outlook":
        return OutlookGraphClient.from_sources(
            credential_file=outlook_credential_file,
            email=email,
            password=password,
        )
    if provider == "outlook_tw":
        if password:
            raise ValueError("outlook.tw temporary mailboxes do not use passwords")
        return OutlookTwClient(email) if email else OutlookTwClient.create_account()
    raise ValueError(f"Unknown mail provider: {provider}")


def screenshot(page: Page | None, name: str) -> Path | None:
    """尽力保存现场；截图失败不能覆盖原始异常。"""
    if page is None:
        return None
    try:
        SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
        path = SCREENSHOT_DIR / f"{name}_{int(time.time())}.png"
        page.screenshot(path=str(path), full_page=True)
        print(f"  Screenshot: {path}")
        return path
    except Exception:
        return None


def save_record(
    email: str,
    password: str,
    workspace: str,
    token: str,
    status: str,
    note: str = "",
) -> None:
    """追加结果，不覆盖已有敏感数据。"""
    needs_header = not OUTPUT_CSV.exists() or OUTPUT_CSV.stat().st_size == 0
    with OUTPUT_CSV.open("a", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        if needs_header:
            writer.writerow(
                ["Time", "Email", "Password", "Workspace", "Token", "Status", "Note"]
            )
        writer.writerow(
            [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                email,
                password,
                workspace,
                token,
                status,
                note,
            ]
        )
    print(f"  Saved to {OUTPUT_CSV}")


def save_domain_token(workspace: str, token: str) -> None:
    """Upsert the release output as exactly Domain,Token."""
    domain = normalize_workspace(workspace)
    if not TOKEN_PATTERN.fullmatch(token or ""):
        raise ValueError("Cannot save an invalid Databricks token")

    records: dict[str, str] = {}
    if KEYS_CSV.exists() and KEYS_CSV.stat().st_size:
        with KEYS_CSV.open("r", newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                legacy_domain = next(
                    (
                        str(row.get(key) or "").strip()
                        for key in ("Domain", "Workspace", "工作区", "工作区域名")
                        if str(row.get(key) or "").strip()
                    ),
                    "",
                )
                legacy_token = str(row.get("Token") or "").strip()
                if not legacy_domain or not TOKEN_PATTERN.fullmatch(legacy_token):
                    continue
                try:
                    legacy_domain = normalize_workspace(legacy_domain)
                except ValueError:
                    continue
                records[legacy_domain] = legacy_token

    records[domain] = token
    temporary = KEYS_CSV.with_name(f"{KEYS_CSV.name}.tmp")
    with temporary.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Domain", "Token"])
        writer.writerows(records.items())
    temporary.replace(KEYS_CSV)
    print(f"  Saved Domain,Token to {KEYS_CSV}")


def visible_locator(page: Page, selectors: tuple[str, ...], timeout: int = 500):
    for selector in selectors:
        try:
            locator = page.locator(selector)
            count = locator.count()
            if count == 1 and locator.is_visible(timeout=timeout):
                return locator
        except Exception:
            continue
    return None


def named_button(page: Page, names: tuple[str, ...], timeout: int = 500):
    for name in names:
        try:
            button = page.get_by_role("button", name=name, exact=True)
            if button.count() == 1 and button.is_visible(timeout=timeout):
                return button
        except Exception:
            continue
    return None


def click_named_button(page: Page, names: tuple[str, ...], timeout: int = 500) -> bool:
    button = named_button(page, names, timeout=timeout)
    if not button:
        return False
    button.click()
    return True


def accept_cookies(page: Page) -> bool:
    return click_named_button(
        page,
        (
            "全部接受",
            "接受所有 Cookie",
            "Accept All",
            "Accept all",
            "Accept all cookies",
            "Allow all",
        ),
        timeout=500,
    )


def click_existing_work_account(page: Page, account_name: str = "") -> bool:
    """在 Welcome back 页面选择已存在的 Work account，不创建新账号。"""
    candidates = []
    if account_name:
        candidates.append(account_name)
    candidates.extend(("You are the account owner", "你是账户所有者"))

    for text in candidates:
        try:
            locator = page.get_by_text(text, exact=True)
            if locator.count() == 1 and locator.is_visible(timeout=300):
                locator.click()
                print(f"  Selected existing Work account: {account_name or text}")
                page.wait_for_timeout(500)
                return True
        except Exception:
            continue
    return False


def wait_until(page: Page, predicate, timeout: float, interval: float = 0.25) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            if predicate():
                return True
        except Exception:
            pass
        page.wait_for_timeout(int(interval * 1000))
    return False


QR_DETECTION_SCRIPT = r"""
() => {
  const bodyText = (document.body?.innerText || '').toLowerCase();
  const hasQrText = /qr\s*code|scan.{0,40}(code|phone)|二维码|扫码|手机扫描/.test(bodyText);
  const selectors = [
    'img[alt*="qr" i]', 'img[src*="qr" i]',
    '[aria-label*="qr" i]', '[data-testid*="qr" i]',
    '[id*="qr" i]', '[class*="qr" i]'
  ];
  const visibleSquare = (element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) return false;
    if (rect.width < 100 || rect.height < 100 || rect.width > 700 || rect.height > 700) return false;
    const ratio = rect.width / rect.height;
    return ratio >= 0.75 && ratio <= 1.25;
  };
  for (const selector of selectors) {
    if (Array.from(document.querySelectorAll(selector)).some(visibleSquare)) return true;
  }
  if (!hasQrText) return false;
  return Array.from(document.querySelectorAll('canvas, svg, img')).some(visibleSquare);
}
"""

LOADING_DETECTION_SCRIPT = r"""
() => {
  const visible = (element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' && style.visibility !== 'hidden' &&
      Number(style.opacity) !== 0 && rect.width > 0 && rect.height > 0;
  };
  const selectors = [
    '[role="progressbar"]', '[aria-busy="true"]',
    '[data-testid*="spinner" i]', '[data-testid*="loading" i]',
    '[class*="spinner" i]', '[class*="loading" i]'
  ];
  if (selectors.some(selector => Array.from(document.querySelectorAll(selector)).some(visible))) return true;
  const text = (document.body?.innerText || '').trim().toLowerCase();
  if (/^(loading|please wait|正在加载|请稍候)[.…]*$/.test(text)) return true;
  if (text.length > 20) return false;
  return Array.from(document.querySelectorAll('svg, canvas')).some(element => {
    if (!visible(element)) return false;
    const rect = element.getBoundingClientRect();
    return rect.width >= 8 && rect.height >= 8 && rect.width <= 160 && rect.height <= 160;
  });
}
"""

ONBOARDING_BLANK_SCRIPT = r"""
() => {
  const visible = (element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' && style.visibility !== 'hidden' &&
      Number(style.opacity) !== 0 && rect.width > 0 && rect.height > 0;
  };
  const text = (document.body?.innerText || '').trim();
  if (text.length > 20) return false;

  const actions = document.querySelectorAll(
    'button, input, select, textarea, a[href], [role="button"], iframe'
  );
  if (Array.from(actions).some(visible)) return false;

  return !Array.from(document.querySelectorAll('img, canvas, svg, video')).some(element => {
    if (!visible(element)) return false;
    const rect = element.getBoundingClientRect();
    const ratio = rect.width / rect.height;
    return rect.width >= 80 && rect.height >= 80 && ratio >= 0.75 && ratio <= 1.25;
  });
}
"""


def qr_challenge_visible(page: Page) -> bool:
    for frame in page.frames:
        try:
            if frame.evaluate(QR_DETECTION_SCRIPT):
                return True
        except Exception:
            continue
    return False


def loading_indicator_visible(page: Page) -> bool:
    for frame in page.frames:
        try:
            if frame.evaluate(LOADING_DETECTION_SCRIPT):
                return True
        except Exception:
            continue
    return False


def onboarding_page_is_blank(page: Page) -> bool:
    """Treat only a blank main onboarding document as a transient loading state."""
    if not is_onboarding_url(page.url):
        return False
    try:
        return bool(page.evaluate(ONBOARDING_BLANK_SCRIPT))
    except Exception:
        return False


def wait_for_qr_resolution(
    page: Page,
    *,
    allow_manual: bool,
    timeout: float = 300,
) -> bool:
    if not qr_challenge_visible(page):
        return False
    screenshot(page, "qr_verification_required")
    if not allow_manual:
        raise RuntimeError("QR verification requires a headed browser")

    print("  QR verification detected. Scan it in Chrome; the script will continue automatically...")
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError("Browser was closed during QR verification")
        if not qr_challenge_visible(page):
            print("  QR verification completed")
            return True
        page.wait_for_timeout(500)
    raise TimeoutError(f"QR verification timed out after {int(timeout)} seconds")


def is_onboarding_url(url: str) -> bool:
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    path = parsed.path.lower()
    if host == "login.databricks.com":
        return path.startswith(
            ("/login/accounts/", "/setup", "/select-account", "/signup")
        )
    return host == "accounts.cloud.databricks.com" and path.startswith("/setup")


def wait_for_navigation_stable(
    page: Page,
    timeout: float = 30,
    stable_for: float = 1.25,
) -> str:
    """等待 Databricks 自动重定向链结束，避免与下一次 goto 竞争。"""
    deadline = time.monotonic() + timeout
    last_url = page.url
    stable_since = time.monotonic()

    while time.monotonic() < deadline:
        try:
            page.wait_for_load_state("domcontentloaded", timeout=750)
        except Exception:
            pass
        current = page.url
        if current != last_url:
            last_url = current
            stable_since = time.monotonic()
        else:
            required = 3.0 if is_onboarding_url(current) else stable_for
            if time.monotonic() - stable_since >= required:
                return current
        page.wait_for_timeout(200)
    return page.url


def resilient_goto(
    page: Page,
    url: str,
    *,
    timeout: int = 45_000,
    attempts: int = 3,
):
    """遇到 Databricks 自己发起的导航时，等待其完成后再重试。"""
    last_error = None
    target = urlparse(url)
    for attempt in range(1, attempts + 1):
        wait_for_navigation_stable(page, timeout=20)
        try:
            return page.goto(url, wait_until="domcontentloaded", timeout=timeout)
        except PlaywrightError as exc:
            error_text = str(exc).lower()
            retryable = any(
                marker in error_text
                for marker in (
                    "interrupted by another navigation",
                    "navigation was interrupted",
                    "net::err_aborted",
                )
            )
            if not retryable:
                raise
            last_error = exc
            print(f"  Databricks redirect in progress; retrying navigation ({attempt}/{attempts})...")
            wait_for_navigation_stable(page, timeout=30)
            current = urlparse(page.url)
            if (
                target.hostname
                and current.hostname == target.hostname
                and (
                    target.path in ("", "/")
                    or current.path.rstrip("/").startswith(target.path.rstrip("/"))
                )
            ):
                print("  Databricks redirect reached the requested page")
                return None
    raise RuntimeError(f"Navigation kept being interrupted: {last_error}")


def otp_is_visible(page: Page) -> bool:
    if visible_locator(page, OTP_SELECTORS, timeout=100):
        return True
    try:
        cells = page.locator('input[aria-label*="character" i]')
        return cells.count() > 0 and cells.first.is_visible(timeout=100)
    except Exception:
        return False


def fill_otp(page: Page, code: str) -> bool:
    clean = re.sub(r"[\s-]+", "", code).strip()
    if not clean:
        return False

    field = visible_locator(page, OTP_SELECTORS, timeout=300)
    if field:
        field.fill(clean)
        return True

    try:
        cells = page.locator('input[aria-label*="character" i]')
        count = cells.count()
        if count == 1 and cells.first.is_visible(timeout=300):
            cells.first.fill(clean)
            return True
        if count > 1 and cells.first.is_visible(timeout=300):
            # 该控件会把从第一个格子输入的字符自动分发到后续格子。
            cells.first.click()
            page.keyboard.type(clean, delay=50)
            return True
    except Exception:
        pass
    return False


def workspace_domain(url: str) -> str:
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    if not host or host == "accounts.cloud.databricks.com":
        return ""
    if host.endswith((".cloud.databricks.com", ".azuredatabricks.net")):
        return f"{parsed.scheme or 'https'}://{host}"
    return ""


def authenticated_workspace(page: Page) -> str:
    current = workspace_domain(page.url)
    if not current:
        return ""
    if urlparse(page.url).path.lower().startswith("/login"):
        return ""
    return current


def workspace_links(page: Page) -> list[tuple[str, str]]:
    """同时处理绝对链接和 Databricks 返回的相对 auto-login 链接。"""
    try:
        hrefs = page.locator("a[href]").evaluate_all(
            "elements => elements.map(element => element.getAttribute('href') || '')"
        )
    except Exception:
        return []

    found: list[tuple[str, str]] = []
    seen: set[str] = set()
    for href in hrefs:
        absolute = urljoin(page.url, href)
        domain = workspace_domain(absolute)
        if domain and domain not in seen:
            found.append((domain, absolute))
            seen.add(domain)
    return found


def normalize_workspace(value: str) -> str:
    value = value.strip().rstrip("/")
    domain = workspace_domain(value)
    if not domain:
        raise ValueError(f"Invalid Databricks workspace URL: {value!r}")
    return domain


def open_known_workspace(
    page: Page,
    value: str,
    *,
    timeout: int = 45_000,
) -> str:
    """Open a known workspace and verify that authentication actually completed."""
    workspace = normalize_workspace(value)
    resilient_goto(page, workspace, timeout=timeout)
    wait_for_navigation_stable(page, timeout=30)

    current = authenticated_workspace(page)
    if current == workspace:
        return current
    return ""


def choose_workspace(
    page: Page,
    preferred_workspace: str | None = None,
    timeout: float = 60,
    allow_manual: bool = True,
    first_name: str = "",
    last_name: str = "",
) -> str:
    wait_for_navigation_stable(page, timeout=30)
    current = authenticated_workspace(page)
    if current:
        return current

    if preferred_workspace:
        workspace = normalize_workspace(preferred_workspace)
        current = open_known_workspace(page, workspace, timeout=45_000)
        if current:
            return current
        if is_onboarding_url(page.url):
            complete_onboarding(
                page,
                first_name,
                last_name,
                allow_manual=allow_manual,
                preferred_workspace=workspace,
            )
            current = open_known_workspace(page, workspace, timeout=45_000)
            if current:
                return current
        raise RuntimeError(f"Could not authenticate to known Workspace: {workspace}")

    print("  Looking for Workspace...")
    for _ in range(3):
        wait_for_navigation_stable(page, timeout=30)
        if is_onboarding_url(page.url):
            print("  Finishing Databricks account setup before Workspace discovery...")
            complete_onboarding(
                page,
                first_name,
                last_name,
                allow_manual=allow_manual,
            )
            wait_for_navigation_stable(page, timeout=30)

        current = authenticated_workspace(page)
        if current:
            return current

        resilient_goto(page, WORKSPACES_URL, timeout=45_000)
        wait_for_navigation_stable(page, timeout=30)
        if not is_onboarding_url(page.url):
            break

    candidates: list[tuple[str, str]] = []
    wait_until(page, lambda: bool(workspace_links(page)), timeout=timeout, interval=0.5)
    candidates = workspace_links(page)

    if candidates:
        if len(candidates) == 1:
            domain, target = candidates[0]
        else:
            domain, target = candidates[0]
            print(f"  Multiple workspaces found; selecting first: {domain}")
        print(f"  -> {domain}")
        resilient_goto(page, target, timeout=60_000)
        wait_until(page, lambda: bool(authenticated_workspace(page)), timeout=60)
        return authenticated_workspace(page) or domain

    if authenticated_workspace(page):
        return authenticated_workspace(page)

    screenshot(page, "workspace_not_found")
    if wait_for_qr_resolution(page, allow_manual=allow_manual):
        return choose_workspace(
            page,
            preferred_workspace=preferred_workspace,
            timeout=timeout,
            allow_manual=allow_manual,
            first_name=first_name,
            last_name=last_name,
        )
    raise RuntimeError("Workspace could not be discovered automatically")


def read_clipboard() -> str:
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
            capture_output=True,
            text=True,
            timeout=5,
            check=True,
        )
        return result.stdout.strip()
    except Exception:
        return ""


def read_browser_clipboard(page: Page) -> str:
    try:
        return (page.evaluate("navigator.clipboard.readText()") or "").strip()
    except Exception:
        return ""


def interaction_scopes(page: Page) -> list:
    scopes = [page]
    try:
        main_frame = page.main_frame
        scopes.extend(frame for frame in page.frames if frame is not main_frame)
    except Exception:
        pass
    return scopes


def first_visible(locator, timeout: int = 500):
    for index in range(locator.count()):
        candidate = locator.nth(index)
        if candidate.is_visible(timeout=timeout):
            return candidate
    return None


def token_from_page(page: Page) -> str:
    """读取生成结果控件；不扫描脚本或隐藏应用状态。"""
    for scope in interaction_scopes(page):
        try:
            values = scope.locator("input, textarea, code").evaluate_all(
                "elements => elements.slice(0, 50).map(element => "
                "('value' in element ? element.value : element.textContent) || '')"
            )
        except Exception:
            continue
        for value in values:
            match = TOKEN_PATTERN.search(str(value).strip())
            if match:
                return match.group(0)
        try:
            visible_text = scope.locator("body").inner_text(timeout=500)
        except Exception:
            continue
        match = TOKEN_PATTERN.search(visible_text or "")
        if match:
            return match.group(0)
    return ""


def wait_for_new_token(page: Page, previous_token: str, timeout: float = 20) -> str:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for candidate in (read_browser_clipboard(page), read_clipboard(), token_from_page(page)):
            match = TOKEN_PATTERN.fullmatch(candidate) or TOKEN_PATTERN.search(candidate)
            if match and match.group(0) != previous_token:
                return match.group(0)
        page.wait_for_timeout(250)
    return ""


def named_action(page: Page, names: tuple[str, ...], timeout: int = 500):
    for name in names:
        for scope in interaction_scopes(page):
            for role in ("button", "link"):
                try:
                    candidate = first_visible(
                        scope.get_by_role(role, name=name, exact=True),
                        timeout=timeout,
                    )
                    if candidate:
                        return candidate
                except Exception:
                    continue
            try:
                candidate = first_visible(
                    scope.get_by_text(name, exact=True),
                    timeout=timeout,
                )
                if candidate:
                    return candidate
            except Exception:
                continue
    return None


def wait_for_exact_text(
    page: Page,
    text: str,
    *,
    allow_manual: bool,
    timeout: float,
    progress_label: str | None = None,
    progress_interval: float = 15,
):
    started = time.monotonic()
    next_progress = max(1, progress_interval)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for scope in interaction_scopes(page):
            try:
                candidate = first_visible(
                    scope.get_by_text(text, exact=True),
                    timeout=300,
                )
                if candidate:
                    return candidate
            except Exception:
                continue
        if wait_for_qr_resolution(page, allow_manual=allow_manual):
            wait_for_navigation_stable(page, timeout=30)
            continue
        elapsed = time.monotonic() - started
        if progress_label and elapsed >= next_progress:
            print(
                f"  Still waiting for {progress_label} "
                f"({int(elapsed)}s/{int(timeout)}s)..."
            )
            next_progress += max(1, progress_interval)
        page.wait_for_timeout(1_000)
    return None


def wait_for_named_action(
    page: Page,
    names: tuple[str, ...],
    *,
    allow_manual: bool,
    timeout: float,
    progress_label: str | None = None,
    progress_interval: float = 15,
):
    started = time.monotonic()
    next_progress = max(1, progress_interval)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        action = named_action(page, names, timeout=300)
        if action:
            return action
        if wait_for_qr_resolution(page, allow_manual=allow_manual):
            wait_for_navigation_stable(page, timeout=30)
            continue
        elapsed = time.monotonic() - started
        if progress_label and elapsed >= next_progress:
            print(
                f"  Still waiting for {progress_label} "
                f"({int(elapsed)}s/{int(timeout)}s)..."
            )
            next_progress += max(1, progress_interval)
        page.wait_for_timeout(1_000)
    return None


def open_ai_gateway(page: Page, *, allow_manual: bool) -> None:
    if "/ml/ai-gateway" in urlparse(page.url).path:
        return
    gateway = wait_for_named_action(
        page,
        (AI_GATEWAY_LABEL,),
        allow_manual=allow_manual,
        timeout=180,
        progress_label="AI Gateway navigation",
    )
    if not gateway:
        screenshot(page, "ai_gateway_missing")
        raise RuntimeError("AI Gateway navigation did not become available")

    try:
        href = gateway.get_attribute("href") or ""
    except Exception:
        href = ""
    if href:
        resilient_goto(page, urljoin(page.url, href), timeout=60_000)
    else:
        gateway.click()
    wait_for_navigation_stable(page, timeout=30)


def glm_model_page_url(url: str) -> bool:
    path = urlparse(url).path.lower().rstrip("/")
    return "glm-5-2" in path and (
        "/ml/ai-gateway/" in path or "/explore/model-services/" in path
    )


def click_model_entry(model) -> None:
    """优先点击模型卡片的可交互祖先，找不到时回退到文字节点。"""
    target = model
    try:
        ancestors = model.locator(
            "xpath=ancestor-or-self::*[self::a or self::button or "
            "@role='button' or @role='link' or @tabindex='0'][1]"
        )
        if ancestors.count() > 0:
            candidate = ancestors.first
            if candidate.is_visible(timeout=500):
                target = candidate
    except Exception:
        target = model

    try:
        target.scroll_into_view_if_needed(timeout=5_000)
    except Exception:
        pass
    target.click(timeout=15_000)


def generate_glm_access_token(
    page: Page,
    workspace: str,
    *,
    allow_manual: bool,
) -> str:
    print("  Opening AI Gateway and waiting for GLM 5.2")
    workspace = normalize_workspace(workspace)
    if authenticated_workspace(page) != workspace:
        current = open_known_workspace(page, workspace, timeout=60_000)
        if not current:
            raise RuntimeError(f"Could not open authenticated Workspace: {workspace}")

    for attempt in range(1, 4):
        open_ai_gateway(page, allow_manual=allow_manual)
        model = wait_for_exact_text(
            page,
            GLM_MODEL_LABEL,
            allow_manual=allow_manual,
            timeout=300,
            progress_label="GLM 5.2 in AI Gateway",
        )
        if not model:
            screenshot(page, "glm_model_missing")
            raise RuntimeError(
                f"GLM 5.2 did not load in AI Gateway; current URL: {page.url}"
            )

        print("  Opening GLM 5.2")
        click_model_entry(model)
        wait_for_navigation_stable(page, timeout=30)
        if not wait_until(page, lambda: glm_model_page_url(page.url), timeout=45, interval=0.5):
            screenshot(page, "glm_model_navigation_failed")
            raise RuntimeError(
                f"Clicking GLM 5.2 did not open its model page; current URL: {page.url}"
            )
        print(f"  GLM 5.2 page ready: {urlparse(page.url).path}")
        generate = wait_for_named_action(
            page,
            (GLM_TOKEN_ACTION,),
            allow_manual=allow_manual,
            timeout=180,
            progress_label="Generate Access Token",
        )
        if generate:
            previous = token_from_page(page)
            print("  Generating GLM access token")
            generate.click()
            token = wait_for_new_token(page, previous_token=previous, timeout=60)
            if token:
                return token
            screenshot(page, "glm_token_value_missing")
            raise RuntimeError("GLM access token was generated but its value was not exposed")

        body = ""
        try:
            body = page.locator("body").inner_text(timeout=1_000)
        except Exception:
            pass
        if "resource not found" not in body.lower() or attempt == 3:
            screenshot(page, "glm_generate_action_missing")
            raise RuntimeError(
                "GLM 5.2 page did not expose Generate Access Token; "
                f"current URL: {page.url}"
            )
        print(f"  GLM 5.2 is still provisioning; retrying from AI Gateway ({attempt}/3)...")
        resilient_goto(page, workspace, timeout=60_000)
        page.wait_for_timeout(10_000)

    raise RuntimeError("GLM 5.2 token generation failed")


def generate_token(
    page: Page,
    workspace: str,
    allow_manual: bool = True,
) -> str:
    return generate_glm_access_token(
        page,
        workspace,
        allow_manual=allow_manual,
    )


def select_first_available_option(page: Page) -> bool:
    selects = page.locator("select")
    count = selects.count()
    for index in range(count):
        select = selects.nth(index)
        try:
            if not select.is_visible(timeout=200):
                continue
            state = select.evaluate(
                "element => ({current: element.value, option: Array.from(element.options)"
                ".map(option => ({value: option.value, label: option.textContent.trim(), "
                "disabled: option.disabled}))"
                ".find(option => option.value && !option.disabled) || null})"
            )
            option = state["option"]
            if not state["current"] and option:
                select.select_option(value=option["value"])
                print(f"  Selected: {option['label']}")
                return True
        except Exception:
            continue
    return False


def fill_optional_profile(page: Page, first_name: str, last_name: str) -> bool:
    changed = False
    account_name = f"{first_name} {last_name}".strip() or "Personal"
    fields = (
        ("input[name='firstName']", first_name),
        ("input[name='lastName']", last_name),
        ("input[name='company']", "Personal"),
        ("input[name='accountName']", account_name),
        ("input[name='accountDisplayName']", account_name),
    )
    for selector, value in fields:
        if not value:
            continue
        try:
            field = page.locator(selector)
            if field.count() == 1 and field.is_visible(timeout=200):
                field.fill(value)
                changed = True
        except Exception:
            continue

    try:
        checkboxes = page.locator('input[type="checkbox"]')
        if checkboxes.count() == 1 and checkboxes.first.is_visible(timeout=200):
            if not checkboxes.first.is_checked():
                checkboxes.first.check()
                changed = True
    except Exception:
        pass
    return changed


def complete_onboarding(
    page: Page,
    first_name: str,
    last_name: str,
    allow_manual: bool = True,
    preferred_workspace: str | None = None,
) -> None:
    """处理当前已知的试用/资料/区域步骤；遇到未知页面时交给用户一次。"""
    accept_cookies(page)
    deadline = time.monotonic() + 180
    idle_rounds = 0
    workspace_attempts = 0

    while time.monotonic() < deadline:
        if authenticated_workspace(page) or workspace_links(page):
            return

        progressed = False
        progressed |= accept_cookies(page)
        progressed |= click_existing_work_account(page, first_name)
        progressed |= fill_optional_profile(page, first_name, last_name)

        if select_first_available_option(page):
            progressed = True

        if click_named_button(
            page,
            (
                "Start trial with express setup",
                "Start trial with express",
                "Start free trial",
                "Continue",
                "Submit",
                "Get started",
                "Next",
                "Create account",
                "Create workspace",
                "Finish setup",
            ),
            timeout=300,
        ):
            progressed = True

        if progressed:
            idle_rounds = 0
            page.wait_for_timeout(750)
        else:
            idle_rounds += 1
            page.wait_for_timeout(500)
            if idle_rounds >= 12:
                wait_for_navigation_stable(page, timeout=15)
                if authenticated_workspace(page) or workspace_links(page):
                    return
                if wait_for_qr_resolution(page, allow_manual=allow_manual):
                    idle_rounds = 0
                    deadline = max(deadline, time.monotonic() + 90)
                    wait_for_navigation_stable(page, timeout=30)
                    continue
                if loading_indicator_visible(page) or onboarding_page_is_blank(page):
                    if preferred_workspace and workspace_attempts < 3:
                        workspace_attempts += 1
                        print(
                            "  Databricks setup is still loading; "
                            f"opening known Workspace ({workspace_attempts}/3)..."
                        )
                        try:
                            if open_known_workspace(page, preferred_workspace):
                                return
                        except Exception as exc:
                            print(f"  Known Workspace is not ready yet: {exc}")
                    else:
                        print("  Databricks setup is still loading; waiting...")
                    idle_rounds = 0
                    page.wait_for_timeout(2_000)
                    continue
                if preferred_workspace and workspace_attempts < 3:
                    workspace_attempts += 1
                    print(f"  Trying known Workspace ({workspace_attempts}/3)...")
                    try:
                        if open_known_workspace(page, preferred_workspace):
                            return
                    except Exception as exc:
                        print(f"  Known Workspace is not ready yet: {exc}")
                    idle_rounds = 0
                    continue
                screenshot(page, "onboarding_unknown")
                raise RuntimeError("Unsupported Databricks onboarding page (no QR challenge detected)")

    raise RuntimeError("Databricks onboarding timed out")


def launch_browser(
    playwright: Playwright,
    headless: bool,
    channel: str | None,
) -> tuple[Browser, BrowserContext, Page]:
    args = ["--disable-blink-features=AutomationControlled"]
    launch_options = {"headless": headless, "args": args}

    if channel:
        try:
            browser = playwright.chromium.launch(channel=channel, **launch_options)
        except Exception as exc:
            print(f"  Browser channel {channel!r} unavailable, using Playwright Chromium: {exc}")
            browser = playwright.chromium.launch(**launch_options)
    else:
        browser = playwright.chromium.launch(**launch_options)

    context = browser.new_context(
        viewport={"width": 1400, "height": 900},
        permissions=["clipboard-read", "clipboard-write"],
    )
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return browser, context, context.new_page()


def submit_email(page: Page, email: str, allow_manual: bool = True) -> None:
    page.goto(SIGNUP_URL, wait_until="domcontentloaded", timeout=45_000)
    accept_cookies(page)

    field = page.locator('input[name="email"]')
    field.wait_for(state="visible", timeout=30_000)
    field.fill(email)

    if not click_named_button(
        page,
        ("Continue with email", "使用邮箱继续", "Continue"),
        timeout=2_000,
    ):
        raise RuntimeError("Continue with email button not found")

    if not wait_until(page, lambda: otp_is_visible(page), timeout=30):
        screenshot(page, "email_submission_blocked")
        if wait_for_qr_resolution(page, allow_manual=allow_manual):
            wait_until(page, lambda: otp_is_visible(page), timeout=30)
    if not otp_is_visible(page):
        raise RuntimeError("Verification-code input did not appear")


def verify_email(
    page: Page,
    mail: VerificationMailClient,
    not_before: datetime,
    timeout: int,
    allow_manual: bool = True,
) -> None:
    code = mail.wait_for_databricks_code(not_before=not_before, timeout=timeout)
    print(f"  Verification code received: {code}")
    if not fill_otp(page, code):
        screenshot(page, "otp_input_unknown")
        if wait_for_qr_resolution(page, allow_manual=allow_manual):
            if not fill_otp(page, code):
                raise RuntimeError("Verification-code input is still unsupported after QR verification")
        else:
            raise RuntimeError("Unknown verification-code input (no QR challenge detected)")
    else:
        click_named_button(page, ("Verify", "Continue", "Sign in"), timeout=1_000)

    if not wait_until(page, lambda: not otp_is_visible(page), timeout=45):
        screenshot(page, "otp_verification_stuck")
        if wait_for_qr_resolution(page, allow_manual=allow_manual):
            if not wait_until(page, lambda: not otp_is_visible(page), timeout=45):
                raise RuntimeError("Verification remained pending after QR verification")
        else:
            raise RuntimeError("Verification remained pending (no QR challenge detected)")


def auto_register(
    email: str | None = None,
    mail_password: str | None = None,
    *,
    headless: bool = False,
    browser_channel: str | None = "chrome",
    mail_provider: str = "outlook",
    outlook_credential_file: str | None = None,
    outlook_success_file: str | None = None,
    resume: bool = False,
    mail_timeout: int = 180,
    preferred_workspace: str | None = None,
) -> tuple[str, str, str, str] | None:
    print(f"\nCreating mail session with provider: {mail_provider}")
    mail = create_mail_client(
        mail_provider,
        email=email,
        password=mail_password,
        outlook_credential_file=outlook_credential_file,
    )
    email = mail.address
    mail_password = mail.password
    print(f"  Email: {email}")

    if resume:
        # 已注册账号通常已有资料；仅在 setup 表单确实为空时使用邮箱本地部分补位。
        first_name = email.split("@", 1)[0]
        last_name = ""
    else:
        first_name = random.choice(FIRST_NAMES)
        last_name = random.choice(LAST_NAMES)
    flow = "resume" if resume else "register"
    playwright = None
    browser = None
    context = None
    page = None
    workspace = ""

    try:
        if preferred_workspace:
            workspace = normalize_workspace(preferred_workspace)
        playwright = sync_playwright().start()
        browser, context, page = launch_browser(
            playwright,
            headless=headless,
            channel=browser_channel,
        )
        otp_time = datetime.now(timezone.utc)
        allow_manual = not headless

        print(f"\n[1/5] Open Databricks email entry ({flow})")
        submit_email(page, email, allow_manual=allow_manual)

        print("[2/5] Wait for verification code")
        verify_email(
            page,
            mail,
            not_before=otp_time,
            timeout=mail_timeout,
            allow_manual=allow_manual,
        )

        print("[3/5] Complete registration/onboarding")
        complete_onboarding(
            page,
            first_name,
            last_name,
            allow_manual=allow_manual,
            preferred_workspace=preferred_workspace if resume else None,
        )

        print("[4/5] Find Workspace")
        workspace = choose_workspace(
            page,
            preferred_workspace=preferred_workspace,
            allow_manual=allow_manual,
            first_name=first_name,
            last_name=last_name,
        )

        print("[5/5] Generate GLM Token")
        token = generate_token(
            page,
            workspace,
            allow_manual=allow_manual,
        )
        if not TOKEN_PATTERN.fullmatch(token or ""):
            raise RuntimeError("Generated value is not a valid dapi Token")

        save_record(
            email,
            mail_password,
            workspace,
            token,
            "OK",
            f"mail_provider={mail_provider}; flow={flow}",
        )
        save_domain_token(workspace, token)
        if mail_provider == "outlook" and outlook_credential_file:
            source = Path(outlook_credential_file).expanduser().resolve()
            destination = Path(
                outlook_success_file or OUTLOOK_SUCCESS_FILE
            ).expanduser().resolve()
            if source != destination:
                if move_outlook_credential(source, destination, email):
                    print(f"  Moved Outlook credential to {destination}")
                else:
                    print(
                        "  Warning: successful Outlook credential was not found "
                        f"in {source}"
                    )
        print(f"\n  Email: {email}")
        print(f"  Workspace: {workspace}")
        print(f"  Token: {token[:8]}...{token[-5:]}")
        return email, mail_password, workspace, token

    except KeyboardInterrupt:
        screenshot(page, "cancelled")
        save_record(
            email,
            mail_password or "",
            workspace,
            "",
            "CANCELLED",
            f"mail_provider={mail_provider}; flow={flow}; User interrupted",
        )
        print("\nCancelled by user.")
        return None
    except Exception as exc:
        print(f"\nFAILED: {exc}")
        screenshot(page, "fail")
        save_record(
            email,
            mail_password or "",
            workspace,
            "",
            "FAIL",
            f"mail_provider={mail_provider}; flow={flow}; {exc}",
        )
        traceback.print_exc()
        return None
    finally:
        if context is not None:
            try:
                context.close()
            except Exception:
                pass
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass
        if playwright is not None:
            try:
                playwright.stop()
            except Exception:
                pass
        try:
            mail.close()
        except Exception:
            pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Thief Cat 2.0 - Databricks GLM 5.2 Access Token automation"
    )
    parser.add_argument("--email", "-e", help="mail address or Outlook record selector")
    parser.add_argument("--password", "-p", help="Outlook password field")
    parser.add_argument(
        "--mail-provider",
        choices=("outlook", "outlook_tw"),
        default="outlook",
        help="mail provider",
    )
    parser.add_argument(
        "--outlook-credential-file",
        help="UTF-8 file containing email----password----refresh_token----client_id",
    )
    parser.add_argument(
        "--outlook-success-file",
        default=str(OUTLOOK_SUCCESS_FILE),
        help="move successful Outlook records to this file",
    )
    parser.add_argument(
        "--mail-test",
        action="store_true",
        help="test the selected mail provider without opening Databricks",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="log in to an existing account and continue setup/Workspace/Token",
    )
    parser.add_argument("--workspace", help="known Databricks workspace URL")
    parser.add_argument("--mail-timeout", type=int, default=180, help="OTP timeout in seconds")
    parser.add_argument("--headless", action="store_true", help="run browser headlessly")
    parser.add_argument(
        "--browser-channel",
        default="chrome",
        help="Playwright browser channel; use 'bundled' for bundled Chromium",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if (
        args.mail_provider == "outlook"
        and
        not args.outlook_credential_file
        and OUTLOOK_PENDING_FILE.exists()
    ):
        args.outlook_credential_file = str(OUTLOOK_PENDING_FILE)
    if args.mail_provider == "outlook_tw":
        if args.outlook_credential_file:
            raise SystemExit("--outlook-credential-file requires --mail-provider outlook")
        if args.password:
            raise SystemExit("--password is not supported by --mail-provider outlook_tw")
        if args.resume and not args.email:
            raise SystemExit("--resume with outlook_tw requires the still-active --email")
        if args.outlook_success_file != str(OUTLOOK_SUCCESS_FILE):
            raise SystemExit("--outlook-success-file requires --mail-provider outlook")
    if args.headless and not args.workspace:
        print("Warning: QR verification requires a visible browser and will fail in headless mode.")

    channel = None if args.browser_channel.lower() == "bundled" else args.browser_channel
    if args.mail_test:
        client = create_mail_client(
            args.mail_provider,
            email=args.email,
            password=args.password,
            outlook_credential_file=args.outlook_credential_file,
        )
        try:
            visible_messages = client.test_connection()
            print(
                f"Mail OK ({args.mail_provider}): {client.address} "
                f"(visible messages: {visible_messages})"
            )
            return 0
        finally:
            client.close()

    result = auto_register(
        email=args.email,
        mail_password=args.password,
        headless=args.headless,
        browser_channel=channel,
        mail_provider=args.mail_provider,
        outlook_credential_file=args.outlook_credential_file,
        outlook_success_file=args.outlook_success_file,
        resume=args.resume,
        mail_timeout=args.mail_timeout,
        preferred_workspace=args.workspace,
    )
    return 0 if result else 1


if __name__ == "__main__":
    sys.exit(main())
