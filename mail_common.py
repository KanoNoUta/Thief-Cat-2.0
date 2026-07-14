"""邮箱 provider 共用的验证码解析函数。"""

import re


def extract_verification_code(text: str) -> str:
    """从邮件标题或正文提取 Databricks 常见验证码。"""
    contextual_patterns = (
        r"(?:verification|security|login|sign[- ]in)\s*(?:code)?\D{0,20}([A-Z0-9]{3}-[A-Z0-9]{3})",
        r"(?:verification|security|login|sign-in)\s*(?:code)?\D{0,30}(\d{6})",
        r"(?:验证码|安全码|登录码)\D{0,20}(\d{6})",
        r"\bcode\D{0,20}(\d{6})\b",
    )
    for pattern in contextual_patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1)

    segmented = re.search(
        r"(?<![A-Z0-9])([A-Z0-9]{3}-[A-Z0-9]{3})(?![A-Z0-9])",
        text,
        re.IGNORECASE,
    )
    if segmented:
        return segmented.group(1)

    numeric = re.search(r"(?<!\d)(\d{6})(?!\d)", text)
    return numeric.group(1) if numeric else ""
