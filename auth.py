"""
Authentication module — Microsoft OAuth login (Entra ID / Azure AD)
- Chỉ cho phép user thuộc tenant công ty (single tenant)
- Chỉ cho phép email đuôi @vng.com.vn (domain check)
- Chỉ cho phép email trong whitelist
"""

import os
import secrets
import requests
from functools import wraps
from urllib.parse import urlencode
from pathlib import Path

from flask import session, redirect, url_for, request, jsonify

# ─── Config helpers ──────────────────────────────────────────────────────────

def _tenant_id() -> str:
    return os.getenv("MS_TENANT_ID", "").strip()


def _ms_auth_url() -> str:
    return f"https://login.microsoftonline.com/{_tenant_id()}/oauth2/v2.0/authorize"


def _ms_token_url() -> str:
    return f"https://login.microsoftonline.com/{_tenant_id()}/oauth2/v2.0/token"


MS_GRAPH_ME_URL = "https://graph.microsoft.com/v1.0/me"


def _allowed_domain() -> str:
    return os.getenv("ALLOWED_EMAIL_DOMAIN", "vng.com.vn").lower().strip()


def _client_creds() -> tuple[str, str]:
    """Microsoft App Registration credentials."""
    cid = os.getenv("MS_CLIENT_ID", "").strip()
    sec = os.getenv("MS_CLIENT_SECRET", "").strip()
    return cid, sec


def _load_whitelist() -> set[str]:
    """
    Đọc whitelist mỗi lần gọi (không cache) để admin sửa file/env xong
    không cần restart app.
    """
    emails: set[str] = set()

    wl_file = os.getenv("WHITELIST_FILE", "whitelist.txt")
    p = Path(wl_file)
    if p.is_file():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip().lower()
            if line and not line.startswith("#"):
                emails.add(line)

    raw = os.getenv("WHITELIST_EMAILS", "")
    for e in raw.split(","):
        e = e.strip().lower()
        if e:
            emails.add(e)

    return emails


# ─── OAuth flow ──────────────────────────────────────────────────────────────

def build_login_url(redirect_uri: str) -> str:
    """Tạo URL redirect người dùng sang Microsoft consent screen."""
    client_id, _ = _client_creds()
    if not client_id:
        raise RuntimeError("MS_CLIENT_ID chưa được cấu hình trong .env")
    if not _tenant_id():
        raise RuntimeError("MS_TENANT_ID chưa được cấu hình trong .env")

    state = secrets.token_urlsafe(24)
    session["oauth_state"] = state

    params = {
        "client_id":     client_id,
        "response_type": "code",
        "redirect_uri":  redirect_uri,
        "response_mode": "query",
        "scope":         "openid email profile User.Read",
        "state":         state,
        "prompt":        "select_account",
    }
    return f"{_ms_auth_url()}?{urlencode(params)}"


def exchange_login_code(code: str, redirect_uri: str) -> dict:
    """Đổi authorization code -> user info qua Microsoft Graph."""
    client_id, client_secret = _client_creds()

    r = requests.post(_ms_token_url(), data={
        "client_id":     client_id,
        "client_secret": client_secret,
        "code":          code,
        "redirect_uri":  redirect_uri,
        "grant_type":    "authorization_code",
        "scope":         "openid email profile User.Read",
    }, timeout=30)
    r.raise_for_status()
    access_token = r.json().get("access_token")
    if not access_token:
        raise RuntimeError("Không nhận được access_token từ Microsoft")

    info = requests.get(
        MS_GRAPH_ME_URL,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )
    info.raise_for_status()
    data = info.json()

    # Graph có nhiều field email khả dĩ; ưu tiên theo thứ tự:
    email = (
        data.get("mail")
        or data.get("userPrincipalName")
        or ""
    ).lower().strip()

    return {
        "email": email,
        "name":  data.get("displayName", ""),
        "id":    data.get("id", ""),
    }


# ─── Authorization ───────────────────────────────────────────────────────────

def is_email_allowed(email: str) -> tuple[bool, str]:
    """Kiểm tra email có được phép truy cập không."""
    if not email:
        return False, "Không lấy được email từ Microsoft."

    email = email.lower().strip()
    domain = _allowed_domain()

    if not email.endswith(f"@{domain}"):
        return False, f"Ứng dụng chỉ chấp nhận email @{domain}."

    whitelist = _load_whitelist()
    if not whitelist:
        return True, ""

    if email not in whitelist:
        return False, "Email của bạn chưa được cấp quyền truy cập. Vui lòng liên hệ admin."

    return True, ""


def login_required(f):
    """Decorator: chặn route nếu user chưa đăng nhập."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("user_email"):
            if request.path.startswith("/api/"):
                return jsonify({"error": "Unauthorized - vui lòng đăng nhập"}), 401
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return wrapper
