"""
VNG Cloud — Incident Notification Web App
Flask backend: Upload Excel → Preview → Fill form → Send Email
+ Google OAuth login (domain + whitelist)
"""

import os
import json
import uuid
import secrets as _secrets
from datetime import datetime, timedelta
from pathlib import Path

from flask import (
    Flask, request, jsonify, render_template,
    session, redirect, url_for,
)
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

# Import core logic từ incident_agent.py
from incident_agent import (
    extract_columns,
    export_filtered_excel,
    build_email_html,
    get_graph_token,
    send_email,
    build_oauth_url,
    exchange_code_for_token,
    MAIL_MODE,
    GMAIL_CLIENT_ID,
)

# Import auth module
from auth import (
    build_login_url,
    exchange_login_code,
    is_email_allowed,
    login_required,
)

# Import translator (gọi 9router để dịch VN→EN)
from translator import translate_fields, is_configured as translator_configured

load_dotenv()

# ─── Config ──────────────────────────────────────────────────────────────────

UPLOAD_FOLDER  = Path("uploads")
HISTORY_FILE   = Path("history/log.json")
ALLOWED_EXT    = {".xlsx", ".xls"}

UPLOAD_FOLDER.mkdir(exist_ok=True)
HISTORY_FILE.parent.mkdir(exist_ok=True)
if not HISTORY_FILE.exists():
    HISTORY_FILE.write_text("[]")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB

# Session / cookie security
app.secret_key = os.getenv("FLASK_SECRET_KEY") or _secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_SECURE   = os.getenv("FLASK_ENV", "").lower() == "production",
    SESSION_COOKIE_HTTPONLY = True,
    SESSION_COOKIE_SAMESITE = "Lax",
    PERMANENT_SESSION_LIFETIME = timedelta(hours=8),
)


# ─── Helpers ─────────────────────────────────────────────────────────────────

def load_history():
    try:
        return json.loads(HISTORY_FILE.read_text())
    except Exception:
        return []


def save_history(entry: dict):
    history = load_history()
    history.insert(0, entry)
    history = history[:100]
    HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2))


def _login_redirect_uri() -> str:
    return os.getenv(
        "MS_REDIRECT_URI",
        request.host_url.rstrip("/") + "/auth/login/callback",
    )


# ═══════════════════════════════════════════════════════════════════════════
# AUTH ROUTES
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/login")
def login():
    """Trang login."""
    if session.get("user_email"):
        return redirect(url_for("index"))
    return render_template(
        "login.html",
        error=request.args.get("error"),
        domain=os.getenv("ALLOWED_EMAIL_DOMAIN", "vng.com.vn"),
    )


@app.route("/auth/login")
def start_login():
    """Bắt đầu Google OAuth flow."""
    try:
        url = build_login_url(_login_redirect_uri())
    except Exception as e:
        return redirect(url_for("login", error=str(e)))
    return redirect(url)


@app.route("/auth/login/callback")
def login_callback():
    """Google redirect về đây với ?code=..."""
    err   = request.args.get("error")
    code  = request.args.get("code")
    state = request.args.get("state")

    if err or not code:
        return redirect(url_for("login", error=err or "Không nhận được code"))

    expected_state = session.pop("oauth_state", None)
    if not expected_state or state != expected_state:
        return redirect(url_for("login", error="State không hợp lệ (CSRF?)"))

    try:
        userinfo = exchange_login_code(code, _login_redirect_uri())
    except Exception as e:
        return redirect(url_for("login", error=f"Lỗi xác thực Google: {e}"))

    email = (userinfo.get("email") or "").lower()
    verified = userinfo.get("verified_email", True)
    if not verified:
        return redirect(url_for("login", error="Email chưa được Google verify."))

    allowed, reason = is_email_allowed(email)
    if not allowed:
        return redirect(url_for("login", error=reason))

    session.permanent = True
    session["user_email"]   = email
    session["user_name"]    = userinfo.get("name", "")
    session["user_picture"] = userinfo.get("picture", "")

    next_url = request.args.get("next") or url_for("index")
    return redirect(next_url)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/api/me")
@login_required
def current_user():
    return jsonify({
        "email":   session.get("user_email"),
        "name":    session.get("user_name"),
        "picture": session.get("user_picture"),
    })


# ═══════════════════════════════════════════════════════════════════════════
# APP ROUTES (đều yêu cầu đăng nhập)
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/teams-config")
@login_required
def teams_config():
    return render_template("teams_config.html")


@app.route("/api/upload", methods=["POST"])
@login_required
def upload_excel():
    if "file" not in request.files:
        return jsonify({"error": "Không tìm thấy file trong request"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Chưa chọn file"}), 400

    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": "Chỉ chấp nhận file .xlsx hoặc .xls"}), 400

    uid      = uuid.uuid4().hex[:8]
    filename = f"{uid}_{secure_filename(file.filename)}"
    filepath = UPLOAD_FOLDER / filename
    file.save(filepath)

    try:
        df = extract_columns(str(filepath))
    except Exception as e:
        filepath.unlink(missing_ok=True)
        return jsonify({"error": f"Lỗi đọc file Excel: {e}"}), 422

    if df.empty:
        filepath.unlink(missing_ok=True)
        return jsonify({"error": "File không có dữ liệu hợp lệ (thiếu cột Salesman Email?)"}), 422

    grouped_preview = []
    for owned_by, group in df.groupby("Owned By"):
        cc = sorted(group["Salesman Email"].dropna().unique().tolist())
        devices = group[["Name", "IP Address", "Floating IP"]].fillna("—").to_dict("records")
        grouped_preview.append({
            "owned_by": owned_by,
            "cc":       cc,
            "devices":  devices,
            "count":    len(devices),
        })

    return jsonify({
        "file_id":      uid,
        "filename":     file.filename,
        "filepath":     str(filepath),
        "total_rows":   len(df),
        "total_groups": len(grouped_preview),
        "columns":      list(df.columns),
        "preview":      grouped_preview,
    })


import re
_DATE_PATTERN = re.compile(r"^\d{2}-\d{2}-\d{4} \d{2}:\d{2}$")


def _validate_date_fields(data: dict, fields: list[str]) -> tuple[bool, str]:
    """Kiểm tra các trường trong `data` có đúng format DD-MM-YYYY HH:MM.

    Chỉ validate những trường có giá trị (skip nếu rỗng — việc required đã check ở chỗ khác).
    Returns (ok, error_message).
    """
    for f in fields:
        val = (data.get(f) or "").strip()
        if not val:
            continue
        if not _DATE_PATTERN.match(val):
            return False, f"Truong '{f}' sai dinh dang. Phai la DD-MM-YYYY HH:MM (vd: 22-04-2026 14:30), nhan duoc: '{val}'"
        # Parse-check thêm: đảm bảo ngày thực sự hợp lệ (vd không có 31-02)
        try:
            datetime.strptime(val, "%d-%m-%Y %H:%M")
        except ValueError as e:
            return False, f"Truong '{f}' khong phai ngay hop le: {val} ({e})"
    return True, ""


@app.route("/api/send", methods=["POST"])
@login_required
def send_incident():
    data = request.json or {}
    category = data.get("category", "incident")
    tt = data.get("template_type", "")

    filepath = data.get("filepath", "")
    if not Path(filepath).is_file():
        return jsonify({"error": f"File khong ton tai: {filepath}"}), 400

    try:
        df = extract_columns(filepath)
    except Exception as e:
        return jsonify({"error": f"Loi doc Excel: {e}"}), 422

    if df.empty:
        return jsonify({"error": "Khong co du lieu hop le"}), 422

    try:
        filtered_path = export_filtered_excel(df, filepath)
    except Exception as e:
        return jsonify({"error": f"Loi xuat Excel: {e}"}), 500

    try:
        token = get_graph_token()
    except Exception as e:
        return jsonify({"error": f"Loi lay Gmail token: {e}"}), 500

    if category == "change":
        required = ["service_name", "change_desc", "change_type", "planned_start", "planned_end", "impact"]
        missing = [f for f in required if not data.get(f)]
        if missing:
            return jsonify({"error": f"Thieu truong: {', '.join(missing)}"}), 400
        if tt not in ("5", "6", "7"):
            return jsonify({"error": "template_type change phai la 5, 6, hoac 7"}), 400
        date_fields = ["planned_start", "planned_end", "actual_start", "actual_end"]
        ok, err = _validate_date_fields(data, date_fields)
        if not ok:
            return jsonify({"error": err}), 400
        info = {k: data.get(k, "") for k in [
            "template_type", "service_name", "change_desc", "change_type",
            "planned_start", "planned_end", "impact", "actual_start", "actual_end"
        ]}
        log_desc = data.get("change_desc", "")
        log_time = data.get("planned_start", "")
        log_end  = data.get("actual_end", "")
    else:
        required = ["template_type", "service_name", "incident_desc", "start_time", "root_cause"]
        missing = [f for f in required if not data.get(f)]
        if missing:
            return jsonify({"error": f"Thieu truong: {', '.join(missing)}"}), 400
        if tt not in ("1", "2", "3", "4"):
            return jsonify({"error": "template_type phai la 1-4"}), 400
        if tt in ("1", "4") and not data.get("end_time"):
            return jsonify({"error": "end_time bat buoc voi template 1 va 4"}), 400
        if tt == "3" and not data.get("solution"):
            return jsonify({"error": "solution bat buoc voi template 3"}), 400
        ok, err = _validate_date_fields(data, ["start_time", "end_time"])
        if not ok:
            return jsonify({"error": err}), 400
        info = {k: data.get(k, "") for k in [
            "template_type", "service_name", "incident_desc",
            "start_time", "end_time", "root_cause", "status", "solution"
        ]}
        log_desc = data.get("incident_desc", "")
        log_time = data.get("start_time", "")
        log_end  = data.get("end_time", "")

    # ── Dịch VN → EN cho các field text (1 API call cho cả batch) ──────────
    # Các field cần dịch theo category. Field rỗng sẽ được skip trong translator.
    if category == "change":
        translate_keys = ["change_desc", "change_type", "impact"]
    else:
        translate_keys = ["incident_desc", "root_cause", "status", "solution"]

    to_translate = {k: info.get(k, "") for k in translate_keys if info.get(k, "").strip()}
    if to_translate and translator_configured():
        try:
            translated = translate_fields(to_translate)
            # Merge vào info dưới key mới "<field>_en" — không đụng field gốc
            for k, v in translated.items():
                info[f"{k}_en"] = v
        except Exception as e:
            # Translator đã có fallback in-built, nhưng defensive coding
            app.logger.warning("Translation step failed: %s", e)
    # Đảm bảo mọi key _en đều tồn tại (fallback = text VN gốc)
    for k in translate_keys:
        info.setdefault(f"{k}_en", info.get(k, ""))

    results = []
    for owned_by_email, group in df.groupby("Owned By"):
        devices = group.to_dict("records")
        cc_list = sorted(group["Salesman Email"].dropna().unique().tolist())
        if "support@vngcloud.vn" not in [e.lower() for e in cc_list]:
            cc_list.append("support@vngcloud.vn")
        try:
            subject, html_body = build_email_html(info, devices)
            ok = send_email(token, owned_by_email, subject, html_body, cc_emails=cc_list)
            results.append({"to": owned_by_email, "cc": cc_list, "count": len(devices), "sent": ok})
        except Exception as e:
            results.append({"to": owned_by_email, "cc": cc_list, "count": len(devices), "sent": False, "error": str(e)})

    sent_count = sum(1 for r in results if r["sent"])
    total      = len(results)

    log_entry = {
        "id":         uuid.uuid4().hex[:8],
        "timestamp":  datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        "category":   category,
        "template":   tt,
        "service":    data.get("service_name", ""),
        "incident":   log_desc,
        "start_time": log_time,
        "end_time":   log_end,
        "sent":       sent_count,
        "total":      total,
        "mail_mode":  MAIL_MODE,
        "sent_by":    session.get("user_email"),   # audit log
        "results":    results,
    }
    save_history(log_entry)

    status = "success" if sent_count == total else ("partial" if sent_count > 0 else "failed")
    return jsonify({
        "status":        status,
        "sent":          sent_count,
        "total":         total,
        "filtered_file": filtered_path,
        "results":       results,
    })


@app.route("/api/history", methods=["GET"])
@login_required
def get_history():
    return jsonify(load_history())


@app.route("/api/history/<entry_id>", methods=["DELETE"])
@login_required
def delete_history(entry_id):
    history = load_history()
    history = [h for h in history if h.get("id") != entry_id]
    HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2))
    return jsonify({"ok": True})


# ─── Gmail sending OAuth2 Routes (admin-only, để cấp token gửi mail) ────────

@app.route("/auth/gmail")
@login_required
def auth_gmail():
    if not GMAIL_CLIENT_ID:
        return jsonify({"error": "GMAIL_CLIENT_ID chưa được cấu hình trong .env"}), 500
    redirect_uri = os.getenv(
        "GMAIL_REDIRECT_URI",
        request.host_url.rstrip("/") + "/auth/gmail/callback"
    )
    url = build_oauth_url(redirect_uri)
    return redirect(url)


@app.route("/auth/gmail/callback")
@login_required
def auth_gmail_callback():
    code = request.args.get("code")
    error = request.args.get("error")
    if error or not code:
        return f"<h3>Lỗi xác thực: {error or 'Không nhận được code'}</h3>", 400
    try:
        redirect_uri = os.getenv(
            "GMAIL_REDIRECT_URI",
            request.host_url.rstrip("/") + "/auth/gmail/callback"
        )
        exchange_code_for_token(code, redirect_uri)
        return """
        <html><body style="font-family:monospace;background:#0a0e14;color:#00e5a0;
                           display:flex;align-items:center;justify-content:center;
                           height:100vh;margin:0;font-size:18px;text-align:center">
          <div>
            <div style="font-size:48px;margin-bottom:16px">✓</div>
            <div>Đăng nhập Gmail thành công!</div>
            <div style="font-size:13px;color:#6b8ab0;margin-top:12px">
              Token đã được lưu. Bạn có thể đóng tab này.
            </div>
            <a href="/" style="display:inline-block;margin-top:24px;
               color:#00c2ff;font-size:13px">← Quay lại ứng dụng</a>
          </div>
        </body></html>"""
    except Exception as e:
        return f"<h3>Lỗi đổi token: {e}</h3>", 500


@app.route("/api/auth/status")
@login_required
def auth_status():
    from incident_agent import _load_token
    tok = _load_token()
    has_token = bool(tok.get("refresh_token"))
    return jsonify({"authenticated": has_token})


if __name__ == "__main__":
    app.run(debug=True, port=5000, host="0.0.0.0")
