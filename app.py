"""
VNG Cloud — Incident Notification Web App
Flask backend: Upload Excel → Preview → Fill form → Send Email
"""

import os
import json
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, render_template, send_from_directory
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


# ─── Helpers ─────────────────────────────────────────────────────────────────

def load_history():
    try:
        return json.loads(HISTORY_FILE.read_text())
    except Exception:
        return []


def save_history(entry: dict):
    history = load_history()
    history.insert(0, entry)
    history = history[:100]   # giữ tối đa 100 bản ghi
    HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2))


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/teams-config")
def teams_config():
    return render_template("teams_config.html")


@app.route("/api/upload", methods=["POST"])
def upload_excel():
    """Upload file Excel, trả về preview dữ liệu."""
    if "file" not in request.files:
        return jsonify({"error": "Không tìm thấy file trong request"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Chưa chọn file"}), 400

    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": "Chỉ chấp nhận file .xlsx hoặc .xls"}), 400

    # Lưu file với tên duy nhất
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

    # Group preview theo Owned By
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
        "file_id":  uid,
        "filename": file.filename,
        "filepath": str(filepath),
        "total_rows":   len(df),
        "total_groups": len(grouped_preview),
        "columns":  list(df.columns),
        "preview":  grouped_preview,
    })


@app.route("/api/send", methods=["POST"])
def send_incident():
    """Validate form → build email → send → log history."""
    data = request.json or {}

    required = ["template_type", "service_name", "incident_desc",
                "start_time", "root_cause", "filepath"]
    missing = [f for f in required if not data.get(f)]
    if missing:
        return jsonify({"error": f"Thiếu trường: {', '.join(missing)}"}), 400

    tt = data["template_type"]
    if tt not in ("1", "2", "3", "4"):
        return jsonify({"error": "template_type phải là 1–4"}), 400
    if tt in ("1", "4") and not data.get("end_time"):
        return jsonify({"error": "end_time bắt buộc với template 1 và 4"}), 400
    if tt == "3" and not data.get("solution"):
        return jsonify({"error": "solution bắt buộc với template 3"}), 400

    filepath = data["filepath"]
    if not Path(filepath).is_file():
        return jsonify({"error": f"File không tồn tại: {filepath}"}), 400

    # Đọc Excel
    try:
        df = extract_columns(filepath)
    except Exception as e:
        return jsonify({"error": f"Lỗi đọc Excel: {e}"}), 422

    if df.empty:
        return jsonify({"error": "Không có dữ liệu hợp lệ"}), 422

    # Xuất filtered Excel
    try:
        filtered_path = export_filtered_excel(df, filepath)
    except Exception as e:
        return jsonify({"error": f"Lỗi xuất Excel: {e}"}), 500

    # Lấy Graph token
    try:
        token = get_graph_token()
    except Exception as e:
        return jsonify({"error": f"Lỗi lấy Microsoft Graph token: {e}"}), 500

    # Build info dict
    info = {k: data.get(k, "") for k in [
        "template_type", "service_name", "incident_desc",
        "start_time", "end_time", "root_cause", "status", "solution"
    ]}

    # Gửi email
    results = []
    for owned_by_email, group in df.groupby("Owned By"):
        devices = group.to_dict("records")
        cc_list = sorted(group["Salesman Email"].dropna().unique().tolist())
        if "support@vngcloud.vn" not in cc_list:
            cc_list.append("support@vngcloud.vn")
        try:
            subject, html_body = build_email_html(info, devices)
            ok = send_email(token, owned_by_email, subject, html_body, cc_emails=cc_list)
            results.append({
                "to":      owned_by_email,
                "cc":      cc_list,
                "count":   len(devices),
                "sent":    ok,
            })
        except Exception as e:
            results.append({
                "to":    owned_by_email,
                "cc":    cc_list,
                "count": len(devices),
                "sent":  False,
                "error": str(e),
            })

    sent_count = sum(1 for r in results if r["sent"])
    total      = len(results)

    # Lưu lịch sử
    log_entry = {
        "id":          uuid.uuid4().hex[:8],
        "timestamp":   datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        "template":    tt,
        "service":     data["service_name"],
        "incident":    data["incident_desc"],
        "start_time":  data["start_time"],
        "end_time":    data.get("end_time", ""),
        "sent":        sent_count,
        "total":       total,
        "mail_mode":   MAIL_MODE,
        "results":     results,
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
def get_history():
    return jsonify(load_history())


@app.route("/api/history/<entry_id>", methods=["DELETE"])
def delete_history(entry_id):
    history = load_history()
    history = [h for h in history if h.get("id") != entry_id]
    HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2))
    return jsonify({"ok": True})


# ─── Gmail OAuth2 Routes ─────────────────────────────────────────────────────

@app.route("/auth/gmail")
def auth_gmail():
    """Redirect user to Google consent screen."""
    if not GMAIL_CLIENT_ID:
        return jsonify({"error": "GMAIL_CLIENT_ID chưa được cấu hình trong .env"}), 500
    redirect_uri = os.getenv(
        "GMAIL_REDIRECT_URI",
        request.host_url.rstrip("/") + "/auth/gmail/callback"
    )
    url = build_oauth_url(redirect_uri)
    from flask import redirect as flask_redirect
    return flask_redirect(url)


@app.route("/auth/gmail/callback")
def auth_gmail_callback():
    """Google redirects here with ?code=..."""
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
def auth_status():
    """Kiểm tra xem Gmail token đã được lưu chưa."""
    from incident_agent import _load_token
    tok = _load_token()
    has_token = bool(tok.get("refresh_token"))
    return jsonify({"authenticated": has_token})



if __name__ == "__main__":
    app.run(debug=True, port=5000, host="0.0.0.0")
