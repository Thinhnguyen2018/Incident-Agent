"""
VNG Cloud — Incident Notification Agent
========================================
Workflow:
  1. Nguoi dung nhap thong tin su co (interactive)
  2. Doc file Excel tu OPTool, loc 5 cot can thiet
  3. Xuat file Excel moi
  4. Soan email theo template BM-SDK-007
  5. Gui qua Gmail API (OAuth2)

Yeu cau: pip install -r requirements.txt
Cau hinh: copy .env.example -> .env roi dien thong tin
"""

import os
import sys
import json
import base64
import re
import requests
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

MAIL_MODE = "gmail"

# --- Cau hinh Gmail OAuth2 ---------------------------------------------------

GMAIL_CLIENT_ID     = os.getenv("GMAIL_CLIENT_ID")
GMAIL_CLIENT_SECRET = os.getenv("GMAIL_CLIENT_SECRET")
GMAIL_TOKEN_FILE    = os.getenv("GMAIL_TOKEN_FILE", "gmail_token.json")
SENDER_EMAIL        = os.getenv("GMAIL_SENDER_EMAIL")

GMAIL_TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_SEND_URL  = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
GMAIL_AUTH_URL  = "https://accounts.google.com/o/oauth2/v2/auth"
GMAIL_SCOPE     = "https://www.googleapis.com/auth/gmail.send"

REQUIRED_COLS = {
    "Owned By":       "Owned By",
    "Name":           "Name",
    "Floating IP":    "Floating IP",
    "IP Address":     "IP Address",
    "Salesman.Email": "Salesman Email",   # rename cho dễ đọc
}

# ─── Loại template ──────────────────────────────────────────────────────────

TEMPLATE_TYPES = {
    "1": "Sự cố xử lý nhanh (≤ 20 phút)",
    "2": "Sự cố xử lý kéo dài (> 20 phút)",
    "3": "Cập nhật tiến độ xử lý",
    "4": "Hoàn tất xử lý sự cố",
}


# ════════════════════════════════════════════════════════════════════════════
# PHẦN 1: NHẬP THÔNG TIN SỰ CỐ
# ════════════════════════════════════════════════════════════════════════════

def prompt(label: str, required: bool = True, default: str = "") -> str:
    """Hỏi người dùng một trường, hỗ trợ giá trị mặc định."""
    hint = f" [{default}]" if default else (" (bắt buộc)" if required else " (tuỳ chọn, Enter để bỏ qua)")
    while True:
        val = input(f"  {label}{hint}: ").strip()
        if not val and default:
            return default
        if not val and required:
            print("    ⚠  Trường này bắt buộc, vui lòng nhập lại.")
            continue
        return val


def collect_incident_info() -> dict:
    now_str = datetime.now().strftime("%d-%m-%Y %H:%M")

    print("\n" + "═" * 60)
    print("  VNG CLOUD — NHẬP THÔNG TIN SỰ CỐ")
    print("═" * 60)

    print("\n▸ Chọn loại thông báo:")
    for k, v in TEMPLATE_TYPES.items():
        print(f"    [{k}] {v}")
    while True:
        choice = input("  Nhập số (1–4): ").strip()
        if choice in TEMPLATE_TYPES:
            template_type = choice
            break
        print("    ⚠  Vui lòng nhập 1, 2, 3 hoặc 4.")

    print(f"\n▸ Template đã chọn: {TEMPLATE_TYPES[template_type]}")
    print("─" * 60)

    info = {"template_type": template_type}

    info["service_name"] = prompt("Tên hệ thống / dịch vụ bị ảnh hưởng")
    info["incident_desc"] = prompt(
        "Mô tả sự cố",
        default="một số khách hàng mất kết nối dịch vụ"
    )
    info["start_time"] = prompt(
        "Thời gian bắt đầu (DD-MM-YYYY HH:MM)",
        default=now_str
    )

    if template_type in ("1", "4"):
        info["end_time"] = prompt("Thời gian kết thúc (DD-MM-YYYY HH:MM)")

    if template_type in ("1", "4", "3"):
        info["root_cause"] = prompt("Nguyên nhân")
    else:
        info["root_cause"] = prompt(
            "Nguyên nhân",
            default="Các kỹ sư của VNG Cloud đang kiểm tra để xác định nguyên nhân chính xác"
        )

    if template_type == "2":
        info["status"] = prompt(
            "Tình trạng xử lý",
            default="VNG Cloud đang tập trung nguồn lực để khắc phục và sẽ thông báo ngay khi dịch vụ hoạt động trở lại"
        )

    if template_type == "3":
        info["solution"] = prompt("Hướng xử lý / giải pháp")

    print("\n▸ Đường dẫn file Excel từ OPTool:")
    while True:
        excel_path = input("  Nhập đường dẫn file .xlsx: ").strip().strip('"')
        if Path(excel_path).is_file():
            info["excel_path"] = excel_path
            break
        print(f"    ⚠  Không tìm thấy file: {excel_path}")

    print("─" * 60)
    return info


# ════════════════════════════════════════════════════════════════════════════
# PHẦN 2: ĐỌC VÀ LỌC FILE EXCEL
# ════════════════════════════════════════════════════════════════════════════

def extract_columns(excel_path: str) -> pd.DataFrame:
    """Đọc file OPTool, lọc và đổi tên các cột cần thiết."""
    print(f"\n[1/4] Đọc file Excel: {excel_path}")
    df_raw = pd.read_excel(excel_path)

    # Tìm các cột có trong file (tên cột có thể khác nhau đôi chút)
    found = {}
    for col_key, col_rename in REQUIRED_COLS.items():
        # tìm chính xác hoặc gần đúng (case-insensitive)
        match = next(
            (c for c in df_raw.columns if c.strip().lower() == col_key.lower()),
            None
        )
        if match:
            found[match] = col_rename
        else:
            print(f"    ⚠  Không tìm thấy cột '{col_key}' trong file. Bỏ qua.")

    df = df_raw[list(found.keys())].rename(columns=found).copy()
    df = df.dropna(subset=["Salesman Email"])

    print(f"    ✓ Đọc xong — {len(df)} dòng, {len(df.columns)} cột được lọc")
    return df


def export_filtered_excel(df: pd.DataFrame, source_path: str) -> str:
    """Xuất file Excel mới chứa 5 cột đã lọc."""
    stem = Path(source_path).stem
    out_path = str(Path(source_path).parent / f"{stem}_filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Devices")
        ws = writer.sheets["Devices"]
        # Auto-width cho các cột
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    print(f"\n[2/4] Xuất file Excel đã lọc: {out_path}")
    return out_path


# ════════════════════════════════════════════════════════════════════════════
# PHẦN 3: SOẠN EMAIL THEO TEMPLATE BM-SDK-007
# ════════════════════════════════════════════════════════════════════════════

def _fmt_date_en(date_str: str) -> str:
    """Chuyển DD-MM-YYYY HH:MM → DD-Mon-YYYY HH:MM cho phần tiếng Anh."""
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    try:
        dt = datetime.strptime(date_str.strip(), "%d-%m-%Y %H:%M")
        return f"{dt.day:02d}-{months[dt.month-1]}-{dt.year} {dt.strftime('%H:%M')}"
    except Exception:
        return date_str  # trả về nguyên nếu không parse được


def _vm_table_html(devices: list[dict]) -> str:
    """Tạo bảng HTML danh sách VM bị ảnh hưởng."""
    rows = ""
    for d in devices:
        floating = d.get("Floating IP") or "—"
        rows += (
            f"<tr>"
            f"<td style='padding:4px 8px;border:1px solid #ddd'>{d.get('Name','')}</td>"
            f"<td style='padding:4px 8px;border:1px solid #ddd'>{d.get('Owned By','')}</td>"
            f"<td style='padding:4px 8px;border:1px solid #ddd'>{d.get('IP Address','')}</td>"
            f"<td style='padding:4px 8px;border:1px solid #ddd'>{floating}</td>"
            f"</tr>"
        )
    return (
        "<table style='border-collapse:collapse;font-size:13px;margin:8px 0'>"
        "<thead><tr style='background:#f5f5f5'>"
        "<th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Tên VM / VM Name</th>"
        "<th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Khách hàng / Customer</th>"
        "<th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>IP Address</th>"
        "<th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Floating IP</th>"
        "</tr></thead>"
        f"<tbody>{rows}</tbody></table>"
    )


def build_email_html(info: dict, devices: list[dict]) -> tuple[str, str]:
    """Trả về (subject, html_body) theo loại template."""
    t   = info["template_type"]
    svc = info["service_name"]
    inc = info["incident_desc"]
    t0  = info["start_time"]
    t0_en = _fmt_date_en(t0)
    rc  = info.get("root_cause", "")
    rc_en = rc  # same text used for EN

    date_label = datetime.now().strftime("%d-%m-%Y %H:%M")

    vm_table = _vm_table_html(devices)
    vm_note  = (
        "<p>Danh sách dịch vụ/thiết bị của Quý Khách bị ảnh hưởng:<br>"
        "<i>List of your affected services/devices:</i></p>"
        + vm_table
    )

    def wrap(subject_vn: str, subject_en: str, body_vn: str, body_en: str) -> tuple[str, str]:
        subject = f"English below|[VNG Cloud][Thông báo] {subject_vn}"
        html = f"""
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#222;line-height:1.6">

<p>Kính gửi Quý Khách hàng,</p>
{body_vn}
{vm_note}
<p>Quý Khách hàng vui lòng kiểm tra lại các dịch vụ trên hệ thống của Quý Khách.
Nếu có thắc mắc hoặc cần hỗ trợ thêm, vui lòng liên hệ bộ phận hỗ trợ Khách hàng qua hotline
<strong>1900 1549</strong> hoặc email <a href="mailto:support@vngcloud.vn">support@vngcloud.vn</a>.</p>
<p>Trân trọng cảm ơn!</p>

<hr/>

<p><strong>[VNG Cloud][Notification] {subject_en}</strong></p>
<p>Dear Valued Customer,</p>
{body_en}
{vm_note}
<p>Please re-check all your services. If you need support, please contact our support team
via hotline <strong>19001549</strong> or email <a href="mailto:support@vngcloud.vn">support@vngcloud.vn</a>.</p>
<p>Thanks &amp; Best Regards,<br/>Support Team</p>
</body></html>
"""
        return subject, html

    # ── Template 1: Xử lý nhanh (đã xong) ──────────────────────────────────
    if t == "1":
        t1    = info.get("end_time", "")
        t1_en = _fmt_date_en(t1)
        return wrap(
            subject_vn=f"SỰ CỐ {svc.upper()} NGÀY {date_label}",
            subject_en=f"INCIDENT OF {svc.upper()} ON {_fmt_date_en(date_label)}",
            body_vn=f"""
<p>VNG Cloud đã hoàn tất việc khắc phục sự cố <strong>{inc}</strong>
gây ảnh hưởng đến dịch vụ <strong>{svc}</strong>.</p>
<p><strong>Thời gian bắt đầu:</strong> {t0} (GMT+7)<br/>
<strong>Thời gian kết thúc:</strong> {t1} (GMT+7)<br/>
<strong>Nguyên nhân:</strong> {rc}</p>
<p><strong>{svc}</strong> của VNG Cloud đã hoạt động trở lại bình thường.</p>
<p>VNG Cloud xin phép thông báo để Quý Khách hàng nắm thông tin và an tâm tiếp tục sử dụng dịch vụ.</p>
""",
            body_en=f"""
<p>VNG Cloud has finished resolving an incident <strong>{inc}</strong>
impacted to <strong>{svc}</strong>.</p>
<p><strong>Outage Start:</strong> {t0_en} (GMT+7)<br/>
<strong>Outage End:</strong> {t1_en} (GMT+7)<br/>
<strong>Root cause:</strong> {rc_en}</p>
<p>Please be informed that <strong>{svc}</strong> on VNG Cloud's infrastructure system has been recovered.</p>
""",
        )

    # ── Template 2: Xử lý kéo dài (đang diễn ra) ───────────────────────────
    if t == "2":
        status    = info.get("status", "")
        return wrap(
            subject_vn=f"SỰ CỐ {svc.upper()} NGÀY {date_label}",
            subject_en=f"INCIDENT OF {svc.upper()} ON {_fmt_date_en(date_label)}",
            body_vn=f"""
<p>Hiện tại VNG Cloud đang gặp sự cố sau đây:</p>
<p><strong>Tên sự cố:</strong> {inc}<br/>
<strong>Dịch vụ ảnh hưởng:</strong> {svc}<br/>
<strong>Thời gian bắt đầu:</strong> {t0} (GMT+7)<br/>
<strong>Nguyên nhân:</strong> {rc}<br/>
<strong>Tình trạng xử lý:</strong> {status}</p>
<p>VNG Cloud chân thành xin lỗi Quý Khách hàng vì sự bất tiện này.
Vui lòng theo dõi các email thông báo từ VNG Cloud về tình hình dịch vụ.</p>
""",
            body_en=f"""
<p>At present, VNG Cloud has occurred an incident as followings:</p>
<p><strong>Incident:</strong> {inc}<br/>
<strong>Impact Service:</strong> {svc}<br/>
<strong>Outage Start:</strong> {t0_en} (GMT+7)<br/>
<strong>Root Cause:</strong> {rc_en}<br/>
<strong>Status:</strong> {status}</p>
<p>We do apologize to you for this inconvenience. Please follow up on the next emails from VNG Cloud.</p>
""",
        )

    # ── Template 3: Cập nhật tiến độ ────────────────────────────────────────
    if t == "3":
        solution = info.get("solution", "")
        return wrap(
            subject_vn=f"CẬP NHẬT TÌNH HÌNH XỬ LÝ SỰ CỐ {svc.upper()} NGÀY {date_label}",
            subject_en=f"INCIDENT UPDATED {svc.upper()} ON {_fmt_date_en(date_label)}",
            body_vn=f"""
<p>Hiện tại, sự cố <strong>{inc}</strong> gây ảnh hưởng đến dịch vụ
<strong>{svc}</strong> đang trong quá trình xử lý.</p>
<p><strong>Thời gian bắt đầu:</strong> {t0} (GMT+7)<br/>
<strong>Nguyên nhân:</strong> {rc}<br/>
<strong>Hướng xử lý:</strong> {solution}</p>
<p>VNG Cloud đang tập trung nguồn lực để khắc phục và sẽ thông báo đến Quý Khách hàng
ngay khi dịch vụ hoạt động trở lại bình thường.</p>
<p>VNG Cloud chân thành xin lỗi Quý Khách hàng vì sự bất tiện này.</p>
""",
            body_en=f"""
<p>At present, incident of <strong>{inc}</strong> impacted to <strong>{svc}</strong> is under processed.</p>
<p><strong>Outage Start:</strong> {t0_en} (GMT+7)<br/>
<strong>Root cause:</strong> {rc_en}<br/>
<strong>Solution:</strong> {solution}</p>
<p>We are focusing on resolving the incident and will keep you updated as soon as possible.</p>
<p>We do apologize to you for this inconvenience. Please follow up on the next emails from VNG Cloud.</p>
""",
        )

    # ── Template 4: Hoàn tất xử lý ──────────────────────────────────────────
    t1    = info.get("end_time", "")
    t1_en = _fmt_date_en(t1)
    return wrap(
        subject_vn=f"HOÀN TẤT XỬ LÝ SỰ CỐ {svc.upper()} NGÀY {date_label}",
        subject_en=f"INCIDENT RESOLVED {svc.upper()} ON {_fmt_date_en(date_label)}",
        body_vn=f"""
<p>VNG Cloud đã hoàn tất việc khắc phục sự cố <strong>{inc}</strong>
gây ảnh hưởng đến dịch vụ <strong>{svc}</strong>.</p>
<p><strong>Thời gian bắt đầu:</strong> {t0} (GMT+7)<br/>
<strong>Thời gian kết thúc:</strong> {t1} (GMT+7)<br/>
<strong>Nguyên nhân:</strong> {rc}</p>
<p><strong>{svc}</strong> của VNG Cloud đã hoạt động trở lại bình thường.</p>
<p>VNG Cloud xin phép thông báo để Quý Khách hàng nắm thông tin và an tâm tiếp tục sử dụng dịch vụ.</p>
""",
        body_en=f"""
<p>VNG Cloud has finished resolving an incident <strong>{inc}</strong>
impacted to <strong>{svc}</strong>.</p>
<p><strong>Outage Start:</strong> {t0_en} (GMT+7)<br/>
<strong>Outage End:</strong> {t1_en} (GMT+7)<br/>
<strong>Root cause:</strong> {rc_en}</p>
<p>Please be informed that <strong>{svc}</strong> on VNG Cloud's infrastructure system has been recovered.</p>
""",
    )


# ════════════════════════════════════════════════════════════════════════════
# PHẦN 4: GỬI EMAIL QUA MICROSOFT GRAPH API
# ════════════════════════════════════════════════════════════════════════════


# ============================================================================
# PHAN 4: GUI EMAIL QUA GMAIL API (OAuth2)
# ============================================================================

def _load_token() -> dict:
    """Doc token da luu tu file."""
    p = Path(GMAIL_TOKEN_FILE)
    if p.exists():
        return json.loads(p.read_text())
    return {}


def _save_token(token: dict):
    Path(GMAIL_TOKEN_FILE).write_text(json.dumps(token, indent=2))


def get_graph_token() -> str:
    """
    Tra ve access_token Gmail con han.
    Tu dong refresh neu het han.
    Raise RuntimeError neu chua co refresh_token (can chay OAuth flow truoc).
    """
    tok = _load_token()
    if not tok.get("refresh_token"):
        raise RuntimeError(
            "Chua co Gmail token. Vui long dang nhap Gmail truoc "
            "bang cach truy cap /auth/gmail tren web app."
        )

    # Thu dung access_token hien tai
    if tok.get("access_token"):
        test = requests.get(
            "https://gmail.googleapis.com/gmail/v1/users/me/profile",
            headers={"Authorization": f"Bearer {tok['access_token']}"},
        )
        if test.status_code == 200:
            return tok["access_token"]

    # Refresh
    resp = requests.post(GMAIL_TOKEN_URL, data={
        "client_id":     GMAIL_CLIENT_ID,
        "client_secret": GMAIL_CLIENT_SECRET,
        "refresh_token": tok["refresh_token"],
        "grant_type":    "refresh_token",
    })
    resp.raise_for_status()
    new_tok = resp.json()
    tok["access_token"] = new_tok["access_token"]
    _save_token(tok)
    return tok["access_token"]


def exchange_code_for_token(code: str, redirect_uri: str) -> dict:
    """Doi authorization code lay token (goi 1 lan duy nhat khi dang nhap)."""
    resp = requests.post(GMAIL_TOKEN_URL, data={
        "code":          code,
        "client_id":     GMAIL_CLIENT_ID,
        "client_secret": GMAIL_CLIENT_SECRET,
        "redirect_uri":  redirect_uri,
        "grant_type":    "authorization_code",
    })
    resp.raise_for_status()
    token = resp.json()
    _save_token(token)
    return token


def build_oauth_url(redirect_uri: str) -> str:
    """Tao URL de nguoi dung click dang nhap Google."""
    import urllib.parse
    params = {
        "client_id":     GMAIL_CLIENT_ID,
        "redirect_uri":  redirect_uri,
        "response_type": "code",
        "scope":         GMAIL_SCOPE,
        "access_type":   "offline",
        "prompt":        "consent",
    }
    return GMAIL_AUTH_URL + "?" + urllib.parse.urlencode(params)


def send_email(token: str, to_email: str, subject: str, html_body: str,
               cc_emails: list | None = None) -> bool:
    """
    Gui email qua Gmail API.
    - token     : access_token Gmail
    - to_email  : dia chi chinh (Owned By)
    - cc_emails : danh sach CC (Salesman Email)
    """
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = to_email

    # Luon CC ve chinh sender de luu ban sao
    cc_all = list(cc_emails or [])
    if "support@vngcloud.vn" not in [e.lower() for e in cc_all]:
        cc_all.append("support@vngcloud.vn")
    msg["Cc"] = ", ".join(cc_all)

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

    resp = requests.post(
        GMAIL_SEND_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        json={"raw": raw},
    )

    if resp.status_code == 200:
        return True
    print(f"    Loi gui toi {to_email}: {resp.status_code} - {resp.text[:200]}")
    return False

def check_env():
    missing = [v for v in ["GMAIL_CLIENT_ID", "GMAIL_CLIENT_SECRET", "GMAIL_SENDER_EMAIL"]
               if not os.getenv(v)]
    if missing:
        print(f"\n⛔  Thiếu biến môi trường: {', '.join(missing)}")
        print("    Hãy copy file .env.example → .env rồi điền đầy đủ.\n")
        sys.exit(1)


def main():
    print("\n" + "═" * 60)
    print("  VNG CLOUD — INCIDENT NOTIFICATION AGENT  v1.0")
    print("═" * 60)

    check_env()

    # Bước 1: Thu thập thông tin
    info = collect_incident_info()

    # Bước 2: Đọc và lọc Excel
    df = extract_columns(info["excel_path"])

    # Bước 3: Xuất file Excel đã lọc
    filtered_path = export_filtered_excel(df, info["excel_path"])

    # Bước 4 + 5: Soạn và gửi email theo từng salesman
    print(f"\n[3/4] Lấy token Microsoft Graph...")
    token = get_graph_token()  # lay Gmail access token
    print("    ✓ Token OK")

    # Nhóm theo Owned By — mỗi khách hàng nhận 1 email
    # Salesman Email của nhóm đó được CC vào
    grouped = df.groupby("Owned By")
    total   = len(grouped)
    sent    = 0

    print(f"\n[4/4] Gửi email cho {total} khách hàng (Owned By)...\n")
    for owned_by_email, group in grouped:
        devices = group.to_dict("records")

        # Lấy danh sách salesman duy nhất trong nhóm này để CC
        cc_list = sorted(group["Salesman Email"].dropna().unique().tolist())

        subject, html_body = build_email_html(info, devices)

        cc_display = ", ".join(cc_list) if cc_list else "—"
        print(f"  → TO: {owned_by_email} ({len(devices)} VM)")
        print(f"     CC: {cc_display} ... ", end="", flush=True)
        ok = send_email(token, owned_by_email, subject, html_body, cc_emails=cc_list)
        if ok:
            print("✓ Đã gửi")
            sent += 1
        else:
            print("✗ Thất bại")

    print("\n" + "═" * 60)
    print(f"  Hoàn tất: {sent}/{total} email đã gửi thành công")
    print(f"  Gửi tới : Owned By (TO) · Salesman (CC)")
    print(f"  File Excel đã lọc: {filtered_path}")
    print("═" * 60 + "\n")


if __name__ == "__main__":
    main()
