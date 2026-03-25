from flask import Flask, request, render_template
from datetime import datetime
import os
import requests
from dotenv import load_dotenv
import base64

load_dotenv()

app = Flask(__name__)

# 限制整个请求大小（防止上传超大文件）
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25MB

# ==============================
# Environment Variables
# ==============================
GRAPH_TENANT_ID = os.getenv("GRAPH_TENANT_ID")
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET")

GRAPH_SENDER = os.getenv("GRAPH_SENDER", "hr@marlugroupwa.com.au")
GRAPH_TO = os.getenv("GRAPH_TO", "accounts@marlugroupwa.com.au")

# ==============================
# Attachment Rules
# ==============================
MAX_FILES = 3
MAX_TOTAL_SIZE_MB = 20
MAX_TOTAL_SIZE = MAX_TOTAL_SIZE_MB * 1024 * 1024

ALLOWED_EXTENSIONS = {
    "pdf", "jpg", "jpeg", "png",
    "doc", "docx", "xls", "xlsx"
}


# ==============================
# Auth
# ==============================
def get_graph_access_token():
    token_url = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"

    data = {
        "client_id": GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    response = requests.post(token_url, data=data, timeout=30)
    response.raise_for_status()
    return response.json()["access_token"]


# ==============================
# File Helpers
# ==============================
def allowed_file(filename):
    if "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in ALLOWED_EXTENSIONS


def build_graph_attachments(files):
    attachments = []

    valid_files = [f for f in files if f and f.filename]

    # 限制数量
    if len(valid_files) > MAX_FILES:
        raise ValueError(f"You can upload up to {MAX_FILES} files only.")

    total_size = 0

    for file in valid_files:
        filename = file.filename.strip()

        if not allowed_file(filename):
            raise ValueError(f"File type not allowed: {filename}")

        file_bytes = file.read()
        total_size += len(file_bytes)

        # 限制总大小
        if total_size > MAX_TOTAL_SIZE:
            raise ValueError("Total attachment size cannot exceed 20MB.")

        ext = filename.rsplit(".", 1)[1].lower()

        content_type_map = {
            "pdf": "application/pdf",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "png": "image/png",
            "doc": "application/msword",
            "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "xls": "application/vnd.ms-excel",
            "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }

        content_type = content_type_map.get(ext, "application/octet-stream")

        attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": filename,
            "contentType": content_type,
            "contentBytes": base64.b64encode(file_bytes).decode("utf-8")
        })

    return attachments


# ==============================
# Send Email
# ==============================
def send_email(name, phone, email, employer, site, pay_period, query_type, description, attachments=None):
    access_token = get_graph_access_token()

    to_recipients = []
    for addr in GRAPH_TO.split(","):
        addr = addr.strip()
        if addr:
            to_recipients.append({
                "emailAddress": {"address": addr}
            })

    body = f"""Payroll Query

Name: {name}
Phone: {phone}
Email: {email}
Employer: {employer}
Site: {site or 'N/A'}
Pay Period: {pay_period}
Type: {query_type}
Description: {description or 'N/A'}
"""

    message = {
        "subject": f"Payroll Query - {name}",
        "body": {
            "contentType": "Text",
            "content": body
        },
        "toRecipients": to_recipients
    }

    if attachments:
        message["attachments"] = attachments

    url = f"https://graph.microsoft.com/v1.0/users/{GRAPH_SENDER}/sendMail"

    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        },
        json={"message": message, "saveToSentItems": True},
        timeout=60
    )

    response.raise_for_status()


# ==============================
# Routes
# ==============================
@app.route("/")
def home():
    return render_template("form.html")


@app.route("/submit", methods=["POST"])
def submit():
    try:
        name = request.form.get("name", "").strip()
        phone = request.form.get("phone", "").strip()
        email = request.form.get("email", "").strip()
        employer = request.form.get("employer", "").strip()
        site = request.form.get("site", "").strip()
        pay_period_start = request.form.get("pay_period_start", "").strip()
        pay_period_end = request.form.get("pay_period_end", "").strip()
        query_type = request.form.get("query_type", "").strip()
        description = request.form.get("description", "").strip()

        # 必填校验
        if not all([name, phone, email, employer, pay_period_start, pay_period_end, query_type]):
            return "Missing required fields", 400

        start_text = datetime.strptime(pay_period_start, "%Y-%m-%d").strftime("%d/%m/%Y")
        end_text = datetime.strptime(pay_period_end, "%Y-%m-%d").strftime("%d/%m/%Y")

        pay_period = f"{start_text} to {end_text}"

        # 处理附件
        uploaded_files = request.files.getlist("attachments")
        graph_attachments = build_graph_attachments(uploaded_files)

        send_email(
            name, phone, email, employer,
            site, pay_period, query_type,
            description, graph_attachments
        )

        return "Submitted successfully ✅"

    except ValueError as e:
        return f"Attachment error: {str(e)}", 400

    except requests.HTTPError as e:
        return f"Graph API error: {e.response.text}", 500

    except Exception as e:
        return f"Error: {str(e)}", 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)