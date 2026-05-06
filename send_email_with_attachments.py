import os
import smtplib
from pathlib import Path
from datetime import datetime, timezone, timedelta
from email.message import EmailMessage

EMAIL_USER = os.environ["EMAIL_USER"]
EMAIL_PASS = os.environ["EMAIL_PASS"]
EMAIL_TO = os.environ["EMAIL_TO"]

OUTPUT_DIR = Path("outputs")

KST = timezone(timedelta(hours=9))
today = datetime.now(KST).strftime("%Y-%m-%d")
now_text = datetime.now(KST).strftime("%Y-%m-%d %H:%M")

workflow_name = os.environ.get("WORKFLOW_NAME", "FUT.GG 자동 크롤링")
subject_prefix = os.environ.get("EMAIL_SUBJECT_PREFIX", "FUT.GG 자동 크롤링 결과")

files = sorted(
    list(OUTPUT_DIR.glob("*.xlsx")) +
    list(OUTPUT_DIR.glob("*.txt")) +
    list(OUTPUT_DIR.glob("*.png")) +
    list(OUTPUT_DIR.glob("*.html"))
)

if not files:
    raise FileNotFoundError("outputs 폴더에서 첨부할 결과 파일을 찾지 못했습니다.")

msg = EmailMessage()
msg["Subject"] = f"{subject_prefix} ({today})"
msg["From"] = EMAIL_USER
msg["To"] = EMAIL_TO

file_list = "\n".join(f"- {f.name}" for f in files)

msg.set_content(
    f"{workflow_name} 실행이 완료되었습니다.\n\n"
    f"기준시각: {now_text} KST\n\n"
    f"첨부파일:\n{file_list}\n\n"
    "GitHub Actions artifact에도 동일한 결과가 저장되어 있습니다."
)

for file in files:
    data = file.read_bytes()

    if file.suffix.lower() == ".xlsx":
        maintype = "application"
        subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif file.suffix.lower() == ".txt":
        maintype = "text"
        subtype = "plain"
    elif file.suffix.lower() == ".html":
        maintype = "text"
        subtype = "html"
    elif file.suffix.lower() == ".png":
        maintype = "image"
        subtype = "png"
    else:
        maintype = "application"
        subtype = "octet-stream"

    msg.add_attachment(
        data,
        maintype=maintype,
        subtype=subtype,
        filename=file.name,
    )

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL_USER, EMAIL_PASS)
    smtp.send_message(msg)

print(f"Email sent to {EMAIL_TO} with {len(files)} attachment(s).")
