import os
import smtplib
from email.message import EmailMessage

def send_email_with_ppt(recipient: str, subject: str, body: str, file_paths: list):
    """
    Send an email with one or more PPT attachments.
    """
    sender = os.getenv("EMAIL_USER", "divyanshi.pal.2408@gmail.com")
    password = os.getenv("EMAIL_PASSWORD", "llpd xolt lubx qqyc")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    # Attach all PPTs
    for file_path in file_paths:
        if not os.path.exists(file_path):
            print(f"⚠️ File not found, skipping: {file_path}")
            continue

        with open(file_path, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(file_path)
            msg.add_attachment(
                file_data,
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
                filename=file_name
            )
            print(f"📎 Attached: {file_name}")

    # Send using Gmail SMTP
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)

    print(f"✅ Email sent to {recipient} with {len(file_paths)} attachments")