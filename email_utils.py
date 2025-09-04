import os
import smtplib
from email.message import EmailMessage

def send_email_with_ppt(recipient: str, subject: str, body: str, file_paths: list):
    """
    Send an email with one or more PPT attachments.
    Includes better error handling and debug logs.
    """
    sender = os.getenv("EMAIL_USER", "pr3101165@gmail.com")
    password = os.getenv("EMAIL_PASSWORD", "vvrc wfvy pkox wsqn")

    if not sender or not password:
        print("❌ ERROR: Email credentials are missing. Set EMAIL_USER and EMAIL_PASSWORD.")
        return False

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

        try:
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
        except Exception as e:
            print(f"❌ Failed to attach {file_path}: {e}")

    try:
        print(f"📧 Connecting to Gmail SMTP as {sender}...")
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(sender, password)
            smtp.send_message(msg)
        print(f"✅ Email sent to {recipient} with {len(file_paths)} attachments")
        return True

    except smtplib.SMTPAuthenticationError:
        print("❌ Authentication failed: Invalid email or password (check App Password settings).")
    except smtplib.SMTPConnectError:
        print("❌ Connection failed: Could not connect to Gmail SMTP server.")
    except Exception as e:
        print(f"❌ Unexpected error while sending email: {e}")

    return False
