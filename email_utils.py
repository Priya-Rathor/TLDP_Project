import os
import base64
from sib_api_v3_sdk import ApiClient, Configuration
from sib_api_v3_sdk.api import transactional_emails_api
from sib_api_v3_sdk.models import SendSmtpEmail, SendSmtpEmailAttachment, SendSmtpEmailTo
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Brevo API key from env
API_KEY = os.getenv("BREVO_API_KEY")

# Initialize API client
configuration = Configuration()
configuration.api_key['api-key'] = API_KEY
api_instance = transactional_emails_api.TransactionalEmailsApi(ApiClient(configuration))


def send_email_with_ppt(
    recipient: str,
    subject: str,
    html_content: str,
    sender_email: str,
    sender_name: str = "Sender",
    ppt_paths: list = None
):
    """
    Send an email via Brevo API with optional PPT attachments.
    
    :param recipient: Recipient email address
    :param subject: Email subject
    :param html_content: HTML content of the email
    :param sender_email: Verified sender email in Brevo
    :param sender_name: Name of the sender
    :param ppt_paths: List of local PPT file paths to attach
    """
    attachments = []
    
    if ppt_paths:
        for path in ppt_paths:
            with open(path, "rb") as f:
                encoded_content = base64.b64encode(f.read()).decode()
            attachments.append(SendSmtpEmailAttachment(content=encoded_content, name=os.path.basename(path)))
    
    email = SendSmtpEmail(
        to=[SendSmtpEmailTo(email=recipient)],
        sender={"email": sender_email, "name": sender_name},
        subject=subject,
        html_content=html_content,
        attachment=attachments if attachments else None
    )
    
    try:
        response = api_instance.send_transac_email(email)
        print("✅ Email sent successfully:", response)
    except Exception as e:
        print("❌ Failed to send email:", e)

