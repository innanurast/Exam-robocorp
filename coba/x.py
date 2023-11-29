from RPA.Email.ImapSmtp import ImapSmtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import os

try:
    account = "innanur16@gmail.com"
    password = "tino hnxh rjry sukd"

    def send_email(recipient, subject, body, attachment_path=None):
        gmail_account = account
        gmail_password = password
        mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
        mail.authorize(account=gmail_account, password=gmail_password)

        message = mail.send_message(
            sender=gmail_account,
            recipients=recipient,
            subject=subject,
            body=body,
        )

        if attachment_path:
            attach_file_to_message(message, attachment_path)

        mail.send_message(message)

except Exception as e:
    print(f"Error sending email: {e}")

def attach_file_to_message(message, file_path):
    multipart_message, _ = message  # Unpack the tuple
    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {os.path.basename(file_path)}",
        )
        message.attach_file(part)
