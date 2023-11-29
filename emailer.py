from RPA.Email.ImapSmtp import ImapSmtp

try :
    account = "innanur16@gmail.com"
    password = "tino hnxh rjry sukd"

    def send_email(recipient, subject, body):
        gmail_account = account
        gmail_password = password
        mail = ImapSmtp(smtp_server="smtp.gmail.com", smtp_port=587)
        mail.authorize(account=gmail_account, password=gmail_password)
        mail.send_message(
            sender=gmail_account,
            recipients=recipient,
            subject=subject,
            body=body
        )
except Exception as e:
    print(f"Error sending email: {e}")