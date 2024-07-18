import aiosmtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import dotenv

dotenv.load_dotenv()
SMTP_MAIL = os.environ.get("SMTP_MAIL")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")


async def send_email(
    subject: str, recipient: str, html_content: str, attachment_path: str = None
):
    message = MIMEMultipart("alternative")
    message["From"] = f"MLSA Event Certificates <{SMTP_MAIL}>"
    message["To"] = recipient
    message["Subject"] = subject

    # Attach the HTML body
    message.attach(MIMEText(html_content, "html"))

    # Attach a file (optional)
    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEApplication(
                attachment.read(), Name=os.path.basename(attachment_path)
            )
        part["Content-Disposition"] = (
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        message.attach(part)

    await aiosmtplib.send(
        message,
        hostname="smtp.gmail.com",  # Gmail SMTP server
        port=587,  # Gmail SMTP port (TLS)
        username=SMTP_MAIL,  # Your Gmail address
        password=SMTP_PASSWORD,  # Your Gmail app password
        start_tls=True,
    )


# # test out the script
# import asyncio

# print(SMTP_MAIL, SMTP_PASSWORD)

# asyncio.run(send_email("Test", "iam.sabyasachi.seal@gmail.com", "<h1>Test</h1>"))
