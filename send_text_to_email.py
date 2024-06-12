import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os


def send_email(sender_email, sender_password, to_address, subject, message, pdf_path, link):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    msg = MIMEMultipart()
    msg['From'] = f"David smith"  # <{sender_email}>
    # msg['To'] = to_address
    msg['Subject'] = subject
    msg["Reply-To"] = "asdfa@gmail.com"  # Set the reply-to address

    # Set the message content as plain text
    msg.attach(MIMEText(message, 'plain'))

    # Set the HTML content
    if link != '':
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <body>
        <a href="{link}" style="font-weight: bold; font-size: small;">Google</a>
        </body>
        </html>
        """
        html_part = MIMEText(html_content, 'html')
        msg.attach(html_part)

    # Attach the PDF file
    if pdf_path != '':
        with open(pdf_path, 'rb') as file:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(file.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', 'attachment; filename="{}"'.format(os.path.basename(pdf_path)))
            msg.attach(attachment)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(from_addr=sender_email, to_addrs=to_address, msg=msg.as_string())
        print('Email sent successfully!')
        return True
    except Exception as e:
        print('Error sending email:', str(e))
        return False
    finally:
        server.quit()


# Read list of recipient emails (csv file should must be contain 100 maximum emails or less)
df = pd.read_csv("derin.csv")
df.dropna(inplace=True)
recipient_emails_list = df['Email'].to_list()

# Update the sender-email, sender email-passowrd, and recipient-email below
sender_email = "hr@afsdfasd.net"
sender_password = ""
recipient_email = recipient_emails_list
subject = "How are you do it now"

# update Notepad file name below. (Note:Keep always saved your Notepad, Should be in same directory.)
with open('email_text_data.txt', 'r') as email_text:
    message = email_text.read()

# pdf file should be in same directory
pdf_path = "demo.pdf"

# URL
link = "https://www.google.com"

# Don't change anything below
send_email(sender_email, sender_password, recipient_email, subject, message, pdf_path, link)
