import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
import configparser
import csv
import time
import pandas as pd
from email.mime.base import MIMEBase
from email import encoders
import os

config = configparser.ConfigParser(interpolation=None)
config.read("./config.ini")

print("Start")
print(config.sections())
sender_name = config["SMTP"]["Sender_name"] # me
sender_email = config["SMTP"]["Sender_email"] # sender@ntplc.co.th
password = config.get("SMTP", "Password") # ""
subject = config["SMTP"]["SUBJECT"] # "แบบสอบถามการใช้ข้อมูลในรายงานทางการเงินในระบบ Datawarehouse"
mail_server = config["SMTP"]["Server"] # "ncmail.ntplc.co.th"
port = config["SMTP"]["Port"] # 465 for SSL
print(sender_name)
print(sender_email)

text_email = """
testing .....

This is a simple plain text email sent from a Python script.
No HTML, no multipart structure. Just the basics.

Regards,
The Script
"""

html_email = """
<html>
  <body>
    <h1>HTML Email Test</h1>
    <p>
      This is a <strong>simple HTML email</strong> sent from a Python script.
    </p>
    <p>
      It supports various tags like:
      <ul>
        <li><b>Bold</b> and <i>Italic</i> text.</li>
        <li>Links, like this one to <a href=\"https://www.python.org\">the Python website</a>.</li>
        <li>And other HTML formatting.</li>
      </ul>
    </p>
    <p>Regards,<br>The Script</p>
  </body>
</html>
"""

# อ่านข้อมูลจากไฟล์ Email.xlsx (sheet แรก)
try:
    recipients_df = pd.read_excel('Email.xlsx')
except FileNotFoundError:
    print("Error: 'Email.xlsx' not found. Please make sure the file is in the correct directory.")
    exit()

# --- ตั้งค่าไฟล์แนบ ---
attachment_filename = "position_.csv" # <--- !!! ระบุชื่อไฟล์ที่ต้องการแนบที่นี่ !!!

# Create a secure SSL context
context = ssl.create_default_context()

print(f"Connecting to server at {mail_server}:{port}...")
with smtplib.SMTP_SSL(mail_server, port, context=context) as server:
    print("Logging in...")
    server.login(sender_email, password)
    print("Login successful.")

    for idx, row in recipients_df.iterrows():
        # ตรวจสอบว่าคอลัมน์ 'name' และ 'email' มีอยู่ใน DataFrame หรือไม่
        if 'name' not in row or 'email' not in row:
            print(f"Skipping row {idx+2} because it's missing 'name' or 'email' column.")
            continue

        recipient_name = row['name']
        recipient_email = row['email']
        
        if pd.isna(recipient_email) or pd.isna(recipient_name):
            print(f"Skipping row {idx+2} due to empty name or email.")
            continue

        # สร้าง message หลักแบบ "mixed" เพื่อให้แนบไฟล์ได้
        message = MIMEMultipart("mixed")
        message["Subject"] = subject
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = formataddr((recipient_name, recipient_email))

        # สร้างส่วนของเนื้อหาอีเมล (HTML และ Text)
        body_part = MIMEMultipart("alternative")
        part1 = MIMEText(text_email, "plain")
        part2 = MIMEText(html_email, "html")
        body_part.attach(part1)
        body_part.attach(part2)

        # แนบส่วนเนื้อหาเข้าไปใน message หลัก
        message.attach(body_part)

        # --- ส่วนของการแนบไฟล์ ---
        try:
            with open(attachment_filename, "rb") as attachment:
                file_part = MIMEBase("application", "octet-stream")
                file_part.set_payload(attachment.read())
            
            encoders.encode_base64(file_part)
            
            file_part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(attachment_filename)}",
            )
            
            # เพิ่มไฟล์แนบเข้าไปใน message
            message.attach(file_part)

        except FileNotFoundError:
            print(f"Attachment file not found: {attachment_filename}. Sending email without attachment.")
        except Exception as e:
            print(f"An error occurred while attaching the file: {e}")

        # --- สิ้นสุดส่วนของการแนบไฟล์ ---

        server.sendmail(
                sender_email, recipient_email, message.as_string()
        )

        print("Sent to:", recipient_name, " ", recipient_email)
        time.sleep(1)

print("End")