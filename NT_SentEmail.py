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
import re

def extract_file_and_sheet(position_str):
    # เช่น "สายงาน ฐ_Form_Q_สฐฐ (FORM_INPUT)" -> ("สายงาน ฐ_Form_Q_สฐฐ", "FORM_INPUT")
    match = re.match(r"(.+?)\s*\((.+)\)", str(position_str).strip())
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return str(position_str).strip(), None

def reset_sent_status(csv_path):
    try:
        df = pd.read_csv(csv_path)
        if 'sent_status' in df.columns:
            df['sent_status'] = ''
            df.to_csv(csv_path, index=False)
            print('Reset sent_status complete.')
    except Exception as e:
        print(f'Error resetting sent_status: {e}')

# --- Reset sent_status ทุกครั้งที่รัน ---
reset_sent_status('position_.csv')

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

# --- อ่านข้อมูลจากไฟล์ position_.csv ---
try:
    df = pd.read_csv('position_.csv')
except FileNotFoundError:
    print("Error: 'position_.csv' not found. Please make sure the file is in the correct directory.")
    exit()

# --- เพิ่มคอลัมน์ sent_status ถ้ายังไม่มี ---
if 'sent_status' not in df.columns:
    df['sent_status'] = ''

# Create a secure SSL context
context = ssl.create_default_context()

print(f"Connecting to server at {mail_server}:{port}...")
with smtplib.SMTP_SSL(mail_server, port, context=context) as server:
    print("Logging in...")
    server.login(sender_email, password)
    print("Login successful.")

    for idx, row in df.iterrows():
        print(f"row {idx}: name={row.get('name')}, email={row.get('email')}, position={row.get('position')}, sent_status={row.get('sent_status', '')}")
        # ข้ามถ้าส่งแล้ว
        if str(row.get('sent_status', '')).strip().upper() == 'SENT':
            continue
        recipient_name = row.get('name', '')
        recipient_email = row.get('email', '')
        if pd.isna(recipient_email) or pd.isna(recipient_name) or not recipient_email or not recipient_name:
            print(f"Skipping row {idx+2} due to empty name or email.")
            continue
        # --- แนบไฟล์ตามฝ่ายและ sheet ---
        position = str(row.get('position', '')).strip()
        base_name, sheet_name = extract_file_and_sheet(position)
        csv_file = f"{base_name}.csv"
        xlsx_file = f"{base_name}.xlsx"
        pdf_file = f"{base_name}.pdf"
        attachments = []
        if os.path.isfile(csv_file):
            attachments.append(csv_file)
        if os.path.isfile(xlsx_file):
            if sheet_name:
                try:
                    temp_attachment = f"temp_{idx}_{sheet_name}.xlsx"
                    with pd.ExcelWriter(temp_attachment, engine='openpyxl') as writer:
                        pd.read_excel(xlsx_file, sheet_name=sheet_name).to_excel(writer, index=False, sheet_name=sheet_name)
                    attachments.append(temp_attachment)
                except Exception as e:
                    print(f"Error extracting sheet {sheet_name} from {xlsx_file}: {e}")
            else:
                attachments.append(xlsx_file)
        if os.path.isfile(pdf_file):
            attachments.append(pdf_file)
        if not attachments:
            print(f"Attachment file not found for position: {position}. Sending email without attachment.")
        # --- สร้าง message ---
        message = MIMEMultipart("mixed")
        message["Subject"] = subject
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = formataddr((recipient_name, recipient_email))
        body_part = MIMEMultipart("alternative")
        part1 = MIMEText(f"""
เรียนคุณ {recipient_name},

ขอเรียนแจ้งให้ท่านทราบว่า ทางบริษัทได้จัดส่งเอกสารสำคัญแนบมาพร้อมกับอีเมลฉบับนี้ กรุณาตรวจสอบไฟล์แนบ หากมีข้อสงสัยหรือปัญหาใด ๆ สามารถติดต่อกลับได้ทันที

ขอแสดงความนับถือ
ฝ่ายเทคโนโลยีสารสนเทศ
บริษัท NT2025
""", "plain")
        part2 = MIMEText(f"""
<html>
  <body>
    <p>เรียนคุณ <b>{recipient_name}</b>,</p>
    <p>ขอเรียนแจ้งให้ท่านทราบว่า ทางบริษัทได้จัดส่งเอกสารสำคัญแนบมาพร้อมกับอีเมลฉบับนี้ กรุณาตรวจสอบไฟล์แนบ<br/>
    หากมีข้อสงสัยหรือปัญหาใด ๆ สามารถติดต่อกลับได้ทันที</p>
    <p>ขอแสดงความนับถือ<br/>
    ฝ่ายเทคโนโลยีสารสนเทศ<br/>
    บริษัท NT2025</p>
  </body>
</html>
""", "html")
        body_part.attach(part1)
        body_part.attach(part2)
        message.attach(body_part)
        # --- แนบไฟล์ทุกประเภท ---
        for attachment_filename in attachments:
            try:
                with open(attachment_filename, "rb") as attachment:
                    file_part = MIMEBase("application", "octet-stream")
                    file_part.set_payload(attachment.read())
                encoders.encode_base64(file_part)
                file_part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(attachment_filename)}",
                )
                message.attach(file_part)
            except Exception as e:
                print(f"An error occurred while attaching the file: {e}")
        # --- ส่งอีเมล ---
        try:
            server.sendmail(sender_email, recipient_email, message.as_string())
            print("Sent to:", recipient_name, recipient_email)
            df.at[idx, 'sent_status'] = 'yes'
        except Exception as e:
            print(f"Failed to send to {recipient_name} {recipient_email}: {e}")
            df.at[idx, 'sent_status'] = 'no'
        time.sleep(1)
        # ลบไฟล์ temp ถ้ามี
        if sheet_name and os.path.isfile(f"temp_{idx}_{sheet_name}.xlsx"):
            try:
                os.remove(f"temp_{idx}_{sheet_name}.xlsx")
            except Exception as e:
                print(f"Error removing temp file temp_{idx}_{sheet_name}.xlsx: {e}")

# --- บันทึกสถานะกลับลงไฟล์ ---
df.to_csv('position_.csv', index=False)

print("End")