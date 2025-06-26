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
import requests

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

    # --- เตรียม mapping URL โฟลเดอร์จาก Email_Folder.xlsx ---
    def get_url_mapping(email_xlsx_path):
        df_email = pd.read_excel(email_xlsx_path)
        mapping = {}
        for _, row in df_email.iterrows():
            folder_name = str(row['File_Name_Sheet']).strip()
            url = str(row['URL'])
            mapping[folder_name] = url
        return mapping

    def download_google_sheet_xlsx(sheet_url, save_path):
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
        if not match:
            print(f"Invalid Google Sheet URL: {sheet_url}")
            return False
        file_id = match.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        try:
            response = requests.get(export_url)
            if response.status_code == 200:
                with open(save_path, 'wb') as f:
                    f.write(response.content)
                print(f"Downloaded: {save_path}")
                return True
            else:
                print(f"Failed to download: {sheet_url}")
                return False
        except Exception as e:
            print(f"Error downloading: {sheet_url} : {e}")
            return False

    # --- เตรียม mapping URL ---
    url_mapping = get_url_mapping('Email_folder.xlsx')

    for idx, row in df.iterrows():
        print(f"row {idx}: name={row.get('name')}, email_nt={row.get('email_nt')}, gmail={row.get('gmail')}, foldername={row.get('foldername(ชื่อFolder)')}, sent_status={row.get('sent_status', '')}")
        # ข้ามถ้าส่งแล้ว
        if str(row.get('sent_status', '')).strip().upper() == 'SENT':
            continue
        recipient_name = row.get('name', '')
        email_nt = row.get('email_nt', '')
        gmail = row.get('gmail', '')
        recipient_emails = []
        if pd.notna(email_nt) and email_nt:
            recipient_emails.append(email_nt)
        if pd.notna(gmail) and gmail:
            recipient_emails.append(gmail)
        if not recipient_emails or pd.isna(recipient_name) or not recipient_name:
            print(f"Skipping row {idx+2} due to empty name or email.")
            continue
        # --- หา URL โฟลเดอร์จาก mapping ---
        folder_name = str(row.get('foldername(ชื่อFolder)', '')).strip()
        print(f"LOOKUP: {folder_name}")
        print(f"URL_MAPPING KEYS (ตัวอย่าง 5): {list(url_mapping.keys())[:5]}")
        folder_url = url_mapping.get(folder_name, None)
        print(f"FOUND URL: {folder_url}")
        if not folder_url:
            print(f"No folder URL found for {folder_name}")
            folder_url = ''
        # --- สร้าง message ---
        message = MIMEMultipart("mixed")
        message["Subject"] = subject
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = ", ".join([formataddr((recipient_name, e)) for e in recipient_emails])
        body_part = MIMEMultipart("alternative")
        part1 = MIMEText(f"""
เรียนคุณ {recipient_name},

ขอเรียนแจ้งให้ท่านทราบว่า ทางบริษัทได้จัดส่งเอกสารสำคัญผ่านโฟลเดอร์ออนไลน์ตามลิงก์ด้านล่างนี้ กรุณาคลิกเพื่อตรวจสอบข้อมูล:
{folder_url}

ขอแสดงความนับถือ
ฝ่ายบัญชีบริหาร
บริษัท NT2025
""", "plain")
        part2 = MIMEText(f"""
<html>
  <body>
    <p>เรียนคุณ <b>{recipient_name}</b>,</p>
    <p>ขอเรียนแจ้งให้ท่านทราบว่า ทางฝ่ายบัญชีบริหารได้จัดส่งเอกสารสำคัญผ่านโฟลเดอร์ออนไลน์ตามลิงก์ด้านล่างนี้ กรุณาคลิกเพื่อตรวจสอบข้อมูล:<br/>
    <a href='{folder_url}'>{folder_url}</a></p>
    <p>ขอแสดงความนับถือ<br/>
    ฝ่ายบัญชีบริหาร<br/>
    บริษัท NT2025</p>
  </body>
</html>
""", "html")
        body_part.attach(part1)
        body_part.attach(part2)
        message.attach(body_part)
        # --- ไม่แนบไฟล์ ---
        # --- ส่งอีเมล ---
        try:
            server.sendmail(sender_email, recipient_emails, message.as_string())
            print("Sent to:", recipient_name, recipient_emails)
            df.at[idx, 'sent_status'] = 'yes'
        except Exception as e:
            print(f"Failed to send to {recipient_name} {recipient_emails}: {e}")
            df.at[idx, 'sent_status'] = 'no'
        time.sleep(1)

# --- บันทึกสถานะกลับลงไฟล์ ---
df.to_csv('position_.csv', index=False)

print("End")
