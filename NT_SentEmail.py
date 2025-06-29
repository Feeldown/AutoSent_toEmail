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

def reset_sent_status_xlsx(xlsx_path):
    try:
        df = pd.read_excel(xlsx_path, header=1)
        if 'sent_status' in df.columns:
            df['sent_status'] = ''
            df.to_excel(xlsx_path, index=False)
            print('Reset sent_status complete.')
    except Exception as e:
        print(f'Error resetting sent_status: {e}')

# --- Reset sent_status ทุกครั้งที่รัน ---
# reset_sent_status('position_.csv')
reset_sent_status_xlsx('ข้อมูลประกอบการส่งเมล์ Q.xlsx')

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

# --- อ่านข้อมูลจากไฟล์ ข้อมูลประกอบการส่งเมล์ Q.xlsx ---
try:
    df = pd.read_excel('ข้อมูลประกอบการส่งเมล์ Q.xlsx', header=1)
except FileNotFoundError:
    print("Error: 'ข้อมูลประกอบการส่งเมล์ Q.xlsx' not found. Please make sure the file is in the correct directory.")
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
        df_email = pd.read_excel(email_xlsx_path, header=0)
        mapping = {}
        for _, row in df_email.iterrows():
            folder_name = str(row['ชื่อฝ่ายสายงาน']).strip()
            url = str(row['URL'])
            mapping[folder_name] = url
        return mapping

    url_mapping = get_url_mapping('Email_Folder.xlsx')

    for idx, row in df.iterrows():
        print(f"row {idx}: ฝ่ายผู้ใช้บริการ={row.get('ส่วนงานผู้ใช้บริการ')}, To={row.get('Email ผู้ใช้บริการ (ระดับฝ่าย)')}, CC={row.get('Email ส่วนงานผู้ให้บริการ')}, sent_status={row.get('sent_status', '')}")
        # ข้ามถ้าส่งแล้ว
        if str(row.get('sent_status', '')).strip().upper() == 'SENT':
            continue
        recipient_name = row.get('ส่วนงานผู้ใช้บริการ', '')
        to_email = row.get('Email ผู้ใช้บริการ (ระดับฝ่าย)', '')
        cc_email = row.get('Email ส่วนงานผู้ให้บริการ', '')
        folder_name = str(row.get('ส่วนงานผู้ใช้บริการ', '')).strip()
        folder_url = url_mapping.get(folder_name, '')
        if pd.isna(to_email) or not to_email:
            print(f"Skipping row {idx+2} due to empty To email.")
            continue
        # --- ดึง template ข้อความ ---
        mail_template = row.get('รูปแบบข้อความในเมล์ที่จัดส่งให้ฝ่ายผู้ใช้บริการ', '')
        if not mail_template or pd.isna(mail_template):
            mail_template = f"เรียน ผจก. {recipient_name},\n\nขอเรียนแจ้งให้ท่านทราบว่า ทางฝ่ายบัญชีบริหาร ({sender_name}) ได้จัดส่งเอกสารสำคัญผ่านโฟลเดอร์ออนไลน์ กรุณาคลิกเพื่อตรวจสอบข้อมูล:\n{folder_url}\n\nขอแสดงความนับถือ\nฝ่ายบัญชีบริหาร\nบริษัท NT2025"
        else:
            import re
            recipient_manager = 'ผจก. ' + str(row.get('ส่วนงานผู้ใช้บริการ', '')).strip()
            user_department = str(row.get('ส่วนงานผู้ใช้บริการ', '')).strip()
            mail_template = re.sub(
                r'เรียน\s*\(ส่วนงานผู้ใช้บริการ.*?\)',
                f'เรียน {recipient_manager}',
                mail_template
            )
            mail_template = mail_template.replace('ส่วนงานผู้ใช้บริการ', user_department)
            mail_template = mail_template.replace('ส่วนงานผู้ให้บริการ', str(row.get('ชื่อพนักงานบันทึกข้อมูล', '')).strip())
        # --- สร้าง message ---
        message = MIMEMultipart("mixed")
        message["Subject"] = subject
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = to_email
        if pd.notna(cc_email) and cc_email:
            message["Cc"] = cc_email
            cc_list = [cc_email]
        else:
            cc_list = []
        body_part = MIMEMultipart("alternative")
        # plain
        part1 = MIMEText(mail_template.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n'), "plain")
        # html (รวมบรรทัดว่างซ้อนที่มี space/tab ให้เหลือบรรทัดเดียว)
        html_body = re.sub(r'\n[ \t]*\n', '\n', mail_template)
        html_body = html_body.replace('\n', '<br>')
        part2 = MIMEText(f"""
<html>
  <body>
    {html_body}
    <br><a href='{folder_url}'>{folder_url}</a>
  </body>
</html>
""", "html")
        body_part.attach(part1)
        body_part.attach(part2)
        message.attach(body_part)
        # --- ส่งอีเมล ---
        try:
            server.sendmail(sender_email, [to_email] + cc_list, message.as_string())
            print("Sent to:", recipient_name, to_email, "CC:", cc_list)
            df.at[idx, 'sent_status'] = 'yes'
        except Exception as e:
            print(f"Failed to send to {recipient_name} {to_email}: {e}")
            df.at[idx, 'sent_status'] = 'no'
        time.sleep(1)

# --- บันทึกสถานะกลับลงไฟล์ ---
df.to_excel('ข้อมูลประกอบการส่งเมล์ Q_sent.xlsx', index=False)

print("End")
