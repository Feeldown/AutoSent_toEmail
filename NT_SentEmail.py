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

# --- เพิ่มคอลัมน์ subject_content ---
subject_content_value = 'เรื่อง ข้อมูลปริมาณการใช้งาน (Q) บริการราคาโอนระหว่างส่วนงาน (Transfer Price)'
df['subject_content'] = subject_content_value

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
        recipient_name = row.get('ส่วนงานผู้ใช้บริการ' '')
        to_email = row.get('Email ผู้ใช้บริการ (ระดับฝ่าย)', '')
        cc_email = row.get('Email ส่วนงานผู้ให้บริการ', '')
        folder_name = str(row.get('ส่วนงานผู้ใช้บริการ', '')).strip()
        folder_url = url_mapping.get(folder_name, '')
        if pd.isna(to_email) or not to_email:
            print(f"Skipping row {idx+2} due to empty To email.")
            continue
        # --- ดึง template ข้อความ ---
        mail_template = row.get('รูปแบบข้อความในเมล์ที่จัดส่งให้ฝ่ายผู้ใช้บริการ', '')
        # ลบข้อความ subject_content ออกจากเนื้อหา (ถ้ามี)
        subject_content = row.get('subject_content', '')
        if subject_content:
            import re
            # ลบ subject_content ออกจากทุกตำแหน่งในเนื้อหา
            mail_template = re.sub(re.escape(subject_content), '', str(mail_template))
            # ลบบรรทัดว่างที่เกิดจากการลบ subject_content
            mail_template = re.sub(r'(^|\n)[ \t]*\n', '\n', mail_template)
        # ลบ 'เรื่อง ...' และ 'เรียน ...' ที่ขึ้นต้นเนื้อหา (ถ้ามี)
        mail_template = re.sub(r'^(เรื่อง.*\n)?(เรียน.*\n)?', '', mail_template, flags=re.IGNORECASE)
        # ขึ้นต้นเนื้อหาด้วย 'เรียน <ชื่อผู้รับ>' โดยคงรูปแบบเว้นวรรคเดิม (ไม่มี ,)
        recipient_manager = str(row.get('ส่วนงานผู้ใช้บริการ', '')).strip()
        mail_template = f'เรียน {recipient_manager}\n' + mail_template.lstrip('\n').lstrip('\r')
        # ลบบรรทัดว่างถัดจาก 'เรียน <ชื่อผู้รับ>' 1 บรรทัด (ถ้ามี)
        mail_template = re.sub(r'^(เรียน [^\n]+)\n\s*\n', r'\1\n', mail_template)
        # แทนที่ 'ส่วนงานผู้ให้บริการ' และ 'ส่วนงานผู้ใช้บริการ' ด้วยค่าจาก column ที่เกี่ยวข้อง
        mail_template = mail_template.replace('ส่วนงานผู้ให้บริการ', str(row.get('ชื่อพนักงานบันทึกข้อมูล', '')).strip())
        mail_template = mail_template.replace('ส่วนงานผู้ใช้บริการ', str(row.get('ส่วนงานผู้ใช้บริการ', '')).strip())
        # --- เพิ่มลิงก์ folder_url ลงในเนื้อหา ---
        link_text_plain = f"\n\nตรวจสอบเอกสารได้ที่: {folder_url}\n"
        link_text_html = f"<br><br><a href='{folder_url}' style='font-size:16px;'>[คลิกเพื่อเปิดเอกสาร]</a><br>"
        # --- จัดรูปแบบข้อความลงท้าย (ชิดซ้าย) ---
        ending_plain = '\n\nจึงเรียนมาเพื่อโปรดพิจารณาดำเนินการ จะขอบคุณยิ่ง\nรบชง. โทร 02-5749831 , 02-5759565'
        ending_html = """
<br><br>
จึงเรียนมาเพื่อโปรดพิจารณาดำเนินการ จะขอบคุณยิ่ง<br>
รบชง. โทร 02-5749831 , 02-5759565
"""
        # ลบข้อความลงท้ายเดิม (ถ้ามี) เพื่อป้องกันซ้ำ
        mail_template = re.sub(r'จึงเรียนมาเพื่อโปรดพิจารณาดำเนินการ.*?รบชง\. โทร.*', '', mail_template, flags=re.DOTALL)
        mail_template = mail_template.rstrip('').rstrip('\r')
        # --- จัดรูปแบบย่อหน้า 'เนื่องด้วย' ---
        import re
        # plain text: เว้นบรรทัดและเพิ่มช่องว่างข้างหน้าคำว่า 'เนื่องด้วย'
        mail_template = re.sub(r'เนื่องด้วย', r'    เนื่องด้วย', mail_template)
        # html: เว้นบรรทัดและเพิ่มช่องว่างข้างหน้าคำว่า 'เนื่องด้วย'
        mail_template_html = re.sub(r'(<br>)*เนื่องด้วย', r'<br><br>&nbsp;&nbsp;&nbsp;&nbsp;เนื่องด้วย', mail_template.replace('\n', '<br>'))
        # --- เพิ่มลิงก์และข้อความลงท้าย ---
        mail_template_plain = mail_template + link_text_plain + ending_plain
        html_body = mail_template_html + link_text_html + ending_html
        # --- สร้าง message ---
        message = MIMEMultipart("mixed")
        message["Subject"] = row.get('subject_content', subject)  # ใช้ subject_content เป็น subject
        message["From"] = formataddr((sender_name, sender_email))
        message["To"] = to_email
        if pd.notna(cc_email) and cc_email:
            message["Cc"] = cc_email
            cc_list = [cc_email]
        else:
            cc_list = []
        body_part = MIMEMultipart("alternative")
        # plain
        part1 = MIMEText(mail_template_plain, "plain")
        # html
        part2 = MIMEText(f"""
<html>
  <body>
    {html_body}
  </body>
</html>
""", "html")
        body_part.attach(part1)
        body_part.attach(part2)
        message.attach(body_part)
        # --- แนบไฟล์ถ้ามี ---
        attachment_path = row.get('ไฟล์แนบ', '')
        if pd.notna(attachment_path) and attachment_path and os.path.isfile(attachment_path):
            with open(attachment_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
            message.attach(part)
        # --- ส่งอีเมล ---
        try:
            server.sendmail(sender_email, [to_email] + cc_list, message.as_string())
            print("Sent to:", recipient_name, to_email, "CC:", cc_list)
            df.at[idx, 'sent_status'] = 'SENT'
        except Exception as e:
            print(f"Failed to send to {recipient_name} {to_email}: {e}")
            df.at[idx, 'sent_status'] = 'no'
        time.sleep(1)

# --- บันทึกสถานะกลับลงไฟล์ ---
df.to_excel('ข้อมูลประกอบการส่งเมล์ Q_sent.xlsx', index=False)

print("End")
