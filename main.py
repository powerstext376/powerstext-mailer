import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pytz
import time
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 1. SETUP & AUTHENTICATION
# ==========================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# 🚨 YAHAN APNI GOOGLE SHEET KA URL DAALEIN 🚨
SHEET_URL = "https://docs.google.com/spreadsheets/d/1zucz8OEsttq8a0g9wWIKk2wtur1nOLsnjrCHkguZLsI/edit?gid=0#gid=0"
sheet = client.open_by_url(SHEET_URL)

ws_accounts = sheet.worksheet("Accounts")
ws_templates = sheet.worksheet("Templates")
ws_leads = sheet.worksheet("Leads")

# ==========================================
# 2. TIME & WORKING HOURS CHECK
# ==========================================
IST = pytz.timezone('Asia/Kolkata')
current_time = datetime.now(IST)
today_date = current_time.date()

print(f"Current Time (IST): {current_time.strftime('%Y-%m-%d %H:%M:%S')}")

# Working Hours: Subah 8 AM se Shaam 8 PM (20:00) tak
if not (0 <= current_time.hour < 24):
    print("⏸️ Abhi working hours nahi hain (8 AM - 8 PM). System paused.")
    exit()

# ==========================================
# 3. ACCOUNTS FETCH & SMART SORTING
# ==========================================
all_accounts = ws_accounts.get_all_records()
active_accounts = []

for i, acc in enumerate(all_accounts, start=2):
    if str(acc.get('Status', '')).strip().lower() == 'active':
        acc['sheet_row'] = i
        active_accounts.append(acc)

if not active_accounts:
    print("❌ Error: Koi bhi Sender Account 'Active' nahi hai.")
    exit()

# Sort logic: Jisne sabse kam bheja hai wo top par aayega
def get_sent_count(account):
    count = str(account.get('Daily_Sent_Count', '')).strip()
    return int(count) if count.isdigit() else 0

active_accounts.sort(key=get_sent_count)
accounts_headers = ws_accounts.row_values(1)
count_col_index = accounts_headers.index('Daily_Sent_Count') + 1

# ==========================================
# 4. TEMPLATES FETCH
# ==========================================
templates_data = ws_templates.get_all_records()
templates = {}
for t in templates_data:
    t_level = str(t.get('Template_Level', '')).strip()
    templates[t_level] = {
        'Subject': str(t.get('Subject_Line', '')),
        'Body': str(t.get('Email_Body_HTML', ''))
    }

# ==========================================
# 5. LEADS FETCH & QUEUE BUILDING (INCLUDES FOLLOW-UP)
# ==========================================
leads_data = ws_leads.get_all_records()
priority_queue = []
normal_queue = []

for i, lead in enumerate(leads_data, start=2):
    status = str(lead.get('Email_Status', '')).strip()
    follow_up_str = str(lead.get('Follow_Up', '')).strip()
    
    lead['sheet_row'] = i
    
    if status.lower() == 'pending':
        normal_queue.append((lead, 'Intro'))
    elif status.lower() == 'in-progress' and follow_up_str:
        try:
            follow_up_date = datetime.strptime(follow_up_str, '%Y-%m-%d').date()
            if today_date >= follow_up_date:
                if str(lead.get('Clicked', '')).strip().lower() == 'yes':
                    priority_queue.append((lead, 'Path_Clicked'))
                elif str(lead.get('Opened', '')).strip().lower() == 'yes':
                    priority_queue.append((lead, 'Path_Opened'))
                else:
                    priority_queue.append((lead, 'Path_Unread'))
        except Exception:
            pass

# MAX UTILIZATION: Ek trigger me 400 mails (Takes approx 27 mins)
MAX_MAILS_PER_RUN = 400
sending_queue = (priority_queue + normal_queue)[:MAX_MAILS_PER_RUN]

if not sending_queue:
    print("✅ Aaj ke liye koi task pending nahi hai!")
    exit()

print(f"🚀 Total leads in queue for this run: {len(sending_queue)}")

# ==========================================
# 6. SENDING ENGINE (ROTATION & LIMITS)
# ==========================================
sender_index = 0

for lead_item in sending_queue:
    lead, template_key = lead_item
    target_email = str(lead.get('Client_Email', '')).strip()
    
    if not target_email:
        continue
        
    template = templates.get(template_key)
    if not template:
        continue
        
    # Account Check: Find an account that has sent less than 10 mails today
    current_sender = None
    attempts = 0
    while attempts < len(active_accounts):
        temp_sender = active_accounts[sender_index]
        sent_count = get_sent_count(temp_sender)
        
        if sent_count < 10:  # 🚨 10 MAILS PER ACCOUNT LIMIT 🚨
            current_sender = temp_sender
            break
        else:
            sender_index = (sender_index + 1) % len(active_accounts)
            attempts += 1
            
    if not current_sender:
        print("🛑 WARNING: Saare active accounts ki 10 mails/day ki limit poori ho gayi hai!")
        break
        
    sender_email = str(current_sender.get('Email_ID', '')).strip()
    sender_pass = str(current_sender.get('App_Password', '')).strip()
    
    # Auto-Detect Hostinger or Gmail Server
    if 'gmail.com' in sender_email.lower():
        smtp_host = 'smtp.gmail.com'
    else:
        smtp_host = 'smtp.hostinger.com'
        
    try:
        # Construct HTML Email
        msg = MIMEMultipart()
        msg['From'] = f"Powerstext Services <{sender_email}>"
        msg['To'] = target_email
        msg['Reply-To'] = "sales@powerstext.com"
        msg['Subject'] = template['Subject']
        msg.attach(MIMEText(template['Body'], 'html'))
        
        # Send Email via SMTP
        server = smtplib.SMTP_SSL(smtp_host, 465)
        server.login(sender_email, sender_pass)
        server.send_message(msg)
        server.quit()
        
        print(f"✅ Sent '{template_key}' to {target_email} via {sender_email}")
        
        # 1. Update Lead Status in Sheet
        if template_key == 'Intro':
            next_follow_up = (today_date + timedelta(days=2)).strftime('%Y-%m-%d')
            ws_leads.update_cell(lead['sheet_row'], 2, 'In-Progress') 
            ws_leads.update_cell(lead['sheet_row'], 3, next_follow_up) 
            ws_leads.update_cell(lead['sheet_row'], 4, today_date.strftime('%Y-%m-%d'))
        else:
            ws_leads.update_cell(lead['sheet_row'], 2, 'Completed')
            
        # 2. Update Sender's Daily Sent Count in Sheet
        new_count = get_sent_count(current_sender) + 1
        current_sender['Daily_Sent_Count'] = new_count
        ws_accounts.update_cell(current_sender['sheet_row'], count_col_index, new_count)
            
        # 3. Rotate to Next Sender Account for the next loop
        sender_index = (sender_index + 1) % len(active_accounts)
        
        # ⚡ FAST DELAY
        delay = random.randint(2, 4)
        print(f"⏳ Sleeping for {delay} seconds...")
        time.sleep(delay)
        
    except Exception as e:
        # Naya Debugger: Yeh batayega exact kyu login fail ho raha hai
        print(f"❌ FAIL -> Target: {target_email} | Sender: {sender_email} | Server: {smtp_host} | Pass Length: {len(sender_pass)}")
        print(f"Error Details: {str(e)}")

print("🎉 Run Completed Successfully! Batch Done.")
              
