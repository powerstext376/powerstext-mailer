import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pytz
import time
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import email.utils

# ==========================================
# 1. SETUP
# ==========================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# 🚨 APNI GOOGLE SHEET KA ASLI LINK YAHAN DAALEIN 🚨
SHEET_URL = "https://docs.google.com/spreadsheets/d/1zucz8OEsttq8a0g9wWIKk2wtur1nOLsnjrCHkguZLsI/edit?usp=drivesdk"
sheet = client.open_by_url(SHEET_URL)

ws_accounts = sheet.worksheet("Accounts")
ws_templates = sheet.worksheet("Templates")
ws_leads = sheet.worksheet("Leads")

# ==========================================
# 2. TIME CHECK
# ==========================================
IST = pytz.timezone('Asia/Kolkata')
current_time = datetime.now(IST)
today_date = current_time.date()

print(f"Current Time (IST): {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
if not (8 <= current_time.hour < 20):
    print("⏸️ Abhi working hours nahi hain (8 AM - 8 PM). System paused.")
    exit()

# ==========================================
# 3. ACCOUNTS SMART SORTING
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

def get_sent_count(account):
    count = str(account.get('Daily_Sent_Count', '')).strip()
    return int(count) if count.isdigit() else 0

active_accounts.sort(key=get_sent_count)
accounts_headers = ws_accounts.row_values(1)
count_col_index = accounts_headers.index('Daily_Sent_Count') + 1
status_col_index = accounts_headers.index('Status') + 1

# ==========================================
# 4. TEMPLATES
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
# 5. LEADS QUEUE (WITH SMART FOLLOW-UP)
# ==========================================
leads_data = ws_leads.get_all_records()
priority_queue = []
normal_queue = []

for i, lead in enumerate(leads_data, start=2):
    status = str(lead.get('Email_Status', '')).strip()
    follow_up_str = str(lead.get('Follow_Up_Level', lead.get('Follow_Up', ''))).strip()
    lead['sheet_row'] = i
    
    if status.lower() == 'pending':
        normal_queue.append((lead, 'Intro'))
    elif status.lower() == 'in-progress' and follow_up_str:
        follow_up_date = None
        date_formats = ['%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d']
        
        for fmt in date_formats:
            try:
                follow_up_date = datetime.strptime(follow_up_str, fmt).date()
                break
            except ValueError:
                continue
                
        if follow_up_date and today_date >= follow_up_date:
            clicked_val = str(lead.get('Clicked', '')).strip()
            opened_val = str(lead.get('Opened', '')).strip()
            
            if clicked_val != '' and clicked_val.lower() != 'no':
                priority_queue.append((lead, 'Path_Clicked'))
            elif opened_val != '' and opened_val.lower() != 'no':
                priority_queue.append((lead, 'Path_Opened'))
            else:
                priority_queue.append((lead, 'Path_Unread'))

MAX_MAILS_PER_RUN = 150
sending_queue = (priority_queue + normal_queue)[:MAX_MAILS_PER_RUN]

if not sending_queue:
    print("✅ Aaj ke liye koi task pending nahi hai!")
    exit()

print(f"🚀 Total leads in queue for this run: {len(sending_queue)}")

# ==========================================
# 6. ENGINE & DYNAMIC TRACKING
# ==========================================
sender_index = 0
TRACKING_BASE_URL = "https://powerstext.com/track.php"

for lead_item in sending_queue:
    if not active_accounts:
        print("🛑 WARNING: Saare active accounts is run ke liye fail ho chuke hain. Pausing until next batch.")
        break

    lead, template_key = lead_item
    target_email = str(lead.get('Client_Email', '')).strip()
    if not target_email: continue
        
    template = templates.get(template_key)
    if not template: 
        print(f"⚠️ Template '{template_key}' nahi mila Sheet me. Skipping {target_email}...")
        continue
        
    current_sender = None
    attempts = 0
    while attempts < len(active_accounts):
        temp_sender = active_accounts[sender_index]
        if get_sent_count(temp_sender) < 10:  
            current_sender = temp_sender
            break
        else:
            sender_index = (sender_index + 1) % len(active_accounts)
            attempts += 1
            
    if not current_sender:
        print("🛑 WARNING: Bachen hue saare accounts ki 10 mails limit poori ho gayi hai!")
        break
        
    sender_email = str(current_sender.get('Email_ID', '')).strip() 
    sender_pass = str(current_sender.get('App_Password', '')).strip()
    
    try:
        # DATA PRIVACY
        if '@' in target_email:
            name_part, domain_part = target_email.split('@')
            masked_target = name_part[:2] + "****@" + domain_part
        else:
            masked_target = "Hidden_Email"

        # TRACKING MAGIC
        custom_body = template['Body'].replace("{{EMAIL}}", target_email)
        cache_buster = random.randint(1000000, 9999999)
        open_pixel = f'<img src="{TRACKING_BASE_URL}?email={target_email}&action=open&rnd={cache_buster}" width="1" height="1" style="display:none;" />'
        final_body = custom_body + open_pixel

        msg = MIMEMultipart()
        msg['From'] = f"Powerstext Services <{sender_email}>"
        msg['To'] = target_email
        msg['Reply-To'] = "sales@powerstext.com" 
        msg['Subject'] = template['Subject']
        
        # HOSTINGER ANTI-SPAM HEADERS
        msg['Date'] = email.utils.formatdate(localtime=True)
        domain_name = sender_email.split('@')[1] if '@' in sender_email else 'powerstext.com'
        msg['Message-ID'] = email.utils.make_msgid(domain=domain_name)

        msg.attach(MIMEText(final_body, 'html'))
        
        # SMART SMTP CONNECTION
        if 'gmail.com' in sender_email.lower():
            smtp_host = 'smtp.gmail.com'
            server = smtplib.SMTP_SSL(smtp_host, 465)
            server.login(sender_email, sender_pass)
        else:
            smtp_host = 'smtp.hostinger.com'
            server = smtplib.SMTP(smtp_host, 587)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender_email, sender_pass)
            
        server.send_message(msg)
        server.quit()
        
        print(f"✅ Sent '{template_key}' to {masked_target} via {sender_email}")
        
        # 🚨 THE FIX: Saving date for BOTH Intro and Follow-up
        if template_key == 'Intro':
            next_follow_up = (today_date + timedelta(days=1)).strftime('%Y-%m-%d')
            ws_leads.update_cell(lead['sheet_row'], 3, next_follow_up) 
            ws_leads.update_cell(lead['sheet_row'], 2, 'In-Progress')  
            ws_leads.update_cell(lead['sheet_row'], 4, today_date.strftime('%Y-%m-%d')) 
        else:
            ws_leads.update_cell(lead['sheet_row'], 2, 'Completed')
            ws_leads.update_cell(lead['sheet_row'], 4, today_date.strftime('%Y-%m-%d'))
            
        new_count = get_sent_count(current_sender) + 1
        current_sender['Daily_Sent_Count'] = new_count
        ws_accounts.update_cell(current_sender['sheet_row'], count_col_index, new_count)
            
        sender_index = (sender_index + 1) % len(active_accounts)
        
        delay = random.randint(5, 8)
        print(f"⏳ Sleeping for {delay} seconds...")
        time.sleep(delay)
        
    except Exception as e:
        error_msg = str(e)
        host_display = 'smtp.gmail.com' if 'gmail.com' in sender_email.lower() else 'smtp.hostinger.com'
        print(f"❌ FAIL -> Target: {masked_target} | Sender: {sender_email} | Server: {host_display}")
        print(f"Error Details: {error_msg}")
        
        if "535" in error_msg or "534" in error_msg or "auth" in error_msg.lower():
            print(f"⚠️ Account {sender_email} Auth Failed. Marking as 'Inactive' in sheet...")
            try:
                ws_accounts.update_cell(current_sender['sheet_row'], status_col_index, 'Inactive')
            except Exception:
                pass
        else:
            print(f"⚠️ Temporary limit for {sender_email}. Skipping for this run only.")
            
        try:
            if current_sender in active_accounts:
                active_accounts.remove(current_sender)
            if len(active_accounts) > 0:
                sender_index = sender_index % len(active_accounts)
        except Exception:
            pass

print("🎉 Run Completed Successfully! Batch Done.")
        
