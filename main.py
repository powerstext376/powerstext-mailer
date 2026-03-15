import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import random
from datetime import datetime
import pytz
import re
import urllib.parse

print("🚀 Powerstext Smart AI Engine Starting...\n")

# ==========================================
# ⚙️ MAIN SETTINGS 
# ==========================================
SHEET_NAME = "Powerstext Mailer"  # 🚨 YAHAN APNI SHEET KA NAAM DAALEIN
REPLY_TO_EMAIL = "sales@powerstext.com"
FOLLOW_UP_GAP_DAYS = 2 
WEBHOOK_URL = "https://powerstext.com/track.php" # Aapka apna Hostinger Server
# ==========================================

IST = pytz.timezone('Asia/Kolkata')

def check_business_hours():
    current_time = datetime.now(IST)
    # 10 AM se 6 PM tak chalega (Abhi 10 baj chuke hain toh yeh chalega)
    if 10 <= current_time.hour < 18:
        return True
    return False

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

try:
    if not check_business_hours():
        print("⏸️ System Paused: Abhi Business Hours (10 AM - 6 PM) nahi hain.")
        exit()

    print("Connecting to Powerstext Database...")
    sheet = client.open(SHEET_NAME)
    
    accounts_tab = sheet.worksheet("Accounts")
    leads_tab = sheet.worksheet("Leads")
    templates_tab = sheet.worksheet("Templates")
    blacklist_tab = sheet.worksheet("Blacklist")

    all_accounts = accounts_tab.get_all_records()
    all_templates = templates_tab.get_all_records()
    blacklist_data = blacklist_tab.get_all_values()
    leads_data = leads_tab.get_all_values()

    blacklisted_emails = set([row[0].strip().lower() for row in blacklist_data if row])
    
    templates_dict = {}
    for t in all_templates:
        level = str(t['Template_Level']).strip()
        templates_dict[level] = {'subject': t['Subject_Line'], 'body': t['Email_Body_HTML']}

    active_accounts = []
    for i, acc in enumerate(all_accounts, start=2): 
        if str(acc.get('Status', '')).strip().lower() == 'active':
            acc['sheet_row'] = i
            active_accounts.append(acc)

    if not active_accounts:
        print("❌ Error: Koi bhi Sender Account 'Active' nahi hai.")
        exit()

    today_date = datetime.now(IST).date()
    priority_queue = [] 
    normal_queue = []   

    print("🔍 Scanning leads behavior and building queue...")
    
    for index, row in enumerate(leads_data[1:], start=2): 
        while len(row) < 6: row.append("") 
            
        email = row[0].strip()
        status = row[1].strip().lower()
        last_date_str = row[3].strip()
        
        has_opened = str(row[4]).strip() != ""
        has_clicked = str(row[5]).strip() != ""

        if email.lower() in blacklisted_emails:
            if status != 'blacklisted':
                leads_tab.update_cell(index, 2, 'Blacklisted')
            continue

        if status == 'completed':
            continue

        if status == 'in-progress':
            try:
                last_date = datetime.strptime(last_date_str, "%Y-%m-%d").date()
                days_passed = (today_date - last_date).days
                
                if days_passed >= FOLLOW_UP_GAP_DAYS:
                    if has_clicked:
                        dynamic_level = 'Path_Clicked'
                    elif has_opened:
                        dynamic_level = 'Path_Opened'
                    else:
                        dynamic_level = 'Path_Unread'
                        
                    priority_queue.append({"row_index": index, "email": email, "level": dynamic_level})
            except ValueError:
                pass 

        elif status == 'pending':
            normal_queue.append({"row_index": index, "email": email, "level": 'Intro'})

    sending_queue = priority_queue + normal_queue

    if not sending_queue:
        print("✅ Aaj ke liye koi task pending nahi hai!")
        exit()

    sender_index = 0

    for task in sending_queue:
        if not check_business_hours():
            print("\n⏰ 6:00 PM baj gaye. Business hours over. Stopping engine...")
            break

        client_email = task['email']
        row_index = task['row_index']
        current_level = task['level']

        if current_level not in templates_dict:
            current_level = 'Intro' 
            
        temp_data = templates_dict.get(current_level)
        if not temp_data:
            continue

        current_sender = active_accounts[sender_index]
        sender_email = current_sender['Email_ID']
        app_password = str(current_sender['App_Password']).replace(" ", "")
        provider = str(current_sender['Provider']).strip().lower()

        print(f"\n➡️ Sending [{current_level}] to: {client_email} via {sender_email}")

        safe_email = urllib.parse.quote(client_email)
        raw_body = temp_data['body']
        
        # 🚀 ANTI-CACHE OPEN TRACKING PIXEL (Gmail ko bypass karne ke liye) 🚀
        rand_num = random.randint(100000, 999999) 
        pixel_url = f"{WEBHOOK_URL}?email={safe_email}&action=open&nocache={rand_num}"
        pixel_html = f'<img src="{pixel_url}" width="1" height="1" style="display:none;" />'
        final_body = raw_body + pixel_html

        # Click Tracking Wrapper
        def wrap_link(match):
            original_url = match.group(1)
            safe_redirect = urllib.parse.quote(original_url, safe='')
            return f'href="{WEBHOOK_URL}?email={safe_email}&action=click&redirect={safe_redirect}"'
            
        final_body = re.sub(r'href="(https://wa\.me/[^"]+)"', wrap_link, final_body)

        msg = MIMEMultipart('alternative')
        msg['Subject'] = temp_data['subject']
        msg['From'] = f"Powerstext Service <{sender_email}>"
        msg['To'] = client_email
        msg.add_header('reply-to', REPLY_TO_EMAIL)
        msg.attach(MIMEText(final_body, 'html'))

        try:
            if provider == 'gmail':
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
            else:
                server = smtplib.SMTP_SSL('smtp.hostinger.com', 465)
                
            server.login(sender_email, app_password)
            server.sendmail(sender_email, client_email, msg.as_string())
            server.quit()
            
            print("✅ Mail Sent Successfully!")

            today_str = today_date.strftime("%Y-%m-%d")
            next_status = 'In-Progress' if current_level == 'Intro' else 'Completed'

            leads_tab.update_cell(row_index, 2, next_status)
            leads_tab.update_cell(row_index, 3, current_level)
            leads_tab.update_cell(row_index, 4, today_str)

            acc_row = current_sender['sheet_row']
            current_count = current_sender.get('Daily_Sent_Count', '')
            new_count = int(current_count) + 1 if str(current_count).strip().isdigit() else 1
            
            accounts_tab.update_cell(acc_row, 4, new_count)
            current_sender['Daily_Sent_Count'] = new_count 

            sender_index = (sender_index + 1) % len(active_accounts)

            delay = random.randint(5, 10)
            print(f"⚡ Fast Sleeping for {delay} seconds...")
            time.sleep(delay)

        except Exception as e:
            print(f"❌ Failed. Error: {e}")
            leads_tab.update_cell(row_index, 2, 'Failed')

    print("\n🎉 Engine Process Completed gracefully!")

except Exception as e:
    print("\n❌ SYSTEM CRASH ERROR:")
    print(e)
            
