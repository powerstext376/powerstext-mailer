import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import random
from datetime import datetime, timedelta
import pytz # Nayi library time check ke liye

print("🚀 Powerstext Smart Engine Starting...\n")

# ==========================================
# ⚙️ MAIN SETTINGS 
# ==========================================
SHEET_NAME = "Powerstext Mailer"  # Apni Google Sheet ka exact naam daalein
REPLY_TO_EMAIL = "sales@powerstext.com"
FOLLOW_UP_GAP_DAYS = 3 # Kitne din baad follow-up bhejna hai
# ==========================================

# IST Timezone setup
IST = pytz.timezone('Asia/Kolkata')

def check_business_hours():
    """Check karta hai ki time 10 AM se 6 PM ke beech hai ya nahi"""
    current_time = datetime.now(IST)
    if 10 <= current_time.hour < 18:
        return True
    return False

# 1. Google Sheets Connection
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

try:
    if not check_business_hours():
        print("⏸️ System Paused: Abhi Business Hours (10 AM - 6 PM) nahi hain.")
        print("Engine agle din subah khud resume karega. Exiting...")
        exit()

    print("Connecting to Powerstext Database...")
    sheet = client.open(SHEET_NAME)
    
    accounts_tab = sheet.worksheet("Accounts")
    leads_tab = sheet.worksheet("Leads")
    templates_tab = sheet.worksheet("Templates")
    blacklist_tab = sheet.worksheet("Blacklist")

    # Data Fetching
    all_accounts = accounts_tab.get_all_records()
    all_templates = templates_tab.get_all_records()
    blacklist_data = blacklist_tab.get_all_values()
    leads_data = leads_tab.get_all_values()

    # 2. Blacklist & Template Setup
    blacklisted_emails = set([row[0].strip().lower() for row in blacklist_data if row])
    
    templates_dict = {}
    for t in all_templates:
        level = str(t['Template_Level']).strip()
        templates_dict[level] = {'subject': t['Subject_Line'], 'body': t['Email_Body_HTML']}

    # Active Accounts Filter
    active_accounts = [acc for acc in all_accounts if str(acc.get('Status', '')).strip().lower() == 'active']
    if not active_accounts:
        print("❌ Error: Koi bhi Sender Account 'Active' nahi hai.")
        exit()

    print(f"✅ Loaded {len(active_accounts)} Active Accounts & {len(templates_dict)} Templates.")

    # 3. Priority Sorting Logic (Follow-ups pehle, Naye Leads baad me)
    today_date = datetime.now(IST).date()
    priority_queue = [] # Follow-ups
    normal_queue = []   # Fresh Leads

    print("🔍 Scanning leads and building sending queue...")
    
    # Heading chhod kar read karna
    for index, row in enumerate(leads_data[1:], start=2): 
        # Safety check agar row choti ho
        while len(row) < 6: row.append("") 
            
        email = row[0].strip()
        status = row[1].strip().lower()
        level = row[2].strip()
        last_date_str = row[3].strip()

        if email.lower() in blacklisted_emails:
            if status != 'blacklisted':
                leads_tab.update_cell(index, 2, 'Blacklisted')
            continue

        if status == 'completed':
            continue

        # Priority 1: Follow-ups
        if status == 'in-progress':
            try:
                last_date = datetime.strptime(last_date_str, "%Y-%m-%d").date()
                days_passed = (today_date - last_date).days
                if days_passed >= FOLLOW_UP_GAP_DAYS:
                    priority_queue.append({"row_index": index, "email": email, "level": level})
            except ValueError:
                pass # Invalid date ignore karega

        # Priority 2: Fresh Leads
        elif status == 'pending':
            normal_queue.append({"row_index": index, "email": email, "level": 'Intro'})

    # Final Sequence: Pehle Follow-up, fir Naye leads
    sending_queue = priority_queue + normal_queue
    print(f"📊 Today's Target: {len(priority_queue)} Follow-ups | {len(normal_queue)} Fresh Leads")

    if not sending_queue:
        print("✅ Aaj ke liye koi task pending nahi hai!")
        exit()

    # 4. Main Sending Engine (The Loop)
    sender_index = 0

    for task in sending_queue:
        # Loop ke beech me 6 PM check
        if not check_business_hours():
            print("\n⏰ 6:00 PM baj gaye. Business hours over. Stopping engine...")
            break

        client_email = task['email']
        row_index = task['row_index']
        current_level = task['level']

        # Agar level ka template exist nahi karta, toh default 'Path_C' manega
        # (Kyunki jab tak Click/Open tracking shuru nahi hoti, sab Path_C (Cold) hain)
        if current_level not in templates_dict:
            current_level = 'Path_C' 
            
        temp_data = templates_dict.get(current_level)
        if not temp_data:
            print(f"⚠️ Template missing for level '{current_level}'. Skipping {client_email}.")
            continue

        # Account Rotation
        current_sender = active_accounts[sender_index]
        sender_email = current_sender['Email_ID']
        app_password = str(current_sender['App_Password']).replace(" ", "")
        provider = str(current_sender['Provider']).strip().lower()

        print(f"\n➡️ Sending [{current_level}] to: {client_email} | Via: {sender_email}")

        # Email Draft
        msg = MIMEMultipart('alternative')
        msg['Subject'] = temp_data['subject']
        msg['From'] = f"Powerstext Service <{sender_email}>"
        msg['To'] = client_email
        msg.add_header('reply-to', REPLY_TO_EMAIL)
        msg.attach(MIMEText(temp_data['body'], 'html'))

        try:
            # SMTP Setup
            if provider == 'gmail':
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
            else:
                server = smtplib.SMTP_SSL('smtp.hostinger.com', 465)
                
            server.login(sender_email, app_password)
            server.sendmail(sender_email, client_email, msg.as_string())
            server.quit()
            
            print("✅ Mail Sent Successfully!")

            # 5. Sheet Updates (State badalna)
            next_level = 'Path_C' if current_level == 'Intro' else 'Completed'
            next_status = 'In-Progress' if next_level != 'Completed' else 'Completed'
            today_str = today_date.strftime("%Y-%m-%d")

            # Update calls (Status, Level, Date)
            leads_tab.update_cell(row_index, 2, next_status)
            leads_tab.update_cell(row_index, 3, next_level)
            leads_tab.update_cell(row_index, 4, today_str)

            sender_index = (sender_index + 1) % len(active_accounts)

            # Random Anti-Spam Delay (5s to 10s)
            delay = random.randint(5, 10)
            print(f"⏳ Sleeping for {delay} seconds to act like a human...")
            time.sleep(delay)

        except Exception as e:
            print(f"❌ Failed. Error: {e}")
            leads_tab.update_cell(row_index, 2, 'Failed')

    print("\n🎉 Engine Process Completed / Stopped gracefully!")

except Exception as e:
    print("\n❌ SYSTEM CRASH ERROR:")
    print(e)

