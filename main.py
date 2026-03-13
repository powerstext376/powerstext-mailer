import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

print("Powerstext System Engine Starting...\n")

# ==========================================
# ⚙️ MAIN SETTINGS (Yahan Details Dalein)
# ==========================================
SHEET_NAME = "Powerstext Mailer"  # Apni Google Sheet ka exact naam
REPLY_TO_EMAIL = "sales@powerstext.com"  
# ==========================================

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

try:
    print("Connecting to Database...")
    sheet = client.open(SHEET_NAME)
    
    accounts_tab = sheet.worksheet("Sender_Accounts")
    leads_tab = sheet.worksheet("Target_Leads")
    templates_tab = sheet.worksheet("Content_Templates")

    # 1. Sirf 'Active' Accounts Ko Filter Karna
    all_accounts = accounts_tab.get_all_records()
    active_accounts = []
    for acc in all_accounts:
        if str(acc.get('Status', '')).strip().lower() == 'active':
            active_accounts.append(acc)

    if len(active_accounts) == 0:
        print("❌ Error: Aapki sheet me koi bhi account 'Active' nahi hai.")
        exit()

    print(f"✅ Total Active Sender Accounts found: {len(active_accounts)}\n")

    templates = templates_tab.get_all_records()
    template_subject = templates[0]['Subject_Line']
    template_body = templates[0]['Email_Body_HTML']
    
    leads_data = leads_tab.get_all_values()
    
    # 2. Rotation Index Set Karna
    sender_index = 0  

    for i in range(1, len(leads_data)): # Heading ko chhod kar
        client_email = leads_data[i][0]
        status = leads_data[i][1]
        
        if status.strip().lower() == 'pending':
            # Current account uthana rotation ke hisaab se
            current_sender = active_accounts[sender_index]
            sender_email = current_sender['Email_ID']
            app_password = str(current_sender['App_Password']).replace(" ", "")
            provider = str(current_sender['Provider']).strip().lower()

            print(f"➡️ Preparing to send to: {client_email} | Using ID: {sender_email}")
            
            # Email Draft karna
            msg = MIMEMultipart('alternative')
            msg['Subject'] = template_subject
            msg['From'] = f"Powerstext Service <{sender_email}>"
            msg['To'] = client_email
            msg.add_header('reply-to', REPLY_TO_EMAIL)
            
            part = MIMEText(template_body, 'html')
            msg.attach(part)
            
            try:
                # 3. SMTP Connection (Har mail ke liye naya taaki rotate ho sake)
                if provider == 'gmail':
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(sender_email, app_password)
                elif provider == 'hostinger':
                    server = smtplib.SMTP_SSL('smtp.hostinger.com', 465)
                    server.login(sender_email, app_password)
                
                # Mail bhejna aur connection close karna
                server.sendmail(sender_email, client_email, msg.as_string())
                server.quit() 
                
                print(f"✅ Mail successfully sent!")
                
                # Google Sheet Update Karna
                leads_tab.update_cell(i + 1, 2, 'Sent')
                
                # 4. Agle mail ke liye account change karna (The Rotation Logic)
                sender_index = (sender_index + 1) % len(active_accounts)
                
                # Delay spam se bachne ke liye
                time.sleep(2)
                print("-" * 40)
                
            except Exception as e:
                print(f"❌ Failed to send via {sender_email}. Error: {e}")
                leads_tab.update_cell(i + 1, 2, 'Failed')
                print("-" * 40)

    print("\n🎉 System task completed successfully!")

except Exception as e:
    print("\n❌ SYSTEM ERROR:")
    print(e)