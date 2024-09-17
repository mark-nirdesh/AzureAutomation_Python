import logging
import msal
import requests
import openpyxl
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import json
import os
import azure.functions as func

# Use /tmp/ directory for Linux-based Azure Functions
TEMP_DIR = "/tmp/"

# Load configuration from config.json
def load_config():
    with open("config.json", "r") as f:
        return json.load(f)

# Replace these values with your actual Azure AD App Registration details
CLIENT_ID = "ad904b76-b7df-4081-a58a-16acdc85c4e9"  # Your Azure app's Client ID
AUTHORITY = "https://login.microsoftonline.com/common"  # 'common' is used for personal Microsoft accounts
SCOPES = ["Files.ReadWrite"]  # Include 'Files.ReadWrite' for OneDrive access
TOKEN_CACHE_FILE = "token_cache.bin"  # File to store the token cache

# MSAL Token Acquisition with Caching
def acquire_token():
    token_cache = msal.SerializableTokenCache()

    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            token_cache.deserialize(f.read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=token_cache)
    accounts = app.get_accounts()

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise ValueError("Failed to create device flow. Err: %s" % json.dumps(flow, indent=4))

    logging.info(f"User code: {flow['user_code']}")
    logging.info(f"Please visit {flow['verification_uri']} and enter the code to authenticate.")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(token_cache.serialize())
        return result["access_token"]
    else:
        return None

# Fetch Excel file from OneDrive Personal using Microsoft Graph API
def fetch_excel_from_onedrive(file_name):
    access_token = acquire_token()

    if access_token:
        headers = {
            "Authorization": f"Bearer {access_token}"
        }
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/Cerco%20time/Retail%20Responce/{file_name}:/content" # change the file path based on your onedrive directory
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            temp_file_path = os.path.join(TEMP_DIR, file_name)  # Save to /tmp/ directory
            with open(temp_file_path, 'wb') as temp_file:
                temp_file.write(response.content)
            logging.info(f"Excel file '{file_name}' downloaded successfully to {temp_file_path}")
            return temp_file_path
        else:
            logging.error(f"Failed to download Excel file. Status code: {response.status_code}")
            return None
    else:
        logging.error("Access token not available")
        return None

# Function to get the previous Monday and next Sunday based on the current date
def get_previous_monday_and_next_sunday():
    today = datetime.now()
    
    # Calculate the previous Monday
    prev_monday = today - timedelta(days=today.weekday())
    
    # Calculate the next Sunday
    next_sunday = prev_monday + timedelta(days=6)
    
    return prev_monday, next_sunday

# Modify the Excel file (update B22 to the previous Monday and M22 to the next Sunday)
def modify_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # Get the previous Monday and next Sunday based on the current date
    prev_monday, next_sunday = get_previous_monday_and_next_sunday()

    # Update the values in the Excel file
    sheet["B22"].value = prev_monday.strftime("%m.%d.%Y")  # US date format: mm.dd.yyyy
    sheet["M22"].value = next_sunday.strftime("%m.%d.%Y")  # US date format: mm.dd.yyyy

    # Save the updated file to the /tmp/ directory
    updated_file_path = os.path.join(TEMP_DIR, f"w-e_{prev_monday.strftime('%m.%d.%Y')}_Retail_Response_Timesheet.xlsx")
    wb.save(updated_file_path)

    logging.info(f"Excel file '{file_path}' modified and saved as '{updated_file_path}'")
    return updated_file_path, prev_monday, next_sunday

# Send email with the modified file as attachment
def send_email(attachment_path, b22_date, m22_date, config):
    from_address = "" # write your from email and password in config.json file
    to_address = "" # where you want to automate the email by sending it in interval
    subject = f"w-e {b22_date.strftime('%d.%m.%Y')}-{m22_date.strftime('%d.%m.%Y')} Retail Response Timesheet"
    
    body = f"""
    Hi there,

    I hope you are doing okay. Please find attached document for my timesheet submitted for {b22_date.strftime('%d.%m.%Y')}-{m22_date.strftime('%d.%m.%Y')} week.

    at Retail Response, Unit 19/20 Sandbeck Park, Sandbeck Lane, Wetherby

    Regards,
    Nirdesh
    """

    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {attachment_path.split('/')[-1]}")
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        logging.info("SMTP connection established")

        email_password = config['email_password']  # Read password from config file
        server.login(from_address, email_password)
        logging.info("Logged into email server")
        
        text = msg.as_string()
        server.sendmail(from_address, to_address, text)
        logging.info(f"Email sent to {to_address}")

        server.quit()
    except Exception as e:
        logging.error(f"Failed to send email: {e}")

# Send Matrix notification using Element.io API
def send_matrix_notification(config, message):
    try:
        url = f"{config['matrix_homeserver_url']}/_matrix/client/r0/rooms/{config['matrix_room_id']}/send/m.room.message?access_token={config['matrix_access_token']}"
        headers = {"Content-Type": "application/json"}
        data = {
            "msgtype": "m.text",
            "body": message
        }

        response = requests.post(url, headers=headers, json=data)
        if response.status_code != 200:
            logging.error(f"Failed to send Matrix notification: {response.status_code}, {response.text}")
        else:
            logging.info(f"Notification sent to Matrix room {config['matrix_room_id']} successfully.")
    except Exception as e:
        logging.error(f"Error while sending Matrix notification: {e}")

# Azure Function Timer Trigger
app = func.FunctionApp()

@app.schedule(schedule="0 30 16 * * 5", arg_name="mytimer", use_monitor=False)  # Every Friday at 4:30 PM
def cerco_timer_trigger(mytimer: func.TimerRequest) -> None:
    logging.info("Timer trigger function executed.")
    
    try:
        # Load the configuration from config.json
        config = load_config()
        logging.info("Configuration loaded successfully")

        # File name as confirmed in Graph Explorer for OneDrive Personal
        file_name = "w-e_02.09.2024_Retail_Response_Timesheet.xlsx"
        downloaded_file = fetch_excel_from_onedrive(file_name)

        if downloaded_file:
            logging.info(f"File downloaded: {downloaded_file}")
            # Modify the Excel file with updated dates
            modified_file, prev_monday, next_sunday = modify_excel(downloaded_file)

            # Send the email with the modified file
            send_email(modified_file, prev_monday, next_sunday, config)

            # Send a notification to the Matrix room
            message = f"Email for timesheet {prev_monday.strftime('%d.%m.%Y')} - {next_sunday.strftime('%d.%m.%Y')} has been sent."
            send_matrix_notification(config, message)

    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")
