import pandas as pd
import msal
import requests
from config import EMAIL_TEMPLATES, CLIENT_ID, CLIENT_SECRET, TENANT_ID, EMAIL_ADDRESS, CC_EMAIL, GRAPH_API_ENDPOINT
import re
import logging
import sys
import time
import tkinter as tk
from tkinter import filedialog
from validate_email import validate_email
import smtplib
import dns.resolver

# -------------------- Template and File Selection -------------------- #

# choose a template to send
def choose_template():
    if len(EMAIL_TEMPLATES) == 1:
        return next(iter(EMAIL_TEMPLATES.values()))
    
    print("\nChoose an email template from the following by entering the corresponding number.")
    
    # Display the templates with their corresponding indices
    for index, key in enumerate(EMAIL_TEMPLATES):
        print(f"{index}: {key}")
    
    # Get the user's choice
    template_choice = input("Template number: ")
    
    # Validate the choice
    while not template_choice.isdigit() or int(template_choice) not in range(len(EMAIL_TEMPLATES)):
        template_choice = input("Enter a valid template choice: ")
    
    # Get the selected template
    template = list(EMAIL_TEMPLATES.values())[int(template_choice)]
    template_name = list(EMAIL_TEMPLATES.keys())[int(template_choice)]
    print(f"You have selected {template_name}")
    return template
    
# a method to select a chosen excel file for automailing
def select_excel_file():
    root = tk.Tk()
    root.lift()
    root.focus_force()

    file_path = filedialog.askopenfilename (
        title = "Select an excel file",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    if file_path:
        print(f"File selected: {file_path}")
        return file_path
    else:
        print("No file selected")
        return None

# -------------------- Logging Configuration -------------------- #

# Configure logging to write INFO and higher level messages to 'email_logs.log'
logging.basicConfig(
    filename='email_logs.log', # Create a logs file to record errors or successes
    level=logging.DEBUG,  # Changed from INFO to DEBUG for more information
    format='%(asctime)s:%(levelname)s:%(message)s'
)

# -------------------- Function Definitions -------------------- #

def is_valid_email(email):
    """
    Validates the email address using a simple regex.
    Returns True if valid, False otherwise.
    """
    regex = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(regex, email) is not None

def acquire_token(client_id, client_secret, tenant_id):
    """
    Acquires an OAuth2 token using the Client Credentials flow.
    This token will be used to authenticate with the Microsoft Graph API.
    """
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    scopes = ["https://graph.microsoft.com/.default"]  # Scopes required by the app

    logging.info("Attempting to acquire OAuth2 token.") 
    result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        logging.info("OAuth2 token acquired successfully.")
        return result["access_token"]
    else:
        error_msg = f"Failed to acquire token: {result.get('error_description')}"
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)  # Exit the script if token acquisition fails

def send_email(access_token, recipient, subject, body):
    """
    Sends an email using the Microsoft Graph API.
    Parameters:
        - access_token: OAuth2 token for authentication
        - recipient: Recipient's email address
        - subject: Subject of the email
        - body: Body content of the email
    """
    # Construct the API endpoint for sending mail
    url = f"{GRAPH_API_ENDPOINT}/users/{EMAIL_ADDRESS}/sendMail"

    # Set up the request headers with the OAuth2 token
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Construct the email payload in JSON format
    email_msg = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",  # Use "HTML" for HTML emails
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                }
            ],
            "ccRecipients": [
                {
                    "emailAddress": {
                        "address": CC_EMAIL  # Static CC recipient
                    }
                }
            ]
        },
        "saveToSentItems": "true"  # true means that emails are saved in sent items
    }

    try:
        # Make a POST request to the Graph API to send the email
        response = requests.post(url, headers=headers, json=email_msg)

        if response.status_code == 202:
            # 202 Accepted indicates the email was accepted for delivery
            logging.info(f"Email successfully sent to {recipient}")
            print(f"Email successfully sent to {recipient}")
        else:
            # Log the error details for troubleshooting
            error_msg = f"Failed to send email to {recipient}: {response.status_code} {response.text}"
            logging.error(error_msg)
            print(error_msg)

    except requests.exceptions.RequestException as e:
        # Handle any network-related errors
        error_msg = f"An error occurred while sending email to {recipient}: {e}"
        logging.error(error_msg)
        print(error_msg)

# -------------------- Main Execution -------------------- #

def main():
    """
    Main function to read contacts from an Excel file and send emails to each contact.
    """
    # Step 1: Acquire OAuth2 Token
    access_token = acquire_token(CLIENT_ID, CLIENT_SECRET, TENANT_ID)

    # Step 2: Read Contacts from Excel
    try:
        excel_file = select_excel_file()
        contacts = pd.read_excel(excel_file)
        logging.info(f"Successfully read {len(contacts)} contacts from '{excel_file}'.")
    except FileNotFoundError:
        error_msg = f"The file '{excel_file}' was not found."
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)  # Exit the script if the Excel file is not found
    except Exception as e:
        error_msg = f"An error occurred while reading '{excel_file}': {e}"
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)  # Exit the script for any other read errors

    # Step 3 : Choose the email template
    email_template = choose_template()

    # Step 4: Iterate Through Contacts and Send Emails
    for index, contact in contacts.iterrows():
        to_email = contact.get('Email')
        name = contact.get('Name', 'Friend')  # Default name if missing
        company = contact.get('Company', 'Valued Partner')  # Default to 'Valued Partner' if 'company' is missing


        # Validate email address
        # if validate_email(to_email):
        #     warning_msg = f"Email address: {to_email} is invalid. Domain may be unreachable."
        #     logging.warning(warning_msg)
        #     print(warning_msg)
        #     continue

        if pd.isna(to_email):
            warning_msg = f"Row {index + 2}: Missing email address. Skipping."
            logging.warning(warning_msg)
            print(warning_msg)
            continue  # Skip to the next contact

        if not is_valid_email(to_email):
            warning_msg = f"Row {index + 2}: Invalid email address '{to_email}'. Skipping."
            logging.warning(warning_msg)
            print(warning_msg)
            continue  # Skip to the next contact

        # Format the email content using the plain text template
        try:
            plain_text_content = email_template.format(name=name, company=company)
        except KeyError as e:
            error_msg = f"Missing placeholder data {e} for email to {to_email}. Skipping."
            logging.error(error_msg)
            print(error_msg)
            continue  # Skip to the next contact if formatting fails
        except Exception as e:
            error_msg = f"Error formatting email for {to_email}: {e}. Skipping."
            logging.error(error_msg)
            print(error_msg)
            continue  # Skip to the next contact for any other formatting errors

        # Define the subject of the email
        if "{company}" in email_template:
            subject = f"{company} - Henrich"
        else:
            subject = f"{name} - Henrich"

        # Send the email using the Graph API
        send_email(access_token, to_email, subject, plain_text_content)

        #Throttle emails to comply with sending limits
        time.sleep(5)

if __name__ == "__main__":
    main()
