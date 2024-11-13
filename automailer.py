import pandas as pd
import msal
import requests
from config import CLIENT_ID, CLIENT_SECRET, TENANT_ID, EMAIL_ADDRESS, CC_EMAIL, GRAPH_API_ENDPOINT, XLSX_PATH, EMAIL_TEMPLATE
import re
import logging
import sys
import time

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
        contacts = pd.read_excel(XLSX_PATH)
        logging.info(f"Successfully read {len(contacts)} contacts from '{XLSX_PATH}'.")
    except FileNotFoundError:
        error_msg = f"The file '{XLSX_PATH}' was not found."
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)  # Exit the script if the Excel file is not found
    except Exception as e:
        error_msg = f"An error occurred while reading '{XLSX_PATH}': {e}"
        logging.error(error_msg)
        print(error_msg)
        sys.exit(1)  # Exit the script for any other read errors

    # Step 3: Iterate Through Contacts and Send Emails
    for index, contact in contacts.iterrows():
        to_email = contact.get('Email')
        name = contact.get('Name', 'Valued Customer')  # Default name if missing

        # Validate email address
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
            plain_text_content = EMAIL_TEMPLATE.format(name=name)
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
        subject = f"{name} - Henrich"

        # Send the email using the Graph API
        send_email(access_token, to_email, subject, plain_text_content)

        #Throttle emails to comply with sending limits
        time.sleep(.5)

if __name__ == "__main__":
    main()
