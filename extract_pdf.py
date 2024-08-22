import os
import re
import shutil
import time
import win32com.client
import extract_msg  # You need to install this via pip (pip install extract-msg)
from win32com.client import Dispatch

# Define the folder where you want to save the extracted PDFs
save_folder = r"C:\Users\Public\Downloads"

# Create the folder if it doesn't exist
if not os.path.exists(save_folder):
    os.makedirs(save_folder)

# Function to sanitize filenames
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename)

# Function to save the .msg file
def save_msg_file(attachment, save_path):
    try:
        attachment.SaveAsFile(save_path)
        print(f"Saved MSG file: {save_path}")
        return True
    except Exception as e:
        print(f"Failed to save MSG file {attachment.FileName}: {e}")
        return False

# Recursive function to extract PDF from nested .msg files
def extract_pdf_from_msg(msg):
    try:
        if isinstance(msg, str):
            print(f"Attempting to extract PDF from: {msg}")
            msg = extract_msg.Message(msg)

        if not msg.attachments:
            print(f"No attachments found in {msg}")
            return False

        for att in msg.attachments:
            print(f"Processing attachment: {att}")
            if att is None:
                print(f"Encountered a NoneType attachment in {msg}, skipping.")
                continue

            print(f"Attachment type: {type(att)}")
            
            # Handle EmbeddedMsgAttachment (nested .msg files)
            if isinstance(att, extract_msg.attachments.emb_msg_att.EmbeddedMsgAttachment):
                nested_msg = att.data  # Get the nested Message object
                print(f"Processing nested MSG attachment.")
                return extract_pdf_from_msg(nested_msg)
            elif att.longFilename and att.longFilename.endswith(".pdf"):
                att.save()  # Save the file to the current directory
                current_dir = os.getcwd()  # Get current directory
                pdf_save_path = os.path.join(save_folder, sanitize_filename(att.longFilename))
                shutil.move(os.path.join(current_dir, att.longFilename), pdf_save_path)  # Move to desired location
                print(f"Extracted PDF: {pdf_save_path}")
                return True

        return False
    except Exception as e:
        print(f"Failed to extract PDF from {msg}: {e}")
        return False

# Connect to Outlook
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get all the accounts
accounts = outlook.Folders

# Specify the name of the account you want to target
account_name = "TEST@EXAMPLE.COM"

# Find the inbox for the specified account
for account in accounts:
    if account.Name == account_name:
        inbox = account.Folders("Inbox")
        break
else:
    print(f"Account '{account_name}' not found.")
    exit()

# Start the extraction process from unread emails in the inbox
for message in inbox.Items:
    if message.UnRead:  # Process only unread emails
        print(f"Processing unread email: {message.Subject}")
        for attachment in message.Attachments:
            if attachment.FileName.endswith(".msg"):
                sanitized_filename = sanitize_filename(attachment.FileName)
                msg_save_path = os.path.join(save_folder, sanitized_filename)
                if save_msg_file(attachment, msg_save_path):
                    if extract_pdf_from_msg(msg_save_path):
                        message.UnRead = False  # Mark the email as read
                        message.Save()  # Save changes to the email
