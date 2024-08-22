# mailboxdownload
Download .msg files from Outlook, then extract .pdf files nested within. Deleted temp .msg file once complete and you are left with only the extracted .pdf file in your specified downloads folder.

# Dependencies
Python3, pip, extract_msg (pip install extract-msg)


This package contains 3 files: extract_pdf.py, extract_pdf_cleanup.py, and run_extract_pdf.ps1

To make this work, you will need to update the following lines in each file according to your local directories, email account, and email inbox/folder.

# extract_pdf.py
Line 10: Update the path where you want the extracted PDF to be stored.
Line 74: Update the email address of the account on your local Outlook you want to extract from.
Line 79: Update the name of the folder where you want to extract the .msg file from (if not the INBOX).

# extract_pdf_cleanup.py
Line 5: Update the path where you want the extracted PDF to be stored.

# run_extract_pdf.ps1
Line 2: Update the path where you want the extracted PDF to be stored.
Line 5: Update the path to the extract_pdf.py script on your local machine.
Line 6: Update the path to the extract_pdf_cleanup.py script on your local machine.
Line 19: Uncomment this line if you need to troubleshoot script output from running the shortcut.
Line 22: Update the path to where your run_extract_pdf.ps1 script will be and use this for the target location in the shortcut.
Line 25: Run this command once in a PowerShell terminal as Administrator. 
