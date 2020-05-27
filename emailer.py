#! python3
# Sends email to given address.

import sys
import win32com.client as win32
from datetime import datetime

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = sys.argv[1]
mail.Subject = sys.argv[2]
mail.Body = f"Email has been sent on {datetime.now().date()} at {datetime.now().time()}"
# mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

# To attach a file to the email (optional):
# attachment  = "Path to the attachment"
# mail.Attachments.Add(attachment)

mail.Send()
print('Email succesfully sent')