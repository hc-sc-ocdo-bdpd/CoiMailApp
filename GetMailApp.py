from pathlib import Path
import win32com.client
import re
import pandas as pd

print("Enter a file path!")
filepath_input = input()
filepath_input.encode('unicode_escape')
# create output folder
output_dir = Path((filepath_input)) 


# connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


# connect to inbox
inbox = outlook.GetDefaultFolder(6).Folders["COI"]
# Inbox folder reference = 6 -> https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders 

# get messages
messages = inbox.Items

# restricting emails by subject
messages = messages.Restrict("[Subject] = 'Conflict of Interest Forms'")

# initializing dictionary to store information
info_dict = {}


# regex recognition pattern
name_pattern = r'Name:\s*([\w\s]+)(?=\s*Telephone/Mobile:|$)'
phone_pattern = r'Telephone/Mobile:\s*([\d-]+)'
email_pattern = r'Email:\s*([\w.-]+@[\w.-]+)'
department_pattern = r'Department\s*([\w\s]+)'
branch_pattern = r'Branch\s*([\w\s]+)'
directorate_pattern = r'Directorate / Centre\s*([\w\s]+)'
request_type_pattern = r'Intranet Request Type:\s*([\w\s]+)'
specify_request_pattern = r'Please specify:\s*(.*?)\n'
priority_pattern = r'Priority level\s*\[([^\]]+)\]'


x=0
# Iterate over each item to collect info
for message in messages:
    #update dictionary
    update_dict = {}

    
    subject = message.Subject
    body = message.body
    attachments = message.Attachments
    # sent = message.sentDateTime

    update_dict["Subject"] = subject
    # info_dict["Date"] = sent

    # initializing variables to store information
    name = ''
    phone = ''
    email = ''
    department = ''
    branch = ''
    directorate = ''
    request_type = ''
    specify_request = ''
    priority = ''
    
    # searching text
    name_match = re.search(name_pattern, body)
    phone_match = re.search(phone_pattern, body)
    email_match = re.search(email_pattern, body)
    department_match = re.search(department_pattern, body)
    branch_match = re.search(branch_pattern, body)
    directorate_match = re.search(directorate_pattern, body)
    request_type_match = re.search(request_type_pattern, body)
    specify_request_match = re.search(specify_request_pattern, body)
    priority_match = re.search(priority_pattern, body)

    # matched text to dict
    if name_match:
        name = name_match.group(1).strip()
        update_dict["Name"] = name
    
    if phone_match:
        phone = phone_match.group(1).strip()
        update_dict["Phone Number"] = phone

    if email_match:
        email = email_match.group(1).strip()
        update_dict["Email"] = email

    if department_match:
        department = department_match.group(1).strip()
        update_dict["Department"] = department

    if branch_match:
        branch = branch_match.group(1).strip()
        update_dict["Branch"] = branch

    if directorate_match:
        directorate = directorate_match.group(1).strip()
        update_dict["Directorate"] = directorate

    if request_type_match:
        request_type = request_type_match.group(1).strip()
        update_dict["Request Type"] = request_type

    if specify_request_match:
        specify_request = specify_request_match.group(1).strip()
        update_dict["Specify Request"] = specify_request

    if priority_match:
        priority = priority_match.group(1).strip()
        update_dict["Priority"] = priority

    
    #updating full dict
    info_dict[x] = update_dict
    x=x+1

    # Make folder for each email
    target_folder = output_dir / str(name)
    target_folder.mkdir(parents=True, exist_ok=True)

    
    # Save Attachments
    for attachment in attachments:
        attachment.SaveAsFile(target_folder / str(attachment))

# dictionary to dataframe
df = pd.DataFrame.from_dict(data= info_dict)

# transpose for format
df = (df.T)


excel_path = output_dir / "COI_Forms.xlsx"
df.to_excel(excel_path)









