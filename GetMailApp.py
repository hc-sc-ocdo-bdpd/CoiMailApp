from pathlib import Path
import win32com.client
import re
import pandas as pd



#print("Enter the file path you wish for the output! (Copy and Paste directly)")
#filepath_input = input()
#filepath_input.encode('unicode_escape')
# create output folder
#output_dir = Path((filepath_input)) 

output_dir = Path(r'C:\Users\DZAIDI\Desktop\dummy')

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
RT_pattern = r'Request Title[:\s]+([\w\s-]+)(?=\s*2. Requestor:)'
name_pattern = r'Name[:\s*]+([\w\s]+)(?=\s*Telephone|$)'
phone_pattern = r'Telephone/Mobile:\s*([\d-]+)'
email_pattern = r'Email[:\s]+([\w.-]+@[\w.-]+)'
department_pattern = r'Department[:\s]+([\w\s]+)(?=\s*Branch)'
branch_pattern = r'Branch[:\s]+([\w\s]+)(?=\s*Directorate)'
directorate_pattern = r'/ Centre[:\s]+([\w\s]+)(?=\s*3. Request Type)'
request_type_pattern = r'Intranet Request Type[:\s]+([\w\s]+)(?=\s*Please )'
specify_request_pattern = r'Please specify[:\s]+(.*?)\n'
priority_pattern = r'Priority level[:\s]+\[([^\]]+)\]'
approval_pattern = r'Approved by:\s*([\w\s]+)(?=\s*Telephone|$)'
approval_phone_pattern = r'Telephone[:\s]+([\d-]+)'
postingdate_pattern = r'Posting date[:\s]+([\d-]+)'
TS_pattern = r'Time sensitive or tied to an event\?([\s\w-]+)(?=\s*7. Audience)'
audience_pattern = r'7\. Audience[\r\n]+Audience[:\s]+([\w\s]+)(?=\s*8. Proposed)'

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
    RT = ''
    name = ''
    phone = ''
    email = ''
    department = ''
    branch = ''
    directorate = ''
    request_type = ''
    specify_request = ''
    priority = ''
    approval = ''
    approval_phone = ''
    postingdate = ''
    TS = ''
    audience = ''
    
    # searching text
    RT_match = re.search(RT_pattern, body)
    name_match = re.search(name_pattern, body)
    phone_match = re.search(phone_pattern, body)
    email_match = re.search(email_pattern, body)
    department_match = re.search(department_pattern, body)
    branch_match = re.search(branch_pattern, body)
    directorate_match = re.search(directorate_pattern, body)
    request_type_match = re.search(request_type_pattern, body)
    specify_request_match = re.search(specify_request_pattern, body)
    priority_match = re.search(priority_pattern, body)
    approval_match = re.search(approval_pattern, body)
    approval_phone_match = re.search(approval_phone_pattern, body)
    postingdate_match = re.search(postingdate_pattern, body)
    TS_match = re.search(TS_pattern, body)
    audience_match = re.search(audience_pattern, body)

    # matched text to dict
    if RT_match:
        RT = RT_match.group(1).strip()
        update_dict["Request Title"] = RT

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

    if approval_match:
        approval = approval_match.group(1).strip()
        update_dict["Approved By"] = approval

    if approval_phone_match:
        approval_phone = approval_phone_match.group(1).strip()
        update_dict["Approved Telephone"] = approval_phone

    if postingdate_match:
        postingdate = postingdate_match.group(1).strip()
        update_dict["Posting Date"] = postingdate

    if TS_match:
        TS = TS_match.group(1).strip()
        update_dict["Time Sensitivity?"] = TS

    if audience_match:
        audience = audience_match.group(1).strip()
        update_dict["Audience"] = audience
    
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
df.drop(df.columns[0], axis=1)
df.columns = df.columns.astype(str)

# desired path for excel
excel_path = output_dir / "COI_Forms.xlsx"

writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")
df.to_excel(writer, sheet_name="sheet1", startrow=1, header=False, index=False)

Workbook = writer.book
worksheet = writer.sheets["sheet1"]

(max_row, max_col) = df.shape

column_settings = [{"header": column} for column in df.columns]

worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})

worksheet.set_column(0, max_col - 1, 12)

writer.close()


















