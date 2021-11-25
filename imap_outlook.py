import imaplib, email, os , xlrd
import pandas as pd

user = 'Type your user id'
password = 'Type your email'
imap_url = 'maileu.mail.xxxxxx.xxx'
#Where you want your attachments to be saved (ensure this directory exists) 
attachment_dir = 'Enter directory for saving attachments '

# sets up the auth
def auth(user,password,imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user,password)
    return con
# extracts the body from the email
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
# allows you to download attachments
def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype()=='multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        df = pd.read_excel (f'C:/Users/P5535106/Downloads/files/data1.xlsx', sheet_name='Sheet1')
        print(df)

        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath,'wb') as f:
                f.write(part.get_payload(decode=True))



# location = f"C:/Users/P5535106/Downloads/files/{fileName}"
# a = xlrd.open_workbook(location)     
# sheet = wb.sheet_by_index(0)         
# sheet.cell_value(5, 5)
 
#search for a particular email
def search(key,value,con):
    result, data  = con.search(None,key,'"{}"'.format(value))
    return data
#extracts emails from byte array
def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs

con = auth(user,password,imap_url)
con.select('INBOX')

result, data = con.fetch(b'type inbox index number','(RFC822)')
raw = email.message_from_bytes(data[0][1])
get_attachments(raw)

