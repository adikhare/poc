import imaplib
import email
import traceback
from dateutil import parser
import pandas as pd

TODAY = '6b'


CLASS = {'6b': { 'email': "XYZ",
                 'password': "ABC"},
         '7b': { 'email': "XYZ",
                 'password': "ABC"}
        }

ORG_EMAIL = "@gmail.com" 
FROM_EMAIL = CLASS[TODAY]['email'] + ORG_EMAIL 
FROM_PWD = CLASS[TODAY]['password'] 
SMTP_SERVER = "imap.gmail.com" 
SMTP_PORT = 993
COLUMNS = ['Name', 'From', 'Sub', 'Date', 'Filename', 'Time', 'Attachment Type']


def write_to_excel(df):
    df.to_excel(TODAY + '.xlsx', engine='xlsxwriter', index=False)


def read_excel():
    try:
        existing_df = pd.read_excel(TODAY + '.xlsx', engine='openpyxl')
        existing_df = existing_df.fillna('None')
        return existing_df
    except Exception:
        return pd.DataFrame()


def last_row(df):
        return df.iloc[-1].to_list()


def list_formation(msg):
    email_from = msg['from']
    email_id = email_from.split('<')[1].rstrip('>')
    email_from_name = email_from[:-(len(email_id)+2)].strip()
    email_subject = 'None' if msg.get('subject', 'None') == '' else msg.get('subject', 'None')
    email_datetime = msg['Date']
    email_date = str(parser.parse(email_datetime, fuzzy=True).date())
    email_time = str(parser.parse(email_datetime, fuzzy=True).time())
    attachment = msg.get_payload()[1]
    ls = [email_from_name, email_id, email_subject, email_date, str(attachment.get_filename()), email_time, attachment.get_content_type()]
    return ls


def read_email_from_gmail():
    try:
        flag = True
        existing_df = read_excel()
        if not existing_df.empty:
            last_mail = existing_df.iloc[0].to_list()
        else:
            last_mail = []
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        data = mail.search(None, 'ALL')
        pd_list = []
        mail_ids = data[1]
        id_list = mail_ids[0].split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        for i in range(latest_email_id,first_email_id, -1):
            if not flag:
                break
            data = mail.fetch(str(i), '(RFC822)' )
            for response_part in data: 
                arr = response_part[0]
                if isinstance(arr, tuple):
                    msg = email.message_from_string(str(arr[1],'utf-8'))
                    ls = list_formation(msg)
                    if ls == last_mail:
                        flag = False
                        break
                    pd_list.append(ls)
        new_df = pd.DataFrame(pd_list, columns=COLUMNS)
        df = new_df.append(existing_df, ignore_index=True)
        write_to_excel(df)
    except Exception as e:
        traceback.print_exc() 
        print(str(e))

read_email_from_gmail()