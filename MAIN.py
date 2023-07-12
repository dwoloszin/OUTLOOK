import win32com.client
import pandas as pd
import os
import sys
from datetime import timezone
import ShortName
import ImportDF
import time
timeexport = time.strftime("%Y%m%d_")
script_dir = os.path.abspath(os.path.dirname(sys.argv[0]) or '.')
csv_path = os.path.join(script_dir, 'export/'+timeexport+'OutLook'+'.csv')
attachment_dir = os.path.join(script_dir, 'attachments')



def save_attachments(mail_item, emails):
    attachments = mail_item.Attachments
    for attachment in attachments:
        if attachment.FileName.startswith("TSSR_"):
            attachment_path = os.path.join(attachment_dir, attachment.FileName)
            if not os.path.exists(attachment_path):  # Check if file doesn't exist
                attachment.SaveAsFile(attachment_path)
            
            # Add attachment name to the emails list
            email_details = {
                'Subject': mail_item.Subject,
                'Sender': mail_item.SenderName,
                'ReceivedTime': mail_item.ReceivedTime,
                'AttachmentName': attachment.FileName
            }
            emails.append(email_details)


def get_outlook_emails():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Retrieve the "TSSR" folder within the Inbox
    inbox = namespace.GetDefaultFolder(6)
    tssr_folder = inbox.Folders.Item("TSSR")

    # Get all items in the folder
    items = tssr_folder.Items

    # Retrieve emails from the "TSSR" folder
    emails = []
    for item in items:
        if item.Class == 43 or item.Class == 5:  # Check if the item is a MailItem or SentItem
            # Extract email details
            email_details = {
                'Subject': item.Subject,
                'Sender': item.SenderName,
                'ReceivedTime': item.ReceivedTime.astimezone(timezone.utc),
            }
            emails.append(email_details)

            save_attachments(item, emails)  # Save attachments if available

    df = pd.DataFrame(emails)  # Convert list of dictionaries to DataFrame


    df['SiteName'] = df['Subject'].str.extract(r'(\w+-\w+)', expand=False)
    df['OC'] = df['Subject'].str.extract(r'(\d+)$')

    df = ShortName.tratarShortNumber(df, 'SiteName')
    df.sort_values(by='ReceivedTime', ascending=False, inplace=True)

    fields = ['NAME', 'LOCATION', 'ANF', 'MUNICIPIO']
    pathImport = '\import\MobileSite'
    MobileSite = ImportDF.ImportDF3(fields, pathImport)
    MobileSite = ShortName.tratarShortNumber(MobileSite, 'NAME')
    MobileSite.drop(['NAME'], axis=1, inplace=True)

    MobileSite.rename(columns={'SHORT': 'SHORT2'}, inplace=True)
    df = pd.merge(df, MobileSite, how='left', left_on=['SHORT'], right_on=['SHORT2'])
    df.drop(['SHORT2'], axis=1, inplace=True)
    df = df.loc[~df['AttachmentName'].isna()]
    df["Rev_Archive"] = df['AttachmentName'].str[-5:-4]
    
    pathImportAnalise = '\import\Analise'
    columnsAnalise = ['Analise','Status','Elemento_ID (Site_ID)','Rev RF','OBS RF']
    Analise = ImportDF.ImportDF_Xlsx(pathImportAnalise,columnsAnalise)
    Analise['Analise'] = pd.to_datetime(Analise['Analise'], format="%Y-%m-%d")

    Analise.sort_values(by='Analise', ascending=False, inplace=True)
    subnetcheck = ['Elemento_ID (Site_ID)']
    Analise.drop_duplicates(subset=subnetcheck, keep='first', inplace=True, ignore_index=False)
    df = pd.merge(df, Analise, how='left', left_on=['SiteName'], right_on=['Elemento_ID (Site_ID)'])
    df.drop(['Elemento_ID (Site_ID)'], axis=1, inplace=True)

    df.loc[(~df['Rev RF'].isna()) & (df['Rev RF'] != df['Rev_Archive']),['OBS']] = 'RevArchiveDiffAnalise'

    df.sort_values(by=['SiteName','Rev_Archive'], ascending=[True,False], inplace=True)
    df.drop_duplicates(subset=['SiteName'], keep='first', inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.sort_values(by='ReceivedTime', ascending=False, inplace=True)



    df.drop_duplicates(inplace=True, ignore_index=True)
    df.to_csv(csv_path, sep=';',encoding='ANSI', index=False)  # Save DataFrame as CSV

    return df


# Create the attachment directory if it doesn't exist
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)

print(get_outlook_emails())