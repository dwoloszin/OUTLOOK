import win32com.client
import pandas as pd
import os
import sys
from datetime import timezone,datetime
import ShortName
import ImportDF
import time
import re
from getFileFromWeb import download_csv_files
from getFileFromFolder import copyFiles
import timeit
import util
credentials = util.getCredentials()



print ('\nprocessing... ')
inicio = timeit.default_timer()


timeexport = time.strftime("%Y%m%d_")
script_dir = os.path.abspath(os.path.dirname(sys.argv[0]) or '.')
csv_path = os.path.join(script_dir, 'export/'+timeexport+'OutLook'+'.csv')
attachment_dir = os.path.join(script_dir, 'attachments')
attachment_dir = credentials['Folder_to_save'] # this code has some problem about encoding especial charactere
attachment_dir = '//internal/FileServer/TBR/Network/NetworkAssurance/RegionalNetworkAssurance-SP/RSQA/10.Organização da rede/_EQUIPE PROJETOS_/TSSR'

def process_email_body(body, num_lines=8):
    # Split the body text into lines based on newline character
    lines = body.split('\n')

    # Remove any leading/trailing whitespace from each line and filter out blank lines
    lines = [line.strip() for line in lines if line.strip()]

    # Return the desired number of lines
    return '\n'.join(lines[:num_lines])



def save_attachments(mail_item, emails):
    attachments = mail_item.Attachments
    
    for attachment in attachments:
        if attachment.FileName.startswith("TSSR_"):
            attachment_path = os.path.join(attachment_dir, attachment.FileName)
            # Save TSSR if necessary
   
            if not os.path.exists(attachment_path):  # Check if file doesn't exist
                attachment.SaveAsFile(attachment_path)
                print(f"Saving TSSR files: {attachment.FileName}")
            # Add attachment name to the emails list
            email_details = {
                'Subject': mail_item.Subject,
                'Sender': mail_item.SenderName,
                'ReceivedTime': mail_item.ReceivedTime,
                'AttachmentName': attachment.FileName,
                'Body': process_email_body(mail_item.Body)   # Include email body
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
            #Save att if necessary
            save_attachments(item, emails)  # Save attachments if available

    df = pd.DataFrame(emails)  # Convert list of dictionaries to DataFrame


    df['SiteName'] = df['Subject'].str.extract(r'(\w+-\w+)', expand=False)
    df['OC'] = df['Subject'].str.extract(r'(\d+)$')

    df = ShortName.tratarShortNumber(df, 'SiteName')
    df.sort_values(by='ReceivedTime', ascending=False, inplace=True)

    fields = ['NAME', 'LOCATION', 'ANF', 'MUNICIPIO']
    pathImport = '\import\MobileSite'
    MobileSite = ImportDF.ImportDF3(fields, pathImport,0)
    MobileSite = ShortName.tratarShortNumber(MobileSite, 'NAME')
    MobileSite.drop(['NAME'], axis=1, inplace=True)

    MobileSite.rename(columns={'SHORT': 'SHORT2'}, inplace=True)
    df = pd.merge(df, MobileSite, how='left', left_on=['SHORT'], right_on=['SHORT2'])
    df.drop(['SHORT2'], axis=1, inplace=True)
    df = df.loc[~df['AttachmentName'].isna()]
    df["Rev_Archive"] = df['AttachmentName'].str[-5:-4]
    
    pathImportAnalise = '\import\Analise'
    columnsAnalise = ['Analise','Status','Site_ID','Rev RF','OBS RF']
    Analise = ImportDF.ImportDF_Xlsx(pathImportAnalise,columnsAnalise)
    Analise['Analise'] = pd.to_datetime(Analise['Analise'], format="%Y-%m-%d")
    Analise['Status'] = Analise['Status'].str.upper()
    print(Analise)

    Analise.sort_values(by='Analise', ascending=False, inplace=True)
    subnetcheck = ['Site_ID']
    Analise.drop_duplicates(subset=subnetcheck, keep='first', inplace=True, ignore_index=False)
    df = pd.merge(df, Analise, how='left', left_on=['SiteName'], right_on=['Site_ID'])
    df.drop(['Site_ID'], axis=1, inplace=True)

    df.loc[(~df['Rev RF'].isna()) & (df['Rev RF'] != df['Rev_Archive']),['OBS']] = 'RevArchiveDiffAnalise'


    pathImportPRIO = '\import\PRIO'
    columnsPRIO = ['SiteName','MOS','STEP']
    PRIO = ImportDF.ImportDF_Xlsx(pathImportPRIO,columnsPRIO)
    PRIO['MOS'] = pd.to_datetime(PRIO['MOS'], format="%Y-%m-%d")
    PRIO.sort_values(by='MOS', ascending=False, inplace=True)
    subnetcheck = ['SiteName']
    PRIO.drop_duplicates(subset=subnetcheck, keep='first', inplace=True, ignore_index=False)
    PRIO.rename(columns={'SiteName': 'SiteName_PRIO'}, inplace=True)



    fields_PMO = ['ORDEM_COMPLEXA', 'STATUS_OC']
    pathImportPMO = '\import/reports'
    PMO = ImportDF.ImportDF3(fields_PMO, pathImportPMO,1)

    #PMO.rename(columns={'SHORT': 'SHORT2'}, inplace=True)
    df = pd.merge(df, PMO, how='left', left_on=['OC'], right_on=['ORDEM_COMPLEXA'])
    df.drop(['ORDEM_COMPLEXA'], axis=1, inplace=True)


   
    #include all PRIO
    #df = pd.merge(df, PRIO, how='left', left_on=['SiteName'], right_on=['SiteName_PRIO'])
    df = pd.merge(df, PRIO, how='outer', left_on=['SiteName'], right_on=['SiteName_PRIO'])



    df.sort_values(by=['SiteName','Rev_Archive'], ascending=[True,False], inplace=True)
    df.drop_duplicates(subset=['SiteName'], keep='first', inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.sort_values(by=['MOS','ReceivedTime'], ascending=[True,True], inplace=True)

    df.loc[df['Status'].isna(),['Status']] = 'SEM 1° ANALISE'

    df.drop_duplicates(inplace=True, ignore_index=True)
    df.to_csv(csv_path, sep=';',encoding='ANSI', index=False)  # Save DataFrame as CSV

    return df



# Create the attachment directory if it doesn't exist
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)






if __name__ == "__main__":
  folder_path = "reports/"
  prefix = "radar_pmo_"
  save_path = "import/reports"  # Change this to your desired save path
  download_csv_files(folder_path,prefix,save_path)
  copyFiles()

  print(get_outlook_emails())
  fim = timeit.default_timer()
  print ('duracao: %.2f' % ((fim - inicio)/60) + ' min') 

