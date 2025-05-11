import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import get_message,get_message_by_filter,get_attachments,search_folder,download_attachment 

def process_attachments(headers,message_id,dir_attachment):
    attachments = get_attachments(headers,message_id)
    for attachment in attachments:
        print('Name', attachment['name'])
        print('Size',f'{attachment['size']/1024:.2f} KB')
        print('Content Type:',attachment['contentType'])

        try:
            download_attachment(headers,message_id,attachment['id'],attachment['name'],dir_attachment)
        except httpx.HTTPStatusError as e:
            print(f"Failed to download {attachment['name']}: {e.response.status_code}")
        print()
        print("-"*150)
        print()



def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']

    dir_attachment = Path('./downloaded')
    dir_attachment.mkdir(parents=True,exist_ok=True)

    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        target_folder = "Inbox"
        folder = search_folder(headers,target_folder)
        folder_id = folder['id']

        # better for filtering emails individually
        messages = get_message(headers,folder_id=folder_id,top=2,max_results=2)

        # better for download attachments by folders
        # messages = get_message_by_filter(headers,filter='hasAttachments eq true',folder_id=folder_id,top=2,max_results=2)

        for message in messages:
            if message['hasAttachments']:
                print("Attachments:")
                print("Subject:",message['subject'])
                process_attachments(headers,message['id'],dir_attachment)
    
    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()