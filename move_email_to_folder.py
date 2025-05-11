import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import search_folder,get_folder,move_email_to_folder,get_message,search_messages,get_sub_folder

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        target_parent_folder_name = 'From Shiv'
        target_parent_folder = search_folder(headers,target_parent_folder_name)
        target_parent_folder_id = target_parent_folder['id']

        target_folder_name = 'Shiv personal'
        target_folders = get_sub_folder(headers,target_parent_folder_id)
        for folder in target_folders:
            if folder['displayName'] == target_folder_name:
                target_folder_id = folder['id']

        messages = search_messages(headers,"patelshiv3123") 

        for message in messages:
            parent_folder_id = message['parentFolderId']
            parent_folder = get_folder(headers,parent_folder_id)
            if parent_folder["displayName"]=='Inbox':
                print('Parent Folder:',parent_folder['displayName'])
                print('subject:',message['subject'])
                print('Received Date Time:',message['receivedDateTime'])

                message_id = message['id']
                status_email_moved = move_email_to_folder(headers,message_id,target_folder_id)
                print(f'Email "{status_email_moved['subject']}" moved to "{target_folder_name}" folder.')
                print()
                print("-"*150)



    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()