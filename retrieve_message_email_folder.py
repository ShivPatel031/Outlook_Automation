import os
import httpx
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import search_folder,get_sub_folder,get_message

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        folder_name = 'Inbox'
        target_folder = search_folder(headers,folder_name)
        folder_id= target_folder['id']

        messages = get_message(headers)

        for message in messages:
            print('Subject:',message['subject'])
            print('-'*50)

        sub_folders = get_sub_folder(headers,folder_id)

        for sub_folder in sub_folders:
            if sub_folder['displayName'].lower() == 'sub folder':
                sub_folder_id = sub_folder['id']
                messages = get_message(headers,sub_folder_id)
                for message in messages:
                    print(f'Sub Folder Name: {sub_folder["displayName"]}')
                    print('Subject:'+message['subject'])
                    print('-'*50)

    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()