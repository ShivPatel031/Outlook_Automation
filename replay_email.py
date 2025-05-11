import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token,MS_GRAPH_BASE_URL
from Outlook import get_message,get_folder,reply_to_message,search_messages

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        messages = search_messages(headers,"mind.your.business")

        for message in messages:
            parent_folder = get_folder(headers,message['parentFolderId'])
            if parent_folder['displayName'] == 'Inbox':
                reply_body = "Thank you for the email."
                reply_to_message(headers,message['id'],reply_body)
                print(f'Replaied to the email "{message['subject']}"')
                print()
                print("-"*150)
                print()

    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()