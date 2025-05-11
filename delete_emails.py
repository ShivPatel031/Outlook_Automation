import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token,MS_GRAPH_BASE_URL
from Outlook import search_folder,get_message,delete_message,search_messages



def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        messages = search_messages(headers,"New app(s) connected to your Microsoft account")

        for message in messages:
            is_deleted = delete_message(headers,message['id'])
            if is_deleted:
                print(f'Deleted: {message['subject']}')
                print('-'*150)

    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()