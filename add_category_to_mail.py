import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token,MS_GRAPH_BASE_URL
from Outlook import search_messages,add_category_to_mail


def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        messages = search_messages(headers,"patelshiv3123")

        for message in messages:
           is_add =  add_category_to_mail(headers,message['id'])
           if is_add:
               print(f'Orange category is add to mail {message['subject']}')

    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()