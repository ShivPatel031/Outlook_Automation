import os
import httpx
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import search_messages

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        search_query="patelshiv3123@gmail.com"
        messages=search_messages(headers,search_query)
        print()

        for indx,mail_message in enumerate(messages):
            print(mail_message)
            print(f'Email {indx+1}')
            print('Subject:',mail_message['subject'])
            print('To:',mail_message['toRecipients'])
            print('From:',mail_message['from']['emailAddress']['name'],f"({mail_message['from']['emailAddress']['address']})")
            print('Received Date Time:',mail_message['receivedDateTime'])
            print('Body Preview:',mail_message["bodyPreview"])
            print()
            print('-'*150)
            print()
    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()