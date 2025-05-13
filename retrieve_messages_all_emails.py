import os
import httpx
from dotenv import load_dotenv
from ms_graph import get_access_token,MS_GRAPH_BASE_URL
from tqdm import tqdm


def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']

    endpoint = f'{MS_GRAPH_BASE_URL}/me/messages'


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        for i in range(0,4,2):
            params = {
                '$top':i+2,
                '$select':'*',
                '$skip':i,
                '$orderBy':'receivedDateTime desc'
            }

            response = httpx.get(endpoint,headers=headers,params=params)

            if response.status_code != 200:
                raise Exception(f'Failed to retrieve emails: {response.text}')
            
            json_response = response.json()

            for mail_message in tqdm(json_response.get('value',[])):
                print(mail_message)
                print()
                if mail_message['isDraft']:
                    print('Subject:',mail_message['subject'])
                    print('To:',mail_message['toRecipients'])
                    print('Is Read:',mail_message['isRead'])
                    print('Received Date Time:',mail_message['receivedDateTime'])
                    print()
                else:
                    print('Subject:',mail_message['subject'])
                    print('To:',mail_message['toRecipients'])
                    print('From:',mail_message['from']['emailAddress']['name'],f"({mail_message['from']['emailAddress']['address']})")
                    print('Is Read:',mail_message['isRead'])
                    print('Received Date Time:',mail_message['receivedDateTime'])
                    print()
                print('-'*150)

    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()
