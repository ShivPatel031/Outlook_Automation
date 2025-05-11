import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token,MS_GRAPH_BASE_URL
from Outlook import create_attachment

def draft_message_body(subject,attachments):
    message = {
        'subject':subject,
        'body':{
            'contentType':'HTML',
            'content':'This is a test email sent from python.'
        },
        'toRecipients':[
            {
                'emailAddress':{
                    'address':'patelshiv3123@gmail.com'
                }
            },
            {
                'emailAddress':{
                    'address':'mind.your.business031@gmail.com'
                }
            }
        ],
        'ccRecipients':[
            {
                'emailAddress':{
                    'address':'210303105085@paruluniversity.ac.in'
                }
            }
        ],
        'importance':'high',
        'attachments':attachments
    }

    return message

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        dir_attachments = Path("./attachments")

        attachments = [create_attachment(attachment) for attachment in dir_attachments.iterdir() if attachment.is_file() ]

        endpoint = f'{MS_GRAPH_BASE_URL}/me/sendMail'

        message = {
            'message':draft_message_body('Test Email with Attachments',attachments),
        }

        response = httpx.post(endpoint,headers=headers,json=message)
        if response.status_code != 202:
            raise Exception(f'Failed to send email : {response.text}')
        
        print('Email sent successfully')

    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()