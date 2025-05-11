import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import create_folder,create_sub_folder

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']

    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        folder_name = "From Shiv"

        status,response = create_folder(headers,folder_name)

        if not status:
            print(f'Error creating folder "{response.json()}".')
            return
        
        folder_metadata = response.json()
        print(f'Folder "{folder_name}" created.')

        parent_folder_id = folder_metadata['id']
        sub_folder_names=['Shiv work','Shiv personal']
        for sub_folder_name in sub_folder_names:
            status,response = create_sub_folder(headers,parent_folder_id,sub_folder_name)
            if status:
                print(f'SubFolder "{sub_folder_name}" created.')
            else:
                print(f"Error creating subfolder '{response.json()}'")
    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()