import os
import httpx
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import get_message_by_filter,last_outlook_check_time,search_folder,create_folder,create_sub_folder,get_sub_folder,get_attachments,download_attachment,add_category_to_mail,move_email_to_folder
import re
from tqdm import tqdm

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def process_attachments(headers,message_id,dir_attachment):
    attachments = get_attachments(headers,message_id)
    for attachment in attachments:
        try:
            download_attachment(headers,message_id,attachment['id'],attachment['name'],dir_attachment)
        except httpx.HTTPStatusError as e:
            print(f"Failed to download {attachment['name']}: {e.response.status_code}")

def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read','Mail.ReadWrite']


    try:
        access_token = get_access_token(application_id=APPLICATION_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)
        headers = {'Authorization':'Bearer '+access_token}

        print()
        print(('-'*50)+f'Pre-Alert automation start'+('-'*50))
        print()
        print('Fetching Last process time......')
        print()

        outlook_last_check_time = last_outlook_check_time()

        print(f"Automation start with time {outlook_last_check_time}.")
        print()
        

        filter_condition = f'receivedDateTime ge {outlook_last_check_time}'

        folder_name = 'Inbox'
        target_folder = search_folder(headers,folder_name)
        folder_id=target_folder['id']

        print("Fetching message from Inbox......")
        print()

        messages = get_message_by_filter(headers,filter_condition,folder_id,top=50,max_results=50)


        if not len(messages):
            print("No mail found")
            return
        
        print("Fetching destination Folders id.....")
        print()

        destination_parent_folder_name = "Pre-Alerts Automation"
        destination_parent_folder = search_folder(headers,destination_parent_folder_name)

        pa_folder_id=None
        q_folder_id=None
        woa_folder_id=None

        if destination_parent_folder:
            destination_parent_folder_id=destination_parent_folder['id']
            sub_folders = get_sub_folder(headers,destination_parent_folder_id)

            for folder in sub_folders:
                if folder['displayName']=='Pre-Alerts':
                    pa_folder_id=folder['id']
                if folder['displayName']=='Query':
                    q_folder_id=folder['id']
                if folder['displayName']=='Without Attachments':
                    woa_folder_id=folder['id']

        else:
            status,response = create_folder(headers,destination_parent_folder_name)
            if not status:
                print(f'Error creating folder "{response.json()}".')
                return
        
            folder_metadata = response.json()
            print(f'Folder "{destination_parent_folder_name}" created.')

            destination_parent_folder_id = folder_metadata['id']

            sub_folder_names=['Pre-Alerts','Query','Without Attachments']

            for sub_folder_name in sub_folder_names:
                status,response = create_sub_folder(headers,destination_parent_folder_id,sub_folder_name)
                folder = response.json()
                if folder['displayName']=='Pre-Alerts':
                    pa_folder_id=folder['id']
                if folder['displayName']=='Query':
                    q_folder_id=folder['id']
                if folder['displayName']=='Without Attachments':
                    woa_folder_id=folder['id']
                if status:
                    print(f'SubFolder "{sub_folder_name}" created.')
                else:
                    print(f"Error creating subfolder '{response.json()}'")
            print()


        

        email_list = [
            "patelshiv3123@gmail.com",
            "210303105085@paruluniversity.ac.in",
            "shivpatel310323@gmail.com"
        ]
        print("Filtering e-mail based on present list....")
        print()
        filtered_emails = [
            email for email in tqdm(messages)
            if email.get("from", {}).get("emailAddress", {}).get("address") in email_list
        ]
        print()

        if not len(filtered_emails):
            print("No pre alert found")
            return


        dir_attachment = Path('./downloaded')
        dir_attachment.mkdir(parents=True,exist_ok=True)

        end_time = None

        print("Emails are in process based on conditions....")
        print()

        pbar = tqdm(total=len(filtered_emails))
        for i, message in (enumerate(filtered_emails)):
            is_last = i == len(filtered_emails) - 1

            if message.get("from", {}).get("emailAddress", {}).get("address") == "210303105085@paruluniversity.ac.in":
                add_category_to_mail(headers, message['id'], ["Yellow category"])
                move_email_to_folder(headers, message['id'],q_folder_id)

            elif message['hasAttachments']:
                subject = sanitize_filename(message['subject']) or "no_subject"
                received_time = sanitize_filename(message['receivedDateTime'])
                folder_name = f"{subject}_{received_time}"

                dir_attachment = Path('./downloaded') / folder_name
                dir_attachment.mkdir(parents=True, exist_ok=True)

                process_attachments(headers, message['id'], dir_attachment)
                add_category_to_mail(headers, message['id'], ["Orange category"])
                move_email_to_folder(headers, message['id'], pa_folder_id)
            else:
                add_category_to_mail(headers, message['id'], ["Orange category", "Yellow category"])
                move_email_to_folder(headers, message['id'], woa_folder_id)

            if is_last:
                end_time = message['receivedDateTime']
                with open('last_outlook_check_time.txt','w') as file:
                    file.write(end_time)
            pbar.update(1)
        pbar.close()
        print()  
        print(('-'*50)+f'Pre-Alert automation end at {end_time}'+('-'*50))
        print()


    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()