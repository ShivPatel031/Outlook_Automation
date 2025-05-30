import os
import httpx
from datetime import datetime
import pytz
from pathlib import Path
from dotenv import load_dotenv
from ms_graph import get_access_token
from Outlook import get_message_by_filter,last_outlook_check_time,search_folder,create_folder,create_sub_folder,get_sub_folder,get_attachments,download_attachment,add_category_to_mail,move_email_to_folder,get_single_message
import re
from tqdm import tqdm
import json


category = {
    "query":["Yellow category"],
    "pre_alert":["Orange category"],
    "no_attachments":["Orange category", "Yellow category"]
}

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def update_status_file(data):
    with open("current_email_process.json", "w") as f:
        json.dump(data, f, indent=4)

def assign_filter_value(message,data):
    if message.get("from", {}).get("emailAddress", {}).get("address") == "210303105085@paruluniversity.ac.in":
        data['filter']="query"
    elif message['hasAttachments']:
        data['filter']="pre_alert"
    else:
        data['filter']="no_attachments"
    
    update_status_file(data)

    return data

def combine_conditions(data):
    if data['filter'] == "pre_alert":
        with open("pre_alert_conditions.json","r") as t:
            data=data | json.load(t)
    else:
        with open("no_pre_alert_conditions.json","r") as t:
            data = data | json.load(t)

    update_status_file(data)

    return data


def add_category(headers,message,data,category):
    add_category_to_mail(headers, message['id'], category[data['filter']])

    data["category_added"]=True

    update_status_file(data)

    return data

def download_attachments(headers,message,data):
    subject = sanitize_filename(message['subject']) or "no_subject"
    received_time = sanitize_filename(message['receivedDateTime'])
    folder_name = f"{subject}_{received_time}"

    dir_attachment = Path('./downloaded') / folder_name
    dir_attachment.mkdir(parents=True, exist_ok=True)

    process_attachments(headers, message['id'], dir_attachment)

    data["attachments_downloaded"]=True

    update_status_file(data)

    return data

def move_to_folder(headers,message,data,folders):
    
    move_email_to_folder(headers, message['id'],folders[data['filter']])

    data["moved_to_folder"]=True
    data['end']=True

    update_status_file(data)

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

        folders={
            "query":"",
            "pre_alert":"",
            "no_attachments":""
        }

        print()
        print(('-'*50)+f'Pre-Alert automation start'+('-'*50))
        print()

        folder_name = 'Inbox'
        target_folder = search_folder(headers,folder_name)
        folder_id=target_folder['id']

        destination_parent_folder_name = "Pre-Alerts Automation"
        destination_parent_folder = search_folder(headers,destination_parent_folder_name)



        if destination_parent_folder:
            destination_parent_folder_id=destination_parent_folder['id']
            sub_folders = get_sub_folder(headers,destination_parent_folder_id)

            for folder in sub_folders:
                if folder['displayName']=='Pre-Alerts':
                    folders["pre_alert"]=folder['id']
                if folder['displayName']=='Query':
                    folders["query"]=folder['id']
                if folder['displayName']=='Without Attachments':
                    folders['no_attachments']=folder['id']

        else:
            status,response = create_folder(headers,destination_parent_folder_name)
            if not status:
                print(f'Error creating folder "{response.json()}".')
                return
        
            folder_metadata = response.json()

            destination_parent_folder_id = folder_metadata['id']

            sub_folder_names=['Pre-Alerts','Query','Without Attachments']

            for sub_folder_name in sub_folder_names:
                status,response = create_sub_folder(headers,destination_parent_folder_id,sub_folder_name)
                folder = response.json()
                if folder['displayName']=='Pre-Alerts':
                    folders["pre_alert"]=folder['id']
                if folder['displayName']=='Query':
                    folders["query"]=folder['id']
                if folder['displayName']=='Without Attachments':
                    folders['no_attachments']=folder['id']
        print()

        outlook_last_check_time = None

        if os.path.exists("current_email_process.json"):
            eps = None
            with open("current_email_process.json",'r') as apf:
                eps = json.load(apf)
            
            if eps and (not eps["end"]):

                message = get_single_message(headers,eps['email_Id'],folder_id)

                if message is not None:
                
                    if not eps['filter']:
                        eps = assign_filter_value(message,eps)
                        eps = combine_conditions(eps)
                    
                    if not eps['category_added']:
                        eps = add_category(headers,message,eps,category)

                    if "attachments_downloaded" in eps:
                        if not eps['attachments_downloaded']:
                            eps = download_attachments(headers,message,eps)

                    if not eps['category_added']:
                        move_to_folder(headers,message,eps,folders)
                
                outlook_last_check_time = eps['time']


        
        if outlook_last_check_time is None:
            outlook_last_check_time = last_outlook_check_time()

        print(f"Automation start with time {outlook_last_check_time}.")
        print()
        

        filter_condition = f'receivedDateTime ge {outlook_last_check_time}'

        

        print("Fetching message from Inbox......")
        print()

        messages = get_message_by_filter(headers,filter_condition,folder_id,top=50,max_results=50)


        if not len(messages):
            print("No mail found")
            return
        
        last_email = messages[-1]
        end_time = last_email['receivedDateTime']        

        email_list = [
            "patelshiv3123@gmail.com",
            "210303105085@paruluniversity.ac.in",
            "shivpatel310323@gmail.com"
        ]
        print("Filtering e-mail based on present list....")
        print()

        for message in messages:
            print(message['subject'])
            print(message['receivedDateTime'])

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

        print("Emails are in process based on conditions....")
        print()

        pbar = tqdm(total=len(filtered_emails))
        for i, message in (enumerate(filtered_emails)):
            data=None
            with open("process_structure.json","r") as file:
                data=json.load(file)
            
            data["email_Id"]=message['id']
            data["time"]=message['receivedDateTime']

            update_status_file(data)

            data = assign_filter_value(message,data)

            data = combine_conditions(data)

            data = add_category(headers,message,data,category)
            

            if "attachments_downloaded" in data:

                data = download_attachments(headers,message,data)

            move_to_folder(headers,message,data,folders)

            pbar.update(1)
        pbar.close()

        with open('last_outlook_check_time.txt','w') as file:
            with open("current_email_process.txt","a") as f:
                file.write(end_time)
                f.write("\nLast email time logged successfully.")
                
        print()  
        print(('-'*50)+f'Pre-Alert automation end at {end_time}'+('-'*50))
        print()


    except httpx.HTTPStatusError as e:
        print(f'HTTP Error: {e}')
    except Exception as e:
        print(f'Error: {e}')

main()