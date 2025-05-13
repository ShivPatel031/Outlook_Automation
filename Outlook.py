import os
import base64
import mimetypes
from pathlib import Path
import httpx
from datetime import datetime, timedelta
import pytz
from ms_graph import MS_GRAPH_BASE_URL

def search_folder(headers,folder_name='Drafts'):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders'
    response = httpx.get(endpoint,headers=headers)

    response.raise_for_status()

    folders = response.json().get('value',[])

    for folder in folders:

        if folder['displayName'].lower() == folder_name.lower():
            return folder
        
    return None    
    
def get_sub_folder(headers,folder_id):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}/childFolders'
    response = httpx.get(endpoint,headers=headers)
    response.raise_for_status()
    return response.json().get('value',[])


def get_message(headers,folder_id=None,fields="*",top=5,order_by='receivedDateTime',order_by_desc=True,max_results=20):

    if folder_id is None:
        endpoint = f'{MS_GRAPH_BASE_URL}/me/messages'
    else:
        endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}/messages'

    params = {
        '$select':fields,
        '$top':min(top,max_results),
        '$orderby':f'{order_by} {"desc" if order_by_desc else "asc"}'
    }

    messages =[]
    next_link = endpoint

    while next_link and len(messages) < max_results:
        response = httpx.get(next_link,headers=headers,params=params)

        if response.status_code != 200:
            raise Exception(f'Failed to retrieve emails: {response.json()}')
        
        json_response = response.json()
        messages.extend(json_response.get('value',[]))
        next_link =  json_response.get('@odata.nextLink',None)
        params = None

        if next_link and len(messages) + top > max_results:
            params = {
                '$top': max_results-len(messages)
            }

    return messages[:max_results] 


def search_messages(headers,searchquery,filter=None,folder_id=None,fields="*",top=5,max_results=100):
    if folder_id is None:
        endpoint=f'{MS_GRAPH_BASE_URL}/me/messages'
    else:
        endpoint=f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}/messages'

    params = {
        '$search':f'"{searchquery}"',
        '$filter':filter,
        '$select':fields,
        '$top':min(top,max_results)
    }

    messages = []
    next_link = endpoint

    while next_link and len(messages)<max_results:
        response = httpx.get(next_link,headers=headers,params=params)
        response.raise_for_status()
        if response.status_code != 200:
            raise Exception(f'Failed to retrieve emails: {response.json()}')
        
        json_response = response.json()
        messages.extend(json_response.get('value',[]))
        next_link=json_response.get('@odata.nextLink',None)
        params = None
        if next_link and len(messages) + top > max_results:
            params = {
                '$top':max_results - len(messages)
            }
    
    return messages[:max_results]

def create_attachment(file_path):
    with open(file_path,'rb') as file:
        content = file.read()
        encoded_content = base64.b64encode(content).decode('utf-8')

    return {
        "@odata.type":"#microsoft.graph.fileAttachment",
        "name": os.path.basename(file_path),
        "contentType":get_mime_type(file_path),
        "contentBytes":encoded_content
    }

def get_mime_type(file_path):
    mime_type,_ = mimetypes.guess_type(file_path)
    return mime_type

def get_message_by_filter(headers,filter,folder_id=None,fields="*",top=5,max_results=20):
    if folder_id is None:
        endpoint = f'{MS_GRAPH_BASE_URL}/me/messages'
    else:
        endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}/messages'

    params = {
        '$filter':filter,
        '$select':fields,
        '$top':min(top,max_results),
    }

    messages = []
    next_link = endpoint

    while next_link and len(messages) < max_results:
        response = httpx.get(next_link,headers=headers,params=params)

        if response.status_code != 200:
            raise Exception(f'Failed to retrieve emails: {response.json()}')
        
        json_response = response.json()
        messages.extend(json_response.get('value',[]))
        next_link=json_response.get('@odata.nextLink',None)
        params = None

        if next_link and len(messages)+top > max_results:
            params = {
                '$top':max_results - len(messages)
            }
    
    return messages[:max_results]

def get_attachments(headers,message_id):
    attachments_endpoint = f'{MS_GRAPH_BASE_URL}/me/messages/{message_id}/attachments'
    response = httpx.get(attachments_endpoint,headers=headers)
    response.raise_for_status()
    return response.json().get('value',[])

def download_attachment(headers,message_id,attachment_id,attachment_name,dir_attachment):
    download_endpoint = f'{MS_GRAPH_BASE_URL}/me/messages/{message_id}/attachments/{attachment_id}/$value'
    response = httpx.get(download_endpoint,headers=headers)
    response.raise_for_status()

    file_path = Path(dir_attachment) / attachment_name
    file_path.write_bytes(response.content)
    return True

def create_folder(headers,folder_name):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders'
    params = {
        'displayName':folder_name
    }

    response = httpx.post(endpoint,headers=headers,json=params)
    return response.status_code == 201,response

def create_sub_folder(headers,parent_folder_id,sub_folder_name):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{parent_folder_id}/childFolders'
    params = {
        'displayName':sub_folder_name
    }
    response = httpx.post(endpoint,headers=headers,json=params)

    return response.status_code == 201,response

def get_folder(headers,folder_id):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/mailFolders/{folder_id}'
    response = httpx.get(endpoint,headers=headers)
    response.raise_for_status()
    return response.json()

def reply_to_message(headers,message_id,reply_body):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/messages/{message_id}/reply'
    data = {
        'comment':reply_body
    }
    response = httpx.post(endpoint,headers=headers,json=data)
    response.raise_for_status()
    return response.status_code == 202

def delete_message(headers,message_id):
    delete_endpoint = f'{MS_GRAPH_BASE_URL}/me/messages/{message_id}'
    response = httpx.delete(delete_endpoint,headers=headers)
    response.raise_for_status()
    return True
  
def move_email_to_folder(headers,message_id,destinatio_folder_id):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/messages/{message_id}/move'
    params = {
        'destinationId':destinatio_folder_id
    }
    response = httpx.post(endpoint,headers=headers,json=params)
    response.raise_for_status()
    return response.json()

def add_category_to_mail(headers,message_id,category):
    endpoint = f'{MS_GRAPH_BASE_URL}/me/messages/{message_id}'
    params = {
    'categories': category
    }
    response = httpx.patch(endpoint,headers=headers,json=params)
    response.raise_for_status()
    return True



def to_graph_datetime_utc(year, month, day, hour, minute=0, second=0, tz_name='Asia/Kolkata'):
    local_tz = pytz.timezone(tz_name)
    local_dt = local_tz.localize(datetime(year, month, day, hour, minute, second))
    utc_dt = local_dt.astimezone(pytz.utc)
    return utc_dt.strftime('%Y-%m-%dT%H:%M:%SZ')

def add_one_second_to_graph_time(iso_datetime_str):
    dt = datetime.strptime(iso_datetime_str, '%Y-%m-%dT%H:%M:%SZ')
    dt_plus_one = dt + timedelta(seconds=1)
    return dt_plus_one.strftime('%Y-%m-%dT%H:%M:%SZ')


def last_outlook_check_time():
    if os.path.exists('last_outlook_check_time.txt'):
        with open('last_outlook_check_time.txt','r') as file:
            check_time = file.read().strip()
        return add_one_second_to_graph_time(check_time)
    else:
        local_timezone = pytz.timezone('Asia/Kolkata')
        now = datetime.now(local_timezone)
        yesterday = now - timedelta(days=2)
        return yesterday.astimezone(pytz.utc).isoformat()




