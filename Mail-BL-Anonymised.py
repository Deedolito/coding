################################# Prerequisites #################################

import requests
import json
import base64
import sys
import datetime
import time
import urllib.parse



CLIENT_ID =  "XXXX"
CLIENT_SECRET = "XXXX"
TENANT_ID = "XXXX"
REDIRECT_URI = 'https://localhost'
hostname = "XXXX.sharepoint.com"
site_path = "sites/Planification"
sharepoint_folder_name = "BL"


################################# End of Prerequisites #################################


            
            
#                                     1111111   
#                                    1::::::1   
#                                   1:::::::1   
#                                   111:::::1   
#                                      1::::1   
#                                      1::::1   
#                                      1::::1   
#                                      1::::l   
#                                      1::::l   
#                                      1::::l   
#                                      1::::l   
#                                      1::::l   
#                                   111::::::111
#                                   1::::::::::1
#                                   1::::::::::1
#                                   111111111111
            
            
            
            

################################# Step 1: Setup and Acquire Token #################################

# This code defines a function called acquire_access_token() that requests an access token from the Microsoft Graph API

print("\n\n\n Step 1 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 1 \n".upper())

print("\n Step 1 ====================================================     - Step 1 : Client application Credential clearance     ==================================================== Step 1 \n".upper())

print("\n Step 1 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 1 \n\n\n".upper())



print("   - Step 1.1 complete : Client application has been activated successfully.")

# Function to acquire an access token
def acquire_access_token():
    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    payload = {
        'client_id': CLIENT_ID,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': CLIENT_SECRET,
        'grant_type': 'client_credentials'
    }
    
    response = requests.post(url, headers=headers, data=payload)
    
    if response.status_code == 200:
        print("   - Step 1.2 complete : Access token acquired")
        return response.json()['access_token']
    else:
        print("/ ! \ ERORR / ! \ Step 1.2 : ", response.status_code)
        print(response.json())
        return None

# Acquire the access token 

access_token = acquire_access_token()

print("   - Step 1.3 complete : Access token stored in variable <access_token> ")

# The acquire_access_token() function sends a POST request to the token endpoint, passing the necessary parameters, and returns the access token if the request is successful. The access token is stored in the variable access_token.

################################# End of Step 1: Setup and Acquire Token #################################


                    
                    
#                                   222222222222222    
#                                  2:::::::::::::::22  
#                                  2::::::222222:::::2 
#                                  2222222     2:::::2 
#                                              2:::::2 
#                                              2:::::2 
#                                           2222::::2  
#                                      22222::::::22   
#                                    22::::::::222     
#                                   2:::::22222        
#                                  2:::::2             
#                                  2:::::2             
#                                  2:::::2       222222
#                                  2::::::2222222:::::2
#                                  2::::::::::::::::::2
#                                  22222222222222222222
#                                                      
                    
                    
                    

################################# Step 2: Make a request to the Microsoft Graph API #################################

print("\n\n\n Step 2 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 2 \n".upper())

print("\n Step 2 ====================================================              - Step 2 : Request to Graph API               ==================================================== Step 2 \n".upper())

print("\n Step 2 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 2 \n\n\n".upper())

#  Takes the access token as an argument. The function sends a GET request to the Microsoft Graph API's /me endpoint to get the user's profile information. 


def get_user_profile(access_token, print_status=True):
 
    user_id = "b2f7e347-4447-4b6d-a6f9-dc9ee6ae6a76"
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
    
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        if print_status:
            print("   - Step 2.1 complete : user profile acquired")
        return response.json()
    else:
        print(f"/ ! \ ERORR / ! \ in get_user_profile: {response.status_code}")
        print(response.json())
        return None


# Get the user's profile information
user_profile = get_user_profile(access_token)

#  store the result in the user_profile variable


#if user_profile:
#    print("User profile:")
#    print(json.dumps(user_profile, indent=2))

print("   - Step 2.2 complete : store profile result in variable ")

################################# End of Step 2: Make a request to the Microsoft Graph API #################################


                   
                   
#                                      333333333333333   
#                                     3:::::::::::::::33 
#                                     3::::::33333::::::3
#                                     3333333     3:::::3
#                                                 3:::::3
#                                                 3:::::3
#                                         33333333:::::3 
#                                         3:::::::::::3  
#                                         33333333:::::3 
#                                                 3:::::3
#                                                 3:::::3
#                                                 3:::::3
#                                     3333333     3:::::3
#                                     3::::::33333::::::3
#                                     3:::::::::::::::33 
#                                      333333333333333   
                   
                   
                   

################################# Step 3: List messages in the mailbox #################################

print("\n\n\n Step 3 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 3 \n".upper())

print("\n Step 3 ====================================================      - Step 3: List messages in the mailbox subfolder      ==================================================== Step 3 \n".upper())

print("\n Step 3 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 3 \n\n\n".upper())

# Get the user's profile information

user_profile = get_user_profile(access_token, print_status=False)

if user_profile:
#    print("User profile:")
#    print(json.dumps(user_profile, indent=2))
    print("   - Step 3.1 complete : define variable user_profile ")
    user_id = user_profile['id']


# Define variable list_message

def list_messages(access_token, user_id, folder_id, is_unread=None):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders/{folder_id}/messages"

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    params = {
        "$select": "subject,receivedDateTime",
        "$orderby": "receivedDateTime DESC"
    }

    if is_unread:
        params["$filter"] = "isRead eq false"  

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        return response.json()["value"]
    else:
        print(f"/ ! \ ERORR / ! \ in list_messages: {response.status_code}")
        print(response.json())
        return None

print("   - Step 3.2 complete : define variable list_message ")

# Define the subfolder_id
subfolder_id = "XXXX="

# Update the list_messages function call to include the is_unread filter
messages = list_messages(access_token, user_id, subfolder_id, is_unread=True)

if messages:
    print("   - Step 3.3 complete: Unread messages listed\n")
    print(f"Unread message count: {len(messages)}\n")  # Added print statement for unread message count
    print("Unread messages list:\n")
    for message in messages:
        print(f" - Subject: {message['subject']}")
else:
    print("   - Step 3.3: No unread messages found")

################################# End of Step 3: List messages in the mailbox #################################


                  
                  
#                                              444444444  
#                                             4::::::::4  
#                                            4:::::::::4  
#                                           4::::44::::4  
#                                          4::::4 4::::4  
#                                         4::::4  4::::4  
#                                        4::::4   4::::4  
#                                       4::::444444::::444
#                                       4::::::::::::::::4
#                                       4444444444:::::444
#                                                 4::::4  
#                                                 4::::4  
#                                                 4::::4  
#                                               44::::::44
#                                               4::::::::4
#                                               4444444444
                  
                  

################################## Step 4: send an alert for unread message #################################


print("\n\n\n Step 4 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 4 \n".upper())

print("\n Step 4 ====================================================              - Step 4: Send an email to alert              ==================================================== Step 4 \n".upper())

print("\n Step 4 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 4 \n\n\n".upper())

def send_email(access_token, user_id, to_email, subject, body):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/sendMail"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email
                    }
                }
            ]
        },
        "saveToSentItems": "true"
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 202:
        print("   - Step 4.2 complete : Email sent successfully")
    else:
        print(f"/ ! \ ERORR / ! \ sending email: {response.status_code}")
        print(response.json())

print("   - Step 4.1 complete : create email")

# Send an email if there's at least one unread message

if messages:
    to_email = "t.XXXX@XXXX.fr"  # Change this to the recipient's email address
    email_subject = f"{len(messages)} Nouveau BL en attente dans le dossier BL"
    email_body = f"Bonjour,<br><br> Il y a {len(messages)} BL en attente."

    send_email(access_token, user_id, to_email, email_subject, email_body)


################################## End ofStep 4: send an alert for unread message #################################


                   
                   
#                                       555555555555555555 
#                                       5::::::::::::::::5 
#                                       5::::::::::::::::5 
#                                       5:::::555555555555 
#                                       5:::::5            
#                                       5:::::5            
#                                       5:::::5555555555   
#                                       5:::::::::::::::5  
#                                       555555555555:::::5 
#                                                   5:::::5
#                                                   5:::::5
#                                       5555555     5:::::5
#                                       5::::::55555::::::5
#                                        55:::::::::::::55 
#                                          55:::::::::55   
#                                            555555555     
                   
                   
                   
                   
                   
                   
                   


################################## Step 5: Add a task to Planner for each unread email ##################################

print("\n\n\n Step 5 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 5 \n".upper())

print("\n Step 5 ====================================================   - Step 5: Add a task to Planner for each unread email    ==================================================== Step 5 \n".upper())

print("\n Step 5 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 5 \n\n\n".upper())

# Define a function to get the group ID
def get_group_id(access_token, group_name):
    url = "https://graph.microsoft.com/v1.0/groups"

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    params = {
        "$filter": f"displayName eq '{group_name}'"
    }

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        groups = response.json()["value"]
        if len(groups) > 0:
            return groups[0]["id"]
        else:
            print("No group found with the specified name.")
            return None
    else:
        print(f"/ ! \ ERORR / ! \ getting group ID: {response.status_code}")
        print(response.json())
        return None



# Define a function to get the plan ID
def get_plan_id(access_token, group_id):
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/planner/plans"

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        plans = response.json()["value"]
        if len(plans) > 0:
            return plans[0]["id"]
        else:
            print("No plans found for the specified group.")
            return None
    else:
        print(f"/ ! \ ERORR / ! \ getting plan ID: {response.status_code}")
        print(response.json())
        return None
    



# Define a function to create a task in Planner
def create_planner_task(access_token, plan_id, bucket_id, title, due_date=None):
    url = "https://graph.microsoft.com/v1.0/planner/tasks"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    json_data = {
        "planId": plan_id,
        "bucketId": bucket_id,
        "title": title,
        "assignments": {}
    }

    if due_date:
        json_data["dueDateTime"] = due_date

    response = requests.post(url, headers=headers, json=json_data)

    if response.status_code == 201:
        task_id = response.json()["id"]
        return task_id

    else:
        print(f"/ ! \ ERORR / ! \ creating Planner task: {response.status_code}")
        print(response.json())
        return None

# define a function to get buckets IDs

def get_bucket_id(access_token, plan_id, bucket_name):
    url = f"https://graph.microsoft.com/v1.0/planner/plans/{plan_id}/buckets"

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        buckets = response.json()["value"]
        for bucket in buckets:
            if bucket["name"] == bucket_name:
                return bucket["id"]
        print("No bucket found with the specified name.")
        return None
    else:
        print(f"/ ! \ ERORR / ! \ getting bucket ID: {response.status_code}")
        print(response.json())
        return None



# Replace 'Your Group Name' with the name of the group associated with the plan
group_name = "XXXX | Planification"

# Get the ID of the group
group_id = get_group_id(access_token, group_name)
if group_id:
    print(f"   - Step 5.1 complete : Group ID found: {group_id}")
else:
    print("/ ! \ ERORR / ! \ Step 5.1: No group found with the specified name")
    exit(1)  # Exit if no group is found

# Get the ID of the plan
plan_id = get_plan_id(access_token, group_id)
if plan_id:
    print(f"   - Step 5.2 complete : Plan ID found: {plan_id}")
else:
    print("/ ! \ ERORR / ! \ Step 5.2: No plans found for the specified group")
    exit(1)  # Exit if no plan is found

# Replace 'Your Bucket Name' with the name of the bucket where you want to add tasks
bucket_name = "Réception des BL"
bucket_id = get_bucket_id(access_token, plan_id, bucket_name)
if bucket_id:
    print(f"   - Step 5.3 complete : Bucket ID found: {bucket_id}")

    for message in messages:
        task_title = f"{message['subject']}"
        task_id = create_planner_task(access_token, plan_id, bucket_id, task_title)

        if task_id:
            print(f"   - Step 5.4 complete : Task created for email: {message['subject']} with Task ID: {task_id}")
        else:
            print(f"/ ! \ ERORR / ! \ creating task for email: {message['subject']}")
else:
    print("/ ! \ ERORR / ! \ Step 5.3: No bucket found with the specified name")
    exit(1)  # Exit if no bucket is found



################################# End of Step 5: Add a task to Planner #################################


                   
                   
#                                              66666666   
#                                             6::::::6    
#                                            6::::::6     
#                                           6::::::6      
#                                          6::::::6       
#                                         6::::::6        
#                                        6::::::6         
#                                       6::::::::66666    
#                                      6::::::::::::::66  
#                                      6::::::66666:::::6 
#                                      6:::::6     6:::::6
#                                      6:::::6     6:::::6
#                                      6::::::66666::::::6
#                                       66:::::::::::::66 
#                                         66:::::::::66   
#                                           666666666     
                   
                   
                     
    

################################### Step 6: Extract attachments from emails and store them in SharePoint #################################

print("\n\n\n Step 6 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 6 \n".upper())

print("\n Step 6 ====================================================              - Step 6: processing attachments              ==================================================== Step 6 \n".upper())

print("\n Step 6 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 6 \n\n\n".upper())

# Define a function to get the attachments of an email.
def get_attachments(access_token, user_id, message_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}/attachments"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        attachments = response.json()["value"]
        filtered_attachments = [attachment for attachment in attachments if 'image001' not in attachment['name']]
        return filtered_attachments       
    else:
        return None

# Define a function to get the site id for the SharePoint site.
def get_site_id(access_token, hostname, site_path):
    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        json_response = response.json()
        if "id" in json_response:
            return json_response["id"]            
        else:
            return None
    else:
        return None

# Define a function to get the folder id for the SharePoint folder.
def get_folder_id(access_token, site_id, folder_name):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        folders = response.json()["value"]
        for folder in folders:
            if folder["name"] == folder_name and folder["folder"]:
                return folder["id"]
        return None
    else:
        return None



# Step 6.1: Get the site ID
site_id = get_site_id(access_token, hostname, site_path)
if site_id:
    print(f"   - Step 6.1 complete: Site ID obtained: {site_id}")
else:
    print("/ ! \ ERORR / ! \ Step 6.1: Failed to get the Site ID")
    exit(1)

# Step 6.2: Get the folder ID
folder_id = get_folder_id(access_token, site_id, sharepoint_folder_name)
if folder_id:
    print(f"   - Step 6.2 complete: Folder ID obtained: {folder_id}")
else:
    print("/ ! \ ERORR / ! \ Step 6.2: Failed to get the Folder ID")
    exit(1)

# Step 6.3 - Define a function to get the attachments of an email.
def get_attachments(access_token, user_id, message_id, print_statement=True):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}/attachments"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        attachments = response.json()["value"]
        filtered_attachments = [attachment for attachment in attachments if 'image001' not in attachment['name']]
        if print_statement:
            print(f"   - Step 6.3 complete: Found {len(filtered_attachments)} attachments for message ID {message_id}")
        return filtered_attachments       
    else:
        return None


# Step 6.4 Define a function to upload the attachments to SharePoint.
def upload_to_sharepoint(access_token, site_id, folder_id, attachment):
    # Skip attachments that are likely part of an email signature
    if 'image001' in attachment['name']:
        print(f"Skipping likely signature image: {attachment['name']}")
        return None
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{folder_id}/children/{attachment['name']}/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/octet-stream"
    }

    # Store the name of the uploaded file
    global file_name  # Make this variable accessible globally
    file_name = attachment['name']

    # Convert contentBytes to binary
    binary_data = base64.b64decode(attachment["contentBytes"])
    response = requests.put(url, headers=headers, data=binary_data)
    if response.status_code == 201:
        print(f"   - Step 6.4 complete: Attachment {attachment['name']} uploaded to SharePoint with ID: {response.json()['id']}")
        return response.json()['id']  # Return the id of the uploaded file
    else:
        return None

# Step 6.5 Check if email has attachments, fetch them if available.
for message in messages:
    if 'hasAttachments' in message and message['hasAttachments']:
        attachments = get_attachments(access_token, user_id, message["id"])
        if attachments is not None:
            print(f"   - Step 6.5: Attachments found for email: {message['subject']}")
        else:
            print(f"Warning: No attachments to fetch for email: {message['subject']}")

# Step 6.6 Upload attachments to SharePoint.
for message in messages:
    attachments = get_attachments(access_token, user_id, message["id"])
    if attachments is not None:
        for attachment in attachments:
            if 'contentBytes' in attachment:  # only for file attachments
                upload_id = upload_to_sharepoint(access_token, site_id, folder_id, attachment)
                if not upload_id:
                    print(f"/ ! \ ERROR / ! \ Step 6.6: Failed to upload attachment: {attachment['name']}")
    else:
        print(f"Warning: No attachments to fetch for email: {message['subject']}")



################################### End of Step 6: Extract attachments from emails and store them in SharePoint #################################


                    
                    
#                                         77777777777777777777
#                                         7::::::::::::::::::7
#                                         7::::::::::::::::::7
#                                         777777777777:::::::7
#                                                    7::::::7 
#                                                   7::::::7  
#                                                  7::::::7   
#                                                 7::::::7    
#                                                7::::::7     
#                                               7::::::7      
#                                              7::::::7       
#                                             7::::::7        
#                                            7::::::7         
#                                           7::::::7          
#                                          7::::::7           
#                                         77777777            
#                                                             
                                 


################################### Step 7: Attach the SharePoint file to the Planner task #################################

print("\n\n\n Step 7 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 7 \n".upper())

print("\n Step 7 ====================================================              - Step 7: attaching SharePoint file to Planner task              ==================================================== Step 7 \n".upper())

print("\n Step 7 <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Step 7 \n\n\n".upper())




# Step 7.1: Get the URL of the uploaded file from SharePoint

headers = {"Authorization": f"Bearer {access_token}"}

# Define the endpoint URL
url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{upload_id}?select=id,@microsoft.graph.downloadUrl"

# Make the GET request to the Microsoft Graph API
response = requests.get(url, headers=headers)

# Check the response status code
if response.status_code == 200:
    # The API returns a 200 OK response
    response_json = response.json()
    file_url = response_json.get('@microsoft.graph.downloadUrl')
    print(f"- Step 7.1 complete: File URL obtained: {file_url}")
else:
    print(f"/ ! \\ ERROR / ! \\ Step 7.1: Failed to get file URL from SharePoint. Status code: {response.status_code}")

# Step 7.2: Add the file reference to the task in Planner

def add_attachment_to_task(access_token, task_id, file_name, file_url):
    print(f" - Step 7.2 start: Adding attachment to task in Planner")

 # Define the URL for the Planner Task Details endpoint
    url = f"https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details"

    # Define the headers for the request
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }


    # Extract the file type from the file name
    file_type = file_name.split('.')[-1] if '.' in file_name else "Unknown"

    # Define the JSON body for the request
    # This includes the file reference to be added to the task
# Using the format() method instead of an f-string
    data = {
        "references": {
            "{}".format(file_url_encoded): {
                "@odata.type": "microsoft.graph.plannerExternalReference",
                "alias": file_name,
                "type": file_type
            }
        }
    }


    # Make the PATCH request to the Planner Task Details endpoint
    response = requests.patch(url, headers=headers, json=data)

    # Check the status code of the response
    if response.status_code == 200:
        print(f"   - Step 7.2 complete: Attachment added to task in Planner")
    else:
        print(f"/ ! \\ ERROR / ! \\ Step 7.2: Failed to add attachment to task in Planner. Status code: {response.status_code}, Response: {response.text}")

# Call the function to add the attachment to the task
# The access_token, task_id, file_name, and file_url_encoded are variables obtained from the previous steps
add_attachment_to_task(access_token, task_id, file_name, file_url)

################################### End of Step 7: Attach the SharePoint file to the Planner task #################################






































                                                                                                                                     
                                                                                                                                     

                                                                  
                                                                  
#                  FFFFFFFFFFFFFFFFFFFFFF     IIIIIIIIII     NNNNNNNN        NNNNNNNN
#                  F::::::::::::::::::::F     I::::::::I     N:::::::N       N::::::N
#                  F::::::::::::::::::::F     I::::::::I     N::::::::N      N::::::N
#                  FF::::::FFFFFFFFF::::F     II::::::II     N:::::::::N     N::::::N
#                    F:::::F       FFFFFF       I::::I       N::::::::::N    N::::::N
#                    F:::::F                    I::::I       N:::::::::::N   N::::::N
#                    F::::::FFFFFFFFFF          I::::I       N:::::::N::::N  N::::::N
#                    F:::::::::::::::F          I::::I       N::::::N N::::N N::::::N
#                    F:::::::::::::::F          I::::I       N::::::N  N::::N:::::::N
#                    F::::::FFFFFFFFFF          I::::I       N::::::N   N:::::::::::N
#                    F:::::F                    I::::I       N::::::N    N::::::::::N
#                    F:::::F                    I::::I       N::::::N     N:::::::::N
#                  FF:::::::FF                II::::::II     N::::::N      N::::::::N
#                  F::::::::FF                I::::::::I     N::::::N       N:::::::N
#                  F::::::::FF                I::::::::I     N::::::N        N::::::N
#                  FFFFFFFFFFF                IIIIIIIIII     NNNNNNNN         NNNNNNN
                                                                  
                                                                  
                                                                  


################################## Final Step : Move the unread messages to "A Planifier" folder #################################

print("\n\n\n Final Step <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Final Step \n".upper())

print("\n Final Step ====================================================  Final Step : Move unread messages to A Planifier folder   ==================================================== Final Step \n".upper())

print("\n Final Step <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Final Step \n\n\n".upper())

# Get the ID of the "A Planifier" folder
a_planifier_folder_id = "AAMkADkyMjc2MWIzLWQwNDUtNGY5ZS04NzJhLTk1OTdiZGQzODliNgAuAAAAAACmQI4CMDM8TJezdB5mMrDEAQCyxrQlaMCVT7gNvobDJrQvAAMFeGg6AAA="

# Define a function to move a message to a specified folder
def move_message(access_token, user_id, message_id, destination_folder_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}/move"
    
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    
    json_data = {
        "destinationId": destination_folder_id
    }
    
    response = requests.post(url, headers=headers, json=json_data)
    
    if response.status_code == 201:
        return response.json()
    else:
        print(f"/ ! \ ERORR / ! \ moving message: {response.status_code}")
        print(response.json())
        return None



# Move the unread messages to the "A Planifier" folder
message_moved = False  # Flag to check if any message has been moved
for message in messages:
    result = move_message(access_token, user_id, message["id"], a_planifier_folder_id)
    if result:
        print(f"Message '{message['subject']}' moved to 'A Planifier' folder")
        message_moved = True  # Update flag if a message has been moved

if not message_moved:
    print("No messages have been moved")

print("\n\n\n Final Step <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Final Step \n".upper())

print("\n Final Step ====================================================    - Final step complete : Move email to destination folder==================================================== Final Step \n".upper())

print("\n Final Step <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><> Final Step \n\n\n".upper())

################################## End of Final Step : Move the unread messages to "A Planifier" folder #################################