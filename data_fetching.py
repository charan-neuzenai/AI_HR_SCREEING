# import streamlit as st
# import os
# import requests
# import tempfile
# import base64
# import io
# from msal import PublicClientApplication
# from google.oauth2.credentials import Credentials
# from google_auth_oauthlib.flow import InstalledAppFlow
# from googleapiclient.discovery import build
# from googleapiclient.http import MediaIoBaseDownload
# from googleapiclient.errors import HttpError
# import time
# from pathlib import Path

# # Constants
# RESUME_TYPES = ['.pdf', '.docx', '.doc']
# RESUME_KEYWORDS = ['resume', 'cv', 'curriculum vitae', 'bio data']
# EXCLUDE_KEYWORDS = ['offer', 'letter', 'terms', 'conditions', 'contest', 'referral', 'policy', 'agreement']
# DOWNLOAD_DIR = r"C:\Users\jaswa\OneDrive\Desktop\HR_AI_Resume\mail_download\downloaded_resumes-3"

# # Ensure download directory exists
# Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)

# # Streamlit UI Configuration
# st.set_page_config(page_title="🌟 Ultimate Resume Downloader", layout="wide", page_icon="📥")

# st.title("🌟 Ultimate Resume Downloader")
# st.markdown("""
# <div style="background-color:#0e1117;padding:20px;border-radius:10px;margin-bottom:20px">
#     <h3 style="color:white;text-align:center;">Download resumes from SharePoint, OneDrive, Outlook, Google Drive & Gmail</h3>
# </div>
# """, unsafe_allow_html=True)

# # Initialize session state
# if 'active_tab' not in st.session_state:
#     st.session_state.active_tab = "Microsoft"

# # Tabs for provider selection
# ms_tab, google_tab = st.tabs(["🔷 Microsoft", "🔷 Google"])

# # ====================== MICROSOFT SECTION ======================
# with ms_tab:
#     st.subheader("🔐 Microsoft Authentication")
#     col1, col2 = st.columns(2)
#     with col1:
#         client_id = st.text_input("Client ID", value="5a56d5a6-fa4c-46a4-2f48-941f44b012a", key="ms_client_id")
#     with col2:
#         tenant_id = st.text_input("Tenant ID", value="common", key="ms_tenant_id")

#     mode = st.radio("Select source", ["SharePoint Folder", "OneDrive Folder", "Outlook Attachments"], key="ms_mode")

#     if mode == "SharePoint Folder":
#         st.subheader("🏢 SharePoint Configuration")
#         domain = st.text_input("SharePoint domain (e.g., company.sharepoint.com)", 
#                              value="buildionsportld.sharepoint.com", 
#                              key="sharepoint_domain")
#         site_name = st.text_input("Site name (e.g., HRTeamSite)", 
#                                  value="NEUZEMATISOLUTIONSPYTLTD", 
#                                  key="sharepoint_site")
#         folder_name = st.text_input("Folder name (e.g., Resumes2024)", 
#                                   value="AL RESUME_FILTER", 
#                                   key="sharepoint_folder")
#     elif mode == "OneDrive Folder":
#         st.subheader("📁 OneDrive Configuration")
#         folder_name = st.text_input("Folder name in OneDrive (e.g., MyResumes)", key="onedrive_folder")
#     else:
#         st.subheader("📧 Outlook Configuration")
#         max_mails = st.number_input("Number of emails to scan", min_value=1, max_value=10000, value=1000, key="ms_max_mails")

#     ms_run_button = st.button("🚀 Download Microsoft Resumes", key="ms_run")

# # ====================== GOOGLE SECTION ======================
# with google_tab:
#     st.subheader("🔐 Google Authentication")
#     creds_file = st.file_uploader("Upload credentials.json", type="json", key="google_creds")
    
#     google_mode = st.radio("Select source", ["Google Drive Folder", "Gmail Attachments"], key="google_mode")
    
#     if google_mode == "Google Drive Folder":
#         st.subheader("📁 Google Drive Configuration")
#         folder_url = st.text_input("Google Drive Folder URL", key="drive_url")
#     else:
#         st.subheader("📧 Gmail Configuration")
#         max_emails = st.number_input("Number of emails to scan", min_value=1, max_value=10000, value=1000, key="gmail_max_emails")
    
#     google_run_button = st.button("🚀 Download Google Resumes", key="google_run")

# # ====================== COMMON FUNCTIONS ======================
# def check_resume(filename):
#     """Enhanced resume detection with detailed rejection reasons"""
#     if not filename:
#         return False, "No filename"
    
#     name = filename.lower()
#     ext = os.path.splitext(filename)[1].lower()
    
#     # Check exclusion keywords
#     for word in EXCLUDE_KEYWORDS:
#         if word in name:
#             return False, f"Excluded keyword: '{word}'"
    
#     # Check file extension
#     if ext not in RESUME_TYPES:
#         return False, f"Invalid file type: {ext}"
    
#     # Check resume keywords
#     for word in RESUME_KEYWORDS:
#         if word in name:
#             return True, ""
    
#     # Check common patterns
#     patterns = ['_resume', '-resume', 'resume_', 'resume-', '_cv', '-cv', 'cv_', 'cv-']
#     if any(p in name for p in patterns):
#         return True, ""
    
#     return False, "No resume keywords found"

# def save_file(content, filename, source):
#     """Save file to download directory with collision handling"""
#     counter = 1
#     name, ext = os.path.splitext(filename)
#     new_filename = filename
    
#     while os.path.exists(os.path.join(DOWNLOAD_DIR, new_filename)):
#         new_filename = f"{name}_{counter}{ext}"
#         counter += 1
    
#     filepath = os.path.join(DOWNLOAD_DIR, new_filename)
    
#     if source == "outlook":
#         with open(filepath, 'wb') as f:
#             f.write(base64.b64decode(content))
#     else:
#         with open(filepath, 'wb') as f:
#             f.write(content)
    
#     return new_filename

# # ====================== MICROSOFT FUNCTIONS ======================
# def ms_authenticate(client_id, tenant_id, scopes):
#     authority = f"https://login.microsoftonline.com/{tenant_id}"
#     app = PublicClientApplication(client_id, authority=authority)
    
#     # Clear any cached accounts
#     accounts = app.get_accounts()
#     for account in accounts:
#         app.remove_account(account)
    
#     # Interactive authentication
#     result = app.acquire_token_interactive(scopes=scopes, prompt="select_account")
    
#     if "access_token" in result:
#         return result["access_token"], result.get("account", {}).get("username")
#     raise Exception("Authentication failed: " + str(result.get("error_description", "Unknown error")))

# def get_site_and_drive_ids(headers, hostname, site_name):
#     site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
#     site_resp = requests.get(site_url, headers=headers)
#     site_resp.raise_for_status()
#     site_id = site_resp.json()["id"]
    
#     drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
#     drive_resp = requests.get(drive_url, headers=headers)
#     drive_resp.raise_for_status()
#     drive_id = drive_resp.json()["value"][0]["id"]
    
#     return site_id, drive_id

# def list_sharepoint_files(headers, domain, site_name, folder_name):
#     try:
#         site_id, drive_id = get_site_and_drive_ids(headers, domain, site_name)
#         folder_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_name}"
#         folder_resp = requests.get(folder_url, headers=headers)
#         folder_resp.raise_for_status()
#         folder_id = folder_resp.json()["id"]
        
#         files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children"
#         files_resp = requests.get(files_url, headers=headers)
#         files_resp.raise_for_status()
        
#         return files_resp.json().get("value", []), f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items"
#     except Exception as e:
#         st.error(f"Error accessing SharePoint: {str(e)}")
#         return [], ""

# def list_onedrive_files(headers, folder_name):
#     try:
#         url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_name}:/children"
#         resp = requests.get(url, headers=headers)
#         resp.raise_for_status()
#         return resp.json().get("value", []), "https://graph.microsoft.com/v1.0/me/drive/items"
#     except Exception as e:
#         st.error(f"Error accessing OneDrive: {str(e)}")
#         return [], ""

# def fetch_outlook_attachments(headers, max_mails):
#     try:
#         files = []
#         url = f"https://graph.microsoft.com/v1.0/me/messages?$top={min(1000, max_mails)}&$orderby=receivedDateTime desc"
        
#         while url and len(files) < max_mails:
#             resp = requests.get(url, headers=headers)
#             resp.raise_for_status()
#             data = resp.json()
#             emails = data.get("value", [])
            
#             for email in emails:
#                 if email.get("hasAttachments") and len(files) < max_mails:
#                     attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{email['id']}/attachments"
#                     att_resp = requests.get(attachments_url, headers=headers)
#                     att_resp.raise_for_status()
                    
#                     for att in att_resp.json().get("value", []):
#                         if att['@odata.type'] == "#microsoft.graph.fileAttachment":
#                             files.append({
#                                 "name": att['name'],
#                                 "content": att['contentBytes']
#                             })
#                             if len(files) >= max_mails:
#                                 break
            
#             url = data.get('@odata.nextLink')
        
#         return files
#     except Exception as e:
#         st.error(f"Error accessing Outlook: {str(e)}")
#         return []

# # ====================== GOOGLE FUNCTIONS ======================
# def google_authenticate(creds_file, scopes):
#     if not creds_file:
#         st.error("Please upload credentials.json file first")
#         return None
    
#     try:
#         # Create temp file
#         with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tf:
#             tf.write(creds_file.getvalue())
#             temp_path = tf.name
        
#         flow = InstalledAppFlow.from_client_secrets_file(temp_path, scopes)
#         creds = flow.run_local_server(port=0)
        
#         # Clean up temp file
#         try:
#             os.unlink(temp_path)
#         except:
#             pass
        
#         return creds
#     except Exception as e:
#         st.error(f"Authentication failed: {str(e)}")
#         return None

# def download_drive_files(service, folder_id, max_files=10000):
#     try:
#         all_files = []
#         next_page_token = None
        
#         while len(all_files) < max_files:
#             results = service.files().list(
#                 q=f"'{folder_id}' in parents",
#                 pageSize=1000,
#                 fields="nextPageToken, files(id, name, mimeType)",
#                 pageToken=next_page_token
#             ).execute()
            
#             files = results.get('files', [])
#             all_files.extend(files)
            
#             next_page_token = results.get('nextPageToken')
#             if not next_page_token or len(all_files) >= max_files:
#                 break
        
#         return all_files
#     except HttpError as error:
#         st.error(f"Google Drive API error: {error}")
#         return []
#     except Exception as e:
#         st.error(f"Error: {str(e)}")
#         return []

# def fetch_gmail_attachments(service, max_emails):
#     try:
#         all_attachments = []
#         next_page_token = None
#         processed_count = 0
        
#         while processed_count < max_emails:
#             results = service.users().messages().list(
#                 userId='me',
#                 q='has:attachment',
#                 maxResults=min(500, max_emails - processed_count),
#                 pageToken=next_page_token
#             ).execute()
            
#             messages = results.get('messages', [])
#             if not messages:
#                 break
            
#             for msg in messages:
#                 if processed_count >= max_emails:
#                     break
                
#                 message = service.users().messages().get(
#                     userId='me',
#                     id=msg['id'],
#                     format='full'
#                 ).execute()
                
#                 parts = message.get('payload', {}).get('parts', [])
#                 for part in parts:
#                     if 'filename' in part and part['filename']:
#                         is_resume, reason = check_resume(part['filename'])
#                         if not is_resume:
#                             continue
                        
#                         attach_id = part['body'].get('attachmentId')
#                         if not attach_id:
#                             continue
                        
#                         att = service.users().messages().attachments().get(
#                             userId='me',
#                             messageId=msg['id'],
#                             id=attach_id
#                         ).execute()
                        
#                         file_data = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
#                         all_attachments.append({
#                             'name': part['filename'],
#                             'content': file_data
#                         })
                
#                 processed_count += 1
#                 if processed_count >= max_emails:
#                     break
            
#             next_page_token = results.get('nextPageToken')
#             if not next_page_token:
#                 break
        
#         return all_attachments
#     except HttpError as error:
#         st.error(f"Gmail API error: {error}")
#         return []
#     except Exception as e:
#         st.error(f"Error: {str(e)}")
#         return []

# # ====================== PROCESSING FUNCTIONS ======================
# def process_microsoft():
#     if not client_id:
#         st.error("Please enter Client ID")
#         return
    
#     try:
#         with st.spinner("🔐 Logging in to Microsoft..."):
#             scopes = ["Mail.Read"] if mode == "Outlook Attachments" else ["Files.Read.All"]
#             token, username = ms_authenticate(client_id, tenant_id, scopes)
#             headers = {"Authorization": f"Bearer {token}"}
#             st.success(f"✅ Authenticated as: {username}")
        
#         total_resumes = 0
#         start_time = time.time()
        
#         if mode == "Outlook Attachments":
#             with st.spinner(f"📨 Scanning {max_mails} Outlook emails..."):
#                 files = fetch_outlook_attachments(headers, max_mails)
#                 st.info(f"🔍 Found {len(files)} attachments")
                
#                 resume_count = 0
#                 skipped_files = []
                
#                 for file in files:
#                     is_resume, reason = check_resume(file["name"])
#                     if is_resume:
#                         new_filename = save_file(file["content"], file["name"], "outlook")
#                         st.success(f"✅ Downloaded: {new_filename}")
#                         resume_count += 1
#                     else:
#                         skipped_files.append((file['name'], reason))
                
#                 total_resumes = resume_count
#                 st.success(f"📦 Total resumes downloaded: {resume_count}/{len(files)}")
                
#                 if skipped_files:
#                     with st.expander("❌ Skipped files details"):
#                         for fname, reason in skipped_files:
#                             st.write(f"- {fname}: {reason}")
        
#         elif mode == "OneDrive Folder":
#             if not folder_name:
#                 st.error("Please enter the folder name.")
#                 return
            
#             with st.spinner("🔍 Searching OneDrive..."):
#                 all_files, download_prefix = list_onedrive_files(headers, folder_name)
#                 st.info(f"📁 Found {len(all_files)} files in folder")
                
#                 resume_count = 0
#                 skipped_files = []
                
#                 for file in all_files:
#                     if "file" not in file:
#                         continue
                    
#                     is_resume, reason = check_resume(file["name"])
#                     if is_resume:
#                         file_url = f"{download_prefix}/{file['id']}/content"
#                         resp = requests.get(file_url, headers=headers)
#                         if resp.status_code == 200:
#                             new_filename = save_file(resp.content, file['name'], "onedrive")
#                             st.success(f"✅ Downloaded: {new_filename}")
#                             resume_count += 1
#                         else:
#                             st.warning(f"❌ Failed to download: {file['name']}")
#                     else:
#                         skipped_files.append((file['name'], reason))
                
#                 total_resumes = resume_count
#                 st.success(f"📦 Total resumes downloaded: {resume_count}/{len(all_files)}")
                
#                 if skipped_files:
#                     with st.expander("❌ Skipped files details"):
#                         for fname, reason in skipped_files:
#                             st.write(f"- {fname}: {reason}")
        
#         elif mode == "SharePoint Folder":
#             if not all([domain, site_name, folder_name]):
#                 st.error("Please complete all SharePoint fields")
#                 return
            
#             with st.spinner("🔍 Searching SharePoint..."):
#                 all_files, download_prefix = list_sharepoint_files(headers, domain, site_name, folder_name)
#                 st.info(f"🏢 Found {len(all_files)} files in SharePoint")
                
#                 resume_count = 0
#                 skipped_files = []
                
#                 for file in all_files:
#                     if "file" not in file:
#                         continue
                    
#                     is_resume, reason = check_resume(file["name"])
#                     if is_resume:
#                         file_url = f"{download_prefix}/{file['id']}/content"
#                         resp = requests.get(file_url, headers=headers)
#                         if resp.status_code == 200:
#                             new_filename = save_file(resp.content, file['name'], "sharepoint")
#                             st.success(f"✅ Downloaded: {new_filename}")
#                             resume_count += 1
#                         else:
#                             st.warning(f"❌ Failed to download: {file['name']}")
#                     else:
#                         skipped_files.append((file['name'], reason))
                
#                 total_resumes = resume_count
#                 st.success(f"📦 Total resumes downloaded: {resume_count}/{len(all_files)}")
                
#                 if skipped_files:
#                     with st.expander("❌ Skipped files details"):
#                         for fname, reason in skipped_files:
#                             st.write(f"- {fname}: {reason}")
        
#         elapsed = time.time() - start_time
#         st.balloons()
#         st.success(f"🎉 Microsoft download completed! Resumes saved to: {DOWNLOAD_DIR}")
#         st.info(f"⏱️ Processed {total_resumes} resumes in {elapsed:.2f} seconds")
    
#     except Exception as e:
#         st.error(f"❌ Error: {str(e)}")

# def process_google():
#     if not creds_file:
#         st.error("Please upload credentials.json file")
#         return
    
#     try:
#         total_resumes = 0
#         start_time = time.time()
        
#         if google_mode == "Google Drive Folder":
#             if not folder_url:
#                 st.error("Please enter Google Drive folder URL")
#                 return
            
#             with st.spinner("🔐 Authenticating with Google Drive API..."):
#                 creds = google_authenticate(creds_file, ['https://www.googleapis.com/auth/drive.readonly'])
#                 if not creds:
#                     return
                
#                 drive = build('drive', 'v3', credentials=creds)
#                 folder_id = folder_url.split("/")[-1].split("?")[0]
                
#                 with st.spinner(f"🔍 Scanning Google Drive folder..."):
#                     all_files = download_drive_files(drive, folder_id, max_files=10000)
#                     st.info(f"📁 Found {len(all_files)} files in folder")
                    
#                     resume_count = 0
#                     skipped_files = []
                    
#                     for file in all_files:
#                         is_resume, reason = check_resume(file['name'])
#                         if is_resume:
#                             request = drive.files().get_media(fileId=file['id'])
#                             fh = io.BytesIO()
#                             downloader = MediaIoBaseDownload(fh, request)
#                             done = False
#                             while not done:
#                                 _, done = downloader.next_chunk()
                            
#                             new_filename = save_file(fh.getvalue(), file['name'], "drive")
#                             st.success(f"✅ Downloaded: {new_filename}")
#                             resume_count += 1
#                         else:
#                             skipped_files.append((file['name'], reason))
                    
#                     total_resumes = resume_count
#                     st.success(f"📦 Total resumes downloaded: {resume_count}/{len(all_files)}")
                    
#                     if skipped_files:
#                         with st.expander("❌ Skipped files details"):
#                             for fname, reason in skipped_files:
#                                 st.write(f"- {fname}: {reason}")
        
#         else:  # Gmail Attachments
#             with st.spinner("🔐 Authenticating with Gmail API..."):
#                 creds = google_authenticate(creds_file, ['https://www.googleapis.com/auth/gmail.readonly'])
#                 if not creds:
#                     return
                
#                 gmail = build('gmail', 'v1', credentials=creds)
                
#                 with st.spinner(f"📨 Scanning {max_emails} emails..."):
#                     attachments = fetch_gmail_attachments(gmail, max_emails)
#                     st.info(f"📧 Found {len(attachments)} attachments")
                    
#                     resume_count = 0
#                     skipped_files = []
                    
#                     for att in attachments:
#                         is_resume, reason = check_resume(att['name'])
#                         if is_resume:
#                             new_filename = save_file(att['content'], att['name'], "gmail")
#                             st.success(f"✅ Downloaded: {new_filename}")
#                             resume_count += 1
#                         else:
#                             skipped_files.append((att['name'], reason))
                    
#                     total_resumes = resume_count
#                     st.success(f"📦 Total resumes downloaded: {resume_count}/{len(attachments)}")
                    
#                     if skipped_files:
#                         with st.expander("❌ Skipped files details"):
#                             for fname, reason in skipped_files:
#                                 st.write(f"- {fname}: {reason}")
        
#         elapsed = time.time() - start_time
#         st.balloons()
#         st.success(f"🎉 Google download completed! Resumes saved to: {DOWNLOAD_DIR}")
#         st.info(f"⏱️ Processed {total_resumes} resumes in {elapsed:.2f} seconds")
    
#     except Exception as e:
#         st.error(f"❌ Error: {str(e)}")

# # ====================== MAIN EXECUTION ======================
# if ms_run_button:
#     process_microsoft()

# if google_run_button:
#     process_google()

# # Footer
# st.markdown("---")
# st.markdown("""
# <div style="text-align:center;color:gray;font-size:0.8em;">
#     <p>🌟 HR AI Resume Downloader | Supports Microsoft 365 and Google Workspace</p>
# </div>
# """, unsafe_allow_html=True)



### these final for fetching the data   



import streamlit as st
import os
import requests
import tempfile
import base64
import io
from msal import PublicClientApplication
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError
import time
from pathlib import Path

# ====================== INSTRUCTION DROPDOWN ======================
st.set_page_config(page_title="🌟 Ultimate Resume Downloader", layout="wide", page_icon="📥")

st.title("🌟 Ultimate Resume Downloader")

instruction_option = st.selectbox(
    "Show credential setup instructions for:",
    ("Select...", "Microsoft (Outlook)", "Google (credentials.json)")
)

if instruction_option == "Microsoft (Outlook)":
    st.info("""
    ### How to get Microsoft (Outlook) Credentials

    🟩 Step 1: Register the Azure App

1. Go to: [https://portal.azure.com]
2. Navigate to:
   ➤ Azure Entra ID
   ➤ App registrations
   ➤ Click New registration
3. Fill out the form:

Name: Any name
   Supported account types: Choose any (usually "Accounts in this organizational directory only")
   Redirect URI: Click "Add a platform" → Select **Public client (mobile & desktop)**
     ➤ Enter: `http://localhost`

4. Click Register

🟩 Step 2: Configure Authentication

1. Go to your app → Authentication tab
2. Add this Redirect URI under *Mobile and desktop applications*:
 Selecte  Add URL :  Add the below URL
    `http://localhost:8000`
3. Scroll to **Advanced settings**

   * Set **“Allow public client flows”** to **Yes**

4. Click **Save**

🟩 Step 3: Add Read-Only API Permissions

Go to your app → API permissions→ + Add a permission → Microsoft Graph → Delegated permissions.

Add the following  read-only permissions:

📧 Outlook (Mail)

* Mail.Read
* Mail.Read.Shared
* Mail.ReadBasic
* Mail.ReadBasic.Shared
* MailboxFolder.Read

Click Add Permissions

### 📁 OneDrive / Files

* `Files.Read`
* `Files.Read.All`
* `Files.Read.Selected`

### 🌐 SharePoint

* `Sites.Read.All`
* `Sites.Selected`

👤 User Info

* `User.Read`
* `User.ReadBasic.All`
* `email`


✅ Once done, click  Grant admin consent (if available)

🟩 Step 4: Get Your Azure App Credentials

Go to the **Overview** tab of your app
Copy the following values:

Application (client) ID** → use as `client_id`
Directory (tenant) ID** → use as `tenant_id`

These will be used in your code or tools (e.g., MSAL).
    """)
elif instruction_option == "Google (credentials.json)":
    st.info("""
    ### How to get Google credentials.json

✅ STEP 1: Open Google Cloud Console 🌐
Open your browser
Go to: 👉 https://console.cloud.google.com
Sign in with your Gmail (the one you want to use for downloading resumes)

✅ STEP 2: Create a New Project 🧱
At the very top bar, click "Select a project"
Then click “New Project”
Give your project a name: ResumeDownloader
Click Create
After it’s done, click the project name to select it

✅ STEP 3: Enable Gmail & Google Drive APIs 📬📁
Now we’ll turn on the tools we need:
🔸 First: Enable Gmail API
On the left side, click:
APIs & Services → Library
Search: Gmail API
Click on it → Click Enable
🔸 Then: Enable Google Drive API
Search: Google Drive API
Click on it → Click Enable

✅ STEP 4: Set Up the Consent Screen 📃
In the sidebar, go to:
APIs & Services → OAuth consent screen
you’ll see a button: “Get Started”
Click it ✅
It will ask for some details – just follow the instructions and click next or continue when needed.

 then click create OAuth client
 then it will application type choise  desktop app  after enter any name
then click create download json

✅ STEP 5: Add Yourself in Audience Option 👤
You need to allow yourself to use this app:

On the OAuth consent screen page → find “Audience” section (scroll down)  their  will be Test user  below that Add users
Click + Add Users
Type your Gmail ID(same id what you register )
Click Save
✅ STEP 6: Add Access Scopes 🔐
You must manually tell Google what parts of your Gmail and Drive the app can read:

S→ Go to: Data Access   click on the button Add or remove Scopes
Scroll to Manually add scopes section

Copy and paste the following two lines (one by one):
https://www.googleapis.com/auth/drive.readonly
https://www.googleapis.com/auth/gmail.readonly
Click “Add to Table” ➕
Then click Update ✅
Then click Save and Continue

    *Keep your `credentials.json` safe and upload it in the Google section of the app.*
    """)

st.markdown("""
<div style="background-color:#0e1117;padding:20px;border-radius:10px;margin-bottom:20px">
    <h3 style="color:white;text-align:center;">Download resumes from SharePoint, OneDrive, Outlook, Google Drive & Gmail</h3>
</div>
""", unsafe_allow_html=True)

# ====================== ORIGINAL APP CODE BELOW ======================

# Constants
RESUME_TYPES = ['.pdf', '.docx', '.doc']
RESUME_KEYWORDS = ['resume', 'cv', 'curriculum vitae', 'bio data']
EXCLUDE_KEYWORDS = ['offer', 'letter', 'terms', 'conditions', 'contest', 'referral', 'policy', 'agreement']
DOWNLOAD_DIR = r"C:\Users\jaswa\OneDrive\Desktop\HR_AI_Resume\mail_download\downloaded_resumes-3" # Use raw string for path

# Ensure download directory exists
try:
    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)
    st.sidebar.success(f"Download directory created/verified: {DOWNLOAD_DIR}")
except Exception as e:
    st.sidebar.error(f"Error creating download directory: {e}")
    DOWNLOAD_DIR = None # Indicate failure

# Initialize session state for counts (useful if the user interacts while processing)
if 'ms_downloaded_count' not in st.session_state:
    st.session_state.ms_downloaded_count = 0
if 'ms_skipped_count' not in st.session_state:
    st.session_state.ms_skipped_count = 0
if 'google_downloaded_count' not in st.session_state:
    st.session_state.google_downloaded_count = 0
if 'google_skipped_count' not in st.session_state:
    st.session_state.google_skipped_count = 0

# Tabs for provider selection
ms_tab, google_tab = st.tabs(["🔷 Microsoft", "🔷 Google"])

# ====================== MICROSOFT SECTION ======================
with ms_tab:
    st.subheader("🔐 Microsoft Authentication")
    col1, col2 = st.columns(2)
    with col1:
        client_id = st.text_input("Client ID", value="5a56d5a6-fa4c-46a4-2f48-941f44b012a", key="ms_client_id")
    with col2:
        tenant_id = st.text_input("Tenant ID", value="common", key="ms_tenant_id")

    mode = st.radio("Select source", ["SharePoint Folder", "OneDrive Folder", "Outlook Attachments"], key="ms_mode")

    if mode == "SharePoint Folder":
        st.subheader("🏢 SharePoint Configuration")
        domain = st.text_input("SharePoint domain (e.g., company.sharepoint.com)",
                             value="buildionsportld.sharepoint.com",
                             key="sharepoint_domain")
        site_name = st.text_input("Site name (e.g., HRTeamSite)",
                                 value="NEUZEMATISOLUTIONSPYTLTD",
                                 key="sharepoint_site")
        folder_name = st.text_input("Folder name (e.g., Resumes2024)",
                                  value="AL RESUME_FILTER",
                                  key="sharepoint_folder")
    elif mode == "OneDrive Folder":
        st.subheader("📁 OneDrive Configuration")
        folder_name = st.text_input("Folder name in OneDrive (e.g., MyResumes)", key="onedrive_folder")
    else: # Outlook Attachments
        st.subheader("📧 Outlook Configuration")
        max_mails = st.number_input("Number of emails to scan", min_value=1, max_value=10000, value=1000, key="ms_max_mails")

    ms_run_button = st.button("🚀 Download Microsoft Resumes", key="ms_run")

# ====================== GOOGLE SECTION ======================
with google_tab:
    st.subheader("🔐 Google Authentication")
    creds_file = st.file_uploader("Upload credentials.json", type="json", key="google_creds")

    google_mode = st.radio("Select source", ["Google Drive Folder", "Gmail Attachments"], key="google_mode")

    if google_mode == "Google Drive Folder":
        st.subheader("📁 Google Drive Configuration")
        folder_url = st.text_input("Google Drive Folder URL", key="drive_url")
    else: # Gmail Attachments
        st.subheader("📧 Gmail Configuration")
        max_emails = st.number_input("Number of emails to scan", min_value=1, max_value=10000, value=1000, key="gmail_max_emails")

    google_run_button = st.button("🚀 Download Google Resumes", key="google_run")

# ====================== COMMON FUNCTIONS ======================
def check_resume(filename):
    """Enhanced resume detection with detailed rejection reasons"""
    if not filename:
        return False, "No filename"

    name = filename.lower()
    ext = os.path.splitext(filename)[1].lower()

    # Check exclusion keywords
    for word in EXCLUDE_KEYWORDS:
        if word in name:
            return False, f"Excluded keyword: '{word}'"

    # Check file extension
    if ext not in RESUME_TYPES:
        return False, f"Invalid file type: {ext}"

    # Check resume keywords
    for word in RESUME_KEYWORDS:
        if word in name:
            return True, ""

    # Check common patterns
    patterns = ['_resume', '-resume', 'resume_', 'resume-', '_cv', '-cv', 'cv_', 'cv-']
    if any(p in name for p in patterns):
        return True, ""

    return False, "No resume keywords found"

def save_file(content, filename, source):
    """Save file to download directory with collision handling"""
    if DOWNLOAD_DIR is None:
        st.error("Download directory not available. Cannot save file.")
        return None

    counter = 1
    name, ext = os.path.splitext(filename)
    new_filename = filename

    while os.path.exists(os.path.join(DOWNLOAD_DIR, new_filename)):
        new_filename = f"{name}_{counter}{ext}"
        counter += 1

    filepath = os.path.join(DOWNLOAD_DIR, new_filename)

    try:
        # Outlook content is base64 encoded, others are raw bytes
        if source == "outlook":
            with open(filepath, 'wb') as f:
                f.write(base64.b64decode(content))
        else:
            with open(filepath, 'wb') as f:
                f.write(content)
        return new_filename
    except Exception as e:
        st.error(f"Error saving file {new_filename}: {e}")
        return None


# ====================== MICROSOFT FUNCTIONS ======================
def ms_authenticate(client_id, tenant_id, scopes):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = PublicClientApplication(client_id, authority=authority)

    # Clear any cached accounts (optional, helps ensure interactive prompt)
    accounts = app.get_accounts()
    for account in accounts:
        app.remove_account(account)

    # Interactive authentication
    st.info("Opening Microsoft login window in your browser...")
    try:
        result = app.acquire_token_interactive(scopes=scopes, prompt="select_account")
    except Exception as e:
         raise Exception(f"Authentication failed: {e}")

    if "access_token" in result:
        return result["access_token"], result.get("account", {}).get("username")
    else:
         # More detailed error handling from MSAL result
        error_description = result.get("error_description", "Unknown error during token acquisition.")
        correlation_id = result.get("correlation_id", "N/A")
        raise Exception(f"Authentication failed. Error: {result.get('error')}, Description: {error_description}, Correlation ID: {correlation_id}")


def get_site_and_drive_ids(headers, hostname, site_name):
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_name}"
    site_resp = requests.get(site_url, headers=headers)
    site_resp.raise_for_status() # Raise HTTPError for bad responses
    site_id = site_resp.json()["id"]

    # Get the default document library drive
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_resp = requests.get(drives_url, headers=headers)
    drives_resp.raise_for_status()

    # Find the 'Documents' drive or the first available drive
    drives = drives_resp.json().get("value", [])
    drive_id = None
    for drive in drives:
        # Look for 'Documents' or similar default drive
        if drive.get('name') == 'Documents': # Often the default doc library
             drive_id = drive['id']
             break
        # Fallback to the first drive if 'Documents' isn't found or accessible
        if not drive_id and drives:
             drive_id = drives[0]['id']

    if not drive_id:
         raise Exception("Could not find a drive for the specified SharePoint site.")

    return site_id, drive_id

def list_sharepoint_files(headers, domain, site_name, folder_name, status_placeholder):
    try:
        status_placeholder.info(f"Resolving SharePoint site '{site_name}' and drive...")
        site_id, drive_id = get_site_and_drive_ids(headers, domain, site_name)
        status_placeholder.info(f"Accessing folder '{folder_name}'...")

        # Get the folder item details first
        folder_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_name}"
        folder_resp = requests.get(folder_url, headers=headers)
        folder_resp.raise_for_status()
        folder_id = folder_resp.json()["id"]

        status_placeholder.info(f"Listing files in folder ID '{folder_id}'...")
        # List children of the folder, can handle pagination
        files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children"

        all_files = []
        while files_url:
            files_resp = requests.get(files_url, headers=headers)
            files_resp.raise_for_status()
            data = files_resp.json()

            items = data.get("value", [])
            # Filter only file items (skip folders, etc.)
            all_files.extend([item for item in items if 'file' in item])

            files_url = data.get('@odata.nextLink') # Get next page URL
            if files_url:
                 status_placeholder.info(f"Fetching next page of files... ({len(all_files)} files found so far)")

        status_placeholder.success(f"Finished listing. Found {len(all_files)} potential files.")
        return all_files, f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items"

    except Exception as e:
        status_placeholder.error(f"Error accessing SharePoint: {str(e)}")
        st.error(f"Error accessing SharePoint: {str(e)}") # Also show a persistent error
        return [], ""

def list_onedrive_files(headers, folder_name, status_placeholder):
    try:
        status_placeholder.info(f"Accessing OneDrive folder '{folder_name}'...")
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_name}:/children"

        all_files = []
        while url:
            resp = requests.get(url, headers=headers)
            resp.raise_for_status() # Raise HTTPError for bad responses
            data = resp.json()

            items = data.get("value", [])
             # Filter only file items (skip folders, etc.)
            all_files.extend([item for item in items if 'file' in item])

            url = data.get('@odata.nextLink') # Get next page URL
            if url:
                 status_placeholder.info(f"Fetching next page of files... ({len(all_files)} files found so far)")

        status_placeholder.success(f"Finished listing. Found {len(all_files)} potential files.")
        return all_files, "https://graph.microsoft.com/v1.0/me/drive/items"
    except Exception as e:
        status_placeholder.error(f"Error accessing OneDrive: {str(e)}")
        st.error(f"Error accessing OneDrive: {str(e)}") # Also show a persistent error
        return [], ""

def fetch_outlook_attachments_stream(headers, max_mails, status_placeholder):
    """Fetches Outlook attachments and yields them one by one."""
    try:
        status_placeholder.info(f"Searching for emails with attachments (up to {max_mails})...")

        # Start with the first page of messages
        url = f"https://graph.microsoft.com/v1.0/me/messages?$top={min(100, max_mails)}&$orderby=receivedDateTime desc&$select=id,subject,hasAttachments" # Fetch messages in chunks

        processed_emails_count = 0
        yielded_attachments_count = 0

        while url and processed_emails_count < max_mails:
            resp = requests.get(url, headers=headers)
            resp.raise_for_status() # Raise HTTPError for bad responses
            data = resp.json()
            messages = data.get("value", [])

            if not messages:
                status_placeholder.info("No more emails found with attachments.")
                break # No more messages

            for email in messages:
                if processed_emails_count >= max_mails:
                    break # Stop if max_mails reached

                processed_emails_count += 1
                status_placeholder.info(f"Processing email {processed_emails_count}/{max_mails}: '{email.get('subject', 'No Subject')}'")

                if email.get("hasAttachments"):
                    attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{email['id']}/attachments"
                    att_resp = requests.get(attachments_url, headers=headers)
                    att_resp.raise_for_status()

                    for att in att_resp.json().get("value", []):
                         # Only process file attachments, skip inline images etc.
                        if att['@odata.type'] == "#microsoft.graph.fileAttachment":
                             yielded_attachments_count += 1
                             # Yield the attachment info for processing
                             yield {
                                 "name": att['name'],
                                 "content": att['contentBytes'],
                                 "email_subject": email.get('subject', 'No Subject'),
                                 "email_id": email['id'],
                                 "attachment_id": att['id']
                             }
                             status_placeholder.info(f"Yielded attachment {yielded_attachments_count}: {att['name']}")


            url = data.get('@odata.nextLink') # Get URL for the next page of messages

        if not url and processed_emails_count < max_mails:
             status_placeholder.info(f"Finished scanning available emails ({processed_emails_count} processed).")


    except Exception as e:
        status_placeholder.error(f"Error during Outlook attachment fetch: {str(e)}")
        st.error(f"Error during Outlook attachment fetch: {str(e)}") # Also show a persistent error


# ====================== GOOGLE FUNCTIONS ======================
def google_authenticate(creds_file, scopes):
    if not creds_file:
        st.error("Please upload credentials.json file first")
        return None

    try:
        # Create temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tf:
            tf.write(creds_file.getvalue())
            temp_path = tf.name

        # Use a fixed port or 0 for automatic selection
        flow = InstalledAppFlow.from_client_secrets_file(temp_path, scopes)
        st.info("Opening Google login window in your browser...")
        creds = flow.run_local_server(port=0, success_message="Authentication successful! You can close this tab.")

        # Clean up temp file
        try:
            os.unlink(temp_path)
        except Exception as e:
            st.warning(f"Could not delete temporary credentials file: {e}")

        return creds
    except Exception as e:
        st.error(f"Authentication failed: {str(e)}")
        return None

def list_drive_files_in_folder(service, folder_id, status_placeholder):
    """Lists files in a Google Drive folder, handling pagination."""
    try:
        status_placeholder.info(f"Listing files in Google Drive folder ID: {folder_id}...")

        all_files = []
        next_page_token = None

        while True:
            results = service.files().list(
                q=f"'{folder_id}' in parents and trashed = false", # Only include files directly in the folder and not trashed
                pageSize=1000, # Max number of results per page
                fields="nextPageToken, files(id, name, mimeType)",
                # --- CORRECTED LINE BELOW ---
                pageToken=next_page_token
                # --------------------------
            ).execute()

            files = results.get('files', [])
            all_files.extend(files)

            next_page_token = results.get('nextPageToken')
            if not next_page_token:
                break
            status_placeholder.info(f"Fetching next page of files... ({len(all_files)} files found so far)")

        status_placeholder.success(f"Finished listing. Found {len(all_files)} potential files in folder.")
        return all_files

    except HttpError as error:
        status_placeholder.error(f"Google Drive API error: {error}")
        st.error(f"Google Drive API error: {error}") # Persistent error
        return []
    except Exception as e:
        status_placeholder.error(f"Error listing Google Drive files: {str(e)}")
        st.error(f"Error listing Google Drive files: {str(e)}") # Persistent error
        return []


def fetch_gmail_attachments_stream(service, max_emails, status_placeholder):
    """Fetches Gmail attachments from emails and yields them one by one."""
    try:
        status_placeholder.info(f"Searching for emails with attachments (up to {max_emails})...")

        # Start with the first page of messages that have attachments
        next_page_token = None
        processed_emails_count = 0
        yielded_attachments_count = 0

        while processed_emails_count < max_emails:
            results = service.users().messages().list(
                userId='me',
                q='has:attachment',
                maxResults=min(500, max_emails - processed_emails_count), # Fetch messages in chunks
                pageToken=next_page_token
            ).execute()

            messages = results.get('messages', [])
            if not messages:
                status_placeholder.info("No more emails found with attachments.")
                break # No more messages

            for msg_summary in messages:
                if processed_emails_count >= max_emails:
                    break # Stop if max_emails reached

                processed_emails_count += 1

                try:
                    # Get the full message details
                    message = service.users().messages().get(
                        userId='me',
                        id=msg_summary['id'],
                        format='full' # Request full message to get attachments
                    ).execute()

                    subject = next((h['value'] for h in message['payload']['headers'] if h['name'] == 'Subject'), 'No Subject')
                    status_placeholder.info(f"Processing email {processed_emails_count}/{max_emails}: '{subject}'")

                    parts = message.get('payload', {}).get('parts', [])

                    # Recursively find parts with 'filename'
                    def find_parts_with_filename(parts_list):
                         found_parts = []
                         if parts_list:
                             for part in parts_list:
                                 if 'filename' in part and part['filename']:
                                     found_parts.append(part)
                                 if 'parts' in part: # Check nested parts (e.g., multipart/mixed)
                                     found_parts.extend(find_parts_with_filename(part['parts']))
                         return found_parts

                    attachments_parts = find_parts_with_filename(parts)

                    for part in attachments_parts:
                        # Check if it's likely a file attachment (not inline, not related)
                        # This is a heuristic; checking disposition is better but 'full' format might not always include it clearly
                        # A common pattern for attachments is missing 'Content-ID' or having 'attachment' disposition (not always present in 'full')
                        # For simplicity with 'full' format, we rely on `filename` and hope it's not an inline image.
                        # A more robust approach might require `format='metadata'` first, then `format='full'` only for emails with attachments, and inspecting headers/disposition more carefully.
                        # However, given the request's focus on resumes, the `check_resume` will filter most non-relevant files.

                        attach_id = part['body'].get('attachmentId')
                        file_size = part['body'].get('size', 0) # Get size if available

                        if not attach_id:
                            # This might be an inline attachment or other non-downloadable part in this format
                            status_placeholder.warning(f"  - Skipping non-downloadable part: {part.get('filename', 'Unnamed file')} in email '{subject}'")
                            continue # Skip parts without an attachmentId

                        status_placeholder.info(f"  - Potential attachment: {part['filename']} (Size: {file_size} bytes)")

                        try:
                             # Fetch the attachment data
                             att = service.users().messages().attachments().get(
                                 userId='me',
                                 messageId=msg_summary['id'],
                                 id=attach_id
                             ).execute()

                             file_data = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
                             yielded_attachments_count += 1
                             # Yield the attachment info and content
                             yield {
                                 'name': part['filename'],
                                 'content': file_data,
                                 'email_subject': subject,
                                 'email_id': msg_summary['id'],
                                 'attachment_id': attach_id # Keep ID for context if needed
                             }
                             status_placeholder.info(f"  - Yielded attachment {yielded_attachments_count}: {part['filename']}")

                        except HttpError as att_error:
                            status_placeholder.warning(f"  - Could not fetch attachment '{part.get('filename', 'Unnamed file')}' (ID: {attach_id}) from email '{subject}': {att_error}")
                        except Exception as att_e:
                            status_placeholder.warning(f"  - Error processing attachment '{part.get('filename', 'Unnamed file')}': {att_e}")


                except HttpError as msg_error:
                     status_placeholder.warning(f"  - Could not fetch full message (ID: {msg_summary['id']}): {msg_error}")
                except Exception as msg_e:
                     status_placeholder.warning(f"  - Error processing message (ID: {msg_summary['id']}): {msg_e}")


            next_page_token = results.get('nextPageToken') # Get URL for the next page of messages
            if not next_page_token:
                status_placeholder.info(f"Finished scanning available emails ({processed_emails_count} processed).")
                break # No more pages


        if processed_emails_count < max_emails:
             status_placeholder.info(f"Reached end of available emails after processing {processed_emails_count}.")


    except HttpError as error:
        status_placeholder.error(f"Gmail API error during message list/fetch: {error}")
        st.error(f"Gmail API error during message list/fetch: {error}") # Persistent error
    except Exception as e:
        status_placeholder.error(f"General error during Gmail processing: {str(e)}")
        st.error(f"General error during Gmail processing: {str(e)}") # Persistent error


# ====================== PROCESSING FUNCTIONS ======================
def process_microsoft():
    if DOWNLOAD_DIR is None:
        st.error("Download directory is not set up correctly. Aborting.")
        return

    if not client_id or not tenant_id:
        st.error("Please enter Client ID and Tenant ID")
        return

    st.session_state.ms_downloaded_count = 0
    st.session_state.ms_skipped_count = 0

    # Create placeholders for real-time updates
    status_placeholder = st.empty()
    count_placeholder = st.empty()
    details_expander = st.expander("Processing Details (Microsoft)")


    try:
        with st.spinner("🔐 Logging in to Microsoft..."):
            scopes = ["Mail.Read"] if mode == "Outlook Attachments" else ["Files.Read.All", "Sites.Read.All"] # Added Sites.Read.All for SP
            token, username = ms_authenticate(client_id, tenant_id, scopes)
            headers = {"Authorization": f"Bearer {token}"}
            status_placeholder.success(f"✅ Authenticated as: {username}")

        start_time = time.time()

        if mode == "Outlook Attachments":
            status_placeholder.info(f"Starting Outlook attachment scan (up to {max_mails} emails)...")
            # Use the streaming fetch function
            attachment_generator = fetch_outlook_attachments_stream(headers, max_mails, status_placeholder)

            for attachment_info in attachment_generator:
                 count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")

                 is_resume, reason = check_resume(attachment_info["name"])

                 if is_resume:
                     new_filename = save_file(attachment_info["content"], attachment_info["name"], "outlook")
                     if new_filename:
                         details_expander.success(f"✅ Saved: {new_filename} (from '{attachment_info.get('email_subject', 'Email')}')")
                         st.session_state.ms_downloaded_count += 1
                     else:
                          details_expander.warning(f"❌ Failed to save: {attachment_info['name']}")
                 else:
                     details_expander.info(f"➡️ Skipped: {attachment_info['name']} - {reason}")
                     st.session_state.ms_skipped_count += 1

            # Update final counts
            count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")


        elif mode == "OneDrive Folder":
            if not folder_name:
                st.error("Please enter the folder name.")
                return

            status_placeholder.info("🔍 Listing files in OneDrive folder...")
            all_files, download_prefix = list_onedrive_files(headers, folder_name, status_placeholder)

            if not all_files:
                 status_placeholder.warning("No files found in the specified OneDrive folder.")
                 st.warning("No files found in the specified OneDrive folder.")
                 return # Exit if no files are found

            total_files = len(all_files)
            for i, file in enumerate(all_files):
                # Update status and counts dynamically
                status_placeholder.info(f"Processing file {i+1}/{total_files}: {file['name']}")
                count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")

                is_resume, reason = check_resume(file["name"])

                if is_resume:
                    file_url = f"{download_prefix}/{file['id']}/content"
                    try:
                        resp = requests.get(file_url, headers=headers, stream=True) # Use stream=True for potentially large files
                        resp.raise_for_status() # Raise HTTPError for bad responses

                        # Download content chunk by chunk (optional, but good practice)
                        file_content = b""
                        for chunk in resp.iter_content(chunk_size=8192):
                            file_content += chunk

                        new_filename = save_file(file_content, file['name'], "onedrive")
                        if new_filename:
                            details_expander.success(f"✅ Saved: {new_filename}")
                            st.session_state.ms_downloaded_count += 1
                        else:
                             details_expander.warning(f"❌ Failed to save: {file['name']}")

                    except Exception as e:
                        details_expander.warning(f"❌ Failed to download {file['name']}: {e}")
                        st.warning(f"❌ Failed to download {file['name']}: {e}") # Persistent warning

                else:
                    details_expander.info(f"➡️ Skipped: {file['name']} - {reason}")
                    st.session_state.ms_skipped_count += 1

            # Update final counts
            count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")


        elif mode == "SharePoint Folder":
            if not all([domain, site_name, folder_name]):
                st.error("Please complete all SharePoint fields")
                return

            status_placeholder.info("🔍 Listing files in SharePoint folder...")
            all_files, download_prefix = list_sharepoint_files(headers, domain, site_name, folder_name, status_placeholder)

            if not all_files:
                 status_placeholder.warning("No files found in the specified SharePoint folder.")
                 st.warning("No files found in the specified SharePoint folder.")
                 return # Exit if no files are found

            total_files = len(all_files)
            for i, file in enumerate(all_files):
                 # Update status and counts dynamically
                status_placeholder.info(f"Processing file {i+1}/{total_files}: {file['name']}")
                count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")

                is_resume, reason = check_resume(file["name"])

                if is_resume:
                    file_url = f"{download_prefix}/{file['id']}/content"
                    try:
                        resp = requests.get(file_url, headers=headers, stream=True) # Use stream=True
                        resp.raise_for_status()

                        file_content = b""
                        for chunk in resp.iter_content(chunk_size=8192):
                            file_content += chunk

                        new_filename = save_file(file_content, file['name'], "sharepoint")
                        if new_filename:
                            details_expander.success(f"✅ Saved: {new_filename}")
                            st.session_state.ms_downloaded_count += 1
                        else:
                             details_expander.warning(f"❌ Failed to save: {file['name']}")

                    except Exception as e:
                        details_expander.warning(f"❌ Failed to download {file['name']}: {e}")
                        st.warning(f"❌ Failed to download {file['name']}: {e}") # Persistent warning

                else:
                    details_expander.info(f"➡️ Skipped: {file['name']} - {reason}")
                    st.session_state.ms_skipped_count += 1

            # Update final counts
            count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")


        elapsed = time.time() - start_time
        status_placeholder.success(f"🎉 Microsoft download completed! Processed {st.session_state.ms_downloaded_count + st.session_state.ms_skipped_count} items.")
        count_placeholder.empty() # Clear the dynamic count
        st.balloons()
        st.success(f"📦 Total resumes downloaded: {st.session_state.ms_downloaded_count}")
        st.info(f"⏱️ Processed {st.session_state.ms_downloaded_count + st.session_state.ms_skipped_count} total items in {elapsed:.2f} seconds. Resumes saved to: {DOWNLOAD_DIR}")


    except Exception as e:
        status_placeholder.error(f"❌ An error occurred: {str(e)}")
        st.error(f"❌ An error occurred: {str(e)}") # Also show a persistent error


def process_google():
    if DOWNLOAD_DIR is None:
        st.error("Download directory is not set up correctly. Aborting.")
        return

    if not creds_file:
        st.error("Please upload credentials.json file")
        return

    st.session_state.google_downloaded_count = 0
    st.session_state.google_skipped_count = 0

    # Create placeholders for real-time updates
    status_placeholder = st.empty()
    count_placeholder = st.empty()
    details_expander = st.expander("Processing Details (Google)")

    try:
        start_time = time.time()

        if google_mode == "Google Drive Folder":
            if not folder_url:
                st.error("Please enter Google Drive folder URL")
                return

            with st.spinner("🔐 Authenticating with Google Drive API..."):
                # Typo fixed here too, should be creds_file
                creds = google_authenticate(creds_file, ['https://www.googleapis.com/auth/drive.readonly'])
                if not creds:
                    status_placeholder.error("Authentication failed.")
                    return
                status_placeholder.success("✅ Authenticated with Google Drive.")

            drive = build('drive', 'v3', credentials=creds)
            try:
                folder_id = folder_url.split("/")[-1].split("?")[0]
            except:
                 st.error("Invalid Google Drive folder URL format.")
                 status_placeholder.error("Invalid Google Drive folder URL format.")
                 return

            status_placeholder.info("🔍 Listing files in Google Drive folder...")
            all_files = list_drive_files_in_folder(drive, folder_id, status_placeholder)

            if not all_files:
                 status_placeholder.warning("No files found in the specified Google Drive folder.")
                 st.warning("No files found in the specified Google Drive folder.")
                 return # Exit if no files are found


            total_files = len(all_files)
            for i, file in enumerate(all_files):
                # Update status and counts dynamically
                status_placeholder.info(f"Processing file {i+1}/{total_files}: {file['name']}")
                count_placeholder.text(f"Downloaded: {st.session_state.google_downloaded_count}, Skipped: {st.session_state.google_skipped_count}")

                is_resume, reason = check_resume(file['name'])

                if is_resume:
                    request = drive.files().get_media(fileId=file['id'])
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    try:
                        while not done:
                            status, done = downloader.next_chunk()
                            # Optional: Update progress bar here if needed
                            # print(f"Download {int(status.progress() * 100)}%.")

                        file_content = fh.getvalue()

                        new_filename = save_file(file_content, file['name'], "drive")
                        if new_filename:
                            details_expander.success(f"✅ Saved: {new_filename}")
                            st.session_state.google_downloaded_count += 1
                        else:
                             details_expander.warning(f"❌ Failed to save: {file['name']}")

                    except HttpError as download_error:
                         details_expander.warning(f"❌ Failed to download {file['name']}: {download_error}")
                         st.warning(f"❌ Failed to download {file['name']}: {download_error}") # Persistent warning
                    except Exception as download_e:
                         details_expander.warning(f"❌ Error during download of {file['name']}: {download_e}")
                         st.warning(f"❌ Error during download of {file['name']}: {download_e}") # Persistent warning


                else:
                    details_expander.info(f"➡️ Skipped: {file['name']} - {reason}")
                    st.session_state.google_skipped_count += 1

            # Update final counts
            count_placeholder.text(f"Downloaded: {st.session_state.google_downloaded_count}, Skipped: {st.session_state.google_skipped_count}")


        else:  # Gmail Attachments
            with st.spinner("🔐 Authenticating with Gmail API..."):
                creds = google_authenticate(creds_file, ['https://www.googleapis.com/auth/gmail.readonly'])
                if not creds:
                    status_placeholder.error("Authentication failed.")
                    return
                status_placeholder.success("✅ Authenticated with Gmail.")

            gmail = build('gmail', 'v1', credentials=creds)

            status_placeholder.info(f"Starting Gmail attachment scan (up to {max_emails} emails)...")
            # Use the streaming fetch function
            attachment_generator = fetch_gmail_attachments_stream(gmail, max_emails, status_placeholder)

            for attachment_info in attachment_generator:
                 count_placeholder.text(f"Downloaded: {st.session_state.google_downloaded_count}, Skipped: {st.session_state.google_skipped_count}")

                 is_resume, reason = check_resume(attachment_info['name'])

                 if is_resume:
                     # The content is already fetched and decoded by the generator
                     new_filename = save_file(attachment_info['content'], attachment_info['name'], "gmail")
                     if new_filename:
                         details_expander.success(f"✅ Saved: {new_filename} (from '{attachment_info.get('email_subject', 'Email')}')")
                         st.session_state.google_downloaded_count += 1
                     else:
                         details_expander.warning(f"❌ Failed to save: {attachment_info['name']}")
                 else:
                     details_expander.info(f"➡️ Skipped: {attachment_info['name']} - {reason}")
                     st.session_state.google_skipped_count += 1

            # Update final counts
            count_placeholder.text(f"Downloaded: {st.session_state.google_downloaded_count}, Skipped: {st.session_state.google_skipped_count}")


        elapsed = time.time() - start_time
        status_placeholder.success(f"🎉 Google download completed! Processed {st.session_state.google_downloaded_count + st.session_state.google_skipped_count} items.")
        count_placeholder.empty() # Clear the dynamic count
        st.balloons()
        st.success(f"📦 Total resumes downloaded: {st.session_state.google_downloaded_count}")
        st.info(f"⏱️ Processed {st.session_state.google_downloaded_count + st.session_state.google_skipped_count} total items in {elapsed:.2f} seconds. Resumes saved to: {DOWNLOAD_DIR}")


    except Exception as e:
        status_placeholder.error(f"❌ An error occurred: {str(e)}")
        st.error(f"❌ An error occurred: {str(e)}") # Also show a persistent error


# ====================== MAIN EXECUTION ======================
if ms_run_button:
    # Reset counts for a new run
    st.session_state.ms_downloaded_count = 0
    st.session_state.ms_skipped_count = 0
    process_microsoft()

if google_run_button:
     # Reset counts for a new run
    st.session_state.google_downloaded_count = 0
    st.session_state.google_skipped_count = 0
    process_google()

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:gray;font-size:0.8em;">
    <p>🌟 HR AI Resume Downloader | Supports Microsoft 365 and Google Workspace</p>
    <p>Download directory: <code>{}</code></p>
</div>
""".format(DOWNLOAD_DIR if DOWNLOAD_DIR else "Error setting up directory"), unsafe_allow_html=True)