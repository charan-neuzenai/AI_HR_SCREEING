import streamlit as st
# For Microsoft services
from microsoft.outlook import process_outlook
from microsoft.onedrive import process_onedrive
from microsoft.sharepoint import process_sharepoint

# For Google services
from google_services.gmail import process_gmail
from google_services.gdrive import process_gdrive 


# Common utilities
from common.utils import check_resume, track_skip_reason

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

   * Set **"Allow public client flows"** to **Yes**

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
Then click "New Project"
Give your project a name: ResumeDownloader
Click Create
After it's done, click the project name to select it

✅ STEP 3: Enable Gmail & Google Drive APIs 📬📁
Now we'll turn on the tools we need:
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
you'll see a button: "Get Started"
Click it ✅
It will ask for some details – just follow the instructions and click next or continue when needed.

 then click create OAuth client
 then it will application type choise  desktop app  after enter any name
then click create download json

✅ STEP 5: Add Yourself in Audience Option 👤
You need to allow yourself to use this app:

On the OAuth consent screen page → find "Audience" section (scroll down)  their  will be Test user  below that Add users
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
Click "Add to Table" ➕
Then click Update ✅
Then click Save and Continue

    *Keep your `credentials.json` safe and upload it in the Google section of the app.*
    """)

st.markdown("""
<div style="background-color:#0e1117;padding:20px;border-radius:10px;margin-bottom:20px">
    <h3 style="color:white;text-align:center;">Download resumes from SharePoint, OneDrive, Outlook, Google Drive & Gmail</h3>
</div>
""", unsafe_allow_html=True)

# ====================== CONFIGURATION ======================
# Initialize session state for counts and statistics
if 'ms_downloaded_count' not in st.session_state:
    st.session_state.ms_downloaded_count = 0
if 'ms_skipped_count' not in st.session_state:
    st.session_state.ms_skipped_count = 0
if 'ms_skip_reasons' not in st.session_state:
    st.session_state.ms_skip_reasons = {}
if 'google_downloaded_count' not in st.session_state:
    st.session_state.google_downloaded_count = 0
if 'google_skipped_count' not in st.session_state:
    st.session_state.google_skipped_count = 0
if 'google_skip_reasons' not in st.session_state:
    st.session_state.google_skip_reasons = {}

# Tabs for provider selection
ms_tab, google_tab = st.tabs(["🔷 Microsoft", "🔷 Google"])

# ====================== MICROSOFT SECTION ======================
with ms_tab:
    st.subheader("🔐 Microsoft Authentication")
    col1, col2 = st.columns(2)
    with col1:
        client_id = st.text_input("Client ID", value="", key="ms_client_id")
    with col2:
        tenant_id = st.text_input("Tenant ID", value="common", key="ms_tenant_id")

    mode = st.radio("Select source", ["SharePoint Folder", "OneDrive Folder", "Outlook Attachments"], key="ms_mode")

    if mode == "SharePoint Folder":
        st.subheader("🏢 SharePoint Configuration")
        domain = st.text_input("SharePoint domain (e.g., company.sharepoint.com)",
                             value="",
                             key="sharepoint_domain")
        site_name = st.text_input("Site name (e.g., HRTeamSite)",
                                 value="",
                                 key="sharepoint_site")
        folder_name = st.text_input("Folder name (e.g., Resumes2024)",
                                  value="",
                                  key="sharepoint_folder")
    elif mode == "OneDrive Folder":
        st.subheader("📁 OneDrive Configuration")
        folder_name = st.text_input("Folder name in OneDrive (e.g., MyResumes)", key="onedrive_folder")
    else: # Outlook Attachments
        st.subheader("📧 Outlook Configuration")
        max_mails = st.number_input("Number of emails to scan", min_value=1, max_value=10000, value=1000, key="ms_max_mails")
        onedrive_path = st.text_input("OneDrive path to save resumes (e.g., Resumes/Incoming)", 
                                     value="Resumes/Incoming",
                                     key="onedrive_path")

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
        gdrive_path = st.text_input("Google Drive folder path to save resumes (e.g., /Resumes/Incoming)", key="gdrive_path")

    google_run_button = st.button("🚀 Download Google Resumes", key="google_run")

# ====================== MAIN EXECUTION ======================
if ms_run_button:
    st.session_state.ms_downloaded_count = 0
    st.session_state.ms_skipped_count = 0
    st.session_state.ms_skip_reasons = {}
    
    if mode == "Outlook Attachments":
        process_outlook(client_id, tenant_id, max_mails, onedrive_path)
    elif mode == "OneDrive Folder":
        process_onedrive(client_id, tenant_id, folder_name)
    elif mode == "SharePoint Folder":
        process_sharepoint(client_id, tenant_id, domain, site_name, folder_name)

if google_run_button:
    st.session_state.google_downloaded_count = 0
    st.session_state.google_skipped_count = 0
    st.session_state.google_skip_reasons = {}
    
    if google_mode == "Gmail Attachments":
        process_gmail(creds_file, max_emails, gdrive_path)
    else:  # Google Drive Folder
        process_gdrive(creds_file, folder_url)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align:center;color:gray;font-size:0.8em;">
    <p>🌟 HR AI Resume Downloader | Supports Microsoft 365 and Google Workspace</p>
</div>
""", unsafe_allow_html=True)

