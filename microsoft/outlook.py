import streamlit as st
import requests
import base64
import time
from msal import PublicClientApplication
from common.utils import check_resume, track_skip_reason

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
        st.error(f"Authentication failed: {e}")
        return None, None

    if "access_token" in result:
        return result["access_token"], result.get("account", {}).get("username")
    else:
        error_description = result.get("error_description", "Unknown error during token acquisition.")
        correlation_id = result.get("correlation_id", "N/A")
        st.error(f"Authentication failed. Error: {result.get('error')}, Description: {error_description}, Correlation ID: {correlation_id}")
        return None, None

def fetch_outlook_attachments(headers, max_mails, status_placeholder):
    """Fetches Outlook attachments and yields them one by one."""
    try:
        status_placeholder.info(f"Searching for emails with attachments (up to {max_mails})...")
        url = f"https://graph.microsoft.com/v1.0/me/messages?$top={min(100, max_mails)}&$orderby=receivedDateTime desc&$select=id,subject,hasAttachments"
        
        processed_emails_count = 0
        yielded_attachments_count = 0

        while url and processed_emails_count < max_mails:
            resp = requests.get(url, headers=headers)
            resp.raise_for_status()
            data = resp.json()
            messages = data.get("value", [])

            if not messages:
                status_placeholder.info("No more emails found with attachments.")
                break

            for email in messages:
                if processed_emails_count >= max_mails:
                    break
                processed_emails_count += 1
                status_placeholder.info(f"Processing email {processed_emails_count}/{max_mails}: '{email.get('subject', 'No Subject')}'")

                if email.get("hasAttachments"):
                    attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{email['id']}/attachments"
                    att_resp = requests.get(attachments_url, headers=headers)
                    att_resp.raise_for_status()

                    for att in att_resp.json().get("value", []):
                        if att['@odata.type'] == "#microsoft.graph.fileAttachment":
                            yielded_attachments_count += 1
                            yield {
                                "name": att['name'],
                                "content": att['contentBytes'],
                                "email_subject": email.get('subject', 'No Subject'),
                                "email_id": email['id'],
                                "attachment_id": att['id']
                            }
                            status_placeholder.info(f"Yielded attachment {yielded_attachments_count}: {att['name']}")

            url = data.get('@odata.nextLink')

    except Exception as e:
        status_placeholder.error(f"Error during Outlook attachment fetch: {str(e)}")
        st.error(f"Error during Outlook attachment fetch: {str(e)}")

def save_to_onedrive(headers, file_content, filename, target_path):
    """Save file to OneDrive at specified path"""
    try:
        target_path = target_path.strip('/')
        full_path = f"{target_path}/{filename}" if target_path else filename
        
        upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{full_path}:/content"
        
        if isinstance(file_content, str):
            file_content = base64.b64decode(file_content)
            
        response = requests.put(
            upload_url,
            headers={
                **headers,
                "Content-Type": "application/octet-stream"
            },
            data=file_content
        )
        
        if response.status_code in (200, 201):
            return True, f"Saved to: {full_path}"
        else:
            return False, f"Upload failed ({response.status_code}): {response.text}"
            
    except Exception as e:
        return False, f"Error saving to OneDrive: {str(e)}"

def process_outlook(client_id, tenant_id, max_mails, onedrive_path):
    status_placeholder = st.empty()
    count_placeholder = st.empty()
    details_expander = st.expander("Processing Details (Microsoft)")
    stats_expander = st.expander("Statistics (Microsoft)")

    try:
        with st.spinner("🔐 Logging in to Microsoft..."):
            scopes = ["Mail.Read", "Files.ReadWrite"]
            token, username = ms_authenticate(client_id, tenant_id, scopes)
            if not token:
                return
            headers = {"Authorization": f"Bearer {token}"}
            status_placeholder.success(f"✅ Authenticated as: {username}")

        start_time = time.time()
        status_placeholder.info(f"Starting Outlook attachment scan (up to {max_mails} emails)...")
        attachment_generator = fetch_outlook_attachments(headers, max_mails, status_placeholder)

        for attachment_info in attachment_generator:
            count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")

            is_resume, reason = check_resume(attachment_info["name"])

            if is_resume:
                success, message = save_to_onedrive(headers, attachment_info["content"], attachment_info["name"], onedrive_path)
                if success:
                    details_expander.success(f"✅ Saved: {attachment_info['name']} (from '{attachment_info.get('email_subject', 'Email')}')")
                    st.session_state.ms_downloaded_count += 1
                else:
                    details_expander.warning(f"❌ Failed to save to OneDrive: {attachment_info['name']} - {message}")
                    track_skip_reason(message)
            else:
                details_expander.info(f"➡️ Skipped: {attachment_info['name']} - {reason}")
                st.session_state.ms_skipped_count += 1
                track_skip_reason(reason)

        elapsed = time.time() - start_time
        
        stats_expander.subheader("📊 Processing Statistics")
        stats_expander.write(f"**Total items processed:** {st.session_state.ms_downloaded_count + st.session_state.ms_skipped_count}")
        stats_expander.write(f"**Resumes found:** {st.session_state.ms_downloaded_count}")
        stats_expander.write(f"**Files skipped:** {st.session_state.ms_skipped_count}")
        
        if st.session_state.ms_skip_reasons:
            stats_expander.subheader("📝 Skip Reasons")
            for reason, count in st.session_state.ms_skip_reasons.items():
                stats_expander.write(f"- {reason}: {count}")
        
        status_placeholder.success(f"🎉 Microsoft download completed! Processed {st.session_state.ms_downloaded_count + st.session_state.ms_skipped_count} items.")
        count_placeholder.empty()
        st.balloons()
        st.success(f"📦 Total resumes downloaded: {st.session_state.ms_downloaded_count}")
        st.info(f"⏱️ Processed {st.session_state.ms_downloaded_count + st.session_state.ms_skipped_count} total items in {elapsed:.2f} seconds.")

    except Exception as e:
        status_placeholder.error(f"❌ An error occurred: {str(e)}")
        st.error(f"❌ An error occurred: {str(e)}")