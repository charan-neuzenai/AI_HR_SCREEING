import streamlit as st
import time
import tempfile
import base64
import io
import os
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

from common.utils import check_resume, track_skip_reason

def google_authenticate(creds_file, scopes):
    if not creds_file:
        st.error("Please upload credentials.json file first")
        return None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tf:
            tf.write(creds_file.getvalue())
            temp_path = tf.name

        flow = InstalledAppFlow.from_client_secrets_file(temp_path, scopes)
        flow.run_local_server(
            port=0,
            prompt='consent',
            authorization_prompt_message='Please authorize access to your Google Drive',
            success_message='Authentication successful! You can close this tab.'
        )

        creds = flow.credentials

        try:
            os.unlink(temp_path)
        except Exception as e:
            st.warning(f"Could not delete temporary credentials file: {e}")

        return creds
        
    except Exception as e:
        st.error(f"Authentication failed: {str(e)}")
        return None

def fetch_gmail_attachments(service, max_emails, status_placeholder):
    try:
        status_placeholder.info(f"Searching for emails with attachments (up to {max_emails})...")
        next_page_token = None
        processed_emails_count = 0
        yielded_attachments_count = 0

        while processed_emails_count < max_emails:
            results = service.users().messages().list(
                userId='me',
                q='has:attachment',
                maxResults=min(500, max_emails - processed_emails_count),
                pageToken=next_page_token
            ).execute()

            messages = results.get('messages', [])
            if not messages:
                status_placeholder.info("No more emails found with attachments.")
                break

            for msg_summary in messages:
                if processed_emails_count >= max_emails:
                    break
                processed_emails_count += 1

                try:
                    message = service.users().messages().get(
                        userId='me',
                        id=msg_summary['id'],
                        format='full'
                    ).execute()

                    subject = next((h['value'] for h in message['payload']['headers'] if h['name'] == 'Subject'), 'No Subject')
                    status_placeholder.info(f"Processing email {processed_emails_count}/{max_emails}: '{subject}'")

                    parts = message.get('payload', {}).get('parts', [])

                    def find_parts_with_filename(parts_list):
                        found_parts = []
                        if parts_list:
                            for part in parts_list:
                                if 'filename' in part and part['filename']:
                                    found_parts.append(part)
                                if 'parts' in part:
                                    found_parts.extend(find_parts_with_filename(part['parts']))
                        return found_parts

                    attachments_parts = find_parts_with_filename(parts)

                    for part in attachments_parts:
                        attach_id = part['body'].get('attachmentId')
                        file_size = part['body'].get('size', 0)

                        if not attach_id:
                            status_placeholder.warning(f"  - Skipping non-downloadable part: {part.get('filename', 'Unnamed file')} in email '{subject}'")
                            continue

                        status_placeholder.info(f"  - Potential attachment: {part['filename']} (Size: {file_size} bytes)")

                        try:
                            att = service.users().messages().attachments().get(
                                userId='me',
                                messageId=msg_summary['id'],
                                id=attach_id
                            ).execute()

                            file_data = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
                            yielded_attachments_count += 1
                            yield {
                                'name': part['filename'],
                                'content': file_data,
                                'email_subject': subject,
                                'email_id': msg_summary['id'],
                                'attachment_id': attach_id
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

            next_page_token = results.get('nextPageToken')
            if not next_page_token:
                status_placeholder.info(f"Finished scanning available emails ({processed_emails_count} processed).")
                break

    except HttpError as error:
        status_placeholder.error(f"Gmail API error during message list/fetch: {error}")
        st.error(f"Gmail API error during message list/fetch: {error}")
    except Exception as e:
        status_placeholder.error(f"General error during Gmail processing: {str(e)}")
        st.error(f"General error during Gmail processing: {str(e)}")

def save_to_gdrive(service, file_content, filename, folder_path):
    try:
        folder_id = ensure_gdrive_path(service, folder_path)
        existing_files = check_gdrive_file_exists(service, folder_id, filename)
        
        if existing_files:
            return False, "File already exists"
            
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        if isinstance(file_content, str):
            file_content = base64.b64decode(file_content)
            
        media = MediaIoBaseUpload(io.BytesIO(file_content), mimetype='application/octet-stream')
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        return True, "File uploaded successfully"
        
    except HttpError as error:
        return False, f"Google Drive API error: {error}"
    except Exception as e:
        return False, f"Error saving to Google Drive: {str(e)}"

def ensure_gdrive_path(service, path):
    try:
        current_id = "root"
        parts = [p for p in path.split('/') if p]
        
        for part in parts:
            query = f"'{current_id}' in parents and mimeType='application/vnd.google-apps.folder' and name='{part}' and trashed=false"
            results = service.files().list(
                q=query,
                spaces='drive',
                fields='files(id, name)'
            ).execute()
            
            items = results.get('files', [])
            
            if items:
                current_id = items[0]['id']
            else:
                file_metadata = {
                    'name': part,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'parents': [current_id]
                }
                
                folder = service.files().create(
                    body=file_metadata,
                    fields='id'
                ).execute()
                
                current_id = folder.get('id')
                
        return current_id
        
    except HttpError as error:
        raise Exception(f"Google Drive API error: {error}")
    except Exception as e:
        raise Exception(f"Error ensuring Google Drive path: {str(e)}")

def check_gdrive_file_exists(service, folder_id, filename):
    try:
        query = f"'{folder_id}' in parents and name='{filename}' and trashed=false"
        results = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)'
        ).execute()
        
        return results.get('files', [])
        
    except HttpError as error:
        raise Exception(f"Google Drive API error: {error}")
    except Exception as e:
        raise Exception(f"Error checking file existence: {str(e)}")

def process_gmail(creds_file, max_emails, gdrive_path):
    status_placeholder = st.empty()
    count_placeholder = st.empty()
    details_expander = st.expander("Processing Details (Google)")
    stats_expander = st.expander("Statistics (Google)")

    try:
        start_time = time.time()

        with st.spinner("🔐 Authenticating with Google API..."):
            scopes = [
                'https://www.googleapis.com/auth/drive',
                'https://www.googleapis.com/auth/gmail.readonly'
            ]
            creds = google_authenticate(creds_file, scopes)
            if not creds:
                status_placeholder.error("Authentication failed.")
                return
            
            drive_service = build('drive', 'v3', credentials=creds)
            gmail_service = build('gmail', 'v1', credentials=creds)
            status_placeholder.success("✅ Authenticated with Google services")

        status_placeholder.info(f"Starting Gmail attachment scan (up to {max_emails} emails)...")
        attachment_generator = fetch_gmail_attachments(gmail_service, max_emails, status_placeholder)

        for attachment_info in attachment_generator:
            count_placeholder.text(f"Downloaded: {st.session_state.google_downloaded_count}, Skipped: {st.session_state.google_skipped_count}")

            is_resume, reason = check_resume(attachment_info["name"])

            if is_resume:
                success, message = save_to_gdrive(
                    drive_service,
                    attachment_info["content"],
                    attachment_info["name"],
                    gdrive_path
                )
                if success:
                    details_expander.success(f"✅ Saved: {attachment_info['name']} (from '{attachment_info.get('email_subject', 'Email')}')")
                    st.session_state.google_downloaded_count += 1
                else:
                    details_expander.warning(f"❌ Failed to save to Google Drive: {attachment_info['name']} - {message}")
                    st.session_state.google_skipped_count += 1
                    track_skip_reason(message, "google")
            else:
                details_expander.info(f"➡️ Skipped: {attachment_info['name']} - {reason}")
                st.session_state.google_skipped_count += 1
                track_skip_reason(reason, "google")

        elapsed = time.time() - start_time
        
        stats_expander.subheader("📊 Processing Statistics")
        stats_expander.write(f"**Total items processed:** {st.session_state.google_downloaded_count + st.session_state.google_skipped_count}")
        stats_expander.write(f"**Resumes downloaded:** {st.session_state.google_downloaded_count}")
        stats_expander.write(f"**Files skipped:** {st.session_state.google_skipped_count}")
        
        if st.session_state.google_skip_reasons:
            stats_expander.subheader("📝 Skip Reasons")
            for reason, count in st.session_state.google_skip_reasons.items():
                stats_expander.write(f"- {reason}: {count}")
        
        status_placeholder.success(f"🎉 Google processing completed in {elapsed:.2f} seconds!")
        st.balloons()
        st.success(f"📦 Total resumes downloaded: {st.session_state.google_downloaded_count}")
        st.info(f"⏱️ Processed {st.session_state.google_downloaded_count + st.session_state.google_skipped_count} total items in {elapsed:.2f} seconds.")

    except Exception as e:
        status_placeholder.error(f"❌ An error occurred: {str(e)}")
        st.error(f"❌ An error occurred: {str(e)}")