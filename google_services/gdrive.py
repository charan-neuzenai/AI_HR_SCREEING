import streamlit as st
import time
import io
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from common.utils import check_resume, track_skip_reason
from .gmail import google_authenticate, ensure_gdrive_path, check_gdrive_file_exists
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def extract_folder_id(url):
    try:
        if 'folders' in url:
            parts = url.split('/')
            folder_id = parts[parts.index('folders') + 1]
        elif 'id=' in url:
            folder_id = url.split('id=')[1].split('&')[0]
        else:
            folder_id = url
            
        folder_id = folder_id.split('?')[0]
        folder_id = folder_id.split('&')[0]
        
        if len(folder_id) < 10 or not folder_id.replace('_', '').replace('-', '').isalnum():
            raise ValueError("Invalid folder ID format")
            
        return folder_id
        
    except Exception as e:
        st.error(f"Could not extract folder ID from URL: {str(e)}")
        return None

def list_drive_files_in_folder(service, folder_id, status_placeholder):
    try:
        try:
            folder = service.files().get(
                fileId=folder_id,
                fields='id,name,mimeType'
            ).execute()
            
            if folder.get('mimeType') != 'application/vnd.google-apps.folder':
                raise ValueError("The specified ID is not a folder")
                
            status_placeholder.info(f"Accessing folder: {folder.get('name')} ({folder_id})")
            
        except HttpError as e:
            if e.resp.status == 404:
                raise Exception(f"Folder not found. Please check:\n"
                               f"1. The folder exists\n"
                               f"2. You have access permissions\n"
                               f"3. The folder ID is correct: {folder_id}")
            raise

        all_files = []
        next_page_token = None
        status_placeholder.info("Listing files in folder...")

        while True:
            try:
                results = service.files().list(
                    q=f"'{folder_id}' in parents and trashed = false",
                    pageSize=1000,
                    fields="nextPageToken, files(id, name, mimeType)",
                    pageToken=next_page_token
                ).execute()

                files = results.get('files', [])
                all_files.extend(files)

                next_page_token = results.get('nextPageToken')
                if not next_page_token:
                    break
                    
                status_placeholder.info(f"Found {len(all_files)} files so far...")
                
            except HttpError as error:
                raise Exception(f"Google Drive API error: {error}")

        status_placeholder.success(f"Found {len(all_files)} files in folder")
        return all_files

    except Exception as e:
        status_placeholder.error(f"Error listing files: {str(e)}")
        raise

def process_gdrive(creds_file, folder_url):
    status_placeholder = st.empty()
    count_placeholder = st.empty()
    details_expander = st.expander("Processing Details (Google)")
    stats_expander = st.expander("Statistics (Google)")

    try:
        start_time = time.time()

        with st.spinner("🔐 Authenticating with Google API..."):
            scopes = ['https://www.googleapis.com/auth/drive']
            creds = google_authenticate(creds_file, scopes)
            if not creds:
                status_placeholder.error("Authentication failed.")
                return
            
            drive_service = build('drive', 'v3', credentials=creds)
            status_placeholder.success("✅ Authenticated with Google services")

        folder_id = extract_folder_id(folder_url)
        if not folder_id:
            st.error("""
            Could not extract folder ID from URL. Please check the format.
            Valid formats:
            - https://drive.google.com/drive/folders/FOLDER_ID
            - https://drive.google.com/open?id=FOLDER_ID
            - Or just the FOLDER_ID itself
            """)
            return

        status_placeholder.info(f"🔍 Accessing Google Drive folder (ID: {folder_id})...")
        all_files = list_drive_files_in_folder(drive_service, folder_id, status_placeholder)

        if not all_files:
            status_placeholder.warning("No files found in the specified Google Drive folder.")
            return

        total_files = len(all_files)
        for i, file in enumerate(all_files):
            status_placeholder.info(f"Processing file {i+1}/{total_files}: {file['name']}")
            count_placeholder.text(f"Downloaded: {st.session_state.google_downloaded_count}, Skipped: {st.session_state.google_skipped_count}")
            
            is_resume, reason = check_resume(file['name'])
            if is_resume:
                try:
                    request = drive_service.files().get_media(fileId=file['id'])
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                        status_placeholder.info(f"Downloading {file['name']}... {int(status.progress() * 100)}%")
                    
                    file_content = fh.getvalue()
                    details_expander.success(f"✅ Downloaded resume: {file['name']}")
                    st.session_state.google_downloaded_count += 1
                except Exception as e:
                    details_expander.warning(f"❌ Failed to download {file['name']}: {str(e)}")
                    st.session_state.google_skipped_count += 1
                    track_skip_reason(str(e), "google")
            else:
                details_expander.info(f"➡️ Skipped: {file['name']} - {reason}")
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


