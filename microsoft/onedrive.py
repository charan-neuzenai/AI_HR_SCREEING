import streamlit as st
import requests
import time
from common.utils import check_resume, track_skip_reason
from .outlook import ms_authenticate

def list_onedrive_files(headers, folder_name, status_placeholder):
    try:
        status_placeholder.info(f"Accessing OneDrive folder '{folder_name}'...")
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_name}:/children"

        all_files = []
        while url:
            resp = requests.get(url, headers=headers)
            resp.raise_for_status()
            data = resp.json()

            items = data.get("value", [])
            all_files.extend([item for item in items if 'file' in item])

            url = data.get('@odata.nextLink')
            if url:
                status_placeholder.info(f"Fetching next page of files... ({len(all_files)} files found so far)")

        status_placeholder.success(f"Finished listing. Found {len(all_files)} potential files.")
        return all_files, "https://graph.microsoft.com/v1.0/me/drive/items"
    except Exception as e:
        status_placeholder.error(f"Error accessing OneDrive: {str(e)}")
        st.error(f"Error accessing OneDrive: {str(e)}")
        return [], ""

def process_onedrive(client_id, tenant_id, folder_name):
    status_placeholder = st.empty()
    count_placeholder = st.empty()
    details_expander = st.expander("Processing Details (Microsoft)")
    stats_expander = st.expander("Statistics (Microsoft)")

    try:
        with st.spinner("🔐 Logging in to Microsoft..."):
            scopes = ["Files.ReadWrite.All"]
            token, username = ms_authenticate(client_id, tenant_id, scopes)
            if not token:
                return
            headers = {"Authorization": f"Bearer {token}"}
            status_placeholder.success(f"✅ Authenticated as: {username}")

        start_time = time.time()
        status_placeholder.info("🔍 Listing files in OneDrive folder...")
        all_files, download_prefix = list_onedrive_files(headers, folder_name, status_placeholder)

        if not all_files:
            status_placeholder.warning("No files found in the specified OneDrive folder.")
            st.warning("No files found in the specified OneDrive folder.")
            return

        total_files = len(all_files)
        for i, file in enumerate(all_files):
            status_placeholder.info(f"Processing file {i+1}/{total_files}: {file['name']}")
            count_placeholder.text(f"Downloaded: {st.session_state.ms_downloaded_count}, Skipped: {st.session_state.ms_skipped_count}")

            is_resume, reason = check_resume(file["name"])

            if is_resume:
                file_url = f"{download_prefix}/{file['id']}/content"
                try:
                    resp = requests.get(file_url, headers=headers, stream=True)
                    resp.raise_for_status()

                    file_content = b""
                    for chunk in resp.iter_content(chunk_size=8192):
                        file_content += chunk

                    details_expander.success(f"✅ Found resume: {file['name']}")
                    st.session_state.ms_downloaded_count += 1
                except Exception as e:
                    details_expander.warning(f"❌ Failed to download {file['name']}: {e}")
                    st.warning(f"❌ Failed to download {file['name']}: {e}")
                    track_skip_reason(str(e))
            else:
                details_expander.info(f"➡️ Skipped: {file['name']} - {reason}")
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