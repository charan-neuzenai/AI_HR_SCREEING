import os
import streamlit as st

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

def track_skip_reason(reason, provider="microsoft"):
    """Track reasons for skipping files"""
    if provider == "microsoft":
        if reason in st.session_state.ms_skip_reasons:
            st.session_state.ms_skip_reasons[reason] += 1
        else:
            st.session_state.ms_skip_reasons[reason] = 1
    else:  # google
        if reason in st.session_state.google_skip_reasons:
            st.session_state.google_skip_reasons[reason] += 1
        else:
            st.session_state.google_skip_reasons[reason] = 1

# Constants
RESUME_TYPES = ['.pdf', '.docx', '.doc']
RESUME_KEYWORDS = ['resume', 'cv', 'curriculum vitae', 'bio data']
EXCLUDE_KEYWORDS = ['offer', 'letter', 'terms', 'conditions', 'contest', 'referral', 'policy', 'agreement']