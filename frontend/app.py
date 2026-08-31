import os
import re
import sys
import streamlit as st
import msal
from dotenv import load_dotenv
from urllib.parse import urlencode
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────
# PATH + ENV
# ─────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from backend.agreement_generator import generate_agreement_pdf
from backend.docuseal_client import create_submission


def _normalize_phone(raw: str) -> str | None:
    """Return an E.164 US number ('+1XXXXXXXXXX') if raw parses as one, else None."""
    digits = re.sub(r"\D", "", raw or "")
    if len(digits) == 10:
        return f"+1{digits}"
    if len(digits) == 11 and digits.startswith("1"):
        return f"+{digits}"
    if raw.startswith("+") and re.fullmatch(r"\+1\d{10}", raw):
        return raw
    return None

BASE_DIR = Path(__file__).resolve().parent
ICON_PATH = BASE_DIR / ".." / "template" / "kaedix_icon.png"

# -----------------------
# ENV + MSAL CONFIG
# -----------------------
load_dotenv()

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read"]

st.set_page_config(
    page_title="Subcontractor Agreement Generator",
    page_icon=str(ICON_PATH),
    layout="centered"
)

st.markdown(
    """
    <div style="text-align:center; margin-bottom: 10px;">
        <div style="
            font-size:36px;
            font-weight:600;
            letter-spacing:0.7px;
            color:#f57c00;
        ">
            KAEDIX
        </div>
        <div style="
            font-size:20px;
            font-weight:500;
            color:#333;
            margin-top:4px;
        ">
            Subcontractor Agreement Generator
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

@st.cache_resource
def build_msal_app():
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )

msal_app = build_msal_app()

# -----------------------
# SESSION TOKEN
# -----------------------
if "token" not in st.session_state:
    st.session_state.token = None

# -----------------------
# HANDLE MICROSOFT REDIRECT
# -----------------------
query_params = st.query_params
auth_code = query_params.get("code")

if auth_code and not st.session_state.token:
    result = msal_app.acquire_token_by_authorization_code(
        auth_code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    if "id_token_claims" not in result:
        st.error("Microsoft authentication failed")
        st.json(result)
        st.stop()

    st.session_state.token = result
    st.query_params.clear()
    st.rerun()

# -----------------------
# LOGIN SCREEN
# -----------------------
if not st.session_state.token:
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="login",
    )
    st.link_button("Sign in", auth_url, use_container_width=True)
    st.markdown(
        """
        <div style="text-align:center; color: #6b6b6b; font-size: 0.85rem;">
            For authorized internal use only. © 2026 KAEDIX
        </div>
        """,
        unsafe_allow_html=True
    )
    st.stop()

# -----------------------
# AUTH SUCCESS
# -----------------------
claims = st.session_state.token["id_token_claims"]
email = claims.get("preferred_username")
name = claims.get("name")

st.success(f"Logged in as {name}")

# Optional domain restriction
if not email.endswith("@kaedix.com"):
    st.error("Unauthorized user")
    st.stop()

st.divider()

# ─────────────────────────────────────────────
# PROJECT REGISTRY
# ─────────────────────────────────────────────
PROJECTS = {
    "KHP003": "8407 E Rancho Vista Dr, Scottsdale, AZ 85251",
    "KHP004": "4949 E Shaw Butte Dr, Scottsdale, AZ 85254",
    "KHP005": "10818 N 43rd St, Phoenix, AZ 85028",
    "KHP006": "3832 N 85th Pl, Scottsdale, AZ 85251",
    "KHP008": "6318 E Paradise Ln, Scottsdale, AZ 85254",
}

# ─────────────────────────────────────────────
# MAIN FORM
# ─────────────────────────────────────────────
st.subheader("Subcontractor Agreement")

# ── Project Information ───────────────────────────────────────────────────
# Lives outside the form so selecting a Project ID can autofill the address.
# (Widgets inside st.form don't trigger reruns / on_change callbacks.)
def _autofill_project_address():
    st.session_state.project_address = PROJECTS.get(
        st.session_state.project_id, st.session_state.project_address
    )

if "project_id" not in st.session_state:
    st.session_state.project_id = next(iter(PROJECTS))
if "project_address" not in st.session_state:
    st.session_state.project_address = PROJECTS[st.session_state.project_id]

st.markdown("#### Project Information")
col1, col2 = st.columns(2)
with col1:
    project_id      = st.selectbox(
        "Project ID",
        options=list(PROJECTS.keys()),
        key="project_id",
        on_change=_autofill_project_address,
    )
    agreement_date  = st.text_input("Agreement Date", value=datetime.today().strftime("%m/%d/%Y"))
    completion_date = st.text_input("Scheduled Completion Date", placeholder="MM/DD/YYYY")
with col2:
    project_address = st.text_input("Project Address", key="project_address")
    start_date      = st.text_input("Scheduled Start Date", placeholder="MM/DD/YYYY")

st.divider()

with st.form("agreement_form"):

    # ── Subcontractor Information ─────────────────────────────────────────
    st.markdown("#### Subcontractor Information")
    col3, col4 = st.columns(2)
    with col3:
        subcontractor_name = st.text_input("Subcontractor Name")
        license_number     = st.text_input("License Number", placeholder="AZ ROC (if applicable)")
    with col4:
        company_name = st.text_input("Company Name", placeholder="If different from individual name")
        sub_email    = st.text_input("Subcontractor Email")
        sub_phone    = st.text_input("Subcontractor Phone", placeholder="e.g. (602) 555-0123")

    st.divider()

    # ── Contract Terms ────────────────────────────────────────────────────
    st.markdown("#### Contract Terms")
    total_amount = st.text_input("Total Subcontract Amount", placeholder="e.g. 15000")

    st.divider()

    # ── KAEDIX Signatory ──────────────────────────────────────────────────
    st.markdown("#### KAEDIX Signatory")
    col5, col6 = st.columns(2)
    with col5:
        signatory_name  = st.text_input("Signatory Name", placeholder="Person signing for KAEDIX")
        signatory_email = st.text_input("Signatory Email")
    with col6:
        signatory_title = st.text_input("Signatory Title", placeholder="e.g. Managing Member")

    st.divider()

    # ── Appendix A ────────────────────────────────────────────────────────
    st.markdown("#### Appendix A")
    appendix_pdfs = st.file_uploader(
        "Upload Appendix PDF(s) — appended to the agreement in the order added",
        type=["pdf"],
        accept_multiple_files=True,
    )

    st.divider()

    # ── Send for signature ──────────────────────────────────────────────────
    st.markdown("#### Send for Signature")
    col7, col8 = st.columns(2)
    with col7:
        send_via_email = st.checkbox("Send via email")
    with col8:
        send_via_text = st.checkbox("Send via text")

    submit = st.form_submit_button("Generate Agreement PDF", use_container_width=True)

# ─────────────────────────────────────────────
# GENERATION
# ─────────────────────────────────────────────
if submit:
    if not subcontractor_name and not company_name:
        st.warning("Please enter at least a Subcontractor Name or Company Name.")
        st.stop()

    normalized_phone = _normalize_phone(sub_phone)

    send_errors = []
    if send_via_email and not sub_email:
        send_errors.append("Subcontractor Email is required to send via email.")
    if send_via_text and not normalized_phone:
        send_errors.append(
            "Subcontractor Phone must be a valid US number to send via text "
            "(e.g. (602) 555-0123 or +16025550123)."
        )
    if send_errors:
        for err in send_errors:
            st.warning(err)
        st.stop()

    if appendix_pdfs:
        st.caption(
            "Appendices will be appended in this order: "
            + ", ".join(f"{i}. {f.name}" for i, f in enumerate(appendix_pdfs, 1))
        )

    with st.spinner("Generating PDF…"):
        try:
            appendix_bytes_list = (
                [f.read() for f in appendix_pdfs] if appendix_pdfs else None
            )
            pdf_bytes, filename = generate_agreement_pdf(
                project_id=project_id,
                project_address=project_address,
                agreement_date=agreement_date,
                start_date=start_date,
                completion_date=completion_date,
                subcontractor_name=subcontractor_name,
                company_name=company_name,
                license_number=license_number,
                sub_email=sub_email,
                total_amount=total_amount,
                signatory_name=signatory_name,
                signatory_title=signatory_title,
                signatory_email=signatory_email,
                appendix_pdf_bytes_list=appendix_bytes_list,
                include_signature_tags=(send_via_email or send_via_text),
            )
        except Exception as e:
            st.error(f"Generation failed: {e}")
            st.stop()

    st.success("Agreement generated!")
    st.download_button(
        label="Download PDF",
        data=pdf_bytes,
        file_name=filename,
        mime="application/pdf",
        use_container_width=True,
    )

    if send_via_email or send_via_text:
        with st.spinner("Sending for signature…"):
            try:
                result = create_submission(
                    pdf_bytes=pdf_bytes,
                    filename=filename,
                    kaedix_name=name,
                    kaedix_email=email,
                    sub_name=company_name or subcontractor_name,
                    sub_email=sub_email or None,
                    sub_phone=normalized_phone,
                    send_email=send_via_email,
                    send_sms=send_via_text,
                )
                st.success(f"Sent for signature — DocuSeal submission ID {result['id']}")
            except Exception as e:
                st.error(f"Sending for signature failed: {e}")

# ─────────────────────────────────────────────
# SIGN OUT
# ─────────────────────────────────────────────
st.divider()
logout_params = urlencode({"post_logout_redirect_uri": REDIRECT_URI})
logout_url    = f"{AUTHORITY}/oauth2/v2.0/logout?{logout_params}"

if st.button("Sign out"):
    st.session_state.clear()
    st.query_params.clear()
    st.markdown(
        f'<meta http-equiv="refresh" content="0;url={logout_url}">',
        unsafe_allow_html=True,
    )
    st.stop()
