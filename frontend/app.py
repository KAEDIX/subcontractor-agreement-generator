import base64
import os
import re
import sys
import textwrap
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
WORDMARK_PATH = BASE_DIR / ".." / "template" / "kdx_wordmark_white.png"


def _img_data_uri(path: Path) -> str:
    try:
        return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode()
    except FileNotFoundError:
        return ""

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

# ─────────────────────────────────────────────
# KAEDIX / PUNCH-AI DESIGN LANGUAGE
# ─────────────────────────────────────────────
st.html(
    textwrap.dedent("""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:ital,opsz,wght@0,9..40,300..700;1,9..40,300..700&display=swap" rel="stylesheet">
    <style>
    :root{
        --kdx-tiger:#FF671F; --kdx-tiger-700:#C94A0E;
        --kdx-navy:#0A294A; --kdx-charcoal:#2D2D2D; --kdx-graphite:#666666; --kdx-iron:#454545;
        --kdx-ink-50:#FAFAF8; --kdx-ink-100:#F2F1EE; --kdx-ink-200:#E5E3DE; --kdx-ink-300:#C9C6BF; --kdx-ink-400:#9A978F;
        --kdx-success:#0F5933;
        --font-display:'DM Serif Display',Georgia,serif;
        --font-sans:'DM Sans',-apple-system,sans-serif;
    }

    /* ground + base type */
    .stApp{background:var(--kdx-ink-50);}
    html, body, [class*="css"]{font-family:var(--font-sans);}
    .block-container{max-width:720px; padding-top:0 !important; padding-bottom:3rem;}
    header[data-testid="stHeader"]{background:var(--kdx-ink-50); height:0; min-height:0;}
    div[data-testid="stAppViewContainer"] > .main{padding-top:0;}
    div[data-testid="stMainBlockContainer"]{padding-top:0 !important;}

    /* app header band — full-bleed navy, flush with the very top of the page */
    .kdx-app-header{
        background:var(--kdx-navy); margin:0 -1rem 1.75rem -1rem; padding:26px 28px 28px;
        display:flex; align-items:center; gap:14px;
    }
    .kdx-app-header img{height:22px; width:auto; display:block;}
    .kdx-app-eyebrow{
        font-family:var(--font-sans); font-size:10px; font-weight:500; letter-spacing:.22em;
        text-transform:uppercase; color:#FFC4A3; margin-bottom:5px;
    }
    .kdx-app-title{font-family:var(--font-display); font-weight:400; font-size:23px; color:#FFFFFF; line-height:1.15;}

    /* uppercase-eyebrow section dividers — no rule line, just tracked label */
    h4{
        font-family:var(--font-sans) !important; font-weight:700 !important; font-size:9.5px !important;
        letter-spacing:.18em !important; text-transform:uppercase !important; color:var(--kdx-ink-400) !important;
        margin-top:1.5rem !important; margin-bottom:.6rem !important;
    }
    div[data-testid="stForm"] hr, section.main hr{border-color:var(--kdx-ink-100) !important; margin:1.1rem 0 !important;}

    /* card-style bordered sections: form + the container just above it */
    div[data-testid="stForm"]{
        border:1px solid var(--kdx-ink-200) !important; border-radius:6px !important;
        background:#FFFFFF; padding:20px 20px 8px 20px;
        box-shadow:0 1px 2px rgba(45,45,45,.04);
    }

    /* inputs */
    label, .stTextInput label, .stSelectbox label{
        font-family:var(--font-sans) !important; font-size:10.5px !important; font-weight:700 !important;
        color:var(--kdx-iron) !important; letter-spacing:.03em !important;
    }
    .stTextInput input, .stSelectbox div[data-baseweb="select"] > div, textarea{
        border-radius:4px !important; border:1px solid var(--kdx-ink-300) !important;
        font-family:var(--font-sans) !important; font-size:13px !important; color:var(--kdx-charcoal) !important;
        background:#FFFFFF !important;
    }
    .stTextInput input:focus{border-color:var(--kdx-tiger) !important; box-shadow:0 0 0 1px var(--kdx-tiger) !important;}

    /* KHP project picker — pill treatment, not a bare dropdown */
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div{
        border-radius:4px !important; padding-left:2px !important;
    }

    /* primary action button (form submit) */
    div[data-testid="stFormSubmitButton"] button{
        background:var(--kdx-tiger) !important; color:#FFFFFF !important; border:none !important;
        border-radius:4px !important; font-family:var(--font-sans) !important; font-weight:700 !important;
        font-size:13px !important; letter-spacing:.02em !important; padding:12px !important;
        box-shadow:0 2px 8px rgba(255,103,31,.28) !important;
    }
    div[data-testid="stFormSubmitButton"] button:hover{background:var(--kdx-tiger-700) !important;}

    /* de-emphasized secondary action ("Download PDF") — corner-link, not a button */
    div[data-testid="stDownloadButton"] button{
        background:transparent !important; color:var(--kdx-graphite) !important;
        border:none !important; box-shadow:none !important; text-decoration:underline !important;
        text-underline-offset:2px !important; font-family:var(--font-sans) !important;
        font-size:11px !important; font-weight:600 !important; padding:4px 0 !important;
        width:auto !important; float:right;
    }
    div[data-testid="stDownloadButton"]{display:flex; justify-content:flex-end;}

    /* sign-in link-button, kept tiger */
    .stLinkButton a{
        background:var(--kdx-tiger) !important; color:#FFFFFF !important; border-radius:4px !important;
        font-family:var(--font-sans) !important; font-weight:700 !important; border:none !important;
    }

    /* "Logged in as X" — a tinted status line, not a default green alert box */
    .kdx-signed-in{
        display:inline-flex; align-items:center; gap:7px; font-family:var(--font-sans);
        font-size:11.5px; font-weight:600; color:var(--kdx-success); margin:0 0 .75rem;
    }
    .kdx-signed-in::before{
        content:""; width:6px; height:6px; border-radius:50%; background:var(--kdx-success);
    }

    /* page title over the header band */
    .kdx-page-title{
        font-family:var(--font-sans) !important; font-weight:700 !important; font-size:15px !important;
        color:var(--kdx-charcoal) !important; letter-spacing:.01em; margin:.25rem 0 1rem;
    }

    /* narrow-viewport safety: keep the header band flush and content padded */
    @media (max-width: 480px){
        .block-container{padding-left:0.9rem; padding-right:0.9rem;}
        .kdx-app-header{margin:0 -0.9rem 1.5rem -0.9rem; padding:20px 18px 22px;}
        .kdx-app-title{font-size:19px;}
    }
    </style>
    """)
)

_wordmark_uri = _img_data_uri(WORDMARK_PATH)
st.html(
    textwrap.dedent(f"""
    <div class="kdx-app-header">
        {'<img src="' + _wordmark_uri + '" alt="KAEDIX" />' if _wordmark_uri else ''}
        <div>
            <div class="kdx-app-eyebrow">KAEDIX Document Portal</div>
            <div class="kdx-app-title">Subcontractor Agreement Generator</div>
        </div>
    </div>
    """)
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

# LOCAL SCREENSHOT BYPASS — not for commit, dev-only auth stub
if os.environ.get("SCREENSHOT_BYPASS_AUTH"):
    st.session_state.token = {
        "id_token_claims": {
            "preferred_username": "seth.porter@kaedix.com",
            "name": "Seth Porter",
        }
    }

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
        '<div style="text-align:center; color:#9A978F; font-size:0.8rem; margin-top:8px; '
        'font-family:\'DM Sans\',-apple-system,sans-serif;">'
        'For authorized internal use only. © 2026 KAEDIX</div>',
        unsafe_allow_html=True
    )
    st.stop()

# -----------------------
# AUTH SUCCESS
# -----------------------
claims = st.session_state.token["id_token_claims"]
email = claims.get("preferred_username")
name = claims.get("name")

st.html(f'<div class="kdx-signed-in">Logged in as {name}</div>')

# Optional domain restriction
if not email.endswith("@kaedix.com"):
    st.error("Unauthorized user")
    st.stop()

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
st.html('<div class="kdx-page-title">Subcontractor Agreement</div>')

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

    send_clicked = st.form_submit_button(
        "Send PDF for Signature", use_container_width=True, type="primary"
    )
    dl_col = st.columns([3, 1])[1]
    with dl_col:
        download_clicked = st.form_submit_button(
            "Download PDF", use_container_width=True
        )

# Demote the "Download PDF" form button to a plain corner link, not a button.
# (More specific + !important so it wins over the primary submit-button styling above.)
st.markdown(
    """
    <style>
    div[data-testid="stFormSubmitButton"]:has(button[kind="secondaryFormSubmit"]) button {
        border: none !important;
        background: none !important;
        color: #6b6b6b !important;
        font-size: 0.85rem !important;
        text-decoration: underline !important;
        text-underline-offset: 2px !important;
        padding: 0 !important;
        box-shadow: none !important;
        float: right;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# GENERATION
# ─────────────────────────────────────────────
if send_clicked or download_clicked:
    if not subcontractor_name and not company_name:
        st.warning("Please enter at least a Subcontractor Name or Company Name.")
        st.stop()

    normalized_phone = _normalize_phone(sub_phone)

    if send_clicked:
        send_errors = []
        if not send_via_email and not send_via_text:
            send_errors.append(
                "Check Send via email or Send via text before sending — "
                "or use the Download PDF link below to just get the file."
            )
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
                include_signature_tags=send_clicked,
            )
        except Exception as e:
            st.error(f"Generation failed: {e}")
            st.stop()

    if download_clicked:
        st.download_button(
            label="Download PDF",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            use_container_width=True,
        )

    if send_clicked:
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
