"""
backend/docuseal_client.py

Thin client for sending the generated agreement PDF to DocuSeal for signature.
Auth token comes from the environment (loaded via python-dotenv in app.py),
never from a config file baked into this repo.
"""

import base64
import os

import requests

BASE_URL = "https://api.docuseal.com"


def _token() -> str:
    token = os.getenv("DOCUSEAL_API_TOKEN")
    if not token:
        raise RuntimeError("DOCUSEAL_API_TOKEN is not set (see .env.example).")
    return token


def create_submission(
    pdf_bytes: bytes,
    filename: str,
    kaedix_name: str,
    kaedix_email: str,
    sub_name: str,
    sub_email: str,
    sub_phone: str,
    send_email: bool,
    send_sms: bool,
    metadata: dict | None = None,
) -> dict:
    """
    Create a DocuSeal submission from the agreement PDF. Signature/date fields
    are placed via {{...}} text tags already embedded in the PDF, so no
    `fields` array is passed here.

    Submitters sign IN ORDER (order=preserved): the KAEDIX representative
    first, the subcontractor only once KAEDIX has signed. DocuSeal holds the
    second invite until the first is complete, so the subcontractor is not
    contacted until the agreement is countersigned. The submitters list below
    IS that order -- reordering it reorders the signing.

    The KAEDIX signer is whoever is logged into the app (their MSAL identity)
    and always gets an email invite, never SMS.
    """
    submitters = [
        {
            "role": "KAEDIX",
            "name": kaedix_name,
            "email": kaedix_email,
            "send_email": True,
            "send_sms": False,
        },
        {
            "role": "Subcontractor",
            "name": sub_name,
            "email": sub_email,
            "phone": sub_phone or None,
            "send_email": send_email,
            "send_sms": send_sms,
        },
    ]

    # Stamp filing metadata on EVERY submitter, identically. DocuSeal carries
    # metadata per submitter, and the autofiler resolves a completed document
    # from whichever party record it reads -- so a stamp on only one of the two
    # would resolve for one signer and not the other. Nothing here reaches the
    # signer: not on the document, not in the filename, not in the email.
    if metadata:
        for submitter in submitters:
            submitter["metadata"] = dict(metadata)

    payload = {
        "name": filename,
        "order": "preserved",
        "documents": [
            {
                "name": filename,
                "file": base64.b64encode(pdf_bytes).decode(),
            }
        ],
        "submitters": submitters,
    }

    resp = requests.post(
        f"{BASE_URL}/submissions/pdf",
        headers={"X-Auth-Token": _token(), "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"DocuSeal API error ({resp.status_code}): {resp.text[:500]}")
    return resp.json()
