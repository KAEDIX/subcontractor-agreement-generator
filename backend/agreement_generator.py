"""
backend/agreement_generator.py

Reads the clean Word template, substitutes all placeholders with
user-supplied values, and exports a PDF via LibreOffice.

Returns: (pdf_bytes: bytes, filename: str)
"""

import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from pypdf import PdfWriter, PdfReader

TEMPLATE = Path(__file__).parent.parent / "template" / "Subcontractor_Agreement_Template_clean.docx"
EXPORTS  = Path(__file__).parent.parent / "exports"
EXPORTS.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _fmt_date(value: str) -> str:
    """Try to parse and return MM/DD/YYYY; fall back to the original string."""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(value.strip(), fmt).strftime("%m/%d/%Y")
        except ValueError:
            pass
    return value


def _fmt_amount(value: str) -> str:
    """Format as $X,XXX.XX if the value looks numeric."""
    try:
        n = float(value.replace(",", "").replace("$", ""))
        return f"${n:,.2f}"
    except (ValueError, AttributeError):
        return value



def _replace_nth(text: str, old: str, new: str, n: int) -> str:
    """Replace the n-th (1-based) occurrence of *old* with *new*."""
    pos = -1
    for _ in range(n):
        pos = text.find(old, pos + 1)
        if pos == -1:
            return text
    return text[:pos] + new + text[pos + len(old):]


_RUN_20 = '<w:r w:rsidRPr="00BE323A"><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>'


def _populate_xml(xml: str, f: dict) -> str:
    """Apply all placeholder substitutions to document.xml content."""

    pid   = f["project_id"]         or ""
    addr  = f["project_address"]     or ""
    date  = _fmt_date(f["agreement_date"])  if f["agreement_date"]  else ""
    sd    = _fmt_date(f["start_date"])      if f["start_date"]      else ""
    cd    = _fmt_date(f["completion_date"]) if f["completion_date"] else ""
    sub   = f["subcontractor_name"]  or ""
    co    = f["company_name"]         or ""
    amt   = _fmt_amount(f["total_amount"])  if f["total_amount"]    else ""
    sname = f["signatory_name"]       or ""
    title = f["signatory_title"]      or ""
    seml  = f["signatory_email"]      or ""
    beml  = f["sub_email"]            or ""
    lic   = f["license_number"]       or ""

    # ── Simple single-run placeholders ─────────────────────────────────────
    xml = xml.replace("[Address]",          addr)
    xml = xml.replace("[MM/DD/YYYY]",       date)
    xml = xml.replace("[Amount]",           amt)
    xml = xml.replace("[Name]",             sname)
    xml = xml.replace("[Subcontractor Name]", sub)
    xml = xml.replace("[Email]",            beml)   # catches the simple (sub) occurrence

    # ── Split-run placeholders (exact raw XML replacement) ──────────────────

    # [KHPXXX] — split across 3 runs
    xml = xml.replace(
        '<w:r w:rsidR="00044E8D" w:rsidRPr="00044E8D"><w:t>[</w:t></w:r>'
        '<w:r w:rsidRPr="00044E8D"><w:t>KHP</w:t></w:r>'
        '<w:r w:rsidR="00044E8D" w:rsidRPr="00044E8D"><w:t>XXX]</w:t></w:r>',
        f'<w:r><w:t>{pid}</w:t></w:r>',
    )

    # [Project Start Date] — split across 3 runs
    xml = xml.replace(
        '<w:r><w:t>[</w:t></w:r>'
        '<w:r w:rsidR="00783D83"><w:t xml:space="preserve">Project </w:t></w:r>'
        '<w:r><w:t>Start Date]</w:t></w:r>',
        f'<w:r><w:t>{sd}</w:t></w:r>',
    )

    # [Project Completion Date] — split across 3 runs
    xml = xml.replace(
        '<w:r><w:t>[</w:t></w:r>'
        '<w:r w:rsidR="00783D83"><w:t xml:space="preserve">Project </w:t></w:r>'
        '<w:r><w:t>Completion Date]</w:t></w:r>',
        f'<w:r><w:t>{cd}</w:t></w:r>',
    )

    # [Company Name] — split with bold runs
    xml = xml.replace(
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>[</w:t></w:r>'
        '<w:r w:rsidR="008E29CB"><w:rPr><w:b/><w:bCs/></w:rPr><w:t>Company</w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve"> Name]</w:t></w:r>',
        f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>{co}</w:t></w:r>',
    )

    # [Title] — split across 3 runs
    xml = xml.replace(
        '<w:r w:rsidR="00B16321" w:rsidRPr="00F13F85"><w:t>[</w:t></w:r>'
        '<w:r w:rsidRPr="00F13F85"><w:t>Title</w:t></w:r>'
        '<w:r w:rsidR="00B16321" w:rsidRPr="00F13F85"><w:t>]</w:t></w:r>',
        f'<w:r><w:t>{title}</w:t></w:r>',
    )

    # [Email] KAEDIX (signatory) — split across 3 runs
    xml = xml.replace(
        '<w:r w:rsidR="008777A2" w:rsidRPr="00F13F85"><w:t>[</w:t></w:r>'
        '<w:r w:rsidRPr="00F13F85"><w:t>Email</w:t></w:r>'
        '<w:r w:rsidR="008777A2" w:rsidRPr="00F13F85"><w:t>]</w:t></w:r>',
        f'<w:r><w:t>{seml}</w:t></w:r>',
    )

    # [License #] — split with comment markup
    xml = xml.replace(
        '<w:r w:rsidRPr="00F13F85"><w:t>[</w:t></w:r>'
        '<w:commentRangeStart w:id="5"/>'
        '<w:r w:rsidRPr="00F13F85"><w:t>License #</w:t></w:r>'
        '<w:commentRangeEnd w:id="5"/>'
        '<w:r w:rsidRPr="00F13F85"><w:rPr><w:rStyle w:val="CommentReference"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
        '<w:commentReference w:id="5"/></w:r>'
        '<w:r w:rsidRPr="00F13F85"><w:t>]</w:t></w:r>',
        f'<w:r><w:t>{lic}</w:t></w:r>',
    )

    return xml


# Invisible (white, 1pt) run properties for DocuSeal text tags — present in the
# rendered PDF's text layer so DocuSeal can locate them, but not visible on the page.
_HIDDEN_TAG_RPR = '<w:rPr><w:color w:val="FFFFFF"/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'

# Same, but without the sz override. A tiny-font run dropped into a
# paragraph that previously had no runs at all becomes that paragraph's
# only sizing signal and collapses its line height (the row shrinks), so
# the blank spacer paragraphs use this instead — invisible via color alone,
# sized like normal text so the row keeps its original height.
_HIDDEN_TAG_RPR_NORMAL_SIZE = '<w:rPr><w:color w:val="FFFFFF"/></w:rPr>'


def _tag_run(tag: str, *, normal_size: bool = False) -> str:
    rpr = _HIDDEN_TAG_RPR_NORMAL_SIZE if normal_size else _HIDDEN_TAG_RPR
    return f'<w:r>{rpr}<w:t>{tag}</w:t></w:r>'


def _add_signature_tags(xml: str) -> str:
    """
    Embed DocuSeal {{...}} text tags in the Signatures section, mirroring the
    document's own signature blocks. Each anchor below is matched on its
    unique paraId so KAEDIX (left column) and Subcontractor (right column)
    get distinct fields.

    DocuSeal anchors a field at the tag's own text position and extends its
    height downward from there. Every row in this table prints its label at
    the TOP of the row, with a blank paragraph below it before the row's own
    ruled line — i.e. the printed value (like Name/Title/Email) sits ABOVE
    the line, in the gap between the label and the rule, not below it. So
    each tag is planted in the blank paragraph that already sits directly
    above the ruled line it should clear — the previous row's spacer, not
    the Signature:/Date: label's own paragraph — instead of relying on a
    guessed height to out-run the wrong line. Measured against the rendered
    PDF: each such gap is ~30pt tall, so height=20 (signature) / height=14
    (date) clears the rule below by ~10pt with room to spare.
    """

    # KAEDIX — Signature: planted in the blank paragraph below Email/KAEDIX,
    # which sits directly above the Signature row's own ruled line.
    xml = xml.replace(
        '<w:p w14:paraId="00000105" w14:textId="00000106" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr></w:p>',
        '<w:p w14:paraId="00000105" w14:textId="00000106" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
        + _tag_run('{{Signature_KAEDIX;role=KAEDIX;type=signature;width=180;height=26}}', normal_size=True)
        + '</w:p>',
    )

    # Subcontractor — Signature: planted in the blank paragraph below
    # Email/Subcontractor, directly above the Signature row's own line.
    xml = xml.replace(
        '<w:p w14:paraId="00000101" w14:textId="7F9B57AD" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr></w:p>',
        '<w:p w14:paraId="00000101" w14:textId="7F9B57AD" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
        + _tag_run('{{Signature_Sub;role=Subcontractor;type=signature;width=180;height=26}}', normal_size=True)
        + '</w:p>',
    )

    # KAEDIX — Date: planted in the blank paragraph below Signature/KAEDIX,
    # directly above the Date row's own ruled line.
    xml = xml.replace(
        '<w:p w14:paraId="00000110" w14:textId="00000111" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr></w:p>',
        '<w:p w14:paraId="00000110" w14:textId="00000111" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
        + _tag_run('{{Date_KAEDIX;role=KAEDIX;type=date;width=120;height=20}}', normal_size=True)
        + '</w:p>',
    )

    # Subcontractor — Date: the Subcontractor column has no separate blank
    # spacer paragraph below Signature/Subcontractor the way the KAEDIX
    # column does (paraId 00000110) — that row is only as tall as it is
    # because the KAEDIX column's spacer forces the shared row height. Add
    # the same kind of spacer paragraph here, mirroring the KAEDIX column,
    # and plant the tag there instead of inline on the label (which would
    # anchor at the label's own top and reproduce the exact bug being fixed).
    xml = xml.replace(
        '<w:p w14:paraId="0000010C" w14:textId="165B9C7D" w:rsidR="008777A2" w:rsidRDefault="00F13F85" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>Signature:</w:t></w:r></w:p></w:tc>',
        '<w:p w14:paraId="0000010C" w14:textId="165B9C7D" w:rsidR="008777A2" w:rsidRDefault="00F13F85" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>Signature:</w:t></w:r></w:p>'
        '<w:p w14:paraId="00000113" w14:textId="00000113" w:rsidR="008777A2" w:rsidRDefault="008777A2" w:rsidP="008777A2">'
        '<w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:pPr>'
        + _tag_run('{{Date_Sub;role=Subcontractor;type=date;width=120;height=20}}', normal_size=True)
        + '</w:p></w:tc>',
    )

    return xml


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_agreement_pdf(
    project_id: str,
    project_address: str,
    agreement_date: str,
    start_date: str,
    completion_date: str,
    subcontractor_name: str,
    company_name: str,
    license_number: str,
    sub_email: str,
    total_amount: str,
    signatory_name: str,
    signatory_title: str,
    signatory_email: str,
    appendix_pdf_bytes_list: list[bytes] = None,
    appendix_pdf_bytes: bytes = None,
    include_signature_tags: bool = False,
) -> tuple[bytes, str]:
    """
    Populate the Word template and convert to PDF.
    Returns (pdf_bytes, filename).

    include_signature_tags: embed invisible DocuSeal {{...}} text tags in the
    Signatures section so the resulting PDF can be sent for e-signature as-is.
    """

    fields = {
        "project_id":        project_id.strip(),
        "project_address":   project_address.strip(),
        "agreement_date":    agreement_date.strip(),
        "start_date":        start_date.strip(),
        "completion_date":   completion_date.strip(),
        "subcontractor_name": subcontractor_name.strip(),
        "company_name":      company_name.strip(),
        "license_number":    license_number.strip(),
        "sub_email":         sub_email.strip(),
        "total_amount":      total_amount.strip(),
        "signatory_name":    signatory_name.strip(),
        "signatory_title":   signatory_title.strip(),
        "signatory_email":   signatory_email.strip(),
    }

    # Build output filename
    date_str = ""
    if fields["agreement_date"]:
        try:
            date_str = datetime.strptime(
                _fmt_date(fields["agreement_date"]), "%m/%d/%Y"
            ).strftime("%Y%m%d")
        except ValueError:
            date_str = datetime.today().strftime("%Y%m%d")
    else:
        date_str = datetime.today().strftime("%Y%m%d")

    pid      = fields["project_id"] or "KHPXXX"
    company  = fields["company_name"] or fields["subcontractor_name"] or "Subcontractor"
    safe_co  = re.sub(r"[^\w\s-]", "", company).strip().replace(" ", "_")
    stem     = f"{date_str}_{pid}_Subcontractor_Agreement_{safe_co}"
    docx_out = EXPORTS / f"{stem}.docx"
    pdf_out  = EXPORTS / f"{stem}.pdf"

    # --- 1. Populate XML inside a temp copy of the template ----------------
    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = Path(tmp) / "output.docx"

        # Copy template into tmp
        shutil.copy(TEMPLATE, tmp_docx)

        # Read, modify, and rewrite document.xml inside the ZIP
        with zipfile.ZipFile(tmp_docx, "r") as zin:
            names  = zin.namelist()
            files  = {name: zin.read(name) for name in names}

        doc_xml = files["word/document.xml"].decode("utf-8")
        doc_xml = _populate_xml(doc_xml, fields)
        if include_signature_tags:
            doc_xml = _add_signature_tags(doc_xml)
        files["word/document.xml"] = doc_xml.encode("utf-8")

        with zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)

        # Copy filled docx to exports
        shutil.copy(tmp_docx, docx_out)

        # --- 2. Convert DOCX → PDF via LibreOffice -------------------------
        result = subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(EXPORTS),
                str(docx_out),
            ],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed:\n{result.stderr}")

    # soffice names the output after the docx stem
    soffice_out = EXPORTS / f"{stem}.pdf"
    if not soffice_out.exists():
        raise FileNotFoundError(f"Expected PDF not found: {soffice_out}")

    # --- 3. Merge appendix PDF(s) if provided, in the order added ----------
    appendices = list(appendix_pdf_bytes_list or [])
    if appendix_pdf_bytes:                 # backward-compat single appendix
        appendices.append(appendix_pdf_bytes)

    if appendices:
        writer = PdfWriter()

        # Add all pages from the generated agreement
        for page in PdfReader(soffice_out).pages:
            writer.add_page(page)

        # Add all pages from each uploaded appendix, in upload order
        for ap_bytes in appendices:
            for page in PdfReader(io.BytesIO(ap_bytes)).pages:
                writer.add_page(page)

        merged_buf = io.BytesIO()
        writer.write(merged_buf)
        merged_bytes = merged_buf.getvalue()

        # Overwrite the file on disk with the merged version
        soffice_out.write_bytes(merged_bytes)
        return merged_bytes, pdf_out.name

    pdf_bytes = soffice_out.read_bytes()
    return pdf_bytes, pdf_out.name
