import streamlit as st
import requests
import os
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from dotenv import load_dotenv

load_dotenv()

API_BASE_URL = os.getenv("DIFY_API_BASE_URL", "https://api.dify.ai/v1")
API_KEY = os.getenv("DIFY_API_KEY", "")

st.set_page_config(page_title="AI-Driven Report Generation", layout="wide")
st.title("AI-Driven SOC Report Generation")

# Sidebar: API configuration
with st.sidebar:
    st.header("API Configuration")
    api_base = st.text_input("API Base URL", value=API_BASE_URL)
    api_key = st.text_input("API Secret Key", value=API_KEY, type="password")


def upload_file_to_dify(file_bytes, filename, api_base, api_key):
    """Upload a file to Dify and return the upload_file_id."""
    url = f"{api_base.rstrip('/')}/files/upload"
    headers = {"Authorization": f"Bearer {api_key}"}
    files = {"file": (filename, file_bytes, "application/octet-stream")}
    data = {"user": "streamlit-user"}
    resp = requests.post(url, headers=headers, files=files, data=data, timeout=120)
    resp.raise_for_status()
    return resp.json()["id"]


def run_workflow(inputs, api_base, api_key):
    """Run the Dify workflow and return the outputs dict."""
    url = f"{api_base.rstrip('/')}/workflows/run"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "inputs": inputs,
        "response_mode": "blocking",
        "user": "streamlit-user",
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=600)
    resp.raise_for_status()
    data = resp.json()
    # Dify wraps outputs under data.outputs
    return data.get("data", {}).get("outputs", {})


def markdown_to_docx(md_text: str) -> bytes:
    """Convert a markdown string to a .docx file and return its bytes."""
    doc = Document()

    # Set default paragraph font
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    lines = md_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]

        # Heading 1
        if line.startswith("# "):
            p = doc.add_heading(line[2:].strip(), level=1)
        # Heading 2
        elif line.startswith("## "):
            p = doc.add_heading(line[3:].strip(), level=2)
        # Heading 3
        elif line.startswith("### "):
            p = doc.add_heading(line[4:].strip(), level=3)
        # Heading 4
        elif line.startswith("#### "):
            p = doc.add_heading(line[5:].strip(), level=4)
        # Bullet list
        elif re.match(r"^[-*+] ", line):
            text = line[2:].strip()
            p = doc.add_paragraph(style="List Bullet")
            _add_inline_formatting(p, text)
        # Numbered list
        elif re.match(r"^\d+\. ", line):
            text = re.sub(r"^\d+\. ", "", line).strip()
            p = doc.add_paragraph(style="List Number")
            _add_inline_formatting(p, text)
        # Horizontal rule
        elif re.match(r"^[-*_]{3,}$", line.strip()):
            doc.add_paragraph("─" * 60)
        # Empty line → paragraph break (skip)
        elif line.strip() == "":
            pass
        # Normal paragraph
        else:
            p = doc.add_paragraph()
            _add_inline_formatting(p, line)

        i += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def _add_inline_formatting(paragraph, text: str):
    """Parse **bold**, *italic*, and plain text within a paragraph."""
    # Split on bold/italic markers
    parts = re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("*") and part.endswith("*"):
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            paragraph.add_run(part)


# ── Main form ──────────────────────────────────────────────────────────────────

st.subheader("Report Parameters")

col1, col2 = st.columns(2)

with col1:
    company_name = st.text_input("Company Name (公司名称) *", max_chars=256)
    co_short_name = st.text_input("Company Short Name (公司简称) *", max_chars=48)
    system_or_service_name = st.text_input("Service/System Name (服务/系统名称) *", max_chars=256)
    service_description = st.text_input("Service Description (服务体系描述) *", max_chars=256)
    period_start = st.text_input("Report Period Start (报告期间第一天) *", placeholder="e.g. 2024-01-01")
    period_end = st.text_input(
        "Report Period End (报告期间最后一天, N/A for Type1)",
        placeholder="e.g. 2024-12-31",
    )

with col2:
    report_type = st.selectbox(
        "Report Type (报告类型) *",
        ["SOC1 TYPE1", "SOC1 TYPE2", "SOC2 TYPE1", "SOC2 TYPE2"],
    )
    output_language = st.selectbox(
        "Output Language (报告语言) *",
        ["English", "中文", "Both"],
        index=0,
    )
    industry = st.selectbox(
        "Industry (公司所在行业)",
        ["Other", "SaaS", "Cloud Service", "AI", "General", "PaaS", "IaaS"],
        index=0,
    )
    scope_of_report = st.selectbox(
        "Scope of Report (报告范围) *",
        ["All carve out", "Inclusive", "None"],
    )
    co_website = st.text_input("Company Website (公司网站)", max_chars=256)

st.subheader("Trust Service Criteria (SOC2)")
tsc_col1, tsc_col2, tsc_col3, tsc_col4, tsc_col5 = st.columns(5)
is_security = tsc_col1.checkbox("Security", value=False)
is_availability = tsc_col2.checkbox("Availability", value=False)
is_processing_integrity = tsc_col3.checkbox("Processing Integrity", value=False)
is_confidentiality = tsc_col4.checkbox("Confidentiality", value=False)
is_privacy = tsc_col5.checkbox("Privacy", value=False)

st.subheader("Optional Fields")
opt_col1, opt_col2 = st.columns(2)

with opt_col1:
    system = st.text_input(
        "Internal Supporting Systems (内部支撑系统)",
        help="If empty, the system will auto-extract from control points.",
        max_chars=256,
    )
    systems_function = st.text_input("Systems Function (内部支撑系统的用途)", max_chars=256)
    domain = st.text_input(
        "Control Domain (控制领域, if not in uploaded document)",
        max_chars=256,
    )

with opt_col2:
    subservice_org = st.text_input(
        "Subservice Organization (子服务机构名称)",
        max_chars=256,
    )
    service_provided_by_sub_org = st.text_input(
        "Service Provided by Subservice Org (子服务机构提供的服务内容)",
        max_chars=256,
    )

st.subheader("Upload Control Matrix File(s) *")
uploaded_files = st.file_uploader(
    "Upload files (Excel, PDF, Word, etc.)",
    accept_multiple_files=True,
)

st.markdown("---")

if st.button("Generate Report", type="primary", use_container_width=True):
    # ── Validation ─────────────────────────────────────────────────────────────
    errors = []
    if not api_key:
        errors.append("API Secret Key is required (set in sidebar).")
    if not company_name:
        errors.append("Company Name is required.")
    if not co_short_name:
        errors.append("Company Short Name is required.")
    if not system_or_service_name:
        errors.append("Service/System Name is required.")
    if not service_description:
        errors.append("Service Description is required.")
    if not period_start:
        errors.append("Report Period Start is required.")
    if not uploaded_files:
        errors.append("At least one file must be uploaded.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    with st.spinner("Uploading file(s) to Dify…"):
        try:
            uploaded_file_ids = []
            for uf in uploaded_files:
                fid = upload_file_to_dify(uf.read(), uf.name, api_base, api_key)
                uploaded_file_ids.append(
                    {
                        "transfer_method": "local_file",
                        "upload_file_id": fid,
                        "type": "document",
                    }
                )
        except requests.HTTPError as exc:
            st.error(f"File upload failed: {exc.response.status_code} — {exc.response.text}")
            st.stop()
        except Exception as exc:
            st.error(f"File upload error: {exc}")
            st.stop()

    inputs = {
        "File_input": uploaded_file_ids,
        "Report_type": report_type,
        "Output_language": output_language,
        "Company_name": company_name,
        "Co_short_name": co_short_name,
        "Industry": industry,
        "System_or_service_name": system_or_service_name,
        "Service_description": service_description,
        "System": system,
        "Systems_function": systems_function,
        "Domain": domain,
        "Period_start": period_start,
        "Period_end": period_end,
        "Subservice_org": subservice_org,
        "Service_provided_by_sub_org": service_provided_by_sub_org,
        "Scope_of_the_report": scope_of_report,
        "Co_website": co_website,
        "is_Security": is_security,
        "is_Availability": is_availability,
        "is_Processing_Integrity": is_processing_integrity,
        "is_Confidentiality": is_confidentiality,
        "is_Privacy": is_privacy,
    }

    with st.spinner("Running workflow — this may take several minutes…"):
        try:
            outputs = run_workflow(inputs, api_base, api_key)
        except requests.HTTPError as exc:
            st.error(f"Workflow error: {exc.response.status_code} — {exc.response.text}")
            st.stop()
        except Exception as exc:
            st.error(f"Workflow error: {exc}")
            st.stop()

    result_text = outputs.get("Result") or outputs.get("Error") or str(outputs)

    if not result_text:
        st.warning("Workflow completed but returned no output.")
        st.stop()

    if outputs.get("Error"):
        st.error(f"Workflow returned an error:\n\n{result_text}")
        st.stop()

    st.success("Report generated successfully!")

    # Show preview
    with st.expander("Preview result", expanded=False):
        st.markdown(result_text)

    # Convert to Word and offer download
    docx_bytes = markdown_to_docx(result_text)
    filename = f"{co_short_name}_{report_type.replace(' ', '_')}_Report.docx"

    st.download_button(
        label="Download Report (.docx)",
        data=docx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
