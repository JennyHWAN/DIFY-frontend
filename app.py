import streamlit as st
import requests
import os
import io
import re
import json
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from dotenv import load_dotenv

load_dotenv()

API_BASE_URL  = os.getenv("DIFY_API_BASE_URL", "https://api.dify.ai/v1")
API_KEY_MAIN  = os.getenv("DIFY_API_KEY_MAIN", "")
API_KEY_SUB1  = os.getenv("DIFY_API_KEY_SUB1", "")
API_KEY_SUB2  = os.getenv("DIFY_API_KEY_SUB2", "")

st.set_page_config(page_title="AI-Driven Report Generation", layout="wide")
st.title("AI-Driven SOC Report Generation")

# ── API config — loaded from bundled .env, never shown in UI ──────────────────
api_base = API_BASE_URL
key_main = API_KEY_MAIN
key_sub1 = API_KEY_SUB1
key_sub2 = API_KEY_SUB2

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("SOC Report Generator")
    st.markdown("AI-Driven Report Generation")
    st.markdown("---")
    if st.button("🔄 Reset All Steps", use_container_width=True):
        for k in ["main_outputs", "sub1_outputs", "final_result", "user_inputs"]:
            st.session_state.pop(k, None)
        st.rerun()

# ── Progress indicator ─────────────────────────────────────────────────────────
main_done  = "main_outputs"  in st.session_state
sub1_done  = "sub1_outputs"  in st.session_state
final_done = "final_result"  in st.session_state

s1 = "✅" if main_done  else "🔵"
s2 = "✅" if sub1_done  else ("🟡" if main_done  else "⚪")
s3 = "✅" if final_done else ("🟡" if sub1_done  else "⚪")

st.markdown(
    f"""
    <div style='display:flex;gap:2rem;padding:0.5rem 0 1.2rem 0;font-size:1rem'>
        <span>{s1} <b>Step 1</b> — MAIN: Extract &amp; Prepare</span>
        <span>→</span>
        <span>{s2} <b>Step 2</b> — SUB1: Entity Level</span>
        <span>→</span>
        <span>{s3} <b>Step 3</b> — SUB2: Final Report</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Helpers ────────────────────────────────────────────────────────────────────

def upload_file(file_bytes, filename, api_base, api_key):
    url = f"{api_base.rstrip('/')}/files/upload"
    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {api_key}"},
        files={"file": (filename, file_bytes, "application/octet-stream")},
        data={"user": "streamlit-user"},
        timeout=120,
        verify=False,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def to_str(v) -> str:
    """Ensure any workflow output value is a plain string before passing as input."""
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


def run_workflow(inputs, api_base, api_key, status_placeholder=None):
    """
    Calls the Dify workflow API in streaming mode to avoid nginx 504 timeouts.
    Parses SSE events and returns the final outputs dict from workflow_finished.
    """
    url = f"{api_base.rstrip('/')}/workflows/run"

    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={"inputs": inputs, "response_mode": "streaming", "user": "streamlit-user"},
        stream=True,
        timeout=(30, 1800),  # (connect timeout, read timeout in seconds)
        verify=False,
    )
    resp.raise_for_status()

    outputs = {}
    node_count = 0
    workflow_finished_received = False

    for raw_line in resp.iter_lines():
        if not raw_line:
            continue
        line = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
        if not line.startswith("data: "):
            continue
        try:
            event = json.loads(line[6:])
        except json.JSONDecodeError:
            continue

        event_type = event.get("event", "")

        if event_type == "node_finished":
            node_count += 1
            node_data = event.get("data", {})
            node_title = str(node_data.get("title", ""))
            node_status = str(node_data.get("status", ""))
            node_error = str(node_data.get("error", "")) if node_status == "failed" else ""
            # Release the (potentially very large) event dict before any
            # Streamlit UI call — the MAIN-Output end-node payload contains
            # all 19 output fields and can be several hundred KB.
            event = node_data = None
            if status_placeholder:
                status_placeholder.info(f"⚙️ Nodes completed: {node_count}   (last: {node_title})")
            if node_status == "failed":
                raise RuntimeError(f"Workflow node failed: {node_error or f'node {node_title!r} (#{node_count})'}")

        elif event_type == "workflow_finished":
            workflow_finished_received = True
            data = event.get("data", {})
            if data.get("status") == "failed":
                raise RuntimeError(f"Workflow failed: {data.get('error', 'unknown error')}")
            outputs = data.get("outputs", {})
            if status_placeholder:
                status_placeholder.empty()
            break

        elif event_type == "error":
            raise RuntimeError(event.get("message", "Streaming error from Dify"))

    if not workflow_finished_received:
        raise RuntimeError(
            f"Workflow stream ended unexpectedly after {node_count} node(s) "
            "without a completion event — the workflow may have crashed or timed out on the Dify side."
        )

    return outputs


FONT_LATIN   = "Times New Roman"
FONT_CHINESE = "黑体"


def _apply_fonts(run):
    """Set Times New Roman for Latin characters, 黑体 for Chinese characters.
    Word automatically picks the right one per character based on Unicode range."""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"),    FONT_LATIN)
    rFonts.set(qn("w:hAnsi"),    FONT_LATIN)
    rFonts.set(qn("w:eastAsia"), FONT_CHINESE)
    rFonts.set(qn("w:cs"),       FONT_LATIN)


def _set_style_fonts(style):
    """Apply the same dual-font setting at the paragraph-style level."""
    rPr = style.element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        style.element.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"),    FONT_LATIN)
    rFonts.set(qn("w:hAnsi"),    FONT_LATIN)
    rFonts.set(qn("w:eastAsia"), FONT_CHINESE)
    rFonts.set(qn("w:cs"),       FONT_LATIN)


def _set_cell_background(cell, fill_hex):
    """Apply a solid background fill to a table cell. fill_hex e.g. 'D9D9D9'."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def _add_table_from_md(doc, table_lines):
    """Parse markdown table lines and add a Word table with a gray header row."""
    if len(table_lines) < 2:
        return

    def parse_row(line):
        return [cell.strip() for cell in line.strip().strip("|").split("|")]

    def is_separator(cells):
        return all(re.match(r"^:?-+:?$", c) for c in cells if c)

    headers = parse_row(table_lines[0])
    num_cols = len(headers)

    data_rows = []
    for line in table_lines[1:]:
        cells = parse_row(line)
        if not is_separator(cells):
            data_rows.append(cells)

    table = doc.add_table(rows=1 + len(data_rows), cols=num_cols)
    table.style = "Table Grid"

    # Header row — gray background, bold text
    for j, header_text in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(header_text)
        run.bold = True
        _apply_fonts(run)
        _set_cell_background(cell, "D9D9D9")

    # Data rows — no background
    for i, row_data in enumerate(data_rows):
        for j in range(num_cols):
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            text = row_data[j] if j < len(row_data) else ""
            _inline(cell.paragraphs[0], text)


def _add_numpipe_table(doc, lines, language="English"):
    """Parse 'N | Term | Description' rows (no leading pipe) and add a Word table."""
    rows = []
    for line in lines:
        parts = [c.strip() for c in line.strip().split("|") if c.strip()]
        if parts:
            rows.append(parts)
    if not rows:
        return
    num_cols = max(len(r) for r in rows)
    if language.startswith("中"):
        hdr = ("编号", "名词/系统名称", "名词解释/系统简介")
    else:
        hdr = ("SN", "Term/Application Name", "Terminology/System Introduction")
    header_row = list(hdr[:num_cols])
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.style = "Table Grid"
    for j, h in enumerate(header_row):
        cell = table.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        _apply_fonts(run)
        _set_cell_background(cell, "D9D9D9")
    for i, row_data in enumerate(rows):
        for j in range(num_cols):
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            text = row_data[j] if j < len(row_data) else ""
            run = cell.paragraphs[0].add_run(text)
            _apply_fonts(run)
    _set_col_widths(table, [1.2, 5.0, 9.7])
    _set_repeat_header(table)


def _set_col_widths(tbl, widths_cm):
    """Force fixed column widths (list of cm values) on a table."""
    tblPr = tbl._tbl.tblPr
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    for row in tbl.rows:
        for j, w_cm in enumerate(widths_cm):
            if j < len(row.cells):
                row.cells[j].width = Cm(w_cm)


def _set_repeat_header(table):
    """Make the first row repeat as a header row when the table spans pages."""
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)


def _add_heading(doc, text, level, italic=False):
    """Add a Heading-N paragraph: black, 11 pt, bold; italic only when caller requests it."""
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run(text)
    _apply_fonts(run)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.size = Pt(11)
    if italic:
        run.font.italic = True
    blank = doc.add_paragraph()
    blank.paragraph_format.space_after  = Pt(0)
    blank.paragraph_format.space_before = Pt(0)


def _next_nonblank_is_heading(lines, i):
    """Return True if the next non-blank line after index i looks like a heading.

    Used to decide whether the *current* heading is a section header (not italic)
    or a subsection header (italic): if the next content is itself a heading, the
    current one is a parent / section-level entry and should not be italicised.
    """
    j = i + 1
    while j < len(lines) and lines[j].strip() == "":
        j += 1
    if j >= len(lines):
        return False
    nxt = lines[j]
    if nxt.startswith("#"):
        return True
    s = nxt.strip()
    return bool(s) and len(s) <= 130 and not nxt.rstrip().endswith(('.', '?', '!', ',', ';', ':'))


def _section_contains_subheadings(lines, i):
    """Return True if there is an H3 or H4 heading anywhere between line i+1
    and the next H1/H2 boundary (i.e. within the body of this H2 section).

    Used to identify section-container H2 headings that should be bold-only even
    when their immediately following line is body text rather than a heading.
    """
    j = i + 1
    while j < len(lines):
        ln = lines[j]
        if ln.startswith("# ") or ln.startswith("## "):
            break
        if ln.startswith("### ") or ln.startswith("#### "):
            return True
        j += 1
    return False


def markdown_to_docx(md_text: str, language: str = "English") -> bytes:
    doc = Document()

    _set_style_fonts(doc.styles["Normal"])
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.styles["Normal"].paragraph_format.space_after  = Pt(0)
    doc.styles["Normal"].paragraph_format.space_before = Pt(0)

    lines = md_text.split("\n")
    i = 0
    # Track whether the last thing added was a blank paragraph so we never
    # emit more than one consecutive blank line regardless of how many '\n'
    # sequences the LLM produced.  Start True so leading blank lines are
    # silently dropped.
    last_was_blank = True

    while i < len(lines):
        line = lines[i]

        # ── Markdown table block ───────────────────────────────────────────
        if line.strip().startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            _add_table_from_md(doc, table_lines)
            last_was_blank = False
            continue

        # ── Number-pipe table  (N | Term | Description) ────────────────────
        if re.match(r"^\d+\s*\|", line.strip()):
            npt_lines = []
            while i < len(lines) and re.match(r"^\d+\s*\|", lines[i].strip()):
                npt_lines.append(lines[i])
                i += 1
            _add_numpipe_table(doc, npt_lines, language)
            last_was_blank = False
            continue

        # ── Headings ───────────────────────────────────────────────────────
        # H4 / H3 — bold only, no italic
        if line.startswith("#### "):
            _add_heading(doc, line[5:].strip(), 4, italic=False)
            last_was_blank = True  # _add_heading appends a blank internally

        elif line.startswith("### "):
            _add_heading(doc, line[4:].strip(), 3, italic=False)
            last_was_blank = True

        # H2 — bold-only when:
        #   (a) the immediately-next non-blank line is itself a heading, OR
        #   (b) this section contains H3/H4 sub-headings (section container).
        # Otherwise italic (content subsection label).
        elif line.startswith("## "):
            italic = not (_next_nonblank_is_heading(lines, i)
                          or _section_contains_subheadings(lines, i))
            _add_heading(doc, line[3:].strip(), 2, italic=italic)
            last_was_blank = True

        # H1 — bold only, never italic
        elif line.startswith("# "):
            _add_heading(doc, line[2:].strip(), 1, italic=False)
            last_was_blank = True

        # ── Bullet lists ───────────────────────────────────────────────────
        elif re.match(r"^[-*+] ", line):
            p = doc.add_paragraph(style="List Bullet")
            _inline_bullet(p, line[2:].strip())
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            last_was_blank = False

        elif re.match(r"^[•·]\s+", line):
            p = doc.add_paragraph(style="List Bullet")
            _inline_bullet(p, re.sub(r"^[•·]\s+", "", line))
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            last_was_blank = False

        elif re.match(r"^\d+\. ", line):
            # Collect the whole consecutive numbered block, keeping the SN
            numbered_items = []
            while i < len(lines) and re.match(r"^\d+\. ", lines[i]):
                m = re.match(r"^(\d+)\. (.*)", lines[i])
                numbered_items.append((m.group(1), m.group(2).strip()))
                i += 1
            # If every item follows "Term: Description", convert to a 3-column table
            _def_pat = re.compile(r"^(?!https?:)([^:]{1,80}):\s+(.*)", re.DOTALL)
            texts = [rest for _, rest in numbered_items]
            if len(texts) >= 2 and all(_def_pat.match(t) for t in texts):
                if language.startswith("中"):
                    hdr = ("编号", "名词/系统名称", "名词解释/系统简介")
                else:
                    hdr = ("SN", "Term/Application Name", "Terminology/System Introduction")
                tbl = doc.add_table(rows=1 + len(numbered_items), cols=3)
                tbl.style = "Table Grid"
                for j, h in enumerate(hdr):
                    cell = tbl.rows[0].cells[j]
                    cell.text = ""
                    r = cell.paragraphs[0].add_run(h)
                    r.bold = True
                    _apply_fonts(r)
                    _set_cell_background(cell, "D9D9D9")
                for idx, (sn, rest) in enumerate(numbered_items):
                    m = _def_pat.match(rest)
                    term, desc = m.group(1).strip(), m.group(2)
                    for j, text in enumerate((sn, term, desc)):
                        cell = tbl.rows[idx + 1].cells[j]
                        cell.text = ""
                        run = cell.paragraphs[0].add_run(text)
                        _apply_fonts(run)
                _set_col_widths(tbl, [1.2, 5.0, 9.7])
                _set_repeat_header(tbl)
            else:
                for _, rest in numbered_items:
                    p = doc.add_paragraph(style="List Number")
                    _inline_bullet(p, rest)
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            last_was_blank = False
            continue  # i already advanced past the block

        # ── Blank line — emit at most one consecutive blank paragraph ──────
        elif line.strip() == "":
            if not last_was_blank:
                blank = doc.add_paragraph()
                blank.paragraph_format.space_after  = Pt(0)
                blank.paragraph_format.space_before = Pt(0)
                last_was_blank = True

        # ── Plain-text heading (no # marker: short line, no terminal punct) ─
        elif line.strip() and len(line.strip()) <= 130 and not line.rstrip().endswith(('.', '?', '!', ',', ';', ':')):
            italic = not (_next_nonblank_is_heading(lines, i)
                          or _section_contains_subheadings(lines, i))
            _add_heading(doc, line.strip(), 2, italic=italic)
            last_was_blank = True

        # ── Normal paragraph ───────────────────────────────────────────────
        else:
            p = doc.add_paragraph()
            _inline(p, line)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_after  = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            last_was_blank = False

        i += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def _inline(paragraph, text):
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
            _apply_fonts(run)
        elif part.startswith("*") and part.endswith("*"):
            run = paragraph.add_run(part[1:-1])
            run.italic = True
            _apply_fonts(run)
        else:
            run = paragraph.add_run(part)
            _apply_fonts(run)


def _inline_bullet(paragraph, text):
    """For bullet items: if markdown bold markers are present use _inline as-is;
    otherwise auto-bold the 'Header: content' pattern before the first colon."""
    if "**" in text:
        _inline(paragraph, text)
        return
    # Match 'Header: content' — colon within 80 chars, not a URL scheme (http/https)
    m = re.match(r"^(?!https?:)([^:]{1,80}):\s*(.*)", text, re.DOTALL)
    if m:
        header, content = m.group(1).strip(), m.group(2)
        run = paragraph.add_run(header + ": ")
        run.bold = True
        _apply_fonts(run)
        if content:
            run2 = paragraph.add_run(content)
            _apply_fonts(run2)
    else:
        _inline(paragraph, text)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — User inputs + MAIN workflow
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📋 Step 1 — Report Parameters & MAIN Workflow", expanded=not main_done):

    st.subheader("Upload Control Matrix File(s)")
    st.caption(
        "1. 请上传Excel大表，其中必须包含Control Matrix及组织架构描述相关sheet；"
        "2. 若有，上传子服务机构及其服务内容Sheet；"
        "3. 若有，上传CUEC内容清单Sheet；"
        "4. 若有，上传子服务机构补充控制清单Sheet；"
        "5. 若有，上传名词解释Sheet；"
        "6. 若适用，上传控制目标Sheet"
    )
    uploaded_files = st.file_uploader(
        "Upload files (Excel, PDF, Word, etc.)",
        accept_multiple_files=True,
    )

    # ── Required fields ───────────────────────────────────────────────────────
    st.subheader("Required Fields")
    req1, req2 = st.columns(2)

    with req1:
        company_name        = st.text_input("Company Name",         max_chars=256)
        co_short_name       = st.text_input("Company Short Name",   max_chars=48)
        system_name         = st.text_input("Service / System Name",max_chars=256)
        period_start    = st.text_input("Report Period Start (as of date if SOC1)", placeholder="e.g. 2025-01-01")
        scope_of_report = st.selectbox("Subservice Organization Testing Strategy",
            ["None", "All carve out", "Inclusive"])

    with req2:
        report_type     = st.selectbox("Report Type",
            ["SOC1 TYPE1", "SOC1 TYPE2", "SOC2 TYPE1", "SOC2 TYPE2"])
        output_language = st.selectbox("Output Language",
            ["English", "中文"])
        service_description = st.text_input("Service Description",  max_chars=256)
        period_end = st.text_input("Report Period End (N/A for Type1)", placeholder="e.g. 2024-12-31")
        subservice_org = st.text_area(
                            "Subservice Organization (N/A if no subservice organization)",
                            placeholder="Alibaba Cloud | Elastic Cloud, Object Storage\nTencent Cloud | Cloud Virtual Machine, TencentDB",
                            help="Required if exist Subservice Organization. One entry per line, format: Organization Name | Services Used\nExample: Alibaba Cloud | Elastic Cloud, Object Storage",
                            height=100,
                        )
        if len(subservice_org) > 256:
            st.warning("⚠️ Subservice Organization exceeds 256 characters. Please shorten it.")


    # ── Optional fields ────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Optional Fields")
    opt1, opt2 = st.columns(2)

    with opt1:
        industry         = st.selectbox("Industry",
            ["Other", "HR", "Cloud Service", "AI", "General"])
        co_website       = st.text_input("Company Website",               max_chars=256)
        system_extra     = st.text_input("Internal Supporting Systems",
                            placeholder="e.g. Feishu Platform, Gitlab Platform, Alibaba Cloud Console",
                            help="Optional. List the internal systems used to support operations. If left blank, the workflow will auto-extract from the Control Matrix.",   
                            max_chars=256)

    with opt2:
        domain         = st.text_input("Control Domain",                  max_chars=256)
        systems_function = st.text_input("Systems Function",
                        placeholder="e.g. workflow approval, code management, cloud resource management",
                        help="Optional. Describe the purpose of the internal supporting systems listed above.",              
                        max_chars=256)

    st.subheader("Trust Service Criteria (SOC2 only)")
    tsc_cols = st.columns(5)
    is_security              = tsc_cols[0].checkbox("Security")
    is_availability          = tsc_cols[1].checkbox("Availability")
    is_processing_integrity  = tsc_cols[2].checkbox("Processing Integrity")
    is_confidentiality       = tsc_cols[3].checkbox("Confidentiality")
    is_privacy               = tsc_cols[4].checkbox("Privacy")

    st.markdown("---")
    run_main = st.button("▶ Run Step 1 — MAIN Workflow", type="primary", use_container_width=True)

    if run_main:
        errors = []
        if not key_main:   errors.append("MAIN Workflow API key is not set. Add DIFY_API_KEY_MAIN to the .env file next to the executable.")
        if not company_name: errors.append("Company Name is required.")
        if not co_short_name: errors.append("Company Short Name is required.")
        if not system_name: errors.append("Service/System Name is required.")
        if not service_description: errors.append("Service Description is required.")
        if not period_start: errors.append("Report Period Start is required.")
        if not uploaded_files: errors.append("At least one file must be uploaded.")
        if len(subservice_org) > 256:
            errors.append("Subservice Organization exceeds 256 characters")
        if (scope_of_report != "None") and (not subservice_org):
            errors.append("Please input Subservice Organizations and its service provided")

        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        # Upload files
        with st.spinner("Uploading file(s) to Dify…"):
            try:
                file_ids = []
                for uf in uploaded_files:
                    fid = upload_file(uf.read(), uf.name, api_base, key_main)
                    file_ids.append({
                        "transfer_method": "local_file",
                        "upload_file_id": fid,
                        "type": "document",
                    })
            except requests.HTTPError as e:
                st.error(f"File upload failed: {e.response.status_code} — {e.response.text}")
                st.stop()
            except Exception as e:
                st.error(f"File upload error: {e}")
                st.stop()

        # Store user inputs for later steps
        st.session_state["user_inputs"] = {
            "Report_type": report_type,
            "Output_language": output_language,
            "Company_name": company_name,
            "Co_short_name": co_short_name,
            "Industry": industry,
            "System_or_service_name": system_name,
            "Service_description": service_description,
            "System": system_extra,
            "Systems_function": systems_function,
            "Domain": domain,
            "Period_start": period_start,
            "Period_end": period_end,
            "Subservice_org": subservice_org,
            "Scope_of_the_report": scope_of_report,
            "Co_website": co_website,
        }

        inputs_main = {
            **st.session_state["user_inputs"],
            "File_input": file_ids,
            "is_Security":             is_security,
            "is_Availability":         is_availability,
            "is_Processing_Integrity": is_processing_integrity,
            "is_Confidentiality":      is_confidentiality,
            "is_Privacy":              is_privacy,
        }

        status = st.empty()
        with st.spinner("Running MAIN workflow — this may take several minutes…"):
            try:
                outputs = run_workflow(inputs_main, api_base, key_main, status)
            except requests.HTTPError as e:
                st.error(f"MAIN workflow error: {e.response.status_code} — {e.response.text}")
                st.stop()
            except Exception as e:
                st.error(f"MAIN workflow error: {e}")
                st.stop()

        if outputs.get("Error") or outputs.get("Error_ORG"):
            st.error(f"MAIN workflow returned an error:\n{outputs.get('Error') or outputs.get('Error_ORG')}")
            st.stop()

        st.session_state["main_outputs"] = outputs
        st.success("✅ Step 1 complete — MAIN workflow finished.")
        st.rerun()

if main_done:
    with st.expander("🔍 MAIN Workflow Outputs (preview)", expanded=False):
        mo = st.session_state["main_outputs"]
        for k, v in mo.items():
            if v:
                st.markdown(f"**{k}**")
                st.text(str(v)[:500] + ("…" if len(str(v)) > 500 else ""))


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — SUB1: Entity Level
# ══════════════════════════════════════════════════════════════════════════════
if main_done:
    with st.expander("🏛 Step 2 — SUB1: Entity Level Controls", expanded=not sub1_done):
        st.write("SUB1 generates the Entity-Level Controls section using MAIN's outputs.")

        run_sub1 = st.button("▶ Run Step 2 — SUB1 Workflow", type="primary", use_container_width=True)

        if run_sub1:
            if not key_sub1:
                st.error("SUB1 API key is required (set in sidebar).")
                st.stop()

            mo = st.session_state["main_outputs"]
            ui = st.session_state.get("user_inputs", {})

            inputs_sub1 = {
                # Sections from MAIN
                "overview_section":               to_str(mo.get("overview_section")),
                "principals_section":             to_str(mo.get("principals_section")),
                "scope_section":                  to_str(mo.get("scope_section")),
                "org_overview_section":           to_str(mo.get("org_overview_section")),
                "clean_pairs":                    to_str(mo.get("clean_pairs")),
                "activity_domains_text":          to_str(mo.get("activity_domains_text")),
                "entity_domain_controls":         to_str(mo.get("entity_domain_controls")),
                "needing_assignment":             to_str(mo.get("needing_assignment")),
                "CO_list_text":                   to_str(mo.get("CO_list_text")),
                "CUEC_json":                      to_str(mo.get("CUEC_json")),
                "CSOC_json":                      to_str(mo.get("CSOC_json")),
                "Terminology_json":               to_str(mo.get("Terminology_json")),
                "control_list_text":              to_str(mo.get("control_list_text")),
                "org_structure":                  to_str(mo.get("org_structure")),
                "website_content":                to_str(mo.get("website_content")),
                "activity_control_objectives_json": to_str(mo.get("activity_control_objectives_json")),
                "entity_control_objective":       to_str(mo.get("entity_control_objective")),
                "entity_domain_packs_direct":     to_str(mo.get("entity_domain_packs_direct")),
                "rag_context":                    to_str(mo.get("rag_context")),
                # User inputs passed through
                "Report_type":            ui.get("Report_type", ""),
                "Output_language":        ui.get("Output_language", ""),
                "Company_name":           ui.get("Company_name", ""),
                "Co_short_name":          ui.get("Co_short_name", ""),
                "Industry":               ui.get("Industry", ""),
                "System_or_service_name": ui.get("System_or_service_name", ""),
                "Subservice_org":         ui.get("Subservice_org", ""),
                "Scope_of_the_report":    ui.get("Scope_of_the_report", ""),
                "Domain":                 ui.get("Domain", ""),
                "Period_start":           ui.get("Period_start", ""),
                "Period_end":             ui.get("Period_end", ""),
            }

            status = st.empty()
            with st.spinner("Running SUB1 workflow — this may take several minutes…"):
                try:
                    outputs = run_workflow(inputs_sub1, api_base, key_sub1, status)
                except requests.HTTPError as e:
                    st.error(f"SUB1 workflow error: {e.response.status_code} — {e.response.text}")
                    st.stop()
                except Exception as e:
                    st.error(f"SUB1 workflow error: {e}")
                    st.stop()

            st.session_state["sub1_outputs"] = outputs
            st.success("✅ Step 2 complete — SUB1 workflow finished.")
            st.rerun()

    if sub1_done:
        with st.expander("🔍 SUB1 Outputs (entity_level_section preview)", expanded=False):
            entity_sec = st.session_state["sub1_outputs"].get("entity_level_section", "")
            st.text(entity_sec[:1000] + ("…" if len(entity_sec) > 1000 else ""))


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — SUB2: Final Report
# ══════════════════════════════════════════════════════════════════════════════
if sub1_done:
    with st.expander("📄 Step 3 — SUB2: Final Report Assembly", expanded=not final_done):
        st.write("SUB2 assembles the complete SOC report from all previous outputs.")

        run_sub2 = st.button("▶ Run Step 3 — SUB2 Workflow", type="primary", use_container_width=True)

        if run_sub2:
            if not key_sub2:
                st.error("SUB2 API key is required (set in sidebar).")
                st.stop()

            so = st.session_state["sub1_outputs"]
            mo = st.session_state["main_outputs"]
            ui = st.session_state.get("user_inputs", {})

            inputs_sub2 = {
                # All outputs from SUB1 (passthrough + entity_level_section)
                "overview_section":               to_str(so.get("overview_section")),
                "principals_section":             to_str(so.get("principals_section")),
                "scope_section":                  to_str(so.get("scope_section")),
                "org_overview_section":           to_str(so.get("org_overview_section")),
                "entity_level_section":           to_str(so.get("entity_level_section")),
                "clean_pairs":                    to_str(so.get("clean_pairs")),
                "activity_domains_text":          to_str(so.get("activity_domains_text")),
                "CO_list_text":                   to_str(so.get("CO_list_text")),
                "CUEC_json":                      to_str(so.get("CUEC_json")),
                "CSOC_json":                      to_str(so.get("CSOC_json")),
                "Terminology_json":               to_str(so.get("Terminology_json")),
                "website_content":                to_str(so.get("website_content")),
                "activity_control_objectives_json": to_str(so.get("activity_control_objectives_json")),
                "rag_context":                    to_str(so.get("rag_context")),
                "Subservice_org":         to_str(so.get("Subservice_org") or ui.get("Subservice_org")),
                "Scope_of_the_report":    to_str(so.get("Scope_of_the_report") or ui.get("Scope_of_the_report")),
                "Period_start":           to_str(so.get("Period_start") or ui.get("Period_start")),
                "Period_end":             to_str(so.get("Period_end") or ui.get("Period_end")),
                "Report_type":            to_str(so.get("Report_type") or ui.get("Report_type")),
                "Output_language":        to_str(so.get("Output_language") or ui.get("Output_language")),
                "Company_name":           to_str(so.get("Company_name") or ui.get("Company_name")),
                "Co_short_name":          to_str(so.get("Co_short_name") or ui.get("Co_short_name")),
                "System_or_service_name": to_str(so.get("System_or_service_name") or ui.get("System_or_service_name")),
                "cuec_preformatted":      to_str(mo.get("cuec_preformatted")),
            }

            status = st.empty()
            with st.spinner("Running SUB2 workflow — this may take several minutes…"):
                try:
                    outputs = run_workflow(inputs_sub2, api_base, key_sub2, status)
                except requests.HTTPError as e:
                    st.error(f"SUB2 workflow error: {e.response.status_code} — {e.response.text}")
                    st.stop()
                except Exception as e:
                    st.error(f"SUB2 workflow error: {e}")
                    st.stop()

            result = outputs.get("Result", "")
            if not result:
                st.warning("SUB2 completed but returned no output.")
                st.stop()

            st.session_state["final_result"] = result
            st.success("✅ Step 3 complete — Report generated successfully!")
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# FINAL RESULT — Preview + Download
# ══════════════════════════════════════════════════════════════════════════════
if final_done:
    st.markdown("---")
    st.success("🎉 Report is ready!")

    result_text = st.session_state["final_result"]
    ui = st.session_state.get("user_inputs", {})

    with st.expander("📖 Preview Report", expanded=True):
        st.markdown(result_text)

    docx_bytes = markdown_to_docx(result_text, ui.get("Output_language", "English"))
    filename = (
        f"{ui.get('Co_short_name', 'Report')}_"
        f"{ui.get('Report_type', '').replace(' ', '_')}_Report.docx"
    )

    st.download_button(
        label="⬇ Download Report (.docx)",
        data=docx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
