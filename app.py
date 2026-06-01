import streamlit as st
import requests
import os
import io
import re
import json
import zipfile
import xml.etree.ElementTree as ET
from copy import deepcopy
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from dotenv import load_dotenv
import openpyxl

load_dotenv()

import sys as _sys

if getattr(_sys, "frozen", False):
    # Running as a PyInstaller exe — templates sit next to the .exe
    _TEMPLATE_BASE = os.path.dirname(_sys.executable)
else:
    # Development — templates are in the same folder as app.py (DIFY-frontend/)
    _TEMPLATE_BASE = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_INDEX  = os.path.join(_TEMPLATE_BASE, "template_index.xlsx")
AR_TEMPLATE_DIR = os.path.join(_TEMPLATE_BASE, "AR_template")
MA_TEMPLATE_DIR = os.path.join(_TEMPLATE_BASE, "MA_template")
EY_FIRM_NAME    = "Ernst & Young Hua Ming LLP"

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
        for k in ["main_outputs", "sub1_outputs", "final_result", "user_inputs", "template_config"]:
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

# ── Template helpers ───────────────────────────────────────────────────────────

def get_standard_options(report_type):
    """Return the list of applicable attestation standards for a given report type."""
    if report_type.startswith("SOC1"):
        return ["SSAE 18", "ISAE 3402", "SSAE 18 & ISAE 3402 Combined"]
    else:
        return ["SSAE 18", "ISAE 3000", "SSAE 18 & ISAE 3000 Combined"]


def resolve_template(report_type, standard, sso, language, sheet):
    """
    Look up template_index.xlsx for a matching row and return (wp_no, filepath|None).
    sheet must be 'AR' or 'MA'.
    """
    # Map UI values → spreadsheet values
    if report_type.startswith("SOC1"):
        category = "SOC 1"
    else:
        category = "SOC 2"

    if "TYPE1" in report_type:
        typ = "Type I"
    else:
        typ = "Type II"

    if "Combined" in standard:
        std_mapped = "Combined"
    else:
        std_mapped = standard  # "SSAE 18", "ISAE 3402", "ISAE 3000"

    sso_map = {"None": "none", "All carve out": "all carve out", "Inclusive": "Inclusive"}
    sso_mapped = sso_map.get(sso, "none")

    lang_map = {"English": "EN", "中文": "CN"}
    lang_mapped = lang_map.get(language, "EN")

    template_dir = AR_TEMPLATE_DIR if sheet == "AR" else MA_TEMPLATE_DIR

    try:
        wb = openpyxl.load_workbook(TEMPLATE_INDEX, read_only=True, data_only=True)
        ws = wb[sheet]
        # Column layout for AR: SN, Category, Type, Standards, SSO, Language, ..., WP No (last)
        # Column layout for MA: SN, Category, Type, Standards, SSO, Language, Comments, WP No.
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        # Find column indices
        col_cat  = header.index("Category")
        col_type = header.index("Type")
        col_std  = header.index("Standards")
        col_sso  = header.index("Sub-service Organization (SSO)")
        col_lang = header.index("Language")
        # WP No column name differs between sheets
        col_wp = next(
            i for i, h in enumerate(header)
            if h and str(h).strip().upper().startswith("WP")
        )

        for row in ws.iter_rows(min_row=2, values_only=True):
            if (str(row[col_cat] or "").strip() == category and
                    str(row[col_type] or "").strip() == typ and
                    str(row[col_std] or "").strip() == std_mapped and
                    str(row[col_sso] or "").strip() == sso_mapped and
                    str(row[col_lang] or "").strip() == lang_mapped):
                wp_val = row[col_wp]
                if wp_val is None:
                    return (None, None)
                wp_no = str(wp_val).strip()
                # Search template dir for a file starting with the WP number
                try:
                    for fname in os.listdir(template_dir):
                        if fname.endswith(".docx") and fname.startswith(wp_no + " "):
                            return (wp_no, os.path.join(template_dir, fname))
                except OSError:
                    pass
                return (wp_no, None)
        return (None, None)
    except Exception:
        return (None, None)


def _normalize_ws(s):
    """Collapse all whitespace sequences to a single space and strip."""
    return " ".join(s.split())


_MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def _format_date(s, language="English"):
    """Convert YYYY-MM-DD to 'Month D, YYYY' (EN) or 'YYYY年M月D日' (CN).
    Strings that do not match YYYY-MM-DD are returned unchanged."""
    if not s:
        return s
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s.strip())
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= mo <= 12:
            if language == "中文":
                return f"{y}年{mo}月{d}日"
            return f"{_MONTHS[mo - 1]} {d}, {y}"
    return s


def _kw_in(comment_text, keyword):
    """Check if keyword appears in comment_text, ignoring all internal whitespace.
    Stripping spaces handles XML-inserted spaces between CJK characters."""
    return keyword.replace(" ", "") in comment_text.replace(" ", "")


# Comment keyword sets for deletion logic
_CUEC_KW     = ["无用户补充", "未识别用户补充", "user entity补充", "无用户实体补充"]
_SSO_CC_KW   = ["子服务机构补偿", "子服务机构补充"]
_TRANS_KW    = ["处理transaction", "processing user entity transaction"]
_SINGLE_KW   = ["single user entity report", "single user entity时", "single user entity 时"]
_OTHER_KW    = ["other information"]  # ALWAYS keep — never delete
_AI_KW       = ["使用到了AI技术", "subject matter中某部分使用到了"]  # AI scope exclusion paragraph


def _build_annotation_maps(docx_bytes, flags):
    """
    Parse the docx XML (from raw bytes) to determine per-paragraph actions.

    Returns:
        del_indices     — set of body-child indices to remove entirely
        single_ue_indices — set of body-child indices where [..] bracket
                            content is conditionally deleted (single user entity)
    """
    ns_w    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    p_tag   = f"{{{ns_w}}}p"
    id_attr = f"{{{ns_w}}}id"

    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
            if "word/comments.xml" in z.namelist():
                comments_xml = z.read("word/comments.xml").decode("utf-8")
            else:
                return set(), set()
    except Exception:
        return set(), set()

    if doc_xml.startswith("\ufeff"):
        doc_xml = doc_xml[1:]
    if comments_xml.startswith("\ufeff"):
        comments_xml = comments_xml[1:]

    # Build comment id → normalised lower-case text map
    try:
        root_c = ET.fromstring(comments_xml)
    except ET.ParseError:
        return set(), set()
    comment_texts = {}
    for comment in root_c.findall(f"{{{ns_w}}}comment"):
        cid = comment.get(id_attr)
        texts = [t.text or "" for t in comment.findall(f".//{{{ns_w}}}t")]
        comment_texts[cid] = _normalize_ws(" ".join(texts)).lower()

    # Build comment id → set of body-child indices (paragraphs only)
    try:
        root_d = ET.fromstring(doc_xml)
    except ET.ParseError:
        return set(), set()
    body = root_d.find(f"{{{ns_w}}}body")
    if body is None:
        return set(), set()

    comment_to_indices = {}
    for i, child in enumerate(body):
        if child.tag != p_tag:
            continue
        for crs in child.findall(f".//{{{ns_w}}}commentRangeStart"):
            cid = crs.get(id_attr)
            if cid:
                comment_to_indices.setdefault(cid, set()).add(i)

    del_indices      = set()
    single_ue_indices = set()

    for cid, text in comment_texts.items():
        # "Other Information" paragraphs are ALWAYS kept
        if any(_kw_in(text, kw.lower()) for kw in _OTHER_KW):
            continue

        indices = comment_to_indices.get(cid, set())
        if not indices:
            continue

        # Single-UE: handle via inline bracket deletion, not whole-para removal
        if any(_kw_in(text, kw.lower()) for kw in _SINGLE_KW):
            single_ue_indices |= indices
            continue

        # AI scope-exclusion paragraph
        if any(_kw_in(text, kw.lower()) for kw in _AI_KW):
            if not flags.get("has_ai_scope_exclusion", False):
                del_indices |= indices
            continue

        should_delete = False
        if not flags.get("cuec_identified", True):
            if any(_kw_in(text, kw.lower()) for kw in _CUEC_KW):
                should_delete = True
        if not flags.get("sso_cc_identified", True):
            if any(_kw_in(text, kw.lower()) for kw in _SSO_CC_KW):
                should_delete = True
        if not flags.get("has_transaction_processing", True):
            if any(_kw_in(text, kw.lower()) for kw in _TRANS_KW):
                should_delete = True

        if should_delete:
            del_indices |= indices

    return del_indices, single_ue_indices


def _clean_docx_bytes(docx_bytes):
    """
    Accept all tracked changes and clear all comments in a docx file.

    Operations performed on word/document.xml via regex (safe for well-formed
    OOXML generated by Word):
      - <w:ins>…</w:ins>  →  unwrapped  (content kept, wrapper removed)
      - <w:del>…</w:del>  →  removed entirely
      - <w:commentRangeStart/> and <w:commentRangeEnd/> self-closing tags → removed
      - <w:r> runs whose only non-rPr child is <w:commentReference/>  → removed

    word/comments.xml is replaced with an empty comments document so that Word
    opens the file without any comment balloons.
    """
    _EMPTY_COMMENTS = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"></w:comments>'
    ).encode("utf-8")

    buf_in  = io.BytesIO(docx_bytes)
    buf_out = io.BytesIO()

    with zipfile.ZipFile(buf_in, "r") as zin:
        with zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    xml_str = data.decode("utf-8").lstrip("\ufeff")
                    # Unwrap <w:ins>…</w:ins>
                    xml_str = re.sub(
                        r"<w:ins\b[^>]*>(.*?)</w:ins>", r"\1",
                        xml_str, flags=re.DOTALL,
                    )
                    # Remove <w:del>…</w:del>
                    xml_str = re.sub(
                        r"<w:del\b[^>]*>.*?</w:del>", "",
                        xml_str, flags=re.DOTALL,
                    )
                    # Remove comment range markers
                    xml_str = re.sub(r"<w:commentRangeStart[^>]*/>", "", xml_str)
                    xml_str = re.sub(r"<w:commentRangeEnd[^>]*/>",   "", xml_str)
                    # Remove runs that only anchor a comment reference
                    xml_str = re.sub(
                        r"<w:r\b[^>]*>\s*(?:<w:rPr>.*?</w:rPr>\s*)?<w:commentReference[^>]*/>\s*</w:r>",
                        "", xml_str, flags=re.DOTALL,
                    )
                    data = xml_str.encode("utf-8")
                elif info.filename == "word/comments.xml":
                    data = _EMPTY_COMMENTS
                zout.writestr(info, data)

    buf_out.seek(0)
    return buf_out.read()


def fill_and_process_template(template_path, subs, flags, language="English"):
    """
    Process an EY MA/AR docx template:
      1. Compute comment-based annotation maps from the original file.
      2. Accept tracked changes and clear comments (regex on raw XML).
      3. Apply placeholder substitutions.
      4. Handle inline square brackets:
           - single-user-entity annotated paras: delete [..] when flag set,
             otherwise strip the brackets and keep the content.
           - [or ..] alternative phrases: always removed.
           - remaining [..] brackets: brackets stripped, content kept.
      5. Delete wholly-conditional paragraphs (CUEC / SSO CC / transaction /
         AI-scope when not applicable).
      6. Standardise fonts: Times New Roman (EN) or 华文楷体 (CN), 11 pt,
         bold removed, italic preserved.

    Returns the modified document as bytes.
    """
    # ── Step 1: annotation maps from the original (comments still present) ──
    with open(template_path, "rb") as fh:
        raw_bytes = fh.read()
    del_indices, single_ue_indices = _build_annotation_maps(raw_bytes, flags)

    # ── Step 2: accept track changes + clear comments ──────────────────────
    cleaned_bytes = _clean_docx_bytes(raw_bytes)
    doc = Document(io.BytesIO(cleaned_bytes))

    # ── Step 3: placeholder substitutions ──────────────────────────────────
    def _apply_subs_to_para(para):
        full_text = "".join(r.text for r in para.runs)
        if not full_text:
            return
        changed = False
        for placeholder, value in subs.items():
            if placeholder in full_text:
                full_text = full_text.replace(placeholder, value)
                changed = True
        if changed and para.runs:
            para.runs[0].text = full_text
            for run in para.runs[1:]:
                run.text = ""

    for para in doc.paragraphs:
        _apply_subs_to_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _apply_subs_to_para(para)

    # ── Step 4: inline bracket handling ────────────────────────────────────
    body_children  = list(doc.element.body)
    body_child_map = {child: i for i, child in enumerate(body_children)}
    _p_tag         = qn("w:p")
    _single_ue_flag = flags.get("single_user_entity", False)

    def _apply_brackets(para, is_single_ue):
        full_text = "".join(r.text for r in para.runs)
        if not full_text or "[" not in full_text:
            return
        original = full_text

        # a) single-user-entity paragraphs
        if is_single_ue:
            if _single_ue_flag:
                # delete the bracketed content entirely
                full_text = re.sub(r"\[[^\]]*\]", "", full_text)
            else:
                # strip the brackets, keep the content
                full_text = re.sub(r"\[([^\]]*)\]", r"\1", full_text)

        # b) [or …] alternative phrases — always remove
        full_text = re.sub(r"\s*\[or [^\]]+\]", "", full_text, flags=re.IGNORECASE)

        # c) strip any remaining brackets, keep content
        full_text = re.sub(r"\[([^\]]*)\]", r"\1", full_text)

        if full_text != original and para.runs:
            para.runs[0].text = full_text
            for run in para.runs[1:]:
                run.text = ""

    # Body-level paragraphs (doc.paragraphs = direct children of <w:body>)
    for para in doc.paragraphs:
        idx        = body_child_map.get(para._element, -1)
        is_sue_para = idx in single_ue_indices
        _apply_brackets(para, is_sue_para)

    # Table-cell paragraphs (never single-UE annotated at body level)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _apply_brackets(para, False)

    # ── Step 5: delete conditionally-excluded paragraphs ───────────────────
    for idx in sorted(del_indices, reverse=True):
        if idx < len(body_children):
            doc.element.body.remove(body_children[idx])

    # ── Step 6: font standardisation ───────────────────────────────────────
    cjk_font = "华文楷体" if language == "中文" else "Times New Roman"

    def _std_run(run):
        run.bold      = False           # strip bold; italic is left untouched
        run.font.size = Pt(11)
        rPr    = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"),    "Times New Roman")
        rFonts.set(qn("w:hAnsi"),    "Times New Roman")
        rFonts.set(qn("w:eastAsia"), cjk_font)
        rFonts.set(qn("w:cs"),       "Times New Roman")

    for para in doc.paragraphs:
        for run in para.runs:
            _std_run(run)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        _std_run(run)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def merge_docx_sections(*docs_bytes):
    """Merge multiple docx byte strings into one document with page breaks between sections."""
    base = Document(io.BytesIO(docs_bytes[0]))
    body = base.element.body
    last_sectPr = body.find(qn("w:sectPr"))

    for extra_bytes in docs_bytes[1:]:
        # Insert a page break before the next section
        p_br = OxmlElement("w:p")
        r_br = OxmlElement("w:r")
        br   = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        r_br.append(br)
        p_br.append(r_br)
        if last_sectPr is not None:
            body.insert(list(body).index(last_sectPr), p_br)
        else:
            body.append(p_br)

        extra = Document(io.BytesIO(extra_bytes))
        for elem in extra.element.body:
            if elem.tag != qn("w:sectPr"):
                if last_sectPr is not None:
                    body.insert(list(body).index(last_sectPr), deepcopy(elem))
                else:
                    body.append(deepcopy(elem))

    buf = io.BytesIO()
    base.save(buf)
    return buf.getvalue()


def build_substitutions(ui, tc):
    """Build the EN + CN placeholder → value substitution dict."""
    language      = ui.get("Output_language", "English")
    company_name  = ui.get("Company_name", "")
    co_short_name = ui.get("Co_short_name", "")
    system_name   = ui.get("System_or_service_name", "")
    report_type   = ui.get("Report_type", "")
    subservice_org = ui.get("Subservice_org", "")
    signing_city  = tc.get("signing_city", "")

    # Format raw YYYY-MM-DD dates to "Month D, YYYY" / "YYYY年M月D日"
    period_start = _format_date(ui.get("Period_start", ""), language)
    period_end   = _format_date(ui.get("Period_end",   ""), language)
    report_date  = _format_date(tc.get("report_date",  ""), language)

    # Determine date placeholders based on report type
    if "TYPE1" in report_type:
        single_date = period_start  # Type I: "as of" date
    else:
        single_date = period_end    # Type II: period end date

    period_str = (
        f"{period_start} to {period_end}"
        if period_start and period_end
        else period_start or period_end
    )

    # Parse first SSO name and services from subservice_org (format: "Name | Services")
    sso_name = ""
    sso_services = ""
    if subservice_org:
        first_line = subservice_org.strip().splitlines()[0]
        if "|" in first_line:
            parts = first_line.split("|", 1)
            sso_name     = parts[0].strip()
            sso_services = parts[1].strip()
        else:
            sso_name = first_line.strip()

    subs = {
        # EN placeholders
        "[Service organization name]":          company_name,
        "[Service organization short name]":     co_short_name,
        "[Service organization\u2019s system]":  system_name,   # right single quote
        "[Service organization's system]":       system_name,   # straight apostrophe
        # Note: [or identification of the function performed by the System] is
        # handled by the [or ..] regex removal in fill_and_process_template.
        "[date] to [date]":                      period_str,
        "[date]":                                single_date,
        "[Date of the service auditor\u2019s report]": report_date,  # right quote
        "[Date of the service auditor's report]":      report_date,  # straight quote
        "[Date of report]":                      report_date,
        "[Ernst & Young Hua Ming LLP]":          EY_FIRM_NAME,
        # City: replace the whole "default_city[alternatives]" pattern
        "Shanghai[Beijing, Shenzhen]":           signing_city,
        "[Beijing, Shenzhen]":                   signing_city,
        "[Subservice organization name]":        sso_name,
        "[identify the function or service provided by the subservice organization]": sso_services,
        "[V]":                                   "V",
        # CN placeholders
        "\u300e\u670d\u52a1\u673a\u6784\u540d\u79f0\u300f": company_name,   # 【服务机构名称】
        "\u300e\u670d\u52a1\u673a\u6784\u7b80\u79f0\u300f": co_short_name,  # 【服务机构简称】
        "\u300e\u670d\u52a1\u673a\u6784\u4f53\u7cfb\u540d\u79f0\u300f": system_name,  # 【服务机构体系名称】
        "\u300e\u65e5\u671f\u300f": single_date,   # 【日期】
        "\u300e\u62a5\u544a\u65e5\u300f": report_date,  # 【报告日】
        # City CN: replace the default+alternatives pattern
        "\u4e2d\u56fd \u4e0a\u6d77\u3010\u6216\u4e2d\u56fd \u5317\u4eac\u6216\u4e2d\u56fd \u6df1\u5733\u3011": f"\u4e2d\u56fd {signing_city}",  # 中国 上海【或中国 北京或中国 深圳】→ 中国 {city}
        "\u3010\u6216\u5b89\u6c38\u534e\u660e\u4f1a\u8ba1\u5e08\u4e8b\u52a1\u6240\uff08\u7279\u6b8a\u666e\u901a\u5408\u4f19\uff09\u3011": "",  # 【或安永华明...】→ empty (keep branch)
    }
    return subs


def build_flags(tc):
    """Build the boolean deletion-flag dict from template_config."""
    return {
        "cuec_identified":            tc.get("cuec_identified", True),
        "sso_cc_identified":          tc.get("sso_cc_identified", True),
        "has_transaction_processing": tc.get("has_transaction_processing", True),
        "single_user_entity":         tc.get("single_user_entity", False),
        "has_ai_scope_exclusion":     tc.get("has_ai_scope_exclusion", False),
    }


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


def _add_heading(doc, text, level):
    """Add a Heading-N paragraph with level-specific formatting:
    H1 = bold only | H2 = bold + italic | H3 = italic + underline | H4+ = bold only"""
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run(text)
    _apply_fonts(run)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.size = Pt(11)
    if level == 1:
        run.font.bold = True
        run.font.italic = False
    elif level == 2:
        run.font.bold = True
        run.font.italic = True
    elif level == 3:
        run.font.bold = False
        run.font.italic = True
        run.font.underline = True
    else:
        run.font.bold = True
        run.font.italic = False
    blank = doc.add_paragraph()
    blank.paragraph_format.space_after  = Pt(0)
    blank.paragraph_format.space_before = Pt(0)



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
        # H1=bold | H2=bold+italic | H3=italic+underline | H4+=bold
        if line.startswith("#### "):
            _add_heading(doc, line[5:].strip(), 4)
            last_was_blank = True  # _add_heading appends a blank internally

        elif line.startswith("### "):
            _add_heading(doc, line[4:].strip(), 3)
            last_was_blank = True

        elif line.startswith("## "):
            _add_heading(doc, line[3:].strip(), 2)
            last_was_blank = True

        elif line.startswith("# "):
            _add_heading(doc, line[2:].strip(), 1)
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
        # Excludes both ASCII and Chinese terminal punctuation so that Chinese
        # sentences ending with 。？！，；：are not misidentified as headings.
        elif line.strip() and len(line.strip()) <= 130 and not line.rstrip().endswith(
            ('.', '?', '!', ',', ';', ':', '。', '？', '！', '，', '；', '：')
        ):
            _add_heading(doc, line.strip(), 2)
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

    # ── Complete report option ─────────────────────────────────────────────────
    generate_complete = st.checkbox(
        "Generate complete report (MA + AR + main sections)",
        value=True,
        help="When checked, the download will include Section I (Management Assertion) and Section II (Independent Auditor's Report) generated from EY templates, followed by the Dify-generated sections. Uncheck to generate Sections III–IV only (existing behaviour).",
    )

    if generate_complete:
        with st.expander("Complete Report Settings", expanded=True):
            cr1, cr2 = st.columns(2)

            # Read current Report Type from session_state key (set by the selectbox below).
            # Falls back to the previous run's saved value, then to default.
            _cur_rt = (
                st.session_state.get("form_report_type")
                or st.session_state.get("user_inputs", {}).get("Report_type", "SOC2 TYPE2")
            )
            _cur_sso = (
                st.session_state.get("form_scope_of_report")
                or st.session_state.get("user_inputs", {}).get("Scope_of_the_report", "None")
            )
            _cur_lang = (
                st.session_state.get("form_output_language")
                or st.session_state.get("user_inputs", {}).get("Output_language", "English")
            )

            with cr1:
                _std_options = get_standard_options(_cur_rt)
                standard = st.selectbox("Standard", _std_options, key="cr_standard")
                report_date  = st.text_input("Report Signing Date", placeholder="e.g. January 30, 2026", key="cr_report_date")
                signing_city = st.text_input("Signing City", placeholder="e.g. Shanghai", key="cr_signing_city")

            with cr2:
                cuec_choice = st.radio(
                    "Complementary User Entity Controls (CUEC)",
                    ["Identified", "Not Identified"],
                    index=0,
                    key="cr_cuec",
                )
                has_transaction_processing = st.checkbox(
                    "Includes transaction processing wording",
                    value=True,
                    key="cr_transaction",
                )
                single_user_entity = st.checkbox(
                    "Single user entity report",
                    value=False,
                    key="cr_single_user",
                )
                has_ai_scope_exclusion = st.checkbox(
                    "Subject matter includes AI technology (audit scope excludes AI-specific functions)",
                    value=False,
                    key="cr_ai_scope",
                    help="When checked, includes the paragraph disclosing that AI technology is used in the subject matter but is not within the audit scope. Leave unchecked if the subject matter does not involve AI technology.",
                )

            # SSO CC — only shown when SSO != None
            if _cur_sso != "None":
                sso_cc_choice = st.radio(
                    "SSO Complementary Controls",
                    ["Identified", "Not Identified"],
                    index=0,
                    key="cr_sso_cc",
                )
            else:
                sso_cc_choice = "Identified"

            # Template resolution preview (computed every render)
            st.markdown("---")
            _ar_wp, _ar_path = resolve_template(_cur_rt, standard, _cur_sso, _cur_lang, "AR")
            _ma_wp, _ma_path = resolve_template(_cur_rt, standard, _cur_sso, _cur_lang, "MA")

            if _ar_wp:
                if _ar_path:
                    st.info(f"AR template: WP No. {_ar_wp} \u2192 {os.path.basename(_ar_path)}")
                else:
                    st.warning(f"AR template: WP No. {_ar_wp} listed but file not found in AR_template/")
            else:
                st.warning("AR template: No matching template found for this combination.")

            if _ma_wp:
                if _ma_path:
                    st.info(f"MA template: WP No. {_ma_wp} \u2192 {os.path.basename(_ma_path)}")
                else:
                    st.warning(f"MA template: WP No. {_ma_wp} listed but file not found in MA_template/")
            else:
                st.warning("MA template: No matching template found for this combination.")

    else:
        standard                  = ""
        report_date               = ""
        signing_city              = ""
        cuec_choice               = "Identified"
        sso_cc_choice             = "Identified"
        has_transaction_processing = True
        single_user_entity         = False
        has_ai_scope_exclusion     = False
        _ar_path                   = None
        _ma_path                   = None

    st.markdown("---")
    st.subheader("Upload Control Matrix File(s)")
    st.caption(
        "1. 请上传Excel大表，其中必须包含Control Matrix sheet；"
        # "2. 若有，上传子服务机构及其服务内容Sheet；"
        # "3. 若有，上传CUEC内容清单Sheet；"
        # "3. 若有，上传子服务机构补充控制清单Sheet；"
        "2. 若有，上传名词解释Sheet；"
        "3. 若适用，上传控制目标Sheet"
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
            ["None", "All carve out", "Inclusive"], key="form_scope_of_report")
        industry = st.selectbox("Industry",
                                ["HR", "IaaS", "AI", "SaaS", "Others"])

    with req2:
        report_type     = st.selectbox("Report Type",
            ["SOC1 TYPE1", "SOC1 TYPE2", "SOC2 TYPE1", "SOC2 TYPE2"], key="form_report_type")
        output_language = st.selectbox("Output Language",
            ["English", "中文"], key="form_output_language")
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
        domain         = st.text_input("Control Domain",                  max_chars=256)
        co_website = st.text_input("Company Website", max_chars=256)
    with opt2:
        systems_function = st.text_input("Systems Function",
                        placeholder="e.g. workflow approval, code management, cloud resource management",
                        help="Optional. Describe the purpose of the internal supporting systems listed above.",
                        max_chars=256)
        system_extra = st.text_input("Internal Supporting Systems",
                         placeholder="e.g. Feishu Platform, Gitlab Platform, Alibaba Cloud Console",
                         help="Optional. List the internal systems used to support operations. If left blank, the workflow will auto-extract from the Control Matrix.",
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

        if generate_complete:
            if not report_date:
                errors.append("Report Signing Date is required when generating a complete report.")
            if not signing_city:
                errors.append("Signing City is required when generating a complete report.")

        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        # Re-resolve templates using the current form's report_type, sso and language
        if generate_complete:
            _ar_wp_final, _ar_path_final = resolve_template(report_type, standard, scope_of_report, output_language, "AR")
            _ma_wp_final, _ma_path_final = resolve_template(report_type, standard, scope_of_report, output_language, "MA")
            if not _ar_path_final:
                st.warning("AR template file not found for this combination — complete report will omit Section II.")
            if not _ma_path_final:
                st.warning("MA template file not found for this combination — complete report will omit Section I.")
        else:
            _ar_path_final = None
            _ma_path_final = None

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

        # Store template config for download step
        st.session_state["template_config"] = {
            "generate_complete":          generate_complete,
            "standard":                   standard,
            "report_date":                report_date,
            "signing_city":               signing_city,
            "cuec_identified":            cuec_choice == "Identified",
            "sso_cc_identified":          sso_cc_choice == "Identified",
            "has_transaction_processing": has_transaction_processing,
            "single_user_entity":         single_user_entity,
            "has_ai_scope_exclusion":     has_ai_scope_exclusion,
            "ar_template_path":           _ar_path_final,
            "ma_template_path":           _ma_path_final,
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
    ui  = st.session_state.get("user_inputs", {})
    tc  = st.session_state.get("template_config", {})

    with st.expander("📖 Preview Report (Dify sections)", expanded=True):
        st.markdown(result_text)

    # Always generate the Dify sections docx
    dify_bytes = markdown_to_docx(result_text, ui.get("Output_language", "English"))

    if (tc.get("generate_complete")
            and tc.get("ar_template_path")
            and tc.get("ma_template_path")):
        subs  = build_substitutions(ui, tc)
        flags = build_flags(tc)

        _lang = ui.get("Output_language", "English")
        try:
            with st.spinner("Generating MA section (Section I)…"):
                ma_bytes = fill_and_process_template(tc["ma_template_path"], subs, flags, _lang)
            with st.spinner("Generating AR section (Section II)…"):
                ar_bytes = fill_and_process_template(tc["ar_template_path"], subs, flags, _lang)
            with st.spinner("Merging all sections…"):
                final_bytes = merge_docx_sections(ma_bytes, ar_bytes, dify_bytes)
            filename = (
                f"{ui.get('Co_short_name', 'Report')}_"
                f"{ui.get('Report_type', '').replace(' ', '_')}_Complete_Report.docx"
            )
        except Exception as exc:
            st.error(f"Failed to generate complete report: {exc}\n\nFalling back to Dify sections only.")
            final_bytes = dify_bytes
            filename = (
                f"{ui.get('Co_short_name', 'Report')}_"
                f"{ui.get('Report_type', '').replace(' ', '_')}_Report.docx"
            )
    else:
        final_bytes = dify_bytes
        filename = (
            f"{ui.get('Co_short_name', 'Report')}_"
            f"{ui.get('Report_type', '').replace(' ', '_')}_Report.docx"
        )

    st.download_button(
        label="⬇ Download Report (.docx)",
        data=final_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
