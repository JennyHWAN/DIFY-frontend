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
    On error returns (None, error_message_string) so callers can surface the problem.
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
        # Do NOT use read_only=True — it can silently fail to iterate rows in
        # some environments.
        wb = openpyxl.load_workbook(TEMPLATE_INDEX, data_only=True)
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        return (None, f"Cannot load template index: {e}")

    if not rows:
        return (None, "Template index sheet is empty")

    header = list(rows[0])
    try:
        col_cat  = header.index("Category")
        col_type = header.index("Type")
        col_std  = header.index("Standards")
        col_sso  = header.index("Sub-service Organization (SSO)")
        col_lang = header.index("Language")
        col_wp   = next(
            i for i, h in enumerate(header)
            if h and str(h).strip().upper().startswith("WP")
        )
    except (ValueError, StopIteration) as e:
        return (None, f"Template index column not found: {e}")

    for row in rows[1:]:
        if (str(row[col_cat]  or "").strip() == category  and
                str(row[col_type] or "").strip() == typ       and
                str(row[col_std]  or "").strip() == std_mapped and
                str(row[col_sso]  or "").strip() == sso_mapped and
                str(row[col_lang] or "").strip() == lang_mapped):
            wp_val = row[col_wp]
            if wp_val is None:
                # This combination has no template — keep iterating in case
                # another row matches (shouldn't happen, but safe).
                continue
            wp_no = str(wp_val).strip()
            try:
                for fname in sorted(os.listdir(template_dir)):
                    if fname.endswith(".docx") and fname.startswith(wp_no + " "):
                        return (wp_no, os.path.join(template_dir, fname))
            except OSError as e:
                return (wp_no, f"Cannot list template directory: {e}")
            return (wp_no, None)

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

        # Single-UE: if comment says "delete" → delete whole para when flag set;
        # otherwise → bracket-replacement path
        if any(_kw_in(text, kw.lower()) for kw in _SINGLE_KW):
            is_delete_cmd = "删除" in text or "delete" in text.lower()
            if is_delete_cmd and flags.get("single_user_entity", False):
                del_indices |= indices
            else:
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
                # Only delete SHORT standalone paragraphs (≤200 chars after
                # stripping whitespace). Long paragraphs containing the
                # transaction phrase as just one embedded clause are handled
                # by inline text substitution in build_substitutions instead.
                body_list = list(body)
                for idx in indices:
                    if idx < len(body_list) and body_list[idx].tag == p_tag:
                        para_txt = "".join(
                            t.text or ""
                            for t in body_list[idx].findall(f".//{{{ns_w}}}t")
                        ).strip()
                        if len(para_txt) <= 200:
                            del_indices.add(idx)
                # (do not set should_delete — individual indices already handled)

        if should_delete:
            del_indices |= indices

    return del_indices, single_ue_indices


def _reject_format_changes(xml_str):
    """
    Reject tracked formatting changes by restoring the OLD formatting:
    - w:pPrChange: replace the parent w:pPr with the OLD w:pPr inside pPrChange
    - w:rPrChange: replace the parent w:rPr with the OLD w:rPr inside rPrChange

    In EY templates, pPrChange/rPrChange record authoring edits that should
    NOT be accepted into the final output — rejecting restores the intended
    original formatting (e.g., list style a. instead of Wingdings bullet l,
    non-bold run instead of bold run).
    """
    def _reject_one(xml_s, change_tag, parent_tag):
        change_re = re.compile(
            rf"<{re.escape(change_tag)}\b[^>]*>(.*?)</{re.escape(change_tag)}>",
            re.DOTALL,
        )
        result = xml_s
        for m in reversed(list(change_re.finditer(result))):
            chg_start, chg_end = m.start(), m.end()

            # Extract old parent element from inside the change block
            old_m = re.search(
                rf"<{re.escape(parent_tag)}\b[^>]*>.*?</{re.escape(parent_tag)}>",
                m.group(1), re.DOTALL,
            )

            # The outer closing tag immediately follows the change block
            rest = result[chg_end:]
            close_m = re.match(rf"\s*</{re.escape(parent_tag)}>", rest)
            if not close_m:
                # Unexpected structure — just drop the change block
                result = result[:chg_start] + result[chg_end:]
                continue
            outer_end = chg_end + close_m.end()

            # The LAST opening parent tag before the change block
            before = result[:chg_start]
            opens = list(re.finditer(rf"<{re.escape(parent_tag)}\b[^>]*>", before))
            if not opens:
                result = result[:chg_start] + result[chg_end:]
                continue
            outer_start = opens[-1].start()

            replacement = old_m.group(0) if old_m else ""
            result = result[:outer_start] + replacement + result[outer_end:]

        return result

    xml_str = _reject_one(xml_str, "w:pPrChange", "w:pPr")
    xml_str = _reject_one(xml_str, "w:rPrChange", "w:rPr")
    return xml_str


def _apply_xml_cleaning(xml_str):
    """
    Apply all tracked-change cleaning to a raw XML string.

    Strategy (matches EY template authoring conventions):
      - w:ins  → ACCEPT:  unwrap, keep content
      - w:del  → ACCEPT:  remove deleted content entirely
      - pPrChange / rPrChange → REJECT:  restore OLD formatting
      - Table / section format-change markers → removed
      - Comment markers → removed
    """
    xml_str = xml_str.lstrip("\ufeff")

    # 1. Accept deletions: remove <w:del>…</w:del> blocks entirely
    xml_str = re.sub(
        r"<w:del\b[^>]*>.*?</w:del>", "",
        xml_str, flags=re.DOTALL,
    )

    # 2. Remove self-closing tracked-change markers
    xml_str = re.sub(r"<w:del\b[^>]*/>",  "", xml_str)
    xml_str = re.sub(r"<w:ins\b[^>]*/>",  "", xml_str)

    # 3. Accept insertions: unwrap <w:ins>…</w:ins>
    xml_str = re.sub(
        r"<w:ins\b[^>]*(?<!/)>(.*?)</w:ins>", r"\1",
        xml_str, flags=re.DOTALL,
    )

    # 5. Reject format changes: restore OLD pPr / rPr
    xml_str = _reject_format_changes(xml_str)

    # 6. Remove table / section format-change tracking elements
    for _ftag in ("w:tblPrChange", "w:trPrChange", "w:tcPrChange", "w:sectPrChange",
                  "w:numChange"):
        xml_str = re.sub(
            rf"<{_ftag}\b[^>]*>.*?</{_ftag}>", "",
            xml_str, flags=re.DOTALL,
        )

    # 7. Remove comment range markers
    xml_str = re.sub(r"<w:commentRangeStart[^>]*/>", "", xml_str)
    xml_str = re.sub(r"<w:commentRangeEnd[^>]*/>",   "", xml_str)

    # 8. Remove runs whose only non-rPr child is a comment reference
    xml_str = re.sub(
        r"<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:commentReference[^>]*/>\s*</w:r>",
        "", xml_str, flags=re.DOTALL,
    )

    return xml_str


def _clean_docx_bytes(docx_bytes):
    """
    Accept/reject tracked changes and clear comments in a docx file.

    Uses _apply_xml_cleaning() on word/document.xml and related XML parts
    (headers, footers, footnotes, endnotes).  word/comments.xml is replaced
    with an empty comments document so Word opens without comment balloons.
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
                data  = zin.read(info.filename)
                fname = info.filename
                if fname == "word/comments.xml":
                    data = _EMPTY_COMMENTS
                elif (
                    fname == "word/document.xml"
                    or re.match(r"word/(header|footer)\d+\.xml$", fname)
                    or fname in ("word/footnotes.xml", "word/endnotes.xml")
                ):
                    data = _apply_xml_cleaning(data.decode("utf-8")).encode("utf-8")
                zout.writestr(info, data)

    buf_out.seek(0)
    return buf_out.read()


def _smart_replace_in_para(para, old, new):
    """Replace one occurrence of *old* with *new* in the paragraph, modifying
    only the minimal run span that contains the match.  Per-run formatting
    (italic, bold, size, etc.) on runs outside that span is fully preserved.
    Returns True if a replacement was made, False otherwise."""
    if not old or not para.runs:
        return False
    texts = [r.text for r in para.runs]
    full  = "".join(texts)
    if old not in full:
        return False
    pos     = full.index(old)
    end_pos = pos + len(old)
    # Cumulative start offset of each run within full
    cum = 0
    starts = []
    for t in texts:
        starts.append(cum)
        cum += len(t)
    # First run whose text overlaps the match
    fi = next((i for i in range(len(texts)) if starts[i] + len(texts[i]) > pos), None)
    if fi is None:
        return False
    # Last run whose text overlaps the match
    li = next((i for i in range(len(texts) - 1, -1, -1) if starts[i] < end_pos), None)
    if li is None:
        return False
    prefix = texts[fi][: pos - starts[fi]]
    suffix = texts[li][end_pos - starts[li] :]
    para.runs[fi].text = prefix + new + suffix
    for k in range(fi + 1, li + 1):
        para.runs[k].text = ""
    return True


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
        if not para.runs:
            return
        # Phase 1: per-run substitution — preserves individual run formatting
        # (bold, italic, etc.) when the placeholder is entirely within one run.
        for run in para.runs:
            for placeholder, value in subs.items():
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
        # Phase 2: smart span replacement for cross-run placeholders.
        # Loop each placeholder until no occurrences remain — a single paragraph
        # may contain the same placeholder more than once (e.g. [Service
        # organization short name] appears multiple times in the MA description).
        for placeholder, value in subs.items():
            while _smart_replace_in_para(para, placeholder, value):
                pass
        # Phase 3: normalize consecutive spaces that may arise from empty
        # substitutions, both within a run and across run boundaries.
        prev_ended_space = False
        for run in para.runs:
            if run.text:
                run.text = re.sub(r"  +", " ", run.text)
                if prev_ended_space and run.text.startswith(" "):
                    run.text = run.text.lstrip(" ")
                prev_ended_space = run.text.endswith(" ")
            # empty run: prev_ended_space unchanged (gap is invisible)

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
        # Process one bracket match per loop iteration using smart replacement so
        # that per-run italic/bold formatting outside the matched span is preserved.
        changed = True
        while changed:
            changed = False
            full_text = "".join(r.text for r in para.runs)
            if not full_text or "[" not in full_text:
                break

            # a) [or …] alternative phrases — always remove (with any leading space)
            m = re.search(r" ?\[or [^\]]+\]", full_text, flags=re.IGNORECASE)
            if m:
                _smart_replace_in_para(para, m.group(0), "")
                changed = True
                continue

            # b) single-user-entity annotated paragraphs
            if is_single_ue:
                if _single_ue_flag:
                    # delete bracketed content entirely
                    m = re.search(r"\[[^\]]*\]", full_text)
                    if m:
                        _smart_replace_in_para(para, m.group(0), "")
                        changed = True
                        continue
                else:
                    # strip brackets, keep content
                    m = re.search(r"\[([^\]]*)\]", full_text)
                    if m:
                        _smart_replace_in_para(para, m.group(0), m.group(1))
                        changed = True
                        continue

            # c) any remaining [..] — strip brackets, keep content
            m = re.search(r"\[([^\]]*)\]", full_text)
            if m:
                _smart_replace_in_para(para, m.group(0), m.group(1))
                changed = True
                continue

        # Normalize spaces introduced by empty removals
        prev_ended_space = False
        for run in para.runs:
            if run.text:
                run.text = re.sub(r"  +", " ", run.text)
                if prev_ended_space and run.text.startswith(" "):
                    run.text = run.text.lstrip(" ")
                prev_ended_space = run.text.endswith(" ")

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

    # Paragraphs whose text matches these patterns keep their bold intact.
    # Checked against the paragraph's full text AFTER substitution.
    _BOLD_KEEP_PATTERNS = [
        "Independent Service Auditor",   # AR section title
        "Management Assertion",          # MA section title (after co-name sub)
        "Ernst & Young",                 # AR signature block (firm name)
        "Hua Ming",                      # AR signature block (firm name variant)
    ]
    # Date paragraph in MA: formatted date string is the entire paragraph text
    _DATE_RE = re.compile(
        r"^(?:[A-Z][a-z]+ \d{1,2}, \d{4}|\d{4}年\d{1,2}月\d{1,2}日)$"
    )
    # Company name used to detect the MA signature line (a short standalone para)
    _company_name_bold = subs.get("[Service organization name]", "")

    def _para_keep_bold(para):
        full = "".join(r.text for r in para.runs).strip()
        # Apply pattern matching only to short paragraphs (section titles,
        # signature-block lines). Without this guard the 10 000-char "We have
        # examined…" AR paragraph — which contains "Management Assertion" in
        # running body text — would be incorrectly made bold in its entirety.
        if len(full) <= 120 and any(pat in full for pat in _BOLD_KEEP_PATTERNS):
            return True
        if _DATE_RE.match(full):
            return True
        # MA signature line: standalone paragraph whose text is exactly the
        # service-organization name (e.g. "ABC Fintech Co., Ltd.")
        if _company_name_bold and full == _company_name_bold:
            return True
        # Short city / country line in the AR signature block
        # e.g. "Shanghai, China" or "中国 上海" (≤50 chars to avoid body text)
        if len(full) <= 50 and ("China" in full or "中国" in full):
            return True
        return False

    def _std_run(run, keep_bold=False):
        rPr = run._r.get_or_add_rPr()
        if not keep_bold:
            run.bold = False            # strip bold; italic is left untouched
            # Also explicitly disable complex-script bold (bCs) so that CJK
            # characters do not inherit bold from a heading paragraph style
            # after the document sections are merged.
            bCs = rPr.find(qn("w:bCs"))
            if bCs is None:
                bCs = OxmlElement("w:bCs")
                rPr.append(bCs)
            bCs.set(qn("w:val"), "0")
        else:
            run.bold = True             # ensure bold is explicitly set
        run.font.size = Pt(11)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"),    "Times New Roman")
        rFonts.set(qn("w:hAnsi"),    "Times New Roman")
        rFonts.set(qn("w:eastAsia"), cjk_font)
        rFonts.set(qn("w:cs"),       "Times New Roman")
        # Explicitly disable underline (overrides any inherited style underline)
        u_el = rPr.find(qn("w:u"))
        if u_el is None:
            u_el = OxmlElement("w:u")
            rPr.append(u_el)
        u_el.set(qn("w:val"), "none")

    def _clear_para_mark_bold(para):
        """Also strip bold from the paragraph-mark rPr (pPr/rPr)."""
        _pPr = para._p.find(qn("w:pPr"))
        if _pPr is None:
            return
        _pRPr = _pPr.find(qn("w:rPr"))
        if _pRPr is None:
            return
        for _bname in ("w:b", "w:bCs"):
            _bel = _pRPr.find(qn(_bname))
            if _bel is None:
                _bel = OxmlElement(_bname)
                _pRPr.append(_bel)
            _bel.set(qn("w:val"), "0")

    for para in doc.paragraphs:
        kb = _para_keep_bold(para)
        for run in para.runs:
            _std_run(run, keep_bold=kb)
        if not kb:
            _clear_para_mark_bold(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    kb = _para_keep_bold(para)
                    for run in para.runs:
                        _std_run(run, keep_bold=kb)
                    if not kb:
                        _clear_para_mark_bold(para)

    # ── Step 7: strip spaces before punctuation ────────────────────────
    _PUNCT = '.,;:!?，。；：！？'

    def _strip_spaces_before_punct(para):
        for run in para.runs:
            if run.text:
                run.text = re.sub(r' +([' + re.escape(_PUNCT) + r'])', r'\1', run.text)
        active = [r for r in para.runs if r.text]
        for i in range(len(active) - 1):
            if active[i].text.endswith(' ') and active[i + 1].text[0] in _PUNCT:
                active[i].text = active[i].text.rstrip(' ')

    for para in doc.paragraphs:
        _strip_spaces_before_punct(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _strip_spaces_before_punct(para)

    buf = io.BytesIO()
    doc.save(buf)
    saved_bytes = buf.getvalue()

    # python-docx may silently drop <w:lvlOverride>/<w:startOverride> elements
    # when serialising the NumberingPart, which breaks lowerLetter list counters
    # (e.g. "a." / "b." become "l" / "l").  Re-inject the original numbering.xml
    # from the template to guarantee the counter-restart overrides are preserved.
    try:
        with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as _z_orig:
            if "word/numbering.xml" in _z_orig.namelist():
                _orig_num = _z_orig.read("word/numbering.xml")
                _buf_in  = io.BytesIO(saved_bytes)
                _buf_out = io.BytesIO()
                with zipfile.ZipFile(_buf_in, "r") as _z_in:
                    with zipfile.ZipFile(_buf_out, "w", zipfile.ZIP_DEFLATED) as _z_out:
                        for _info in _z_in.infolist():
                            _data = _z_in.read(_info.filename)
                            if _info.filename == "word/numbering.xml":
                                _data = _orig_num
                            _z_out.writestr(_info, _data)
                _buf_out.seek(0)
                return _buf_out.read()
    except Exception:
        pass
    return saved_bytes


def _remap_extra_numbering(base_bytes, extra_bytes):
    """
    Remap abstractNumId and numId values in extra_bytes so they do not clash
    with numbering definitions already present in base_bytes.

    Both documents' numbering.xml are inspected to find the highest IDs in
    the base; every ID in the extra is offset by that amount so no two
    definitions share an ID after the merge.

    Returns modified extra_bytes, or the original if either document has no
    numbering.xml or the extra has no numbered lists.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(base_bytes), "r") as z:
            if "word/numbering.xml" not in z.namelist():
                return extra_bytes
            base_num_xml = z.read("word/numbering.xml").decode("utf-8")
    except Exception:
        return extra_bytes

    try:
        with zipfile.ZipFile(io.BytesIO(extra_bytes), "r") as z:
            if "word/numbering.xml" not in z.namelist():
                return extra_bytes
            extra_num_xml = z.read("word/numbering.xml").decode("utf-8")
            extra_doc_xml = z.read("word/document.xml").decode("utf-8")
    except Exception:
        return extra_bytes

    # Determine maximum IDs in base
    base_abs  = [int(x) for x in re.findall(r'w:abstractNumId="(\d+)"', base_num_xml)]
    base_nums = [int(x) for x in re.findall(r'<w:num\b[^>]*w:numId="(\d+)"', base_num_xml)]
    base_max_abs = max(base_abs)  if base_abs  else -1
    base_max_num = max(base_nums) if base_nums else  0

    # Collect IDs present in extra (process largest first to avoid partial hits)
    extra_abs  = sorted(
        set(int(x) for x in re.findall(r'<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"', extra_num_xml)),
        reverse=True,
    )
    extra_nums = sorted(
        set(int(x) for x in re.findall(r'<w:num\b[^>]*w:numId="(\d+)"', extra_num_xml)),
        reverse=True,
    )

    if not extra_nums:
        return extra_bytes  # nothing numbered in extra

    abs_offset = base_max_abs + 1
    num_offset = base_max_num

    new_num_xml = extra_num_xml
    new_doc_xml = extra_doc_xml

    # Remap abstractNumId in numbering.xml (definition + cross-references)
    for old in extra_abs:
        new = old + abs_offset
        # <w:abstractNum w:abstractNumId="N"> — definition attribute
        new_num_xml = new_num_xml.replace(
            f'w:abstractNumId="{old}"', f'w:abstractNumId="{new}"'
        )
        # <w:abstractNumId w:val="N"/> — reference inside <w:num>
        new_num_xml = new_num_xml.replace(
            f'<w:abstractNumId w:val="{old}"', f'<w:abstractNumId w:val="{new}"'
        )

    # Remap numId in numbering.xml (definition) and in document.xml (references)
    for old in extra_nums:
        new = old + num_offset
        # <w:num w:numId="N"> — definition attribute
        new_num_xml = new_num_xml.replace(f'w:numId="{old}"', f'w:numId="{new}"')
        # <w:numId w:val="N"/> — paragraph numPr reference
        new_doc_xml = re.sub(
            rf'(<w:numId\s+w:val="){old}(")',
            rf'\g<1>{new}\g<2>',
            new_doc_xml,
        )

    # Write modified extra docx
    buf_in  = io.BytesIO(extra_bytes)
    buf_out = io.BytesIO()
    with zipfile.ZipFile(buf_in, "r") as zin:
        with zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/numbering.xml":
                    data = new_num_xml.encode("utf-8")
                elif info.filename == "word/document.xml":
                    data = new_doc_xml.encode("utf-8")
                zout.writestr(info, data)
    buf_out.seek(0)
    return buf_out.read()


def _inject_numbering(merged_bytes, extra_bytes):
    """
    Append all abstractNum and num definitions from extra_bytes into the
    word/numbering.xml of merged_bytes.

    This is called after the document bodies have been merged so that the
    remapped numId values in the extra's paragraphs resolve to actual
    definitions in the merged file.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(extra_bytes), "r") as z:
            if "word/numbering.xml" not in z.namelist():
                return merged_bytes
            extra_num_xml = z.read("word/numbering.xml").decode("utf-8")
    except Exception:
        return merged_bytes

    abs_blocks = re.findall(r"<w:abstractNum\b.*?</w:abstractNum>", extra_num_xml, re.DOTALL)
    num_blocks  = re.findall(r"<w:num\b.*?</w:num>",                extra_num_xml, re.DOTALL)

    if not abs_blocks and not num_blocks:
        return merged_bytes

    inject = "\n".join(abs_blocks) + "\n" + "\n".join(num_blocks)

    buf_in  = io.BytesIO(merged_bytes)
    buf_out = io.BytesIO()
    with zipfile.ZipFile(buf_in, "r") as zin:
        with zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/numbering.xml":
                    num_xml = data.decode("utf-8")
                    num_xml = num_xml.replace("</w:numbering>", inject + "\n</w:numbering>")
                    data = num_xml.encode("utf-8")
                zout.writestr(info, data)
    buf_out.seek(0)
    return buf_out.read()


def merge_docx_sections(*docs_bytes):
    """
    Merge multiple docx byte strings into one document with page breaks between
    sections.

    Numbering IDs (abstractNumId / numId) in each extra document are remapped
    to avoid conflicts with the base document's numbering definitions, then the
    extra definitions are injected into the merged file's numbering.xml.  This
    preserves the original list styles (bullets stay bullets, alpha lists stay
    alpha lists) across section boundaries.
    """
    base_bytes = docs_bytes[0]

    # Remap numbering in each extra so its IDs don't clash with the base
    extras_remapped = [
        _remap_extra_numbering(base_bytes, eb) for eb in docs_bytes[1:]
    ]

    # Merge document bodies with python-docx
    base = Document(io.BytesIO(base_bytes))
    body = base.element.body
    last_sectPr = body.find(qn("w:sectPr"))

    for extra_bytes in extras_remapped:
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
    merged_bytes = buf.getvalue()

    # Inject remapped numbering definitions from each extra into the merged doc
    for extra_bytes in extras_remapped:
        merged_bytes = _inject_numbering(merged_bytes, extra_bytes)

    return merged_bytes


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

    # Addressee line: replace the combined "Management of/Board of Directors of"
    # placeholder before the generic [Service organization name] sub runs.
    addressee = tc.get("addressee_choice", "Management")
    if addressee == "Board of Directors":
        addr_label = "Board of Directors"
    else:
        addr_label = "Management"
    subs = {
        # Addressee line — must come BEFORE the generic [Service organization name] sub
        "To the Management of/Board of Directors of [Service organization name]":
            f"To the {addr_label} of {company_name}",
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
    # When transaction-processing wording is excluded, remove the inline
    # transaction phrases and substitute "[or identification of the function
    # performed by the System]" with the user-supplied system function description.
    if not tc.get("has_transaction_processing", True):
        sys_fn = ui.get("Systems_function", "")
        # EN variants — remove transaction-processing phrases
        subs["for processing user entities\u2019 transactions"] = ""   # right-quote
        subs["for processing user entities' transactions"] = ""        # straight-quote
        subs["for processing their transactions"] = ""
        # "in processing or reporting transactions" → "in" so that the following
        # [or identification...] substitution produces "in <system function>"
        # rather than "in processing or reporting <system function>" (issue 8)
        subs["in processing or reporting transactions"] = "in"
        subs["[or identification of the function performed by the System]"] = sys_fn
        subs["[or identification of the function performed by the system]"] = sys_fn
        # Remove the "auditors" clause from the "intended solely for…" paragraph.
        # That clause only applies when the system processes user-entity transactions.
        # Both right-quote (U+2019) and straight-apostrophe variants are covered.
        subs[
            ", and their auditors who audit and report on such user entities\u2019 "
            "financial statements or internal control over financial reporting"
        ] = ""
        subs[
            ", and their auditors who audit and report on such user entities' "
            "financial statements or internal control over financial reporting"
        ] = ""
    return subs


def build_flags(tc):
    """Build the boolean deletion-flag dict from template_config."""
    return {
        "cuec_identified":            tc.get("cuec_identified", True),
        "sso_cc_identified":          tc.get("sso_cc_identified", True),
        "has_transaction_processing": tc.get("has_transaction_processing", True),
        "single_user_entity":         tc.get("single_user_entity", False),
        "has_ai_scope_exclusion":     tc.get("has_ai_scope_exclusion", False),
        "addressee_choice":           tc.get("addressee_choice", "Management"),
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
                addressee_choice = st.radio(
                    "AR Addressee",
                    ["Management", "Board of Directors"],
                    index=0,
                    key="cr_addressee",
                    help="Controls whether the AR opens 'To the Management of' or 'To the Board of Directors of' followed by the service organization name.",
                )
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
                if not has_transaction_processing:
                    st.caption(
                        "⚠️ When unchecked, the 'Systems Function' field below "
                        "(Required Fields section) is used to describe the system's "
                        "function in place of transaction-processing wording."
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

            def _show_template_status(label, wp, path, dir_name):
                if wp is None and isinstance(path, str):
                    # path carries the error message when wp is None
                    st.error(f"{label}: {path}")
                elif wp is None:
                    st.warning(f"{label}: No matching template found for this combination.")
                elif path and os.path.isfile(path):
                    st.info(f"{label}: WP No. {wp} \u2192 {os.path.basename(path)}")
                elif isinstance(path, str) and path.startswith("Cannot"):
                    st.error(f"{label}: WP No. {wp} — {path}")
                else:
                    st.warning(f"{label}: WP No. {wp} listed but .docx not found in {dir_name}/")

            _show_template_status("AR template", _ar_wp, _ar_path, "AR_template")
            _show_template_status("MA template", _ma_wp, _ma_path, "MA_template")

            # ── Test mode: generate MA+AR without running the Dify workflow ──────
            _ar_ok = _ar_path and os.path.isfile(_ar_path)
            _ma_ok = _ma_path and os.path.isfile(_ma_path)
            if _ar_ok and _ma_ok:
                st.markdown("---")
                st.caption(
                    "Use the button below to test Section I + II template generation "
                    "without running the full Dify workflow. "
                    "Fill in the company fields below first, then click the button."
                )
                with st.expander("Test Template Generation (no Dify needed)", expanded=False):
                    t1, t2 = st.columns(2)
                    with t1:
                        _t_company   = st.text_input("Company Name",       key="test_company",   placeholder="e.g. ABC Fintech Co., Ltd.")
                        _t_short     = st.text_input("Short Name",          key="test_short",     placeholder="e.g. ABC")
                        _t_system    = st.text_input("System Name",         key="test_system",    placeholder="e.g. Payment Processing System")
                        _t_svc_desc  = st.text_input("Service Description", key="test_svc_desc",  placeholder="e.g. payment processing services")
                    with t2:
                        _t_period_s  = st.text_input("Period Start (YYYY-MM-DD)", key="test_period_s", placeholder="e.g. 2024-01-01")
                        _t_period_e  = st.text_input("Period End   (YYYY-MM-DD)", key="test_period_e", placeholder="e.g. 2024-12-31")
                        _t_sso_name  = st.text_input("SSO Name (if any)",    key="test_sso_name",  placeholder="leave blank if none")
                        _t_sys_fn    = st.text_input("System Function (if no transaction processing)", key="test_sys_fn", placeholder="e.g. providing cloud-based payment infrastructure")
                    if st.button("Generate test MA + AR sections", key="test_template_btn"):
                        _test_ui = {
                            "Company_name":          _t_company or "Test Organization",
                            "Co_short_name":         _t_short   or "TestOrg",
                            "System_or_service_name": _t_system or "Test System",
                            "Service_description":   _t_svc_desc or "test services",
                            "Period_start":          _t_period_s or "2024-01-01",
                            "Period_end":            _t_period_e or "2024-12-31",
                            "Report_type":           _cur_rt,
                            "Output_language":       _cur_lang,
                            "Subservice_org":        _t_sso_name or "None",
                            "Systems_function":      _t_sys_fn or "",
                        }
                        _test_tc = {
                            "report_date":                report_date  or "January 1, 2025",
                            "signing_city":               signing_city or "Shanghai",
                            "cuec_identified":            cuec_choice == "Identified",
                            "sso_cc_identified":          sso_cc_choice == "Identified",
                            "has_transaction_processing": has_transaction_processing,
                            "single_user_entity":         single_user_entity,
                            "has_ai_scope_exclusion":     has_ai_scope_exclusion,
                            "addressee_choice":           addressee_choice,
                        }
                        try:
                            with st.spinner("Generating test MA + AR sections…"):
                                _t_subs  = build_substitutions(_test_ui, _test_tc)
                                _t_flags = build_flags(_test_tc)
                                _t_ma    = fill_and_process_template(_ma_path, _t_subs, _t_flags, _cur_lang)
                                _t_ar    = fill_and_process_template(_ar_path, _t_subs, _t_flags, _cur_lang)
                                _t_merged = merge_docx_sections(_t_ma, _t_ar)
                            _t_fname = (
                                f"{(_t_short or 'Test')}_{_cur_rt.replace(' ','_')}"
                                f"_MA_AR_test.docx"
                            )
                            st.download_button(
                                label="⬇ Download test MA + AR (.docx)",
                                data=_t_merged,
                                file_name=_t_fname,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )
                        except Exception as _exc:
                            st.error(f"Test generation failed: {_exc}")

    else:
        standard                  = ""
        report_date               = ""
        signing_city              = ""
        cuec_choice               = "Identified"
        sso_cc_choice             = "Identified"
        has_transaction_processing = True
        single_user_entity         = False
        has_ai_scope_exclusion     = False
        addressee_choice           = "Management"
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
            if not has_transaction_processing and not systems_function:
                errors.append(
                    "Systems Function is required when 'Includes transaction processing wording' "
                    "is unchecked — please fill in the Systems Function field."
                )

        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        # Re-resolve templates using the current form's report_type, sso and language
        if generate_complete:
            _ar_wp_final, _ar_path_final = resolve_template(report_type, standard, scope_of_report, output_language, "AR")
            _ma_wp_final, _ma_path_final = resolve_template(report_type, standard, scope_of_report, output_language, "MA")
            # Treat error strings (non-file paths) as missing
            if _ar_path_final and not os.path.isfile(_ar_path_final):
                st.warning(f"AR template issue: {_ar_path_final} — complete report will omit Section II.")
                _ar_path_final = None
            elif not _ar_path_final:
                st.warning("AR template file not found for this combination — complete report will omit Section II.")
            if _ma_path_final and not os.path.isfile(_ma_path_final):
                st.warning(f"MA template issue: {_ma_path_final} — complete report will omit Section I.")
                _ma_path_final = None
            elif not _ma_path_final:
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
            "addressee_choice":           addressee_choice,
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
