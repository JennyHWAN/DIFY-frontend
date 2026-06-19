---
name: soc-report-generator
description: Generate SOC report's section I to section III.For section I and section II, 
  python codes are used for authoring and editing the EY Word (.docx) templates in 
  DIFY-frontend/MA_template/ and AR_template/, based on template_index.xlsx. 
  Covers the placeholder syntax, the [or ..] / 【】 / （注：…）conditional markers, 
  Word-comment annotation markers (CUEC, SSO-CC, single-user-entity, AI-scope, Other-Information), 
  and the WP-number lookup. Also to format the DIFY generated seciton III. 
  Use when adding/changing a template, a placeholder, or a conditional rule
  that app.py's fill_and_process_template must honor.
---

# EY Template Authoring

The MA (Management Assertion) and AR (Auditor's Report) sections of a SOC report
are pre-authored EY `.docx` files filled at runtime by
`fill_and_process_template()` in `DIFY-frontend/app.py`. **The template is the
contract** — what you type into the `.docx` (placeholders, bracket markers, Word
comments) is exactly what the code keys off. Get the markers wrong and the fill
silently does nothing or deletes the wrong thing.

## The three authoring mechanisms

1. **Placeholders** — literal bracketed tokens the code replaces with company /
   date / SSO values. Must match `build_substitutions()` exactly, character for
   character (including curly vs straight quotes and fullwidth 【】 in CN).
2. **Inline `[or ..]` / `（注：…）` markers** — alternative phrasing the code
   strips unconditionally.
3. **Word comments** — anchored to a paragraph or sentence span; their *text*
   (keyword-matched) tells the code to conditionally keep or delete that span
   based on the user's complete-report toggles (`build_flags()`).

## Hard rules when authoring

- A new placeholder does nothing until it's also added to `build_substitutions()`.
- Compound placeholders (`[date] to [date]`, `【日期】至【日期】`) must be authored
  as-is — they're consumed before the single `[date]` / `【日期】`.
- Comment markers are **keyword-matched, case-insensitive** against fixed keyword
  lists (`_CUEC_KW`, `_SSO_CC_KW`, etc. in app.py ~line 271). Use the established
  Chinese/English phrasings or the comment is ignored.
- Partial-sentence deletion needs the commented span to be ≥30 chars; short
  transaction paragraphs delete only when ≤200 chars. Don't author marginal spans.
- `template_index.xlsx` (`MA`/`AR` sheets) maps selections → a WP number;
  `resolve_template` matches the file whose name **starts with** that WP number.

## References — load when relevant

- `references/placeholders.md` — full placeholder token → value table (EN + CN).
- `references/conditional-markers.md` — comment keywords, flags, and the
  `[or ..]` / `（注：…）` inline rules.
- `references/adding-a-template.md` — checklist for adding a new template file.

## Assets — golden examples

Use these in `assets/` as fixtures when changing the fill pipeline — diff new
output against them to catch regressions:

- `template_before.png` / `output_after.png` — before/after of a template fill,
  showing how placeholders and conditional markers resolve.
