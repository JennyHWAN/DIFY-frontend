# Conditional Markers (reference)

Two ways to make template content conditional. Both are read in
`_build_annotation_maps()` / `fill_and_process_template()` in `app.py`.

## 1. Inline markers — always stripped (no flag)

- `[or ...]` / `【或...】` — alternative phrasing. **Always removed** (the
  preceding default phrase stays). Used e.g. for alternative firm names, the
  `[or identification of the function performed by the System]` clause.
- Remaining `[...]` brackets not in the placeholder table: brackets removed,
  inner text kept.

## 2. Word-comment markers — conditional on a flag

Anchor a Word **comment** over the paragraph (or one sentence) you want to make
conditional. The code matches the comment's *text* against fixed keyword lists
(case-insensitive substring match) and acts based on `build_flags(tc)`:

| Comment keyword (any of) | Flag (`template_config`) | Behavior when flag is OFF |
|---|---|---|
| 无用户补充 / 未识别用户补充 / user entity补充 / 无用户实体补充 | `cuec_identified` (default True) | Delete the commented paragraph(s) |
| 子服务机构补偿 / 子服务机构补充 | `sso_cc_identified` (default True) | Delete the commented paragraph(s) |
| 处理transaction / processing user entity transaction | `has_transaction_processing` (default True) | Delete **only short** standalone paras (≤200 chars); long paras handled by inline substitution |
| single user entity report / single user entity时 | `single_user_entity` (default False → OFF) | When flag ON + comment says 删除/delete: delete span (≥30 chars) or whole para (if covers all / says 本段); else just strip the `[..]` brackets |
| 使用到了AI技术 / subject matter中某部分使用到了 | `has_ai_scope_exclusion` (default False) | Delete the AI scope-exclusion paragraph (kept only when flag ON) |
| other information | `has_other_information` (default True) | Delete only if comment also says 删除/delete |

Keyword lists live at `app.py` ~line 271 (`_CUEC_KW`, `_SSO_CC_KW`, `_TRANS_KW`,
`_SINGLE_KW`, `_AI_KW`, `_OTHER_KW`). To add a new conditional category you must
add a keyword list, wire it into `_build_annotation_maps`, and add the flag to
`build_flags`.

## Authoring gotchas

- Match the established Chinese/English wording exactly — a typo means the
  comment is ignored and the content always ships.
- Don't anchor a single-UE delete comment to a span shorter than ~30 chars, or
  to a single word — the ≥30-char guard turns it into a whole-paragraph delete.
- `（注：…）` / `(注：…)` author notes are guidance for the editor; the runtime
  conditional logic is driven by **Word comments**, not parenthetical notes,
  except the Other-Information case which reads the comment text.
- Tracked `pPrChange`/`rPrChange` format edits are **rejected** (old formatting
  restored), and all comments are stripped, during cleaning — so leave intended
  formatting as the accepted/base state, not as a tracked change.
