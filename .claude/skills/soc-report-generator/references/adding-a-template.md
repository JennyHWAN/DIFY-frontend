# Adding / Changing a Template (checklist)

1. **Place the file** in `MA_template/` or `AR_template/`. The filename must
   **start with the WP number** (e.g. `10.1 ...docx`) — `resolve_template`
   matches on that prefix.

2. **Register it in `template_index.xlsx`** (`MA` or `AR` sheet): add the row
   mapping `(Category, Type, Standards, SSO, Language)` → that WP number. If the
   lookup finds no row, the UI shows an error and no section is produced.

3. **Use canonical placeholders** (see `placeholders.md`). If you need a token
   that doesn't exist yet, add it to `build_substitutions()` in `app.py`.

4. **Author conditionals as Word comments / `[or ..]` markers** (see
   `conditional-markers.md`), not as free text.

5. **Formatting** is normalized on fill: Times New Roman (EN) / 华文楷体 (CN),
   11 pt, bold removed, italic kept. Don't rely on bold; do keep intended
   formatting as the base (accepted) state, not a tracked change.

6. **Verify** by running the app, selecting the matching report config, and using
   "Generate MA + AR only" to inspect just the template output before involving
   Dify.

## Quick lookups

- `resolve_template` — `app.py` ~line 126
- `_build_annotation_maps` — ~line 296
- `fill_and_process_template` — ~line 641
- `build_substitutions` — ~line 1840
- `build_flags` — ~line 2066
