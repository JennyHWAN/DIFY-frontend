# Placeholder Tokens (reference)

Every token below is replaced by `build_substitutions(ui, tc)` in `app.py`
(~line 1840). Author them in the `.docx` **exactly** as written — quote style and
fullwidth brackets matter. A token not in this table is left untouched (except
`[..]` brackets, which are stripped with content kept).

## English placeholders

| Token in template | Filled with |
|---|---|
| `[Service organization name]` | Company full name |
| `[Service organization short name]` | Company short name |
| `[Service organization's system]` / `[Service organization's system]` | System name (both quote styles) |
| `[type or name of system]` | System name |
| `[date] to [date]` | Period range (`start to end`) — **author the full compound** |
| `[date]` | Single date (Type I = period start "as of"; Type II = period end) |
| `[Date of the service auditor's report]` / `[Date of report]` | Report date |
| `[Service organization name]` addressee line | `To the {Management|Board of Directors} of {company}` (author the combined `To the Management/Board of Directors of [Service organization name]`) |
| `[Subservice organization name]` / `[... short name]` | First SSO entry |
| `[Subservice organization A/B name]`, `[... A/B/C short name]` | Multi-SSO entries (one per `|`-separated line in the SSO field) |
| `[identify the function or service provided by the subservice organization]` | SSO services text |
| `Shanghai[Beijing, Shenzhen]` | Signing city (author default+alternatives pattern) |
| `[Ernst & Young Hua Ming LLP]` | Removed (the branch line above is kept) |

## Chinese placeholders (fullwidth 【】 = U+3010/U+3011)

| Token | Filled with |
|---|---|
| `【服务机构名称】` / `【服务机构简称】` | Company name / short name |
| `【服务机构体系名称】` / `【服务机构服务体系名称】` | System name |
| `【日期】至【日期】` | Period range — **author the full compound** (consumed before single `【日期】`) |
| `【日期】` / `【报告日】` | Single date / report date |
| `【子服务机构名称】` / `【子服务机构简称】` / `【子服务机构A/B名称】` … | SSO entries |
| `中国 上海【或中国 北京或中国 深圳】` | Signing city (CN default+alternatives pattern) |
| `董事会/管理层` | Chosen addressee (董事会 or 管理层) |

## Author's rule

Adding a new placeholder = two edits: put the token in the `.docx`, **and** add
the `token → value` entry to `build_substitutions()`. Ordering matters when one
token is a substring of another — put the longer/compound one first.
