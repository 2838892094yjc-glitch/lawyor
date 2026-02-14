---
name: contract-template-transformation
description: Convert a filled Chinese contract into a template-ready variable specification for WPS/Word add-ins. Use when asked to turn normal contract text into reusable template variables, preserve fixed legal text, extract context/prefix/placeholder/suffix anchors, classify field type and format function, and identify optional clauses for paragraph toggles.
---

# Contract Template Transformation

Transform contract text into a strict JSON variable list that can be parsed by `ai-parser.js` and embedded by `wps_adapter.js`.

## Follow This Workflow

1. Read full text and split by clause/paragraph boundaries.
2. Run three-layer recognition:
   - Layer 1: explicit placeholders (`____`, `【】`, blank forms)
   - Layer 2: implicit concrete values that should be templatized (company, date, amount, address, ratio)
   - Layer 3: optional/conditional clauses that may disappear entirely
3. For each variable, isolate fixed vs variable text:
   - Put fixed leading text in `prefix`
   - Put variable body only in `placeholder`
   - Put fixed trailing text in `suffix`
4. Build robust anchors:
   - `context` must include enough nearby text to uniquely locate the variable
   - Prefer 20-60 chars with both left and right signals when possible
5. Infer UI and formatting:
   - `type`: `text|number|date|select|radio|textarea`
   - `formatFn`: `none|dateUnderline|dateYearMonth|chineseNumber|chineseNumberWan|amountWithChinese|articleNumber|percentageChinese`
   - `mode`: `insert|paragraph`
6. Create deterministic tags:
   - Use PascalCase Pinyin
   - Keep semantic uniqueness (avoid generic tags like `JinE` when multiple amounts exist)
7. Deduplicate:
   - Merge duplicates by semantic meaning and anchor uniqueness
   - Keep the most specific anchor set
8. Output only final JSON object: `{ "variables": [...] }`

## Hard Constraints

- Keep legal fixed text intact; never move fixed units into `placeholder`.
- `select`/`radio` must include non-empty `options`.
- For `paragraph` mode, variable should represent a whole optional clause, not a token.
- Prefer `formatFn: "none"` unless formatting conversion is truly required.
- Do not output markdown, explanation, or code fences.

## Quality Gates Before Final Output

- All required keys exist: `context,prefix,placeholder,suffix,label,tag,type,formatFn,mode`.
- Every variable has unique `tag`.
- Every variable can be found by at least one strong anchor strategy:
  - `context`
  - `prefix + placeholder + suffix`
  - `placeholder`
- No variable swallows neighboring fixed words.

## Failure Handling

If ambiguity exists, still output best-effort JSON and:
- Lower `confidence` (`medium`/`low`)
- Add concise reason in `reason`
- Keep structure valid (never omit required keys)

Load detailed constraints from:
- `references/output-schema.md`
- `references/quality-checklist.md`
- `references/tag-naming.md`
