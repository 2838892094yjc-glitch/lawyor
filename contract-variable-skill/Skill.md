---
name: contract-variable-recognition
description: Recognize and structure contract variables from Chinese legal/commercial contracts into JSON for form generation and document embedding. Use when converting contract text into context/prefix/placeholder/suffix anchors with field type, format function, and insert/paragraph mode.
---

# Contract Variable Recognition

Use this legacy skill when output must match the existing `ai-parser.js` runtime schema.

## Output Contract

Return only:

```json
{
  "variables": []
}
```

Each variable must include:
- `context,prefix,placeholder,suffix,label,tag,type,formatFn,mode`

Allowed enums:
- `type`: `text|number|date|select|radio|textarea`
- `formatFn`: `none|dateUnderline|dateYearMonth|chineseNumber|chineseNumberWan|amountWithChinese|articleNumber|percentageChinese`
- `mode`: `insert|paragraph`

## Recognition Strategy

- Layer 1: explicit placeholders
- Layer 2: implicit concrete values that should be templatized
- Layer 3: optional clauses suitable for paragraph toggles

Keep fixed legal wording in `prefix/suffix`; keep only variable body in `placeholder`.
