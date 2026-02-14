# Output Schema (Parser-Compatible)

Return a JSON object:

```json
{
  "variables": [
    {
      "context": "string",
      "prefix": "string",
      "placeholder": "string",
      "suffix": "string",
      "label": "string",
      "tag": "PascalCasePinyin",
      "type": "text|number|date|select|radio|textarea",
      "options": ["string"],
      "formatFn": "none|dateUnderline|dateYearMonth|chineseNumber|chineseNumberWan|amountWithChinese|articleNumber|percentageChinese",
      "mode": "insert|paragraph",
      "layer": 1,
      "confidence": "high|medium|low",
      "reason": "string"
    }
  ]
}
```

Required keys per variable:
- `context,prefix,placeholder,suffix,label,tag,type,formatFn,mode`

Conditional key:
- `options` is required when `type` is `select` or `radio`.

Recommendations:
- Set `layer` to 1/2/3 based on recognition source.
- Set `confidence` explicitly for downstream review.
- Keep `reason` concise and evidence-based.
