# Contract Template Transformation Skill

This skill turns a normal Chinese contract into parser-compatible template variables.

## Files

- `SKILL.md`: trigger + workflow instructions
- `references/output-schema.md`: strict output contract
- `references/quality-checklist.md`: pre-output QA checklist
- `references/tag-naming.md`: deterministic tag naming
- `scripts/validate_output.js`: local validator for generated JSON

## Validate Example

```bash
node contract-template-transformation-skill/scripts/validate_output.js /path/to/output.json
```
