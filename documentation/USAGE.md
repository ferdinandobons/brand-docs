# Usage details

The [README quick start](../README.md#quick-start) covers the agent flow and the
three `extract → verify → generate` CLI commands. This page collects the common
use cases and the structured input format the generator expects.

## Common use cases

- **Consulting and operations reports**: generate branded Word reports, memos,
  briefs and status updates from the approved corporate template.
- **Sales and marketing decks**: create PowerPoint presentations from real masters
  and layouts instead of asking an AI to invent approximate slides.
- **Finance and planning workbooks**: fill named Excel inputs and regions while
  preserving formulas and workbook structure.
- **Repeatable agent workflows**: give Claude Code, Codex or another agent a
  reusable Brand Profile instead of re-explaining the brand for every document.

## Input format: the IntermediateDocument

The content you pass to `generate` (`idoc.json`) is an **IntermediateDocument** -
brand-agnostic typed blocks. Notice there is **no style, color or font anywhere**:
the profile resolves all of that.

```json
{
  "cover": { "title": "Quarterly Review", "fields": { "doc_id": "RPT-001" } },
  "blocks": [
    { "type": "heading", "level": 1, "text": "Highlights" },
    { "type": "paragraph", "text": "This paragraph resolves to the brand body style." },
    { "type": "callout", "intent": "info", "text": "The profile chooses the callout style." },
    { "type": "list", "items": [{ "text": "List styling comes from the profile." }] },
    { "type": "table", "columns": ["Area", "Status"], "rows": [["Pipeline", "Healthy"], ["Delivery", "Green"]] }
  ]
}
```

PowerPoint uses the same `IntermediateDocument`; Excel uses a `GridDocument`
(named-region fills, formulas preserved).

QA depth is explicit via `--qa fast|auto|deep|strict`; `deep`/`strict` write a
visual manifest for render-based review and targeted repair. Run
`python scripts/brandkit/cli.py doctor` to preflight dependencies.
