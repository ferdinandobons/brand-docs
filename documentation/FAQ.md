# FAQ

**What is BrandDocs in one line?**
An open-source [Claude Code](https://www.anthropic.com/claude-code) skill bundle
that turns a company's Word, PowerPoint or Excel template into unlimited on-brand
documents of the same format.

**How do I generate on-brand Word, PowerPoint and Excel documents from a company template?**
Point BrandDocs at one branded `.docx`, `.pptx` or `.xlsx`. It `extract`s a
reusable **Brand Profile** (theme colors, fonts, named styles, document structure,
layouts, logos, tables, formulas), then `generate`s new documents from the
original template shell. See the [quick start](../README.md#quick-start).

**How is this different from asking ChatGPT or Claude to "use this template"?**
General-purpose document skills only loosely imitate a reference file, so fonts
drift, the palette wanders and the corporate cover → contents → body structure is
lost. BrandDocs is faithful **by construction**: generators never write a literal
style name, hex color or font, and `verify` refuses any profile that points at
something the template doesn't define. See the
[comparison table](ARCHITECTURE.md#why-not-just-ask-an-ai-to-use-this-template).

**Does it work with Codex or other agents, not just Claude Code?**
Yes. The three skills (`brand-docx`, `brand-pptx`, `brand-xlsx`) are plain agent
skills; [installation](INSTALLATION.md) covers both Claude Code and Codex. The
underlying engine is also usable as a direct Python CLI.

**Is it free and open source?**
Yes. **MIT licensed**, self-contained, pure `python-docx` / `python-pptx` /
`openpyxl` + OOXML. No cloud, no external services, no vendor lock-in.

**Does it keep my templates private?**
Everything runs locally. `brand-kit/` and `generated/` are git-ignored, and a test
fails the build if any real Office binary is committed. Never commit real company
templates; use synthetic fixtures.

**Can it preserve Excel formulas and the template's structure?**
Yes. Excel generation fills named cells and regions while **preserving formulas**,
and Word/PowerPoint generation keeps the template's ordered skeleton (cover → table
of contents → body). See
[Structure-aware](ARCHITECTURE.md#structure-aware-not-just-style-aware).

---

**Keywords:** AI document generator · on-brand document generation · template to
document · Claude Code skill · Codex skill · AI agent skill · brand template
automation · corporate template to document · docx / pptx / xlsx generator · Word /
PowerPoint / Excel automation · Office automation · OOXML · python-docx ·
python-pptx · openpyxl · brand profile · brand kit · document automation.
