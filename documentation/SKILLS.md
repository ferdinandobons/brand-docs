# The three skills & project status

## The three skills

| Skill | Format | Generates |
|---|---|---|
| **`brand-docx`** | Word `.docx` | reports, letters, memos: cover, headings, paragraphs, callouts, quotes, captions, lists, tables, in the template's structural order |
| **`brand-pptx`** | PowerPoint `.pptx` | decks: title / section / content slides from the template's real masters & layouts, with real bullet levels and long-text splitting |
| **`brand-xlsx`** | Excel `.xlsx` | workbooks: fills named cells & regions while **preserving formulas** and workbook structure |

All three expose the same core verbs: **`extract` → `verify` → `generate`**. Each skill is self-contained and **same-format** (a Word template makes Word
documents, never a deck or a sheet). They share one engine: a single profile
schema, resolver, OOXML layer and QA gate underpin all three formats.

On top of the deterministic core sit the **model-assisted verbs**, each one
fail-closed (the model only NAMES captured facts; the engine validates every
proposal and authors every value):

- **`comprehend`** - the model reads a bounded summary of the template and
  records what each structure is **for**: cover slots, demo-vs-real regions,
  derived-index conventions, brand-color naming (plus optional palette
  **aliases** for off-theme accents), faked-heading adjudication
  (`promote_appearance`), and recurring layouts proposed as reusable
  component/section fragments. Frozen into the Brand Profile; generation works
  with no comprehension at all (the deterministic path).
- **`learn`** - deterministically distills RECURRING QA findings from the
  cross-run generation history into a shell-frozen overrides lesson (advisory
  until `--accept`).
- **`propose-overrides`** - the model proposes corrections for the ambiguous
  recurring remainder `learn` could not bind, through the same fail-closed sink.
- **`refine`** - turns end-of-generation user feedback (text or a screenshot)
  into a validated comprehension delta, improving FUTURE generations.

---

## Project status

**Alpha, maturing.** The Word vertical (`brand-docx`) is the reference
implementation, verified end-to-end on real templates; PowerPoint and Excel share
the engine and are catching up.

| Area | Status |
|---|---|
| Shared engine (profile schema, resolver, OOXML, CLI, dual store) | ✅ working |
| `brand-docx`: extract → verify → generate | ✅ working |
| Document **structure** extraction & order-aware generation | ✅ working |
| Brand-guarantee enforcement (`verify` fails on missing artifacts) | ✅ working |
| Deterministic QA (L0: styles, palette, residual text, tables, formula and XLSX extension preservation, language) | ✅ working |
| `brand-pptx`: roles from real layouts, native charts / SmartArt / merged tables | ✅ working (fidelity still catching up to docx) |
| `brand-xlsx`: named-region fills, formula-preserving, native charts | ✅ working (fidelity still catching up to docx) |
| Visual QA (LibreOffice render + manifest-driven repair loop) | 🚧 implemented with graceful degraded mode |
| Native charts (DOCX / PPTX / XLSX), SmartArt diagrams (DOCX / PPTX), merged tables | ✅ working |
| Native Word `toc` (authored field, or deferral to a preserved outline TOC) with a Word-faithful refreshed cache (authored bookmarks, per-level styles, `PAGEREF`; template demo entries never survive) | ✅ working |
| Excel semantic number formats (`number.<family>` resolved to the template's mask) | ✅ working |
| Model-driven reusable-fragment population (`comprehend` → `components` / `sections`, with `{{slot}}` substitution) | ✅ working |
| Brand appearance capture: typography (all 3 formats) + paragraph geometry / table conditional formats / list numbering (docx), dominant-sampled and shell-verified | ✅ working |
| Learn-from-errors loop (cross-run generation report → `learn` / `propose-overrides`, advisory until `--accept`) | ✅ working |
| Faked-heading detection + model adjudication (`pseudo_headings` → `promote_appearance`) | ✅ working (docx-first) |
| Cover synthesis when the template has no cover anchor (`anchors.cover.kind == NONE`) | ✅ working (docx + pptx; xlsx N/A by design) |
| PyMuPDF PDF raster fallback | ✅ working |
| Optional OCR rendered-text residual scan | ✅ working when Tesseract is installed |
| Template-based skill eval set (DOCX/PPTX/XLSX) | ✅ working in CI |
| Strict visual mode (`--qa strict`) | ✅ working |
| Multi-template profile blending (`extract --blend`, same-format, value-facts only, pointers never cross shells) | ✅ working |
| Cross-template brand-drift report (`compare-profiles`, read-only, exit 1 on drift) | ✅ working |
| Role-first authoring surface (PROFILE.md role / palette-role / hints sections + per-skill authoring guidance) | ✅ working |
| Local-corpus fidelity benchmark (`scripts/corpus_benchmark.py`, real templates outside the repo) | ✅ working |
| Richer image analysis | 🔭 planned |

Visual Word overflow needs LibreOffice, since Word lays out at render time.
