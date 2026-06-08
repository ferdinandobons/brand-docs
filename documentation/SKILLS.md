# The three skills & project status

## The three skills

| Skill | Format | Generates |
|---|---|---|
| **`brand-docx`** | Word `.docx` | reports, letters, memos: cover, headings, paragraphs, callouts, quotes, captions, lists, tables, in the template's structural order |
| **`brand-pptx`** | PowerPoint `.pptx` | decks: title / section / content slides from the template's real masters & layouts, with real bullet levels and long-text splitting |
| **`brand-xlsx`** | Excel `.xlsx` | workbooks: fills named cells & regions while **preserving formulas** and workbook structure |

All three expose the same three verbs: **`extract` → `verify` → `generate`**. Each skill is self-contained and **same-format** (a Word template makes Word
documents, never a deck or a sheet). They share one engine: a single profile
schema, resolver, OOXML layer and QA gate underpin all three formats.

---

## Project status

**Alpha.** The Word vertical (`brand-docx`) is the reference implementation,
verified end-to-end on real templates; PowerPoint and Excel share the engine and
are catching up.

| Area | Status |
|---|---|
| Shared engine (profile schema, resolver, OOXML, CLI, dual store) | ✅ working |
| `brand-docx`: extract → verify → generate | ✅ working |
| Document **structure** extraction & order-aware generation | ✅ working |
| Brand-guarantee enforcement (`verify` fails on missing artifacts) | ✅ working |
| Deterministic QA (L0: styles, palette, residual text, tables, formula preservation, language) | ✅ working |
| `brand-pptx`: roles from real layouts, basic generation | 🚧 early |
| `brand-xlsx`: named-region fills, formula-preserving | 🚧 early |
| Visual QA (LibreOffice render + manifest-driven repair loop) | 🚧 implemented with graceful degraded mode |
| Native PPTX charts / SmartArt / richer component regeneration | 🔭 catalogued, regeneration staged |
| PyMuPDF PDF raster fallback | ✅ working |
| Optional OCR rendered-text residual scan | ✅ working when Tesseract is installed |
| Template-based skill eval set (DOCX/PPTX/XLSX) | ✅ working in CI |
| Strict visual mode (`--qa strict`) | ✅ working |
| Richer image analysis | 🔭 planned |

Visual Word overflow needs LibreOffice, since Word lays out at render time.
