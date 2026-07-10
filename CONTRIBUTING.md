<!-- SPDX-License-Identifier: MIT -->
# Contributing to BrandDocs

Thanks for considering a contribution. BrandDocs has one structural promise,
"off-brand output is impossible by construction", and every rule below exists
to keep that promise true while many hands touch the code.

## Dev setup

```bash
git clone https://github.com/ferdinandobons/brand-docs.git
cd brand-docs
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements-ci.txt
```

Optional but recommended, the visual QA renderers (LibreOffice + Poppler,
Tesseract for OCR). One auto-detecting installer:

```bash
bash scripts/setup_visual_qa.sh
python scripts/brandkit/cli.py doctor   # confirms what is available
```

Without them the engine still works; visual QA degrades to deterministic-only
(level L0).

## Running the tests

The suite has three lanes. The first is the one every PR must keep green:

```bash
# 1) Full model-free suite (what CI runs on 3.10 / 3.11 / 3.12 / 3.13)
PYTHONPATH=scripts python -m pytest -q

# 2) Real-render lane (needs LibreOffice + Poppler; CI runs it in its own job)
BRANDDOCS_RUN_REAL_RENDER=1 PYTHONPATH=scripts python -m pytest tests/test_visual_qa.py -q

# 3) Lint + format (CI pins ruff; ruff.toml is the source of truth)
ruff check . && ruff format --check .
```

Mirror of the CI workflow: [.github/workflows/ci.yml](.github/workflows/ci.yml).
Run all three locally before pushing; a red gate never ships.

## The rules that gate every change

These are non-negotiable; CI enforces most of them with dedicated tests.

1. **Fail-closed.** Every sink the model can write to (comprehension, audit,
   triage, overrides, aliases, promotions) validates shape + membership in one
   all-or-nothing transaction. A bad proposal rejects cleanly; nothing
   load-bearing is ever half-written.
2. **Byte-identity when absent.** A new capture axis or feature must leave
   output byte-identical when the template does not exercise it.
   `tests/test_appearance_refactor_anchor.py` freezes this with a hash anchor;
   if your change moves the anchor, the PR must explain exactly why.
3. **No brand literals outside the profile.** No writer or IntermediateDocument
   block ever contains a literal style name, hex color, or font. The resolver
   (`scripts/brandkit/profile/resolver.py`) is the only reader of brand facts.
4. **Universal, never template-tuned.** Logic and prompts are generic; only the
   profile is specific. No matching rule may hardcode a single template's
   content.
5. **Three-format thinking.** The engine is shared across docx / pptx / xlsx.
   A new axis lands with an honest fail-closed check and declares its
   `realized_axes` per backend; format gaps are declared, not silent.
6. **Frozen vocabulary.** Names (verbs, role ids, check ids, schema fields) are
   owned by `scripts/brandkit/profile/schema.py` and mirrored in
   [CONVENTIONS.md](CONVENTIONS.md). Docs quote them verbatim;
   `tests/test_check_registry.py` and `tests/test_reference_sync.py` guard
   drift. The three skills' `reference/comprehension.md` files must stay
   byte-identical.
7. **Never commit real templates.** Company `.docx` / `.pptx` / `.xlsx` files
   stay out of the repo; `tests/test_no_proprietary.py` fails the build if an
   Office binary is tracked outside `tests/fixtures/` or `examples/`. Use the
   synthetic builders in `examples/builders/` for fixtures.

## Where things live

| Area | Entry point |
|---|---|
| Engine internals map | [scripts/brandkit/README.md](scripts/brandkit/README.md) |
| Frozen vocabulary & contracts | [CONVENTIONS.md](CONVENTIONS.md) |
| Dev & release process | [documentation/DEVELOPMENT.md](documentation/DEVELOPMENT.md) |
| Architecture & FAQ | [documentation/ARCHITECTURE.md](documentation/ARCHITECTURE.md) |
| Agent-facing workflow | [documentation/PLUGIN_WORKFLOW.md](documentation/PLUGIN_WORKFLOW.md) |

## Pull request checklist

- [ ] `ruff check .` and `ruff format --check .` pass.
- [ ] `PYTHONPATH=scripts python -m pytest -q` is green locally.
- [ ] New behavior comes with tests (including the fail-closed rejection path).
- [ ] Output is byte-identical for inputs that do not exercise the change, or
      the anchor move is justified in the PR description.
- [ ] Docs that quote a touched name (verb, check id, path, flag) are updated
      in the same PR; the three `reference/comprehension.md` stay identical.
- [ ] No real company template or proprietary string enters the repo.

## Sized first contributions

- **S**: add a deterministic QA check in `scripts/brandkit/qa/` (register it in
  `CHECK_REGISTRY`, write the honest fail path test).
- **M**: capture a new appearance fact on docx end to end (extract -> profile
  -> resolver -> generate -> check), byte-identical when absent.
- **L**: port an existing docx appearance axis to pptx or xlsx and declare it
  in that backend's `realized_axes`.

Open a "good first issue" on GitHub if you want one of these scoped for you.

## Releases

Maintainer-driven; the step-by-step checklist lives in
[documentation/DEVELOPMENT.md](documentation/DEVELOPMENT.md#release-checklist).
Contributors never need to tag: every merge to `main` rides the next release.
