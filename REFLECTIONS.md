<!-- SPDX-License-Identifier: MIT -->
# Strategic Reflections

> Living document. Last updated: 2026-06-10, after the B4 + D1-D3 + E1-E2 wave,
> before the v0.8.0 release. This is the honest self-assessment that drives the
> post-roadmap hardening phase: what the project gets right, where its structural
> ceilings are, and what it takes to become the open-source standard for
> AI-driven corporate document generation.

## Direction

BrandDocs stays **open source (MIT)** with one explicit ambition: become the
**standard** for people who want to create corporate documents with AI. Not a
personal tool, not (for now) a commercial product. Every priority below follows
from that choice: trust, reproducibility, onboarding, distribution, and a
fidelity guarantee nobody else offers.

## 1. What the project gets structurally right

These are the load-bearing decisions. They are the moat; do not weaken them.

- **The inversion.** Almost every "LLM writes documents" tool lets the model
  author formatted output, making brand fidelity best-effort. BrandDocs inverts
  the relationship: the model can only **name** facts captured from the
  template; the deterministic engine is the only author of values, and every
  pointer is membership-validated fail-closed against what the shell actually
  contains. "Off-brand impossible by construction" is a structural property,
  not a slogan.
- **Byte-identity as regression armor.** Deterministic regeneration plus the
  frozen-hash anchor test means every new capture axis must prove the
  no-capture path did not move by a single byte. This is what made it safe to
  land six major features in two days.
- **Fail-closed everywhere.** Every model-writable sink (comprehension, audit,
  triage, overrides, aliases, promotions) validates shape + membership in one
  all-or-nothing transaction. A bad proposal rejects cleanly; nothing
  load-bearing is ever half-written.
- **Format-uniform seams.** The resolver/appearance seam scaled from 3 axes
  (font/size/color) to 6 (geometry, table, numbering) without a redesign, each
  with its own honest fail-closed check. The pattern is proven extensible.
- **The development process.** Incremental clusters, an independent gate per
  cluster (ruff, full suite, real-render, anchor, reference-sync), commits to
  main without premature tags. The discipline is what makes AI-driven
  development velocity safe; treat it as non-negotiable.

## 2. Structural ceilings (the honest part)

Not defects; design choices that will eventually present a bill.

1. **The single-template profile is the quality ceiling.** The anti-overfitting
   rule protects the *code*, but the *profile* is by definition learned from
   one file. A sparse template yields a sparse profile (the dominance floors
   have too few samples); real brands live across several templates plus brand
   guidelines. The profile JSON design already permits the next leap:
   **multi-template extraction with reconciliation** (one profile referencing
   several shells, precedence rules for conflicts). This is the most valuable
   architectural investment after the current roadmap.
2. **The guarded half is now far stronger than the unguarded half.** Brand
   fidelity (resolver, 12+ checks, 3-level QA) is hardened; the
   **IntermediateDocument authoring** (which roles to use, how to structure
   content, when to reuse fragments) is the least-guarded stage and is where a
   real user perceives "correct document" vs "great document". After Cluster E,
   the marginal hour invests better in authoring intelligence than in a seventh
   appearance axis.
3. **Complexity is sustainable only while the discipline holds.** Six axes,
   ten comprehension sinks, overrides, triage, audit, fragments, 800+ tests,
   one maintainer. This is not over-engineering (the rigor is exactly what
   enables the velocity), but it is a high-wire act: the first feature merged
   without its check and its tests starts the decay. Conventions and CI gates
   must stay mandatory.
4. **One rendering engine for visual QA.** LibreOffice is not Word; the visual
   gate certifies "faithful according to LibreOffice". The renderer
   cross-check was cut for sound reasons (one layout engine, rasterization
   noise), so this stays a known ceiling: document it for users, and prefer
   templates that render identically across viewers.
5. **Cross-format profile drift is undetected.** The same business template
   extracted as docx and pptx can yield profiles that disagree (a theme slot in
   one, a raw hex in the other). A `compare-profiles` verb that reports
   mismatches in captured colors, fonts, and roles would close this.

## 3. Process verdict for the 0.8.0 wave

- **What worked, keep it:** cluster-per-cluster delivery with an independent
  gate caught real bugs before they shipped (the xlsx theme-index swap, the
  dead `learn` on real producer data, an illegal alias token in a design spec).
  12 commits after v0.7.0, remote CI green on every one.
- **Lesson, monolithic release surface:** ~250 changelog lines in one
  [Unreleased] block is hard to verify end-to-end, narrate, and roll back. The
  consolidation rule stays, but future waves should cut smaller, thematic
  releases (learning / fidelity / robustness) at natural stable points.
- **Non-negotiable before the v0.8.0 tag:** an end-to-end validation on a
  REAL template with a re-extracted profile (extract, comprehend, generate,
  strict visual QA). All the new capture axes change output only for
  re-extracted profiles, and that path has so far been validated only on
  synthetic fixtures. The byte-identity anchor protects the old path, not the
  new one.
- **E4 risk note:** cover synthesis for `AnchorKind.NONE` is qualitatively
  different from everything else in the wave: it *creates* structure instead of
  capturing or referencing it, and it touches the historically most fragile
  path (covers). It deserves its own QA cycle and is a legitimate candidate to
  ship in the release after v0.8.0.

## 4. OSS-standard gap analysis (fresh-eyes review, 2026-06-10)

An independent cold-read of the repo through the lens of "what does it take for
an OSS project to become a category standard". Condensed, actionable findings:

### Onboarding
- The quick start does not point at the shipped synthetic example templates
  (`examples/templates/`); a cold reader has no file to plug in. Add a
  "try with the included example" one-liner.
- Visual QA messaging mixes "core" and "optional"; reframe as core-but-degrades
  (L0-only) so first-run users are not surprised by `visual.unavailable`.
- Installation paths (Claude Code / Codex / CLI-only) should be one anchor line
  in the README instead of a split across two files.
- No visual of what extraction produces (PROFILE.md role table, QA output).
  A single annotated screenshot would answer "does this solve my problem".

### Contributor experience
- `CONTRIBUTING.md` is a 10-line stub: needs dev setup, test-running guide, CI
  expectations, conventions map, and a release checklist.
- No `.github/ISSUE_TEMPLATE/` (bug / feature / question) and no PR template.
- No "good first issues" onramp with sized tasks (S: add a QA check; M: capture
  a new axis on docx; L: port an axis to pptx).
- No engine-internal map: a `scripts/brandkit/README.md` with the module layout
  (`formats/`, `ir/`, `ooxml/`, `profile/`, `qa/`, `common/`), the five verbs'
  call path, and the fail-closed pattern.

### Trust signals
- "Alpha" undersells a system with three QA lanes and 800+ tests; reframe as
  maturity-per-format (Word robust, PowerPoint/Excel catching up).
- CITATION.cff exists but is not linked from the README.
- NOTICE should state explicitly which OOXML logic is custom and which
  delegates to python-docx / python-pptx / openpyxl.
- A written release checklist (the one already practiced: local CI mirror,
  README sync, CHANGELOG, tag) belongs in DEVELOPMENT.md.

### Distribution (the highest-impact gap)
- The engine is consumable only via clone + PYTHONPATH; there is no documented
  Python API (`brandkit/__init__.py` exports), no PyPI package, no Docker
  image. For CI/CD users, researchers, and tool builders, this is the adoption
  blocker. Plan: documented API surface first, then a `brand-docs-engine` PyPI
  package, then a Dockerfile (python + LibreOffice + Poppler) on ghcr.io.
- requirements are unpinned and Python 3.13 is untested; pin and add an
  allowed-to-fail 3.13 CI lane.
- The two-phase nature (deterministic-first, model-assisted on top) is the
  project's best argument and is not discoverable from the README.

### Documentation hygiene (applied with this commit)
- `documentation/FAQ.md` merged into `documentation/ARCHITECTURE.md` (FAQ
  section) and deleted: same answers lived in both files.
- `documentation/DIRECTORY_SUBMISSIONS.md` (internal marketing kit) moved to
  `.github/DIRECTORY_SUBMISSIONS.md` and dropped from the LLM-facing index.
- `documentation/README.md` kept (it is the folder's landing page on GitHub)
  but slimmed to match. `PLUGIN_WORKFLOW.md` needs a regenerated diagram that
  includes the model-in-the-loop and learning paths (flagged, not yet done).

## 5. Prioritized action plan

The order of the autonomous hardening phase, after E3/E4 and the performance
exploration close the current roadmap:

- **P0, before the v0.8.0 tag:** real-template end-to-end validation
  (re-extract, comprehend, generate, strict visual QA); README/docs sync for
  every new verb and behavior in the release; release checklist written down.
- **P1, solidity:** CONTRIBUTING expansion + issue/PR templates + good first
  issues; quick start wired to the shipped example templates; engine module
  README; PLUGIN_WORKFLOW diagram refresh with a drift guard; maturity-status
  reframing; CITATION link; NOTICE clarity.
- **P2, reach:** documented Python API surface (`brandkit` exports +
  docstrings), `brand-docs-engine` on PyPI, Dockerfile + ghcr image, pinned
  requirements, Python 3.13 CI lane, two-phase workflow section in README.
- **P3, architecture:** multi-template profile blending (design first, ship
  behind the same fail-closed discipline); `compare-profiles` verb;
  IntermediateDocument authoring intelligence (fragment suggestion, role-choice
  guidance); a heterogeneous real-template corpus for fidelity benchmarking
  (local-only, never committed) with Word-vs-LibreOffice notes for users.

Items in section 4/5 that touch the engine go through the same rules as
everything else: fail-closed, byte-identical when absent, universal across the
three formats, never tuned to a single template.
