# SPDX-License-Identifier: MIT
"""Frozen-hash anchor for the cross-format appearance refactor (Cluster A, PR-0).

The docx capture/apply appearance logic is being lifted into format-neutral shared
modules (``common.typography`` / ``common.appearance``) that pptx/xlsx will also
drive. The HARD invariant is that the docx adapter still produces BYTE-IDENTICAL
output: it delegates to the shared engine but emits its current WordprocessingML
tokens verbatim, with the set-only-when-unset guards unchanged.

This module pins a ``sha256`` of a docx ``generate()`` output for a FIXED
``(profile, shell, idoc)`` triple into a frozen constant and asserts it stays
identical. The shell is the committed, byte-stable synthetic ``acme_complex.docx``
fixture (NOT a freshly built ``Document()`` whose ``core.xml`` would carry a
wall-clock timestamp); the profile is an inline, self-contained dict that exercises
every refactored apply path:

  - body font + body size (``theme.fonts.body.latin`` / ``size_hp``);
  - body color (``theme.text.body.color`` as a hex);
  - a role-level THEME-token color (``heading.1`` -> ``accent1``), enriched to a hex
    by the resolver from ``theme.colors`` and applied through the WML theme-color map;
  - a per-run palette color TOKEN (``color: "accent1"``) on a hyperlink run, the
    raw-XML hyperlink path that stays inside docx.

Keeping the profile inline (instead of re-extracting it each run) isolates the proof
to ``generate()``'s output bytes, so this anchor breaks ONLY if the refactor changed
what the docx writer emits - which is exactly what PR-0 must never do.

If a LATER, intentional change to the docx writer alters these bytes, recompute the
constant deliberately (the test prints the actual hash on failure) - never silence
the assertion.
"""

from __future__ import annotations

import hashlib
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.formats.docx import generate as docx_generate
from brandkit.ir import model as ir
from brandkit.profile import schema

# The committed, byte-stable synthetic shell (also used by the complex-fidelity
# suite). Its on-disk bytes do not change between runs, so the generated output hash
# is reproducible across processes.
_SHELL = (
    Path(__file__).resolve().parents[0] / "fixtures" / "complex" / "acme_complex.docx"
)

# The frozen output hash. Computed from the CURRENT docx writer; the refactor must
# keep it identical. (Recompute deliberately only on an intentional writer change.)
_FROZEN_SHA256 = "c96548539684d65df6e91f5ee52009df191ad09670b1e1498672e2add16fa878"


def _anchor_profile() -> dict:
    """A fixed, self-contained profile exercising every refactored apply path."""
    prof = schema.build_envelope("docx", {"name": "anchor"})
    prof["surface"] = {"docx": {}}
    prof["roles"] = {
        "_index": ["paragraph", "heading.1"],
        "paragraph": {
            "resolver": {
                "type": "named_style",
                "style_id": "Normal",
                "style_name": "Normal",
            },
        },
        "heading.1": {
            "resolver": {
                "type": "named_style",
                "style_id": "Heading1",
                "style_name": "Heading 1",
            },
            # A theme-TOKEN color: the resolver enriches it with the concrete hex from
            # theme.colors and the writer applies it via the WML theme-color map.
            "appearance": {"color": {"kind": "theme", "theme": "accent1"}},
        },
    }
    prof["theme"] = {
        "colors": {"accent1": {"hex": "4F81BD"}},
        "fonts": {"body": {"latin": "Roboto", "size_hp": 22}},
        "text": {"body": {"color": {"kind": "hex", "hex": "1F4E79"}}},
        "palette": {
            "accent1": {
                "ref": {"kind": "theme", "theme": "accent1"},
                "provenance": [],
                "frequency": "rare",
                "name": None,
                "purpose": None,
                "use_when": None,
            }
        },
    }
    return prof


def _anchor_idoc() -> ir.IntermediateDocument:
    """A fixed IR exercising headings, a body paragraph with mixed runs, and a
    hyperlink run carrying a per-run palette color TOKEN (the raw-XML link path)."""
    return ir.IntermediateDocument(
        blocks=[
            ir.Heading(level=1, runs=[{"t": "Anchor Title"}]),
            ir.Paragraph(
                runs=[
                    {"t": "Body paragraph with "},
                    {"t": "bold", "b": True},
                    {"t": " text."},
                ]
            ),
            ir.Paragraph(
                runs=[
                    {"t": "A link: "},
                    {"t": "site", "link": "https://example.com", "color": "accent1"},
                ]
            ),
        ]
    )


def _generate_hash() -> str:
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.docx"
        docx_generate.generate(_anchor_profile(), _SHELL, _anchor_idoc(), out)
        return hashlib.sha256(out.read_bytes()).hexdigest()


@unittest.skipUnless(_SHELL.exists(), "complex docx fixture missing")
class AppearanceRefactorAnchorTest(unittest.TestCase):
    def test_docx_generate_output_matches_frozen_hash(self):
        """The docx writer's output for the fixed triple is byte-for-byte unchanged."""
        actual = _generate_hash()
        self.assertEqual(
            actual,
            _FROZEN_SHA256,
            "docx generate() output bytes changed: the appearance refactor must "
            "keep the docx adapter byte-identical. If this is an INTENTIONAL writer "
            f"change, update _FROZEN_SHA256 to {actual!r} deliberately.",
        )

    def test_docx_generate_is_byte_idempotent(self):
        """Two generations of the fixed triple hash identically (determinism guard)."""
        self.assertEqual(_generate_hash(), _generate_hash())


if __name__ == "__main__":
    unittest.main()
