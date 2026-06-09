# SPDX-License-Identifier: MIT
"""DOCX brand typography capture (font family, size, and color).

The brand's REAL visible typography often lives as DIRECT run-level formatting
(``w:rPr/w:rFonts`` / ``w:sz`` / ``w:color``) on the template's content rather than
in the named styles or the theme: a designed template may put everything in
``Normal`` with a direct Roboto / Montserrat override at 22 half-points in accent1.
Role inference (``roles.py``) and theme extraction read only named styles and
``theme1.xml``, so those direct values are never captured and a generated document
falls back to the ``docDefaults`` font/size/color.

This module captures the DOMINANT direct run typography, deterministically, as
THREE INDEPENDENT axes (font family, size, color) sampled in a SINGLE pass:

  - per role: the dominant explicit value among the runs that use the role's style
    -> ``role['appearance']['font'] = {'latin': <name>}`` /
    ``role['appearance']['size_hp'] = <int>`` /
    ``role['appearance']['color'] = {'kind': ...}``;
  - the document's effective body typography: the dominant explicit value across all
    body runs -> ``theme['fonts']['body']['latin'/'size_hp']`` and
    ``theme['text']['body']['color']`` - the fallbacks the generator applies to a
    paragraph whose role carries no captured value.

Each axis is independent: a role may carry a captured size but no captured font
(or vice versa). Only a clear DOMINANT is recorded per axis (at least
:data:`_MIN_RUNS` explicit values and a winner covering at least
:data:`_MIN_DOMINANCE` of them), with its dominance stored as a per-axis confidence
(``confidence`` for font, ``size_confidence`` for size, ``color_confidence`` for
color). Capture is deterministic (model-free).

The brand guarantee is preserved: every captured value is a FACT observed in the
template, stored in the profile, applied only via the resolver, and re-validated
against what the shell proves it contains by ``check_appearance_targets``
(fail-closed). This module is purely additive - it only populates the already-
reserved ``appearance`` field and additive ``theme.fonts.body`` / ``theme.text``
keys; a template with no dominant direct value leaves all of them untouched, so
behavior is unchanged.
"""

from __future__ import annotations

from typing import Optional

from docx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR

from brandkit.common.typography import (
    PALETTE_WHERE,
)
from brandkit.common.typography import (
    capture_appearance as _capture_appearance,
)
from brandkit.common.typography import (
    capture_palette_facts as _capture_palette_facts,
)
from brandkit.ooxml import names

# The generalizable capture engine (``capture_appearance`` / ``capture_palette_facts``
# and the helpers/constants they lean on) now lives in the format-neutral
# ``common.typography`` engine; this module keeps ONLY the docx-specific run readers
# (``_run_size_hp`` / ``_run_color`` / ``_iter_para_runs`` / ``_is_link_run``), wraps
# them into the structural ``RunFacts`` view the engine consumes, and delegates. The
# docx adapter emits its CURRENT WordprocessingML tokens VERBATIM (``_run_color``
# keeps emitting ``'text1'`` / ``'accent1'``; the palette key namespace is unchanged):
# nothing here re-keys to ``THEME_SLOTS`` - that normalization is the pptx/xlsx
# adapters' job only, which is exactly what keeps the docx frozen-hash anchor green.
__all__ = [
    "capture_fonts",
    "capture_palette",
    "PALETTE_WHERE",
]

# python-docx's sentinel for a theme color that maps to no real slot; its
# ``xml_value`` is the truthy string ``"UNMAPPED"``. It is not a brand token (verify
# has no slot for it and apply cannot realize it), so it is never captured.
_UNMAPPED_THEME_TOKEN = MSO_THEME_COLOR.NOT_THEME_COLOR.xml_value

# WordprocessingML qualified-name builder for the link-color helper (it walks the
# raw ``w:hyperlink`` ancestor / ``w:rStyle`` to detect a link run).
_W = names.make_qn("w")

# The hyperlink character-style ids whose presence on a run marks it as link text
# even when it is NOT physically nested under a ``w:hyperlink`` element (a manually
# styled cross-reference). Closed, spec-fixed style ids (NOT brand literals).
_HYPERLINK_RSTYLES: frozenset[str] = frozenset({"Hyperlink", "FollowedHyperlink"})


def _run_size_hp(run) -> Optional[int]:
    """The run's EXPLICIT size as half-points (``w:sz@w:val``), or ``None``.

    ``run.font.size`` is an explicit-only ``Length`` (``None`` when the size is
    inherited from the style/theme), so a run that inherits its size contributes
    nothing. The half-point bucket ``round(pt * 2)`` matches OOXML's ``w:sz`` unit.
    """
    try:
        size = run.font.size
        if size is None:
            return None
        return round(size.pt * 2)
    except Exception:
        # A malformed measure the OOXML layer refuses to parse contributes nothing
        # to this axis - capture must never crash the extraction.
        return None


def _run_color(run) -> Optional[tuple[str, ...]]:
    """The run's EXPLICIT color as a hashable bucket key, or ``None``.

    ``run.font.color`` is a ``ColorFormat`` whose ``.type`` is ``None`` when the
    color is inherited. An RGB color buckets as ``('hex', <RRGGBB>)``; a THEME color
    buckets as ``('theme', <wordprocessingml token>)`` from the slot's
    ``.theme_color.xml_value`` (e.g. ``'accent1'``, ``'text1'``). AUTO / None / an
    unmapped theme slot contributes nothing (it is not a captured brand value).
    """
    try:
        color = run.font.color
        ctype = color.type
        if ctype == MSO_COLOR_TYPE.RGB and color.rgb is not None:
            return ("hex", str(color.rgb))
        if ctype == MSO_COLOR_TYPE.THEME:
            token = getattr(color.theme_color, "xml_value", None)
            # Drop the UNMAPPED sentinel: it is not a verifiable/appliable brand
            # token, so it must never enter the profile (keeps apply/verify in sync).
            if token and token != _UNMAPPED_THEME_TOKEN:
                return ("theme", token)
    except Exception:
        # A spec-valid-but-unmappable themeColor (e.g. 'none'/'phClr') makes
        # python-docx raise on access; that run contributes nothing to this axis
        # rather than crashing the extraction.
        return None
    return None


class _DocxRunFact:
    """A docx run, reduced to the structural :class:`~brandkit.common.typography.RunFacts`
    view the shared capture engine consumes.

    The three axes are read EXACTLY as the v1 inline capture did (``run.font.name`` or
    ``None`` / :func:`_run_size_hp` / :func:`_run_color`), so the engine sees the
    byte-identical votes. ``color`` keeps the docx WordprocessingML token namespace
    verbatim (``'text1'`` / ``'accent1'`` from ``_run_color``) - NO normalization to
    ``THEME_SLOTS`` - which is what keeps the docx palette keys (and the frozen-hash
    anchor) unchanged."""

    __slots__ = ("style_key", "text", "font_name", "size_hp", "color", "is_link")

    def __init__(self, run, style_key, *, is_link: bool) -> None:
        self.style_key = style_key
        self.text = run.text or ""
        self.font_name = run.font.name or None
        self.size_hp = _run_size_hp(run)
        self.color = _run_color(run)
        self.is_link = is_link


def _para_style_key(para) -> tuple[Optional[str], Optional[str]]:
    """The ``(style_id, style_name)`` of a paragraph's effective style, crash-safe.

    python-docx resolves a paragraph's effective style (a paragraph with no explicit
    ``pStyle`` reports the document's default style), so this is the real bucket key
    the per-role fold matches against. A reader that refuses to resolve the style
    yields ``(None, None)`` - the run then votes only toward the document body."""
    try:
        style = para.style
        sid = getattr(style, "style_id", None) if style is not None else None
        sname = getattr(style, "name", None) if style is not None else None
    except Exception:
        sid = sname = None
    return (sid, sname)


def _font_run_facts(doc):
    """Yield a :class:`_DocxRunFact` for every DIRECT ``w:r`` run in document order
    (``doc.paragraphs`` then ``para.runs``), tagged with the paragraph's style key.

    This is the font/size/color sampling pass: it deliberately does NOT widen to the
    hyperlink runs (matching v1 ``capture_fonts``, which read ``para.runs`` only). A
    single ordered generator keeps the ``Counter`` insertion order deterministic."""
    for para in doc.paragraphs:
        style_key = _para_style_key(para)
        for run in para.runs:
            yield _DocxRunFact(run, style_key, is_link=False)


def capture_fonts(doc, roles: dict, theme: dict) -> None:
    """Capture dominant direct run typography (font, size, color) into ``roles``
    (per role ``appearance``) and the document defaults (``theme['fonts']['body']``
    for font/size, ``theme['text']['body']`` for color), mutating both in place.

    Reads only the EXPLICIT run value per axis (``run.font.name`` / ``run.font.size``
    / ``run.font.color``); a run that inherits an axis from the style/theme
    contributes nothing to THAT axis (the three axes are sampled independently).
    python-docx resolves a paragraph's effective style (a paragraph with no explicit
    ``pStyle`` reports the document's default style), so runs are bucketed by their
    real style id/name. This now builds a docx ``RunFacts`` generator over the direct
    runs and delegates to the format-neutral
    :func:`~brandkit.common.typography.capture_appearance`; the default
    ``role_style_key`` reproduces the docx ``named_style`` OR-match byte-identically.
    """
    _capture_appearance(_font_run_facts(doc), roles, theme)


# ---------------------------------------------------------------------------
# theme.palette capture (model-free; the UNDERSTAND half of model-driven color)
# ---------------------------------------------------------------------------
def _iter_para_runs(para):
    """Yield every run in ``para``: its direct ``w:r`` runs AND the runs nested under
    its ``w:hyperlink`` elements.

    python-docx's ``para.runs`` exposes only the direct ``w:r`` children, so a link
    run (nested under ``w:hyperlink``) is otherwise invisible to capture. Newer
    python-docx (>= 1.x) surfaces ``para.hyperlinks[*].runs``; this widens the pass
    to include them, crash-safe (a degraded reader yields the direct runs only).
    """
    for run in para.runs:
        yield run
    try:
        for hyperlink in para.hyperlinks:
            for run in hyperlink.runs:
                yield run
    except Exception:
        # An older python-docx without ``para.hyperlinks`` simply contributes no
        # nested link runs - capture must never crash on a missing attribute.
        return


def _is_link_run(run) -> bool:
    """True if ``run`` is hyperlink text: nested under a ``w:hyperlink`` ancestor OR
    carrying a ``Hyperlink``/``FollowedHyperlink`` ``w:rStyle``.

    Wrapped fully crash-safe by the single caller; reads only structural OOXML
    (no brand literals). A run python-docx cannot introspect contributes nothing.
    """
    try:
        rpr = run._r.find(_W("rPr"))
        if rpr is not None:
            rstyle = rpr.find(_W("rStyle"))
            if rstyle is not None and rstyle.get(_W("val")) in _HYPERLINK_RSTYLES:
                return True
        node = run._r.getparent()
        while node is not None:
            if names.local_name(node.tag) == "hyperlink":
                return True
            node = node.getparent()
    except Exception:
        return False
    return False


def capture_palette(doc, roles: dict, theme: dict) -> None:
    """Capture the template's brand PALETTE into ``theme['palette']`` (mutated in
    place), additively and deterministically.

    The palette is a map keyed by a TEMPLATE-DERIVED id - a theme slot token
    (``accent1`` / ``text1`` / ...) for a theme color, or ``hex:RRGGBB`` for an
    observed off-theme run color. Each entry carries:

      - ``ref``: the byte-identical :func:`_color_obj` (``{kind:theme,theme}`` |
        ``{kind:hex,hex}``);
      - ``provenance``: a list of observed ``{where, detail}`` facts from the
        closed :data:`PALETTE_WHERE` vocabulary, sorted ``(where, detail)``;
      - ``frequency``: a COARSE bucket (``dominant`` | ``accent`` | ``rare``),
        never raw counts;
      - ``name`` / ``purpose`` / ``use_when``: ``null`` in this deterministic path
        (``comprehend`` is the only writer that fills them).

    Provenance is built ONLY from observed facts:
      (a) the theme-color slots the template actually carries (seed theme-keyed
          entries; existence, not a where-fact);
      (b) explicit ``w:color`` on runs (a SINGLE pass via :func:`_run_color`),
          INCLUDING a low-floor accent aggregation - a color on at least
          :data:`_MIN_ACCENT_RUNS` runs that is NOT the document-dominant body
          color is an ``accent`` entry (no dominance gate, accents are sparse);
      (c) the per-role ``appearance.color`` already captured (``role.appearance``);
      (d) link-run colors (runs under a ``w:hyperlink`` ancestor / ``Hyperlink``
          style), wrapped crash-safe, falling back to the theme ``hlink`` /
          ``folHlink`` slot when no explicit link color is observed (``link.color``).

    The hardcoded, template-INVARIANT ``theme.palette_roles`` map is NOT trusted as
    brand evidence; it is recorded only as a non-authoritative ``palette_role``
    where-entry on the slot it names. Deterministic and byte-identical on
    re-extract; a template with no observed color leaves an empty ``{}`` palette.

    This now builds a docx ``RunFacts`` generator over the WIDENED run pass (direct
    ``w:r`` runs AND the runs nested under ``w:hyperlink`` - source d) and delegates to
    :func:`~brandkit.common.typography.capture_palette_facts`. The bucket namespace is
    the docx WordprocessingML tokens VERBATIM (``_run_color``); nothing re-keys to
    ``THEME_SLOTS``, so the palette keys are byte-identical to v1.
    """
    _capture_palette_facts(_palette_run_facts(doc), roles, theme)


def _palette_run_facts(doc):
    """Yield a :class:`_DocxRunFact` for every run in the WIDENED palette pass
    (``doc.paragraphs`` then :func:`_iter_para_runs`, which adds the hyperlink runs),
    each tagged with :func:`_is_link_run`.

    The palette pass reads only ``color`` / ``is_link`` (font/size/style_key are
    irrelevant here), but the same ``_DocxRunFact`` view carries them. A single
    ordered generator keeps the dominance ``Counter`` insertion order deterministic,
    identical to v1 ``capture_palette``."""
    for para in doc.paragraphs:
        for run in _iter_para_runs(para):
            yield _DocxRunFact(run, None, is_link=_is_link_run(run))
