# SPDX-License-Identifier: MIT
"""PPTX brand typography capture: the run-fact view the shared engine consumes.

This is the PPTX peer of ``formats/docx/typography.py``: it walks the deck's runs
and yields the structural :class:`~brandkit.common.typography.RunFacts` view so the
format-neutral capture engine (``capture_appearance`` / ``capture_palette_facts``)
fills the SAME profile shape for pptx that it fills for docx - the document body
font/size (``theme.fonts.body``), the body color (``theme.text.body``), the per-role
``appearance``, and the brand ``theme.palette``.

Two PPTX specifics, both UNIVERSAL (driven off captured facts, never a tuned slot):

  - **Unit:** python-pptx exposes a run's explicit size as an ``Emu``/``Pt`` length;
    the half-point bucket ``round(pt * 2)`` matches the engine's docx unit so a body
    size captured from a deck and from a doc compare in the same space.
  - **Token namespace:** a SCHEME color's ``MSO_THEME_COLOR.xml_value`` is mostly
    already a :data:`~brandkit.common.color.THEME_SLOTS` name (``accent1`` / ``dk1`` /
    ``hlink`` / ...), but text/background slots surface as the WordprocessingML
    aliases ``tx1`` / ``bg1`` / ``tx2`` / ``bg2``. Those are normalized to their
    ``dk1`` / ``lt1`` / ``dk2`` / ``lt2`` slot via the SAME alias map
    ``color.resolve_theme_color`` uses, and any token that does not land on a real
    slot (``NOT_THEME_COLOR`` / ``MIXED`` -> ``''``) contributes nothing. So the
    palette keys pptx writes are canonical clrScheme slots, exactly the namespace the
    verify check's direct-slot branch validates - no docx WML re-keying.

Capture is model-free and deterministic: ``iter_run_facts`` is a SINGLE ordered pass
(slides -> shapes -> paragraphs -> runs, the deck's own order), never a set, so the
engine's ``Counter`` tie-breaks stay stable. Most brand color in a real deck lives in
layout/master placeholders that ``run.font.color`` reports as inherited (``None``);
the ``seed_theme_palette`` floor remains the primary palette source for pptx, and this
pass adds whatever DIRECT run typography the deck actually carries on top of it.
"""

from __future__ import annotations

from typing import Iterator, Optional

from pptx.enum.dml import MSO_COLOR_TYPE

from brandkit.common.color import THEME_SLOTS

# The WordprocessingML text/background aliases python-pptx's ``MSO_THEME_COLOR``
# emits for the four text/background slots, mapped to the canonical clrScheme slot.
# This is the SAME table ``color.resolve_theme_color`` carries (single behavior, no
# new lexicon); every other ``xml_value`` token is already a THEME_SLOTS slot.
_PPTX_TOKEN_ALIAS: dict[str, str] = {
    "tx1": "dk1",
    "bg1": "lt1",
    "tx2": "dk2",
    "bg2": "lt2",
}

# The set of THEME_SLOTS membership (frozenset for the O(1) guard below).
_SLOTS: frozenset[str] = frozenset(THEME_SLOTS)


def _normalize_theme_token(token: Optional[str]) -> Optional[str]:
    """Normalize a pptx ``MSO_THEME_COLOR.xml_value`` to a canonical clrScheme slot,
    or ``None`` when it is not a real brand slot.

    ``tx1`` / ``bg1`` / ``tx2`` / ``bg2`` map to ``dk1`` / ``lt1`` / ``dk2`` / ``lt2``
    (the same aliasing ``resolve_theme_color`` does); ``accent1-6`` / ``dk*`` / ``lt*``
    / ``hlink`` / ``folHlink`` pass through unchanged. ``NOT_THEME_COLOR`` / ``MIXED``
    surface as ``''`` and yield ``None`` - they are not verifiable/appliable brand
    tokens, so they must never enter the profile (keeps capture/verify in sync)."""
    if not token:
        return None
    slot = _PPTX_TOKEN_ALIAS.get(token, token)
    return slot if slot in _SLOTS else None


def _run_size_hp(run) -> Optional[int]:
    """The run's EXPLICIT size as half-points, or ``None`` when it inherits.

    ``run.font.size`` is an explicit-only ``Length`` (``None`` when inherited). The
    half-point bucket ``round(pt * 2)`` matches the engine's docx unit. A malformed
    measure contributes nothing rather than crashing the extraction."""
    try:
        size = run.font.size
        if size is None:
            return None
        return round(size.pt * 2)
    except Exception:
        return None


def _run_color(run) -> Optional[tuple[str, ...]]:
    """The run's EXPLICIT color as a hashable bucket key, or ``None``.

    An RGB color buckets as ``('hex', <RRGGBB>)``; a SCHEME (theme) color buckets as
    ``('theme', <slot>)`` with the token normalized to a canonical clrScheme slot via
    :func:`_normalize_theme_token`. An inherited color (``type is None``), an unmapped
    theme token, or any non-RGB/non-SCHEME color contributes nothing - it is not a
    captured brand value."""
    try:
        color = run.font.color
        ctype = color.type
        if ctype == MSO_COLOR_TYPE.RGB and color.rgb is not None:
            return ("hex", str(color.rgb))
        if ctype == MSO_COLOR_TYPE.SCHEME:
            token = getattr(color.theme_color, "xml_value", None)
            slot = _normalize_theme_token(token)
            if slot is not None:
                return ("theme", slot)
    except Exception:
        return None
    return None


class _PptxRunFact:
    """A pptx run reduced to the structural ``RunFacts`` view the engine consumes.

    pptx runs carry no per-run named-style identity that maps to a placeholder role
    resolver, so ``style_key`` is ``None``: the run votes toward the document body
    (and the palette), never a specific role's per-style fold. The body capture is the
    primary pptx appearance source."""

    __slots__ = ("style_key", "text", "font_name", "size_hp", "color", "is_link")

    def __init__(self, run) -> None:
        self.style_key = None
        self.text = run.text or ""
        self.font_name = run.font.name or None
        self.size_hp = _run_size_hp(run)
        self.color = _run_color(run)
        try:
            self.is_link = bool(run.hyperlink.address)
        except Exception:
            self.is_link = False


def iter_run_facts(prs) -> Iterator[_PptxRunFact]:
    """Yield one :class:`_PptxRunFact` per run across the deck, in a SINGLE ordered
    pass (slides -> text-frame shapes -> paragraphs -> runs, the deck's own order).

    A non-text shape (no ``text_frame``) contributes nothing; a shape python-pptx
    cannot introspect is skipped crash-safe (capture must never crash the extraction).
    Ordered, never a set, so the engine's ``Counter`` tie-breaks stay deterministic."""
    for slide in prs.slides:
        for shape in slide.shapes:
            try:
                if not shape.has_text_frame:
                    continue
                paragraphs = shape.text_frame.paragraphs
            except Exception:
                continue
            for paragraph in paragraphs:
                try:
                    runs = paragraph.runs
                except Exception:
                    continue
                for run in runs:
                    yield _PptxRunFact(run)
