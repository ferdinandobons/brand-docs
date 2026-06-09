# SPDX-License-Identifier: MIT
"""XLSX brand typography capture: the run-fact view the shared engine consumes.

This is the XLSX peer of ``formats/docx/typography.py``: it walks the workbook's
styled cells and yields the structural
:class:`~brandkit.common.typography.RunFacts` view so the format-neutral capture
engine (``capture_appearance`` / ``capture_palette_facts``) fills the SAME profile
shape for xlsx that it fills for docx/pptx - the document body font/size
(``theme.fonts.body``), the body color (``theme.text.body``), the per-role
``appearance``, and the brand ``theme.palette``.

A cell IS the single run for capture: a styled spreadsheet has no run nesting, so
each non-empty STYLED cell contributes exactly ONE fact (mirroring the artifact-catalog
text rule - the cell's literal value is the captured text). The ``has_style`` gate
matters: openpyxl SYNTHESIZES a default ``Calibri`` / ``theme=1`` font on a cell that
carries no style at all, so an UNSTYLED cell would otherwise vote a phantom default
font/color that is not a brand fact. Skipping ``has_style is False`` cells keeps
capture to OBSERVED styling only (the brand guarantee: facts, never library
defaults). Two xlsx specifics, both UNIVERSAL (driven off captured facts /
:data:`~brandkit.common.color.THEME_SLOTS`, never a tuned slot):

  - **Unit:** openpyxl exposes a cell font size as a point float ``cell.font.sz``;
    the half-point bucket ``round(sz * 2)`` matches the engine's docx unit so a body
    size captured from a workbook compares in the same space as a doc/deck.
  - **Token namespace:** an openpyxl theme color is an INTEGER index into the
    SpreadsheetML cell theme-color order, which is NOT clrScheme document order:
    Excel swaps the first two dark/light pairs (``0=lt1 1=dk1 2=lt2 3=dk2``, then
    ``4..=accent1..6 hlink folHlink``; ECMA-376 / MS-OI29500). The index is therefore
    mapped through :data:`_XLSX_THEME_INDEX` (NOT a raw
    :data:`~brandkit.common.color.THEME_SLOTS` lookup, which would mislabel a default
    ``theme=1`` text color as ``lt1`` and brand body text white). The openpyxl
    ``.tint`` transform is dropped - the BASE slot is recorded, matching docx's
    transform-blind buckets (a known parity simplification, not a bug). An RGB color
    is an 8-digit ARGB string whose leading alpha pair is stripped before normalizing
    to ``RRGGBB``. An indexed / auto / out-of-range color contributes nothing - it is
    not a verifiable/appliable brand value.

Capture is model-free and deterministic: ``iter_run_facts`` is a SINGLE ordered pass
over each worksheet's materialized cells (never a set), so the engine's ``Counter``
tie-breaks stay stable. As with pptx, the ``seed_theme_palette`` floor remains the
primary palette source; this pass folds whatever DIRECT cell font color the workbook
carries on top of it.
"""

from __future__ import annotations

from typing import Iterator, Optional

from brandkit.common import color as colorutil
from brandkit.common.color import THEME_SLOTS

# SpreadsheetML cell theme-color index -> clrScheme slot. Excel's <color theme="N">
# index is NOT clrScheme document order (THEME_SLOTS): it swaps the first two
# dark/light pairs (0=lt1, 1=dk1, 2=lt2, 3=dk2), then 4..=accent1..6, hlink, folHlink
# (ECMA-376 / MS-OI29500). Indexing THEME_SLOTS directly would capture the default
# text color (theme=1) as lt1 and brand body text white. Derived from THEME_SLOTS so
# the two never drift.
_XLSX_THEME_INDEX: tuple[str, ...] = (
    THEME_SLOTS[1],  # 0 -> lt1 (Background 1)
    THEME_SLOTS[0],  # 1 -> dk1 (Text 1)
    THEME_SLOTS[3],  # 2 -> lt2 (Background 2)
    THEME_SLOTS[2],  # 3 -> dk2 (Text 2)
    *THEME_SLOTS[4:],  # 4.. -> accent1..6, hlink, folHlink (unchanged)
)


def _cell_size_hp(cell) -> Optional[int]:
    """The cell font's EXPLICIT size as half-points, or ``None`` when absent.

    ``cell.font.sz`` is the point size; the half-point bucket ``round(sz * 2)``
    matches the engine's docx unit. A missing / malformed size contributes nothing."""
    try:
        sz = cell.font.sz
        if sz is None:
            return None
        return round(float(sz) * 2)
    except Exception:
        return None


def _cell_color(cell) -> Optional[tuple[str, ...]]:
    """The cell font's color as a hashable bucket key, or ``None``.

    An ``'rgb'`` color buckets as ``('hex', <RRGGBB>)`` - the 8-digit ARGB string's
    leading alpha pair is stripped and the remainder normalized. A ``'theme'`` color
    buckets as ``('theme', <slot>)`` by mapping the openpyxl theme index through
    :data:`_XLSX_THEME_INDEX` (the Excel cell theme-index order, NOT raw
    :data:`THEME_SLOTS` - see the constant; the ``.tint`` is dropped: the base slot is
    recorded). A ``None`` / indexed / auto / out-of-range color contributes nothing
    (not a captured brand value); a malformed color contributes nothing rather than
    crashing the pass."""
    try:
        font = cell.font
        c = getattr(font, "color", None)
        if c is None:
            return None
        ctype = getattr(c, "type", None)
        if ctype == "rgb":
            rgb = c.rgb
            if not isinstance(rgb, str):
                return None
            hexpart = rgb[2:] if len(rgb) == 8 else rgb
            return ("hex", colorutil.normalize_hex(hexpart))
        if ctype == "theme":
            idx = c.theme
            if isinstance(idx, int) and 0 <= idx < len(_XLSX_THEME_INDEX):
                return ("theme", _XLSX_THEME_INDEX[idx])
    except Exception:
        return None
    return None


class _XlsxRunFact:
    """A styled cell reduced to the structural ``RunFacts`` view the engine consumes.

    The cell is its own single run. ``style_key`` is ``(named_style, named_style)``
    so a future named-style role fold can match by name; today the xlsx roles use
    named-range / cell-style / number-format resolvers (no ``named_style`` type), so
    the engine's default ``role_style_key`` yields ``None`` for them and the body
    capture is the primary xlsx appearance source."""

    __slots__ = ("style_key", "text", "font_name", "size_hp", "color", "is_link")

    def __init__(self, cell) -> None:
        try:
            style = cell.style
        except Exception:
            style = None
        self.style_key = (style, style)
        self.text = str(cell.value)
        self.font_name = (cell.font.name or None) if cell.font else None
        self.size_hp = _cell_size_hp(cell)
        self.color = _cell_color(cell)
        self.is_link = cell.hyperlink is not None


def iter_run_facts(wb) -> Iterator[_XlsxRunFact]:
    """Yield one :class:`_XlsxRunFact` per non-empty cell across the workbook, in a
    SINGLE ordered pass over each worksheet's materialized cells.

    Iterates ``ws._cells.values()`` (materialized cells only - large corporate models
    are sparse over broad dimensions) and skips a cell with no value OR no style (an
    unstyled cell carries only openpyxl's synthetic default font/color, not a brand
    fact - see the ``has_style`` note in the module docstring). A cell openpyxl cannot
    introspect is skipped crash-safe. Ordered, never a set, so the engine's
    ``Counter`` tie-breaks stay deterministic."""
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            try:
                if cell.value is None or not cell.has_style:
                    continue
            except Exception:
                continue
            yield _XlsxRunFact(cell)
