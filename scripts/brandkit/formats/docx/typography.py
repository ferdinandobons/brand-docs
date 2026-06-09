# SPDX-License-Identifier: MIT
"""DOCX brand typography capture (font family).

The brand's REAL visible fonts often live as DIRECT run-level formatting
(``w:rPr/w:rFonts``) on the template's content rather than in the named styles or
the theme: a designed template may put everything in ``Normal`` with a direct
Roboto / Montserrat override. Role inference (``roles.py``) and theme extraction
read only named styles and ``theme1.xml``, so those fonts are never captured and a
generated document falls back to the ``docDefaults`` font (typically Arial).

This module captures the DOMINANT direct run font family, deterministically:

  - per role: the dominant explicit font among the runs that use the role's style
    -> ``role['appearance']['font'] = {'latin': <name>}``;
  - the document's effective body font: the dominant explicit font across all body
    runs -> ``theme['fonts']['body'] = {'latin': <name>}``, the fallback the
    generator applies to a paragraph whose role carries no captured font.

Only a clear DOMINANT is recorded (at least :data:`_MIN_RUNS` explicitly-fonted
runs and a winner covering at least :data:`_MIN_DOMINANCE` of them), with its
dominance stored as ``confidence``. Capture is deterministic (model-free).

The brand guarantee is preserved: a captured font is a FACT observed in the
template, stored in the profile, applied only via the resolver, and re-validated
against the shell's available fonts by ``check_appearance_targets`` (fail-closed).
This module is purely additive - it only populates the already-reserved
``appearance`` field and an additive ``theme.fonts.body`` key; a template with no
dominant direct font leaves both untouched, so behavior is unchanged.
"""

from __future__ import annotations

from collections import Counter
from typing import Optional

from brandkit.profile import schema

# A capture is only trusted when it is a clear convention, not noise.
_MIN_RUNS = 3  # need at least this many explicitly-fonted runs to call a winner
_MIN_DOMINANCE = 0.6  # the winner must cover >= 60% of those runs


def _dominant(counter: Counter) -> Optional[tuple[str, float]]:
    """Return ``(font, dominance)`` for the most common font when it is a clear
    convention, else ``None``."""
    total = sum(counter.values())
    if total < _MIN_RUNS:
        return None
    font, n = counter.most_common(1)[0]
    ratio = n / total
    if ratio < _MIN_DOMINANCE:
        return None
    return font, ratio


def capture_fonts(doc, roles: dict, theme: dict) -> None:
    """Capture dominant direct run fonts into ``roles`` (per role ``appearance``)
    and ``theme['fonts']['body']`` (the document default), mutating both in place.

    Reads only the explicit run font (``run.font.name``); a run that inherits its
    font from the style/theme contributes nothing. python-docx resolves a
    paragraph's effective style (a paragraph with no explicit ``pStyle`` reports the
    document's default style), so runs are bucketed by their real style id/name.
    """
    per_style: dict[tuple[Optional[str], Optional[str]], Counter] = {}
    overall: Counter = Counter()

    for para in doc.paragraphs:
        try:
            style = para.style
            sid = getattr(style, "style_id", None) if style is not None else None
            sname = getattr(style, "name", None) if style is not None else None
        except Exception:
            sid = sname = None
        for run in para.runs:
            if not (run.text or "").strip():
                continue
            font = run.font.name  # explicit ascii/hAnsi typeface, or None if inherited
            if not font:
                continue
            overall[font] += 1
            if sid or sname:
                per_style.setdefault((sid, sname), Counter())[font] += 1

    body = _dominant(overall)
    if body is not None:
        fonts = theme.setdefault("fonts", {})
        fonts["body"] = {"latin": body[0], "confidence": round(body[1], 3)}

    for rid, entry in roles.items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        resolver = entry.get("resolver") or {}
        if resolver.get("type") != schema.ResolverType.NAMED_STYLE.value:
            continue
        sid = resolver.get("style_id")
        sname = resolver.get("style_name")
        counter: Counter = Counter()
        for (k_sid, k_sname), c in per_style.items():
            if (sid and k_sid == sid) or (sname and k_sname == sname):
                counter.update(c)
        dom = _dominant(counter)
        if dom is not None:
            appearance = entry.setdefault("appearance", {})
            appearance["font"] = {"latin": dom[0]}
            appearance["confidence"] = round(dom[1], 3)
