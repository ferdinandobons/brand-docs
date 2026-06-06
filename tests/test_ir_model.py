# SPDX-License-Identifier: MIT
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.common import text as textutil
from brandkit.ir import model as ir


def test_list_items_accept_plain_string_shortcut() -> None:
    doc = ir.parse_idoc(
        {
            "blocks": [
                {
                    "type": "list",
                    "items": [
                        "Plain bullet",
                        {"text": "Nested parent", "items": ["Nested child"]},
                    ],
                }
            ]
        }
    )

    block = doc.blocks[0]
    assert isinstance(block, ir.ListBlock)
    assert textutil.runs_to_text(block.items[0].runs) == "Plain bullet"
    assert block.items[0].level == 0
    assert textutil.runs_to_text(block.items[1].items[0].runs) == "Nested child"
    assert block.items[1].items[0].level == 1
