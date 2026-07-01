#!/usr/bin/env python3
"""Generate composition-api-setup.xlsx from composition-api-setup.md."""

from __future__ import annotations

import sys

from md_xlsx_common import ROOT, generate_xlsx

MD_PATH = ROOT / "composition-api-setup.md"
OUTPUT_PATH = ROOT / "composition-api-setup.xlsx"

STRUCTURE: list[str | tuple[str, str]] = [
    ("section", "Основы setup()"),
    "setup()",
    "setup(props)",
    "toRefs(props) / toRef(props, key)",
    "async setup()",
    ("section", "Setup Context"),
    "setup(props, context)",
    "context.attrs",
    "context.slots",
    "context.emit",
    "context.expose()",
    ("section", "Render functions"),
    "setup() → render function",
]

CATEGORY_BY_API: dict[str, str] = {
    "setup()": "Точка входа Composition API",
    "setup(props)": "Доступ к props",
    "toRefs(props) / toRef(props, key)": "Деструктуризация props",
    "async setup()": "Асинхронный setup (Suspense)",
    "setup(props, context)": "Setup Context",
    "context.attrs": "Fallthrough attributes",
    "context.slots": "Слоты",
    "context.emit": "Emit событий",
    "context.expose()": "Expose для template refs",
    "setup() → render function": "Render function из setup",
}


def main() -> int:
    return generate_xlsx(
        md_path=MD_PATH,
        output_path=OUTPUT_PATH,
        structure=STRUCTURE,
        categories=CATEGORY_BY_API,
        sheet_title="Composition setup()",
    )


if __name__ == "__main__":
    raise SystemExit(main())
