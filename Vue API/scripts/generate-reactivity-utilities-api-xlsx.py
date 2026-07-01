#!/usr/bin/env python3
"""Generate reactivity-utilities-api.xlsx from reactivity-utilities-api.md."""

from __future__ import annotations

import sys

from md_xlsx_common import ROOT, generate_xlsx

MD_PATH = ROOT / "reactivity-utilities-api.md"
OUTPUT_PATH = ROOT / "reactivity-utilities-api.xlsx"

STRUCTURE: list[str | tuple[str, str]] = [
    ("section", "Нормализация ref и value"),
    "isRef()",
    "unref()",
    "toRef()",
    "toValue()",
    "toRefs()",
    ("section", "Проверки proxy"),
    "isProxy()",
    "isReactive()",
    "isReadonly()",
]

CATEGORY_BY_API: dict[str, str] = {
    "isRef()": "Проверка ref",
    "unref()": "Разворачивание ref",
    "toRef()": "Создание/нормализация ref",
    "toValue()": "Нормализация value/ref/getter (3.3+)",
    "toRefs()": "Деструктуризация reactive",
    "isProxy()": "Проверка proxy",
    "isReactive()": "Проверка reactive proxy",
    "isReadonly()": "Проверка readonly",
}


def main() -> int:
    return generate_xlsx(
        md_path=MD_PATH,
        output_path=OUTPUT_PATH,
        structure=STRUCTURE,
        categories=CATEGORY_BY_API,
        sheet_title="Reactivity Utilities",
    )


if __name__ == "__main__":
    raise SystemExit(main())
