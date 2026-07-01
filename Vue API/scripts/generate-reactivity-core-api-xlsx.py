#!/usr/bin/env python3
"""Generate reactivity-core-api.xlsx from reactivity-core-api.md."""

from __future__ import annotations

from md_xlsx_common import ROOT, generate_xlsx

MD_PATH = ROOT / "reactivity-core-api.md"
OUTPUT_PATH = ROOT / "reactivity-core-api.xlsx"

STRUCTURE: list[str | tuple[str, str]] = [
    ("section", "Базовые реактивные примитивы"),
    "ref()",
    "computed()",
    "reactive()",
    "readonly()",
    ("section", "Отслеживание изменений (watchers)"),
    "watchEffect()",
    "watchPostEffect()",
    "watchSyncEffect()",
    "watch()",
    "onWatcherCleanup()",
]

CATEGORY_BY_API: dict[str, str] = {
    "ref()": "Примитивная реактивность",
    "computed()": "Вычисляемые значения",
    "reactive()": "Реактивный объект",
    "readonly()": "Только чтение",
    "watchEffect()": "Автоматическое отслеживание зависимостей",
    "watchPostEffect()": "watchEffect после render",
    "watchSyncEffect()": "Синхронный watchEffect",
    "watch()": "Явное отслеживание источников",
    "onWatcherCleanup()": "Cleanup watcher (3.5+)",
}


def main() -> int:
    return generate_xlsx(
        md_path=MD_PATH,
        output_path=OUTPUT_PATH,
        structure=STRUCTURE,
        categories=CATEGORY_BY_API,
        sheet_title="Reactivity Core",
    )


if __name__ == "__main__":
    raise SystemExit(main())
