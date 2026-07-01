#!/usr/bin/env python3
"""Generate lifecycle-hooks-api.xlsx from lifecycle-hooks-api.md."""

from __future__ import annotations

import sys
from pathlib import Path

from md_xlsx_common import ROOT, generate_xlsx

MD_PATH = ROOT / "lifecycle-hooks-api.md"
OUTPUT_PATH = ROOT / "lifecycle-hooks-api.xlsx"

STRUCTURE: list[str | tuple[str, str]] = [
    ("section", "Composition API"),
    "onBeforeMount()",
    "onMounted()",
    "onBeforeUpdate()",
    "onUpdated()",
    "onBeforeUnmount()",
    "onUnmounted()",
    "onErrorCaptured()",
    "onRenderTracked()",
    "onRenderTriggered()",
    "onActivated()",
    "onDeactivated()",
    "onServerPrefetch()",
    ("section", "Options API"),
    "beforeCreate",
    "created",
    "beforeMount",
    "mounted",
    "beforeUpdate",
    "updated",
    "beforeUnmount",
    "unmounted",
    "errorCaptured",
    "renderTracked",
    "renderTriggered",
    "activated",
    "deactivated",
    "serverPrefetch",
]

CATEGORY_BY_API: dict[str, str] = {
    "onBeforeMount()": "Монтирование (Composition API)",
    "onMounted()": "Монтирование (Composition API)",
    "onBeforeUpdate()": "Обновление DOM (Composition API)",
    "onUpdated()": "Обновление DOM (Composition API)",
    "onBeforeUnmount()": "Размонтирование (Composition API)",
    "onUnmounted()": "Размонтирование (Composition API)",
    "onErrorCaptured()": "Обработка ошибок (Composition API)",
    "onRenderTracked()": "Отладка реактивности (Composition API, dev)",
    "onRenderTriggered()": "Отладка реактивности (Composition API, dev)",
    "onActivated()": "KeepAlive (Composition API)",
    "onDeactivated()": "KeepAlive (Composition API)",
    "onServerPrefetch()": "SSR (Composition API)",
    "beforeCreate": "Создание экземпляра (Options API)",
    "created": "Создание экземпляра (Options API)",
    "beforeMount": "Монтирование (Options API)",
    "mounted": "Монтирование (Options API)",
    "beforeUpdate": "Обновление DOM (Options API)",
    "updated": "Обновление DOM (Options API)",
    "beforeUnmount": "Размонтирование (Options API)",
    "unmounted": "Размонтирование (Options API)",
    "errorCaptured": "Обработка ошибок (Options API)",
    "renderTracked": "Отладка реактивности (Options API, dev)",
    "renderTriggered": "Отладка реактивности (Options API, dev)",
    "activated": "KeepAlive (Options API)",
    "deactivated": "KeepAlive (Options API)",
    "serverPrefetch": "SSR (Options API)",
}


def main() -> int:
    return generate_xlsx(
        md_path=MD_PATH,
        output_path=OUTPUT_PATH,
        structure=STRUCTURE,
        categories=CATEGORY_BY_API,
        sheet_title="Lifecycle Hooks",
    )


if __name__ == "__main__":
    raise SystemExit(main())
