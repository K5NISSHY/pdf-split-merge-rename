"""ファイル名の自然順（数字の連なりは数値として比較）を提供する.

Explorer の名前順では `10.pdf` が `9.pdf` より前に来てしまうため、
名前を「数字の連なり」と「それ以外」に分解して比較キーを作る
（docs/PROJECT_DOCUMENT.md 1章「自然順」）。

- 数字連なり → `(1, 数値, "")`
- それ以外のテキスト → `(0, 0, 小文字化テキスト)`
- 末尾に全文の casefold と原文を置き、同順位でも決定的に揃える
- 桁数が極端に大きい数字連なりは整数化せず、文字列比較へフォールバックする（7章の留意点）
"""

from __future__ import annotations

import re
from typing import Iterable

# これを超える桁数の数字連なりは整数にせず、文字列比較へフォールバックする
MAX_DIGITS_AS_INT = 18

_DIGIT_RUN = re.compile(r"[0-9]+")


def natural_sort_key(name: str) -> tuple:
    """名前の自然順比較キーを返す（すべての要素は (種別, 数値, テキスト) の3つ組）."""
    blocks: list[tuple[int, int, str]] = []
    position = 0
    for match in _DIGIT_RUN.finditer(name):
        if match.start() > position:
            blocks.append((0, 0, name[position:match.start()].casefold()))
        run = match.group()
        if len(run) <= MAX_DIGITS_AS_INT:
            blocks.append((1, int(run), ""))
        else:
            blocks.append((0, 0, run))  # 極端に長い数字連なりは文字列比較へ
        position = match.end()
    if position < len(name):
        blocks.append((0, 0, name[position:].casefold()))
    blocks.append((2, 0, name.casefold()))  # 全文の大小無視（同順位を決める）
    blocks.append((3, 0, name))  # 原文の文字コード順（決定的な最終キー）
    return tuple(blocks)


def sort_pdf_names(names: Iterable[str]) -> list[str]:
    """.pdf（大文字小文字無視）だけを自然順に並べたリストを返す."""
    pdf_names = [name for name in names if name.lower().endswith(".pdf")]
    return sorted(pdf_names, key=natural_sort_key)
