"""ファイル名プレビューの生成（読み取り専用。書き出し・名前変更は一切行わない）.

FR-14 / FR-M13 / FR-R11 の判定ロジック（`[既存]` / `[重複]` / `[読込不可]` /
`[変更不要]` / `[衝突]`）はすべてこのファイルに置き、
GUI（pdf_split_merge_rename.py）は受け取った行を表示するだけにする。
ファイル名生成は実行時と同一の関数（modes.build_filenames）を共有する。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Sequence, TypeVar

import merge_core
import rename_core
from modes import build_filenames
from split_core import load_pdf

MAX_ANNOUNCED = 100  # 一覧にそのまま並べる件数。超分は「ほか N 件」にまとめる

Row = TypeVar("Row")  # 行種別（PreviewRow / MergePreviewRow / RenamePreviewRow）を問わず使える


@dataclass(frozen=True)
class PreviewRow:
    """プレビュー1行. exists=出力先に既存, duplicate=大文字小文字無視で同一名が複数."""

    page_no: int
    filename: str
    exists: bool
    duplicate: bool

    @property
    def status(self) -> str:
        """表示用の状態文字列（`[既存]` `[重複]`）."""
        marks = []
        if self.exists:
            marks.append("[既存]")
        if self.duplicate:
            marks.append("[重複]")
        return " ".join(marks)

    def display_cells(self) -> tuple[str, str, str]:
        """Treeview へ渡す (ページ番号, ファイル名, 状態) タプル."""
        return (str(self.page_no), self.filename, self.status)


def truncate_display(rows: Sequence[Row]) -> tuple[list[Row], int]:
    """表示用の先頭 MAX_ANNOUNCED 件と、まとめられた「ほか」件数を返す.

    rows（データ）自体は全件保持され、この関数は書き出しや破壊的変更をしない。
    行種別（分割 / 結合 / リネーム）を問わず同じ件数でまとめる。
    """
    if len(rows) <= MAX_ANNOUNCED:
        return rows, 0
    return rows[:MAX_ANNOUNCED], len(rows) - MAX_ANNOUNCED


def build_preview(
    excel_path: Path,
    sheet_name: str,
    cell: str,
    mode: str,
    output_dir: Path,
    pdf_path: Path,
) -> list[PreviewRow]:
    """現在の入力からプレビュー行を全件組み立てる.

    PDFとExcelを読むだけで、ファイルの書き出し・既存ファイルの変更は行わない。
    FR-6 相当の不正（空欄・不正文字・予約名・数式キャッシュ無し）は
    ValueError をそのまま伝搬する（GUI のエラー欄に表示される）。
    """
    reader = load_pdf(pdf_path)
    # 実行時と同一の生成関数を使う。プレビューでは重複をエラーにせず、下の重複判定で `[重複]` 印を付ける。
    filenames = build_filenames(mode, excel_path, sheet_name, cell, len(reader.pages), abort_on_duplicates=False)

    counts: dict[str, int] = {}
    for filename in filenames:
        key = filename.casefold()
        counts[key] = counts.get(key, 0) + 1

    return [
        PreviewRow(
            page_no=page_no,
            filename=filename,
            exists=(output_dir / filename).exists(),
            duplicate=counts[filename.casefold()] > 1,
        )
        for page_no, filename in enumerate(filenames, start=1)
    ]


@dataclass(frozen=True)
class MergePreviewRow:
    """結合プレビュー1行. exists=出力先に結合後ファイル名が既存, unreadable=開けないPDF."""

    order: int
    filename: str
    page_count: int
    exists: bool = False
    unreadable: bool = False

    @property
    def status(self) -> str:
        """表示用の状態文字列（`[既存]` `[読込不可]`）."""
        marks = []
        if self.exists:
            marks.append("[既存]")
        if self.unreadable:
            marks.append("[読込不可]")
        return " ".join(marks)

    def display_cells(self) -> tuple[str, str, str, str]:
        """Treeview へ渡す (順序, 入力PDF, ページ数, 状態) タプル."""
        pages = "-" if self.unreadable else str(self.page_count)
        return (str(self.order), self.filename, pages, self.status)


@dataclass(frozen=True)
class RenamePreviewRow:
    """リネームプレビュー1行. unchanged=現在名と同じ, conflict=対象以外の既存名と衝突."""

    order: int
    old_name: str
    new_name: str
    unchanged: bool = False
    conflict: bool = False
    duplicate: bool = False

    @property
    def status(self) -> str:
        """表示用の状態文字列（`[変更不要]` `[衝突]` `[重複]`）."""
        marks = []
        if self.unchanged:
            marks.append("[変更不要]")
        if self.conflict:
            marks.append("[衝突]")
        if self.duplicate:
            marks.append("[重複]")
        return " ".join(marks)

    def display_cells(self) -> tuple[str, str, str, str]:
        """Treeview へ渡す (順序, 現在のファイル名, 新しいファイル名, 状態) タプル."""
        return (str(self.order), self.old_name, self.new_name, self.status)


def build_merge_preview(folder: Path, out_name: str, output_dir: Path) -> list[MergePreviewRow]:
    """結合プレビュー行を全件組み立てる（PDFを開くだけで書き出しはしない）.

    FR-M13 の「合計 N ファイル / M ページ → 出力名」は GUI がこの行から作る。
    開けないPDFは例外にせず `[読込不可]` 行として残す（実行時は merge_core が中止する）。
    """
    sources = merge_core.list_merge_sources(folder)
    name = merge_core.validate_output_name(out_name)
    output_exists = output_dir.is_dir() and (output_dir / name).exists()

    rows: list[MergePreviewRow] = []
    for order, source in enumerate(sources, start=1):
        try:
            reader = load_pdf(source)
        except ValueError:
            rows.append(MergePreviewRow(order, source.name, 0, exists=output_exists, unreadable=True))
            continue
        try:
            page_count = len(reader.pages)
        finally:
            try:
                reader.close()
            except Exception:
                pass
        rows.append(MergePreviewRow(order, source.name, page_count, exists=output_exists))
    return rows


def build_rename_preview(folder: Path, excel_path: Path, sheet_name: str, cell: str, mode: str) -> list[RenamePreviewRow]:
    """リネームプレビュー行を全件組み立てる（名前変更は一切しない）.

    実行時と同一の rename_core.build_rename_plan を使い、プレビューでは
    重複をエラーにせず `[重複]` 印を付ける（FR-R11）。
    """
    items = rename_core.build_rename_plan(
        folder, excel_path, sheet_name, cell, mode, abort_on_duplicates=False,
    )
    conflict_names = {name.casefold() for name in rename_core.find_rename_conflicts(items, folder)}

    counts: dict[str, int] = {}
    for item in items:
        key = item.new_name.casefold()
        counts[key] = counts.get(key, 0) + 1

    return [
        RenamePreviewRow(
            order=item.order,
            old_name=item.source.name,
            new_name=item.new_name,
            unchanged=item.unchanged,
            conflict=item.new_name.casefold() in conflict_names,
            duplicate=counts[item.new_name.casefold()] > 1,
        )
        for item in items
    ]
