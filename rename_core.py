"""ExcelでPDFをリネームするコア処理（FR-R1〜FR-R13）.

入力フォルダーの直下 `.pdf` を自然順に並べ、既存の命名関数
（`modes.build_filenames`）へ「ページ数」の代わりに「ファイル数」を渡して
新名を作り、同じ順序で割り当てる。実行は2段階リネーム
（①同一フォルダー内の一時名 `.pdf_rename_tmp_<連番>.tmp` へ ②最終名へ）で行い、
どちらの段階で失敗しても変更済みのファイルをすべて元の名前へ戻してからエラーを出す。
PDF・Excelの内容は一切変更せず、名前だけを変える。
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import modes
from merge_core import TEMP_NAME_PREFIXES
from natural_sort import sort_pdf_names

TEMP_RENAME_PREFIX = ".pdf_rename_tmp_"  # 2段階リネームの一時名（FR-R9）
MAX_LISTED = 8  # エラー一覧にそのまま並べる件数。超分は「ほか N 件」にまとめる


@dataclass(frozen=True)
class RenameItem:
    """リネーム1件（order は自然順の連番、unchanged は現在名と新名が同じ）."""

    order: int
    source: Path
    new_name: str
    unchanged: bool


@dataclass(frozen=True)
class RenameResult:
    """リネーム実行結果（実際に名前を変えた件数と変更不要だった件数）."""

    renamed: int
    unchanged: int


def list_rename_targets(folder: Path) -> list[Path]:
    """入力フォルダー直下の `.pdf` を自然順で返す（0件は ValueError）."""
    if not folder.is_dir():
        raise ValueError("入力フォルダーが見つかりません。")
    names = [
        path.name
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf" and not path.name.startswith(TEMP_NAME_PREFIXES)
    ]
    ordered = sort_pdf_names(names)
    if not ordered:
        raise ValueError("PDFが見つかりません。入力フォルダーの直下に .pdf ファイルを置いてください。")
    return [folder / name for name in ordered]


def build_rename_plan(
    folder: Path,
    excel_path: Path,
    sheet_name: str,
    cell: str,
    mode: str,
    *,
    abort_on_duplicates: bool = True,
) -> list[RenameItem]:
    """自然順の対象ファイルへ、Excel由来の新名を同じ順序で対応付ける.

    生成個数にはページ数ではなく対象ファイル数を渡す（FR-R2）。
    プレビューのみ abort_on_duplicates=False を渡して重複行を残す。
    """
    targets = list_rename_targets(folder)
    filenames = modes.build_filenames(
        mode, excel_path, sheet_name, cell, len(targets), abort_on_duplicates=abort_on_duplicates,
    )
    return [
        RenameItem(
            order=order,
            source=source,
            new_name=new_name,
            unchanged=source.name.casefold() == new_name.casefold(),
        )
        for order, (source, new_name) in enumerate(zip(targets, filenames), start=1)
    ]


def find_rename_conflicts(items: list[RenameItem], folder: Path) -> list[str]:
    """リネーム対象以外の既存ファイル（非PDF含む）と衝突する新名を返す（読み取りのみ）."""
    target_names = {item.source.name.casefold() for item in items}
    conflicts: list[str] = []
    reported: set[str] = set()
    for item in items:
        key = item.new_name.casefold()
        if key in target_names or key in reported:
            continue
        if (folder / item.new_name).exists():
            conflicts.append(item.new_name)
            reported.add(key)
    return conflicts


def check_rename_plan(items: list[RenameItem], folder: Path) -> None:
    """衝突があれば名前変更を1つも始める前に FileExistsError で中止する（FR-R7/FR-R10）."""
    conflicts = find_rename_conflicts(items, folder)
    if conflicts:
        preview = "\n".join(conflicts[:MAX_LISTED])
        extra = f"\nほか {len(conflicts) - MAX_LISTED} 件" if len(conflicts) > MAX_LISTED else ""
        raise FileExistsError(
            "新しいファイル名が既存ファイル（リネーム対象以外）と衝突します。上書きは行いません。\n" + preview + extra
        )


def apply_renames(items: list[RenameItem]) -> RenameResult:
    """2段階リネームを実行する。失敗時は元の名前へ戻してから ValueError（FR-R9）."""
    targets = [item for item in items if not item.unchanged]
    unchanged_count = len(items) - len(targets)
    if not targets:
        return RenameResult(renamed=0, unchanged=unchanged_count)

    folder = targets[0].source.parent
    temp_paths = {
        item.source: folder / f"{TEMP_RENAME_PREFIX}{order}.tmp"
        for order, item in enumerate(targets, start=1)
    }
    current: dict[RenameItem, Path] = {}
    rollback_ok = True
    try:
        for item in targets:  # ① 一時名へ
            temp = temp_paths[item.source]
            os.rename(item.source, temp)
            current[item] = temp
        for item in targets:  # ② 最終名へ
            final = folder / item.new_name
            os.rename(current[item], final)
            current[item] = final
    except Exception as exc:
        for item in targets:  # 変更済みだけを元の名前へ戻す
            path = current.get(item)
            if path is None or path == item.source:
                continue
            try:
                os.rename(path, item.source)
            except OSError:
                rollback_ok = False
        raise ValueError(f"ファイル名を変更できませんでした。変更済みの名前は元に戻しました。\n{exc}") from exc
    finally:
        if rollback_ok:  # 一時名の残骸を掃除する（戻せなかったファイルを消さない）
            for temp in temp_paths.values():
                if temp.exists():
                    try:
                        temp.unlink()
                    except OSError:
                        pass
    return RenameResult(renamed=len(targets), unchanged=unchanged_count)
