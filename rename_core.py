"""ExcelでPDFをリネームするコア処理（FR-R1〜FR-R17）.

適用先は2つ選べる:

- **入力フォルダー内で名前を変える**（従来どおり。コピーは作らない）
- **出力先フォルダーへコピーしながら名前を変える**（FR-R16。入力フォルダーのPDFは無傷のまま）

入力フォルダーの直下 `.pdf` を自然順に並べ、既存の命名関数
（`modes.build_filenames`）へ「ページ数」の代わりに「ファイル数」を渡して
新名を作り、同じ順序で割り当てる。実行は2段階リネーム
（①同一フォルダー内の一時名 `~pdf_rename_tmp_<連番>.tmp` へ ②最終名へ）で行い、
どちらの段階で失敗しても変更済みのファイルをすべて元の名前へ戻してからエラーを出す。
名前変更は pathops.py 経由で行うため、ネットワークドライブ（NAS共有）では
一時的な共有違反の再試行や、共有が拒否する名前の代替（`.pdf_rename_tmp_`）を使う。
一時名は共有で隠れない「~」始まりを先に選ぶ。
名前変更のあとに新しい名前のファイルが揃っているかを確かめ、揃わなければ
その場でエラーとして報告する（共有では成功したように見えて残らないことがあるため）。
コピー実行後は、出力先に残っていることとサイズが元ファイルと一致することを1件ずつ確かめる。
PDF・Excelの内容は一切変更せず、名前だけを変える（コピーでは元ファイルも残す）。

"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import modes
import pathops
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
    items = [
        RenameItem(
            order=order,
            source=source,
            new_name=new_name,
            unchanged=source.name.casefold() == new_name.casefold(),
        )
        for order, (source, new_name) in enumerate(zip(targets, filenames), start=1)
    ]
    temp_like = find_temp_like_names(items)
    if temp_like:
        preview = "\n".join(temp_like[:MAX_LISTED])
        extra = f"\nほか {len(temp_like) - MAX_LISTED} 件" if len(temp_like) > MAX_LISTED else ""
        raise ValueError(
            "Excelのセル値が一時ファイルの名前（~pdf_rename_ / .pdf_rename_ など）に似ています。"
            "この名前にすると以後の一覧から消えてしまうため、別の名前を使ってください。\n" + preview + extra
        )
    return items


def find_rename_conflicts(items: list[RenameItem], folder: Path) -> list[str]:
    """リネーム対象以外の既存ファイル（非PDF含む）と衝突する新名を返す（読み取りのみ）.

    名前を変えるファイルは一時名へ動かすので、その名前への衝突は問題にならない。
    一方、名前を変えないファイル（unchanged）は動かさないため、その名前を
    別のファイルが使うと上書きになる。名前を変えないファイルの名前も障害物として扱う。
    """
    moved_names = {item.source.name.casefold() for item in items if not item.unchanged}
    conflicts: list[str] = []
    reported: set[str] = set()
    for item in items:
        if item.unchanged:  # 自分自身の名前と衝突することにはならない
            continue
        key = item.new_name.casefold()
        if key in moved_names or key in reported:
            continue
        if (folder / item.new_name).exists():
            conflicts.append(item.new_name)
            reported.add(key)
    return conflicts


def find_temp_like_names(items: list[RenameItem]) -> list[str]:
    """一時ファイルの名前と紛らわしい新名を返す（以後の一覧から消えるため使わせない）."""
    like = [item.new_name for item in items if item.new_name.startswith(TEMP_NAME_PREFIXES)]
    return list(dict.fromkeys(like))


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
    temp_paths = dict(
        zip(
            (item.source for item in targets),
            pathops.temp_paths_for(folder, TEMP_RENAME_PREFIX, len(targets)),
        )
    )
    current: dict[RenameItem, Path] = {}
    rollback_ok = True
    try:
        for item in targets:  # ① 一時名へ
            temp = temp_paths[item.source]
            pathops.move_into(item.source, temp)
            current[item] = temp
        for item in targets:  # ② 最終名へ
            final = folder / item.new_name
            pathops.move_into(current[item], final)
            current[item] = final
    except Exception as exc:
        for item in targets:  # 変更済みだけを元の名前へ戻す
            path = current.get(item)
            if path is None or path == item.source:
                continue
            try:
                pathops.move_into(path, item.source)
            except OSError:
                rollback_ok = False
        raise ValueError(
            f"ファイル名を変更できませんでした。変更済みの名前は元に戻しました。\n{pathops.describe_os_error(exc, folder)}"
        ) from exc
    finally:
        if rollback_ok:  # 一時名の残骸を掃除する（戻せなかったファイルを消さない）
            for temp in temp_paths.values():
                if temp.exists():
                    pathops.remove_file(temp)
    # 名前変更が共有上で正しく完了したかを確かめる（共有では成功したように見えて残らないことがある）
    missing = [item.new_name for item in targets if not (folder / item.new_name).exists()]
    if missing:
        preview = "\n".join(missing[:MAX_LISTED])
        extra = f"\nほか {len(missing) - MAX_LISTED} 件" if len(missing) > MAX_LISTED else ""
        raise ValueError(
            "名前を変更したはずのファイルが入力フォルダーに見当たりません。ファイルが消えたわけではありません。\n"
            "同じフォルダーをエクスプローラーで開き、Excelのセル値と同じ名前のファイル、"
            "および ~pdf_rename_tmp_ や .pdf_rename_tmp_ で始まるファイルを確認してください。\n"
            "ネットワークドライブの場合は共有の再接続、またはフォルダーをF5で更新して再確認してください。\n"
            + preview
            + extra
        )
    return RenameResult(renamed=len(targets), unchanged=unchanged_count)


# --- 出力先へコピーしながら名前を変える実行（元ファイルは変更しない） ---


@dataclass(frozen=True)
class CopyItem:
    """コピー1件（元ファイルはそのままで、出力先に新しい名前で書き出す）."""

    order: int
    source: Path
    new_name: str
    destination: Path


@dataclass(frozen=True)
class CopyResult:
    """コピー実行結果（コピーした件数と、同じ場所・同じ名前で変更不要だった件数）."""

    copied: int
    unchanged: int


def is_same_file(left: Path, right: Path) -> bool:
    """同じフォルダーの同じ名前かどうか（Windowsは大文字小文字と区切り記号を区別しない）."""
    return os.path.normcase(str(left)) == os.path.normcase(str(right))


def build_copy_plan(items: list[RenameItem], out_folder: Path) -> list[CopyItem]:
    """リネーム計画を「出力先へコピーする計画」へ読み替える（FR-R16）."""
    return [
        CopyItem(
            order=item.order,
            source=item.source,
            new_name=item.new_name,
            destination=out_folder / item.new_name,
        )
        for item in items
    ]


def find_copy_conflicts(plan: list[CopyItem]) -> list[str]:
    """出力先に既にあってコピーできない名前を返す（元ファイル自身は除く）."""
    conflicts: list[str] = []
    reported: set[str] = set()
    for item in plan:
        key = item.new_name.casefold()
        if key in reported or not item.destination.exists():
            continue
        if is_same_file(item.source, item.destination):  # 元ファイルと同じ場所・同じ名前
            continue
        conflicts.append(item.new_name)
        reported.add(key)
    return conflicts


def check_copy_plan(plan: list[CopyItem], out_folder: Path) -> None:
    """コピーを1つも始める前に、出力先の状態を検査して中止する（FR-R17）."""
    if not out_folder.is_dir():
        raise ValueError("出力先フォルダーが見つかりません。存在するフォルダーを指定してください。")
    conflicts = find_copy_conflicts(plan)
    if conflicts:
        preview = "\n".join(conflicts[:MAX_LISTED])
        extra = f"\nほか {len(conflicts) - MAX_LISTED} 件" if len(conflicts) > MAX_LISTED else ""
        raise FileExistsError(
            "出力先に同じ名前のファイルが既にあります。上書きは行いません。\n" + preview + extra
        )


def apply_copies(plan: list[CopyItem]) -> CopyResult:
    """出力先へコピーする。入力フォルダーのPDFは変更しないため、失敗しても戻す作業は不要（FR-R16）."""
    copied: list[CopyItem] = []
    unchanged = 0
    for item in plan:
        if is_same_file(item.source, item.destination):
            unchanged += 1
            continue
        try:
            pathops.copy_into(item.source, item.destination)
        except OSError as exc:
            done = f"\n{len(copied)} 件はコピー済みです。" if copied else ""
            raise ValueError(
                "PDFをコピーできませんでした。入力フォルダーのPDFは無傷です。\n"
                + pathops.describe_os_error(exc, item.destination.parent)
                + done
            ) from exc
        copied.append(item)
    problems = verify_copies(copied)
    if problems:
        preview = "\n".join(problems[:MAX_LISTED])
        extra = f"\nほか {len(problems) - MAX_LISTED} 件" if len(problems) > MAX_LISTED else ""
        raise ValueError(
            "コピーしたはずのファイルが出力先に正しく残っていません。入力フォルダーのPDFは無傷です。\n"
            "出力先フォルダーをF5で更新して確認し、ネットワークドライブの場合は共有の再接続もお試しください。\n"
            + preview
            + extra
        )
    return CopyResult(copied=len(copied), unchanged=unchanged)


def verify_copies(copied: list[CopyItem]) -> list[str]:
    """コピー結果を1件ずつ確かめる（共有では「書いたのに残らない」「途中で切れる」が起きうる）."""
    problems: list[str] = []
    for item in copied:
        if not item.destination.is_file():
            problems.append(f"{item.new_name}（出力先に見当たりません）")
            continue
        try:
            written = item.destination.stat().st_size
            original = item.source.stat().st_size
        except OSError:
            problems.append(f"{item.new_name}（確認できません）")
            continue
        if written != original:
            problems.append(f"{item.new_name}（サイズが一致しません: {written} / {original} バイト）")
    return problems
