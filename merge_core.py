"""PDF結合のコア処理（FR-M1〜FR-M14）.

入力フォルダーの直下にある `.pdf` を自然順に並べ、**全件を開いて検査してから**
1つのPDFへ連結する。暗号化・破損があれば書き始める前に中止し、
出力は出力先内の一時フォルダー（`.pdf_merge_*`）へ排他作成（`xb`）で書いて
成功後にだけ最終名へ移動する。既存ファイルの上書きは一切行わない。
一時フォルダーの作成・最終名への移動・後片付けは pathops.py 経由で行い、
ネットワークドライブ（NAS共有）でも同じ結果にする。
エラー文言は split_core.py の既存文言と同じ文体を保つ。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pypdf import PdfWriter

import pathops
from natural_sort import sort_pdf_names
from split_core import load_pdf, validate_filename_base

# 分割・結合・リネームが一時的に作る名前（対象ファイルから常に除外する）
TEMP_NAME_PREFIXES = pathops.TEMP_NAME_PREFIXES
MAX_LISTED = 8  # エラー一覧にそのまま並べる件数。超分は「ほか N 件」にまとめる


@dataclass(frozen=True)
class MergePlan:
    """検査済みの結合プラン（sources は自然順、page_counts は同じ順序）."""

    sources: tuple[Path, ...]
    page_counts: tuple[int, ...]
    total_pages: int
    output_path: Path


def is_temp_name(name: str) -> bool:
    """分割・結合・リネームの一時フォルダー／一時ファイル名かどうかを返す."""
    return name.startswith(TEMP_NAME_PREFIXES)


def list_merge_sources(folder: Path) -> list[Path]:
    """入力フォルダー直下の `.pdf` を自然順で返す（0件・1件は ValueError）."""
    if not folder.is_dir():
        raise ValueError("入力フォルダーが見つかりません。")
    names = [
        path.name
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() == ".pdf" and not is_temp_name(path.name)
    ]
    ordered = sort_pdf_names(names)
    if not ordered:
        raise ValueError("PDFが見つかりません。入力フォルダーの直下に .pdf ファイルを置いてください。")
    if len(ordered) == 1:
        raise ValueError("結合には2つ以上のPDFが必要です。")
    return [folder / name for name in ordered]


def default_output_name(sources: list[Path]) -> str:
    """結合後ファイル名の初期値（自然順で最初のPDFのファイル名・拡張子を除く）を返す."""
    return sources[0].stem


def validate_output_name(name: str) -> str:
    """結合後ファイル名を検査し、拡張子 `.pdf` を補った正規名を返す."""
    candidate = name.strip()
    if not candidate:
        raise ValueError("結合後のファイル名を入力してください。")
    suffix = Path(candidate).suffix
    if suffix and suffix.lower() != ".pdf":
        raise ValueError(f"結合後のファイル名には .pdf だけを使えます（「{candidate}」の拡張子は指定できません）。")
    base = candidate[: len(candidate) - len(suffix)] if suffix else candidate
    base = validate_filename_base(base, "結合後のファイル名")
    return f"{base}.pdf"


def check_output_target(output_dir: Path, out_name: str, sources: list[Path]) -> None:
    """出力先の既存ファイル・入力PDFと同名（大小無視）を検査して中止する."""
    if not output_dir.is_dir():
        raise ValueError("出力先フォルダーが見つかりません。")
    if any(source.name.casefold() == out_name.casefold() for source in sources):
        raise ValueError(f"入力PDFを出力名にできません。別のファイル名を使ってください（{out_name}）。")
    if (output_dir / out_name).exists():
        raise FileExistsError(f"同名の出力ファイルが既にあります。上書きは行いません。\n{out_name}")


def build_merge_plan(folder: Path, out_name: str, output_dir: Path) -> MergePlan:
    """全PDFを開いて検査してから結合プランを作る（ここでは一切書き出さない）."""
    sources = list_merge_sources(folder)
    name = validate_output_name(out_name)
    check_output_target(output_dir, name, sources)

    page_counts: list[int] = []
    unreadable: list[str] = []
    readers = []
    try:
        for source in sources:
            try:
                reader = load_pdf(source)
                readers.append(reader)
                page_counts.append(len(reader.pages))
            except ValueError as exc:
                page_counts.append(0)
                unreadable.append(f"{source.name}: {exc}")
    finally:
        for reader in readers:
            try:
                reader.close()
            except Exception:
                pass
    # 全件を検査してからまとめてエラーにする（途中で書き始めない）
    if unreadable:
        preview = "\n".join(unreadable[:MAX_LISTED])
        extra = f"\nほか {len(unreadable) - MAX_LISTED} 件" if len(unreadable) > MAX_LISTED else ""
        raise ValueError("読み込めないPDFがあります。暗号化や破損を解消してください。\n" + preview + extra)

    return MergePlan(
        sources=tuple(sources),
        page_counts=tuple(page_counts),
        total_pages=sum(page_counts),
        output_path=output_dir / name,
    )


def merge_pdfs(plan: MergePlan) -> Path:
    """プランどおり自然順で全ページを連結し、一時フォルダー経由で保存する."""
    output_dir = plan.output_path.parent
    try:
        staging_dir = pathops.make_staging_dir(output_dir, ".pdf_merge_")
    except OSError as exc:
        raise ValueError(f"結合ファイルを保存できませんでした。\n{pathops.describe_os_error(exc, output_dir)}") from exc
    staged = staging_dir / plan.output_path.name
    readers = []
    try:
        writer = PdfWriter()
        for source in plan.sources:
            reader = load_pdf(source)
            readers.append(reader)
            for page in reader.pages:
                writer.add_page(page)
        pathops.write_exclusive(staged, writer.write)  # 排他作成（既存名へは書かない）
        pathops.move_into(staged, plan.output_path)
    except Exception as exc:
        pathops.remove_file(staged)
        raise ValueError(f"結合ファイルを保存できませんでした。\n{pathops.describe_os_error(exc, plan.output_path)}") from exc
    finally:
        for reader in readers:
            try:
                reader.close()
            except Exception:
                pass
        pathops.remove_tree(staging_dir)
    return plan.output_path
