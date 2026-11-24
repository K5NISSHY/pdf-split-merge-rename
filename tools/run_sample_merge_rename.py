r"""サンプル資料でPDF結合・Excelリネームを実行し、結果を samples\検証結果 に保存する検証スクリプト.

実行: py -3 tools\run_sample_merge_rename.py

- 結合: samples\結合入力 のPDF5個（001 / 003 / 008 / 10 / 9）を samples\検証結果\結合 へコピーして
  実際の処理系（merge_core）で結合する。結合順が自然順（001 → 003 → 008 → 9 → 10）であること、
  ページ数が合計8ページであること、本文文字が順序どおり残っていることを機械判定する
- リネーム: samples\リネーム入力 のPDF5個（名前がばらばら）を samples\検証結果\リネーム へコピーして
  実際の処理系（rename_core / modes）で名前を変える。samples 自体は無変更
- 再実行（結合は上書き中止・リネームは変更不要扱い）、衝突中止、一時名残骸の有無も検査し、
  最後に TOTAL: <成功>/<総数> を出す（全成功で終了コード0）
"""

from __future__ import annotations

import hashlib
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pypdf import PdfReader  # noqa: E402

import make_samples  # noqa: E402
import merge_core  # noqa: E402
import modes  # noqa: E402
import preview  # noqa: E402
import rename_core  # noqa: E402

SAMPLE_DIR = ROOT / "samples"
MERGE_SRC_DIR = SAMPLE_DIR / "結合入力"
RENAME_SRC_DIR = SAMPLE_DIR / "リネーム入力"
XLSX_PATH = SAMPLE_DIR / "sample_names.xlsx"
RESULT_ROOT = SAMPLE_DIR / "検証結果"
MERGE_DIR = RESULT_ROOT / "結合"
MERGE_IN_DIR = MERGE_DIR / "入力"
RENAME_DIR = RESULT_ROOT / "リネーム"
SHEET_NAME = "サンプル"

RESULTS: list[tuple[bool, str]] = []


def check(ok: bool, message: str) -> None:
    RESULTS.append((bool(ok), message))
    print(("PASS: " if ok else "FAIL: ") + message)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def prepare_dir(folder: Path) -> None:
    """結果フォルダーを作り直す（同名既存ファイルで中止する規則を避けるため）."""
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True)


def copy_pdfs(source: Path, target: Path) -> list[Path]:
    """サンプルPDFを結果フォルダーへコピーしてコピー後の一覧を返す."""
    copied = []
    for path in sorted(source.glob("*.pdf")):
        shutil.copy2(path, target / path.name)
        copied.append(target / path.name)
    return copied


def page_texts(path: Path) -> list[str]:
    """PDFの各ページから本文文字列を取り出す."""
    reader = PdfReader(str(path))
    try:
        return [(page.extract_text() or "").replace("\n", " / ") for page in reader.pages]
    finally:
        try:
            reader.close()
        except Exception:
            pass


def verify_merge() -> None:
    """結合タブ: 自然順での結合・ページ数・既存名での中止・一時名残骸."""
    # 入力フォルダーと出力先を分ける（同じだと結合結果そのものが入力PDFになり、
    # FR-M9「入力PDFと同名の出力名」の中止になる。実運用でも分ける想定）
    prepare_dir(MERGE_DIR)
    MERGE_IN_DIR.mkdir(parents=True)
    sources = copy_pdfs(MERGE_SRC_DIR, MERGE_IN_DIR)
    before = {p.name: digest(p) for p in sources}

    listed = [p.name for p in merge_core.list_merge_sources(MERGE_IN_DIR)]
    check(listed == make_samples.MERGE_NATURAL_ORDER,
          f"FR-M2 結合順序は自然順（Explorer順 001/003/008/10/9 ではない）: {' → '.join(listed)}")

    rows = preview.build_merge_preview(MERGE_IN_DIR, "結合結果", MERGE_DIR)
    check([r.page_count for r in rows] == [2, 1, 3, 1, 1],
          f"FR-M13 プレビューのページ数: {[r.page_count for r in rows]}")

    plan = merge_core.build_merge_plan(MERGE_IN_DIR, "結合結果", MERGE_DIR)
    check(len(plan.sources) == 5 and plan.total_pages == make_samples.MERGE_TOTAL_PAGES,
          f"FR-M14 結合対象 {len(plan.sources)} ファイル / {plan.total_pages} ページ → {plan.output_path.name}")

    saved = merge_core.merge_pdfs(plan)
    texts = page_texts(saved)
    expected = [
        f"Page {page_no} of {make_samples.MERGE_PAGES_BY_NAME[name]} in {name}"
        for name in make_samples.MERGE_NATURAL_ORDER
        for page_no in range(1, make_samples.MERGE_PAGES_BY_NAME[name] + 1)
    ]
    check(len(texts) == len(expected) and all(want in got for want, got in zip(expected, texts)),
          f"FR-M11 全ページが自然順のまま連結されている: {texts[0]} … {texts[-1]}")
    check({p.name: digest(p) for p in sources} == before,
          "FR-M11 入力PDFは無変更（ハッシュ一致）")
    check([p.name for p in MERGE_DIR.iterdir() if p.is_file()] == ["結合結果.pdf"],
          f"FR-M12 出力先のファイル構成: {' / '.join(sorted(p.name for p in MERGE_DIR.iterdir()))}")

    try:
        merge_core.merge_pdfs(merge_core.build_merge_plan(MERGE_IN_DIR, "結合結果", MERGE_DIR))
        check(False, "FR-M10 同名出力への再結合を中止していない")
    except FileExistsError as exc:
        check("上書きは行いません" in str(exc), f"FR-M10 同名出力への再結合は中止: {str(exc).splitlines()[-1]}")

    try:
        merge_core.build_merge_plan(MERGE_IN_DIR, "001.pdf", MERGE_DIR)
        check(False, "FR-M9 入力PDFと同名の出力名を中止していない")
    except ValueError as exc:
        check("入力PDFを出力名にできません" in str(exc), f"FR-M9 入力PDFと同名の出力名は中止: {exc}")

    leftovers = [folder.name / p.name for folder in (MERGE_DIR, MERGE_IN_DIR)
                 for p in folder.iterdir() if merge_core.is_temp_name(p.name)]
    check(leftovers == [], "FR-M12 一時フォルダー（.pdf_merge_*）は残っていない")

    with tempfile.TemporaryDirectory() as tmp:
        one = Path(tmp) / "one"
        one.mkdir()
        shutil.copy2(MERGE_SRC_DIR / "001.pdf", one / "001.pdf")
        empty = Path(tmp) / "empty"
        empty.mkdir()
        try:
            merge_core.list_merge_sources(empty)
            check(False, "FR-M1 PDF0件を中止していない")
        except ValueError as exc:
            check("PDFが見つかりません" in str(exc), f"FR-M1 PDF0件は中止: {exc}")
        try:
            merge_core.list_merge_sources(one)
            check(False, "FR-M1 PDF1件を中止していない")
        except ValueError as exc:
            check(str(exc) == "結合には2つ以上のPDFが必要です。", f"FR-M1 PDF1件は中止: {exc}")


def verify_rename() -> None:
    """リネームタブ: 自然順への割り当て・2ルール・変更不要・衝突中止・一時名残骸."""
    prepare_dir(RENAME_DIR)
    copied = copy_pdfs(RENAME_SRC_DIR, RENAME_DIR)
    src_digest = {p.name: digest(p) for p in copied}

    targets = [p.name for p in rename_core.list_rename_targets(RENAME_DIR)]
    check(targets == make_samples.RENAME_NATURAL_ORDER,
          f"FR-R1/FR-R3 リネーム対象は自然順: {' → '.join(targets)}")

    rows = preview.build_rename_preview(RENAME_DIR, XLSX_PATH, SHEET_NAME, "B1", modes.PER_CELL)
    check([r.new_name for r in rows] == [f"{name}.pdf" for name in make_samples.PAGE_NAMES],
          f"FR-R11 プレビューの新ファイル名（B1から下行き）: {' / '.join(r.new_name for r in rows)}")

    items = rename_core.build_rename_plan(RENAME_DIR, XLSX_PATH, SHEET_NAME, "B1", modes.PER_CELL)
    rename_core.check_rename_plan(items, RENAME_DIR)
    result = rename_core.apply_renames(items)
    check((result.renamed, result.unchanged) == (5, 0),
          f"FR-R12 変更 {result.renamed} 件 / 変更不要 {result.unchanged} 件")
    check(sorted(p.name for p in RENAME_DIR.iterdir()) == sorted(f"{name}.pdf" for name in make_samples.PAGE_NAMES),
          f"FR-R2 開始セルから下の行のセル値で名前変更: {' / '.join(sorted(p.name for p in RENAME_DIR.iterdir()))}")
    check(all(digest(RENAME_DIR / f"{old}.pdf") == src_digest[new]
              for old, new in zip(make_samples.PAGE_NAMES, make_samples.RENAME_NATURAL_ORDER)),
          "FR-R13 PDFの中身は無変更（ハッシュ一致）")

    again = rename_core.apply_renames(
        rename_core.build_rename_plan(RENAME_DIR, XLSX_PATH, SHEET_NAME, "B1", modes.PER_CELL)
    )
    check((again.renamed, again.unchanged) == (0, 5),
          f"FR-R6 再実行はすべて変更不要: 変更 {again.renamed} 件 / 変更不要 {again.unchanged} 件")

    fixed = rename_core.build_rename_plan(RENAME_DIR, XLSX_PATH, SHEET_NAME, "A1", modes.BASE_SEQ)
    check([it.new_name for it in fixed] == [f"{make_samples.BASE_NAME}_{n:03d}.pdf" for n in range(1, 6)],
          f"FR-R2 固定セル値＋連番の新ファイル名: {fixed[0].new_name}〜{fixed[-1].new_name}")
    rename_core.check_rename_plan(fixed, RENAME_DIR)
    second = rename_core.apply_renames(fixed)
    check((second.renamed, second.unchanged) == (5, 0),
          f"FR-R9 2段階リネームで全件変更: 変更 {second.renamed} 件")

    decoy = RENAME_DIR / "既存の資料.pdf"
    decoy.write_bytes(b"this is not a pdf")
    conflict_items = [
        rename_core.RenameItem(order=1, source=RENAME_DIR / fixed[-1].new_name, new_name=decoy.name, unchanged=False)
    ]
    names_before = sorted(p.name for p in RENAME_DIR.iterdir())
    try:
        rename_core.check_rename_plan(conflict_items, RENAME_DIR)
        check(False, "FR-R7 既存名との衝突を中止していない")
    except FileExistsError as exc:
        check("衝突します" in str(exc) and decoy.name in str(exc),
              f"FR-R7 リネーム対象以外の既存名と衝突すると中止: {str(exc).splitlines()[-1]}")
    check(sorted(p.name for p in RENAME_DIR.iterdir()) == names_before,
          "FR-R7 衝突の検査では名前も中身も変わっていない")
    decoy.unlink()  # このスクリプトが作った一時ファイルだけ削除する

    leftovers = [p.name for p in RENAME_DIR.iterdir() if merge_core.is_temp_name(p.name)]
    check(leftovers == [], "FR-R9 一時ファイル（.pdf_rename_*）は残っていない")
    check(sorted(p.name for p in RENAME_SRC_DIR.iterdir()) == sorted(make_samples.RENAME_NATURAL_ORDER),
          "FR-R13 samples\\リネーム入力 の元サンプルは無変更")


def main() -> int:
    missing = not (
        MERGE_SRC_DIR.is_dir() and RENAME_SRC_DIR.is_dir() and XLSX_PATH.is_file()
        and (SAMPLE_DIR / "sample_5pages.pdf").is_file()
    )
    if missing:
        make_samples.main()
    verify_merge()
    verify_rename()
    succeeded = sum(1 for ok, _ in RESULTS if ok)
    print(f"TOTAL: {succeeded}/{len(RESULTS)}")
    return 0 if succeeded == len(RESULTS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
