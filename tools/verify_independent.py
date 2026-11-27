"""独立検証スクリプト（FR-13）。

プロジェクト内の一時フォルダー（`.pdf_split_verify_*`）に実データ
（3ページ空白PDF・A1〜A3 に あいう/かきく/さしす のExcel）を作り、
両モードで分割を実行して機械的に合否を出す。表示用数値は手で書かない。
後半（15〜）は PDF結合（docs 2.3 FR-M1〜FR-M14）と Excelリネーム（docs 2.4 FR-R1〜FR-R13）
の項目を、同じ実データ生成で1項目1行として判定する。

実行: `py -3 tools\\verify_independent.py`
出力: PASS/FAIL を1行ずつ、最終行に `TOTAL: <成功>/<総数>`。
終了コードは全件成功時 0、そうでない時 0 以外。
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

import merge_core
import modes
import pathops
import preview
import rename_core
import split_core

RESULTS: list[tuple[str, bool]] = []


def check(label: str, ok: bool) -> None:
    RESULTS.append((label, bool(ok)))


def make_pdf(path: Path, pages: int) -> None:
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=72, height=72)
    with path.open("wb") as stream:
        writer.write(stream)


def make_excel(path: Path, cells: dict[str, str]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    for address, value in cells.items():
        sheet[address] = value
    workbook.save(path)
    workbook.close()


def page_count(path: Path) -> int:
    return len(PdfReader(str(path)).pages)


def top_level_pdf_names(output_dir: Path) -> list[str]:
    return sorted(p.name for p in output_dir.glob("*.pdf"))


def all_dir_names(output_dir: Path) -> list[str]:
    return sorted(p.name for p in output_dir.iterdir())


def make_marked_pdf(path: Path, marks: list[str]) -> None:
    """各ページに英字マーカーを描いたPDFを作る（結合順序の検証用）."""
    writer = PdfWriter()
    for mark in marks:
        page = writer.add_blank_page(width=200, height=120)
        page[NameObject("/Resources")] = DictionaryObject(
            {
                NameObject("/Font"): DictionaryObject(
                    {
                        NameObject("/F1"): DictionaryObject(
                            {
                                NameObject("/Type"): NameObject("/Font"),
                                NameObject("/Subtype"): NameObject("/Type1"),
                                NameObject("/BaseFont"): NameObject("/Helvetica"),
                            }
                        )
                    }
                )
            }
        )
        stream = DecodedStreamObject()
        stream.set_data(f"BT /F1 24 Tf 20 60 Td ({mark}) Tj ET".encode("ascii"))
        page[NameObject("/Contents")] = stream
    with path.open("wb") as stream:
        writer.write(stream)


def page_marks(path: Path) -> list[str]:
    """PDFの各ページから本文の英字マーカーを取り出して並べる."""
    reader = PdfReader(str(path))
    try:
        return [(page.extract_text() or "").strip() for page in reader.pages]
    finally:
        try:
            reader.close()
        except Exception:
            pass


def digest(folder: Path) -> dict[str, str]:
    """フォルダー直下の全ファイルの中身ハッシュ（名前は含まない値）."""
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.iterdir() if p.is_file()}


def temp_names(folder: Path) -> list[str]:
    """一時フォルダー・一時ファイルの名前（残っていてはいけないもの）."""
    return [p.name for p in folder.iterdir() if merge_core.is_temp_name(p.name)]


def verify(root: Path) -> None:
    # 1. 実データの生成
    pdf_path = root / "input.pdf"
    excel_path = root / "names.xlsx"
    make_pdf(pdf_path, 3)
    make_excel(excel_path, {"A1": "あいう", "A2": "かきく", "A3": "さしす"})
    book = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        sheet = book["Sheet"]
        excel_values = [sheet[f"A{row}"].value for row in (1, 2, 3)]
        sheet_names = book.sheetnames
    finally:
        book.close()
    data_ok = page_count(pdf_path) == 3 and sheet_names == ["Sheet"] and excel_values == ["あいう", "かきく", "さしす"]
    check("1. 3ページの空白PDFとA1〜A3にあいう/かきく/さしすを持つExcelを生成", data_ok)

    # 2. base_seq（A1）
    out_base = root / "out_base_seq"
    out_base.mkdir()
    split_core.split_pdf(pdf_path, excel_path, out_base, "Sheet", "A1", modes.build_filenames_base_seq)
    base_names = top_level_pdf_names(out_base)
    check(
        "2. base_seq（A1）→ あいう_001.pdf / あいう_002.pdf / あいう_003.pdf が生成される",
        base_names == ["あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf"],
    )
    check(
        "3. base_seqの各PDFはちょうど1ページ",
        len(base_names) == 3 and all(page_count(out_base / name) == 1 for name in base_names),
    )

    # 3. per_cell（A1）
    out_cell = root / "out_per_cell"
    out_cell.mkdir()
    split_core.split_pdf(pdf_path, excel_path, out_cell, "Sheet", "A1", modes.build_filenames_per_cell)
    cell_names = top_level_pdf_names(out_cell)
    check(
        "4. per_cell（A1）→ あいう.pdf / かきく.pdf / さしす.pdf が生成される",
        cell_names == sorted(["あいう.pdf", "かきく.pdf", "さしす.pdf"]),
    )
    check(
        "5. per_cellの各PDFはちょうど1ページ",
        len(cell_names) == 3 and all(page_count(out_cell / name) == 1 for name in cell_names),
    )

    # 4. 同じ出力先で再実行 → FileExistsError で中止・既存の中身は無変更
    before_bytes = {name: (out_cell / name).read_bytes() for name in cell_names}
    guard_base = False
    try:
        split_core.split_pdf(pdf_path, excel_path, out_base, "Sheet", "A1", modes.build_filenames_base_seq)
    except FileExistsError:
        guard_base = True
    check("6. base_seqの再実行はFileExistsErrorで中止（上書きなし）", guard_base)
    guard_cell = False
    try:
        split_core.split_pdf(pdf_path, excel_path, out_cell, "Sheet", "A1", modes.build_filenames_per_cell)
    except FileExistsError:
        guard_cell = True
    check("7. per_cellの再実行はFileExistsErrorで中止（上書きなし）", guard_cell)
    after_bytes = {name: (out_cell / name).read_bytes() for name in top_level_pdf_names(out_cell)}
    tidy_ok = (
        after_bytes == before_bytes
        and all_dir_names(out_base) == [f"あいう_{n:03d}.pdf" for n in (1, 2, 3)]
        and all_dir_names(out_cell) == cell_names  # 一時フォルダー（.pdf_split_*）も残さない
    )
    check("8. 再実行後も既存ファイルの中身が変わっておらず、後片付けされている", tidy_ok)
    # 5. 空欄セル（タブ2でA2を空）→ 実行前に中止し、出力は0件
    excel_blank = root / "blank.xlsx"
    make_excel(excel_blank, {"A1": "p1", "A3": "p3"})  # A2 は作らない（空欄）
    out_blank = root / "out_blank"
    out_blank.mkdir()
    blank_abort = False
    try:
        split_core.split_pdf(pdf_path, excel_blank, out_blank, "Sheet", "A1", modes.build_filenames_per_cell)
    except ValueError as exc:
        blank_abort = "が空です" in str(exc)
    check("9. 空欄セル（A2）は実行前のエラーで中止", blank_abort)
    check("10. 空欄セルのとき出力ファイルは1つも作られていない", all_dir_names(out_blank) == [])

    # 6. 重複名（Report / report）→ 重複エラーで中止
    excel_dup = root / "dup.xlsx"
    make_excel(excel_dup, {"A1": "Report", "A2": "report", "A3": "solo"})
    out_dup = root / "out_dup"
    out_dup.mkdir()
    dup_abort = False
    try:
        split_core.split_pdf(pdf_path, excel_dup, out_dup, "Sheet", "A1", modes.build_filenames_per_cell)
    except ValueError as exc:
        dup_abort = "同じファイル名" in str(exc)
    check("11. 重複名（Report / report）は重複エラーで中止", dup_abort)
    check("12. 重複名のとき出力ファイルは1つも作られていない", all_dir_names(out_dup) == [])

    # 7. プレビュー: 4の状況で [既存]、6の状況で [重複]、書き出しなし
    dup_rows = preview.build_preview(excel_dup, "Sheet", "A1", modes.PER_CELL, out_dup, pdf_path)
    dup_marks = [row.status for row in dup_rows] == ["[重複]", "[重複]", ""]
    check(
        "13. プレビューは重複状況で [重複] を付け、書き出しが発生していない",
        dup_marks and all_dir_names(out_dup) == [],
    )
    exist_rows = preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, out_cell, pdf_path)
    exist_marks = len(exist_rows) == 3 and all("[既存]" in row.status for row in exist_rows)
    preview_untouched = {name: (out_cell / name).read_bytes() for name in top_level_pdf_names(out_cell)} == before_bytes
    check(
        "14. プレビューは既存ファイル状況で [既存] を付け、既存の中身と一覧は無変更",
        exist_marks and preview_untouched and all_dir_names(out_cell) == cell_names,
    )


def verify_merge(root: Path) -> None:
    """FR-M1〜FR-M14: 結合（自然順・全件検査・排他作成・ロールバック）."""
    # FR-M1: 直下の .pdf だけ（大文字小文字無視・サブフォルダーは対象外）
    m1 = root / "m1"
    (m1 / "sub").mkdir(parents=True)
    make_pdf(m1 / "001.pdf", 1)
    make_pdf(m1 / "002.PDF", 1)
    make_pdf(m1 / "sub" / "003.pdf", 1)
    (m1 / "notes.txt").write_text("not pdf", encoding="utf-8")
    check(
        "15. FR-M1 結合対象は入力フォルダー直下の .pdf のみ（大文字小文字無視・サブフォルダーは除外）",
        [p.name for p in merge_core.list_merge_sources(m1)] == ["001.pdf", "002.PDF"],
    )

    # 結合の材料（Explorer順に作り、自然順へ並び替わることを検査する）
    src = root / "merge_in"
    src.mkdir()
    marks = {
        "001.pdf": ["ONE-1", "ONE-2"],
        "003.pdf": ["THREE"],
        "008.pdf": ["EIGHT-1", "EIGHT-2", "EIGHT-3"],
        "10.pdf": ["TEN"],
        "9.pdf": ["NINE"],
    }
    for name in ("001.pdf", "003.pdf", "008.pdf", "10.pdf", "9.pdf"):
        make_marked_pdf(src / name, marks[name])
    out = root / "merge_out"
    out.mkdir()
    src_before = digest(src)

    check(
        "16. FR-M2 結合順序は自然順（001→003→008→9→10。Explorer順 001/003/008/10/9 ではない）",
        [p.name for p in merge_core.list_merge_sources(src)] == ["001.pdf", "003.pdf", "008.pdf", "9.pdf", "10.pdf"],
    )

    empty = root / "merge_empty"
    empty.mkdir()
    try:
        merge_core.list_merge_sources(empty)
        check("17. FR-M3 0件は「PDFが見つかりません」で実行しない", False)
    except ValueError as exc:
        check("17. FR-M3 0件は「PDFが見つかりません」で実行しない", "PDFが見つかりません" in str(exc))
    one = root / "merge_one"
    one.mkdir()
    make_pdf(one / "a.pdf", 1)
    try:
        merge_core.list_merge_sources(one)
        check("18. FR-M3 1件は「結合には2つ以上のPDFが必要です。」で実行しない", False)
    except ValueError as exc:
        check("18. FR-M3 1件は「結合には2つ以上のPDFが必要です。」で実行しない",
              str(exc) == "結合には2つ以上のPDFが必要です。")

    check(
        "19. FR-M4 結合後ファイル名の初期値は自然順で最初のPDF名（拡張子を除く）",
        merge_core.default_output_name(merge_core.list_merge_sources(src)) == "001",
    )

    name_ok = True
    for bad, expect in (
        ("", "結合後のファイル名を入力してください"),
        ("   ", "結合後のファイル名を入力してください"),
        ("a<b", "ファイル名に使えない文字"),
        ("CON", "予約名"),
        ("trail.", "拡張子は指定できません"),
        ("x.txt", "拡張子は指定できません"),
    ):
        try:
            merge_core.validate_output_name(bad)
            name_ok = False
        except ValueError as exc:
            name_ok = name_ok and expect in str(exc)
    check(
        "20. FR-M7 空 / 不正文字 / 予約名 / 末尾ピリオド / .pdf以外の拡張子を実行前にエラー（前後の空白は除いて検査）",
        name_ok and merge_core.validate_output_name(" 結合結果 ") == "結合結果.pdf",
    )
    check(
        "21. FR-M7 .pdf を除いた名前と .pdf 付きの名前は同じ正規名に補う",
        merge_core.validate_output_name("報告書") == "報告書.pdf"
        and merge_core.validate_output_name("報告書.pdf") == "報告書.pdf",
    )

    same_ok = False
    try:
        merge_core.build_merge_plan(src, "001.PDF", out)
    except ValueError as exc:
        same_ok = "入力PDFを出力名にできません" in str(exc)
    check("22. FR-M8 入力PDFと同名（大文字小文字無視）の出力名は実行しない", same_ok)

    bad_in = root / "merge_bad_in"
    bad_in.mkdir()
    bad_out = root / "merge_bad_out"
    bad_out.mkdir()
    make_pdf(bad_in / "000_good.pdf", 1)
    for index in range(1, 10):
        if index % 2:
            (bad_in / f"00{index}_broken.pdf").write_bytes(b"not a pdf")
        else:
            writer = PdfWriter()
            writer.add_blank_page(width=72, height=72)
            writer.encrypt("secret")
            with (bad_in / f"00{index}_enc.pdf").open("wb") as stream:
                writer.write(stream)
    bad_message = ""
    try:
        merge_core.build_merge_plan(bad_in, "結合結果", bad_out)
    except ValueError as exc:
        bad_message = str(exc)
    check(
        "23. FR-M10 暗号化・破損は全件検査してからエラー一覧（先頭8件＋「ほか 1 件」）を出す",
        "読み込めないPDFがあります" in bad_message
        and bad_message.count(".pdf: ") == merge_core.MAX_LISTED
        and "ほか 1 件" in bad_message,
    )
    check("24. FR-M10 読み込めないPDFがあるとき出力ファイルは1つも作られていない", all_dir_names(bad_out) == [])

    rows = preview.build_merge_preview(src, "結合結果", out)
    check(
        "25. FR-M13 プレビューは 順序/入力PDF/ページ数/状態 を全件出し、書き出していない",
        [r.page_count for r in rows] == [2, 1, 3, 1, 1]
        and all(len(r.display_cells()) == 4 for r in rows)
        and all_dir_names(out) == [],
    )

    plan = merge_core.build_merge_plan(src, "結合結果", out)
    saved = merge_core.merge_pdfs(plan)
    check(
        "26. FR-M11 全ページが自然順のまま連結されている（本文マーカーで順序確認）",
        page_marks(saved) == ["ONE-1", "ONE-2", "THREE", "EIGHT-1", "EIGHT-2", "EIGHT-3", "NINE", "TEN"],
    )
    check("27. FR-M11 入力PDFは一切変更されていない（ハッシュ一致）", digest(src) == src_before)
    check(
        "28. FR-M14 結合結果のページ数はプランの合計（完了メッセージの N ファイル / M ページと同じ値）",
        page_count(saved) == plan.total_pages == 8 and len(plan.sources) == 5,
    )
    check("29. FR-M12 一時フォルダー（.pdf_merge_*）は成功後に残らない", temp_names(out) == [])

    class FailingWriter(PdfWriter):
        def add_page(self, page):  # 書き出し途中の失敗を再現する
            raise OSError("検証による意図的な失敗")

    real_writer = merge_core.PdfWriter
    merge_core.PdfWriter = FailingWriter
    fail_message = ""
    try:
        merge_core.merge_pdfs(merge_core.build_merge_plan(src, "結合結果2", out))
    except ValueError as exc:
        fail_message = str(exc)
    finally:
        merge_core.PdfWriter = real_writer
    check(
        "30. FR-M12 書き出しに失敗したときは作成済みファイルを削除してロールバックする",
        "結合ファイルを保存できませんでした" in fail_message
        and all_dir_names(out) == ["結合結果.pdf"]
        and temp_names(out) == [],
    )

    existing_ok = False
    keep = digest(out)
    try:
        merge_core.build_merge_plan(src, "結合結果", out)
    except FileExistsError as exc:
        existing_ok = "上書きは行いません" in str(exc)
    marked = preview.build_merge_preview(src, "結合結果", out)
    check(
        "31. FR-M9 出力先に同名の既存ファイルがあるときは実行せず、プレビューは [既存]",
        existing_ok
        and all("[既存]" in r.status for r in marked)
        and digest(out) == keep
        and temp_names(out) == [],
    )


def verify_rename(root: Path) -> None:
    """FR-R1〜FR-R13: ExcelでPDFをリネーム（自然順・2段階・ロールバック）."""
    rn = root / "rename_in"
    rn.mkdir()
    for name in ("10_c.pdf", "2_a.pdf", "8_b.pdf"):  # 名前がばらばらな状態
        make_pdf(rn / name, 1)
    rn_excel = root / "rename_names.xlsx"
    make_excel(rn_excel, {"A1": "報告書", "B1": "一", "B2": "二", "B3": "三"})
    excel_before = hashlib.sha256(rn_excel.read_bytes()).hexdigest()
    contents_before = sorted(digest(rn).values())

    check(
        "32. FR-R1 リネーム対象は入力フォルダー直下の .pdf を自然順に並べたもの",
        [p.name for p in rename_core.list_rename_targets(rn)] == ["2_a.pdf", "8_b.pdf", "10_c.pdf"],
    )

    plan_a = rename_core.build_rename_plan(rn, rn_excel, "Sheet", "A1", modes.BASE_SEQ)
    check(
        "33. FR-R2/FR-R4 ルールAは {固定セル値}_{001..N}.pdf（ページ数ではなくファイル数で生成）",
        [item.new_name for item in plan_a] == ["報告書_001.pdf", "報告書_002.pdf", "報告書_003.pdf"],
    )
    check(
        "34. FR-R3 自然順の i 番目に新名の i 番目を当てる",
        [(item.order, item.source.name, item.new_name) for item in plan_a]
        == [(1, "2_a.pdf", "報告書_001.pdf"), (2, "8_b.pdf", "報告書_002.pdf"), (3, "10_c.pdf", "報告書_003.pdf")],
    )
    plan_b = rename_core.build_rename_plan(rn, rn_excel, "Sheet", "B1", modes.PER_CELL)
    check(
        "35. FR-R2/FR-R4 ルールBは開始セルから同列を下行きN個（連番なし）",
        [item.new_name for item in plan_b] == ["一.pdf", "二.pdf", "三.pdf"],
    )

    short_excel = root / "rename_short.xlsx"
    make_excel(short_excel, {"A1": "報告書", "B1": "一", "B2": "二"})
    bad_excel = root / "rename_bad.xlsx"
    make_excel(bad_excel, {"A1": "報告書", "B1": "一", "B2": "不正:名前", "B3": "三"})
    short_ok = bad_ok = False
    for expect, path, attr in (("が空です", short_excel, "B1"), ("ファイル名に使えない文字", bad_excel, "B1")):
        try:
            rename_core.build_rename_plan(rn, path, "Sheet", attr, modes.PER_CELL)
        except ValueError as exc:
            if expect in str(exc):
                if path is short_excel:
                    short_ok = True
                else:
                    bad_ok = True
    check("36. FR-R5 セル不足・不正文字は分割タブと同じ文言で実行前にエラー", short_ok and bad_ok)

    dup_excel = root / "rename_dup.xlsx"
    make_excel(dup_excel, {"A1": "報告書", "B1": "dup", "B2": "DUP", "B3": "solo"})
    dup_ok = False
    try:
        rename_core.build_rename_plan(rn, dup_excel, "Sheet", "B1", modes.PER_CELL)
    except ValueError as exc:
        dup_ok = "同じファイル名" in str(exc)
    names_at_dup = sorted(p.name for p in rn.iterdir())
    check(
        "37. FR-R6/FR-R10 新名同士の重複（大小無視）は実行前に中止し、名前変更は1つも起きていない",
        dup_ok and names_at_dup == ["10_c.pdf", "2_a.pdf", "8_b.pdf"],
    )
    dup_rows = preview.build_rename_preview(rn, dup_excel, "Sheet", "B1", modes.PER_CELL)
    check(
        "38. FR-R6 プレビューの重複行は [重複]（実行はしていない）",
        [r.status for r in dup_rows] == ["[重複]", "[重複]", ""] and all(len(r.display_cells()) == 4 for r in dup_rows),
    )

    decoy = rn / "一.pdf"
    decoy.mkdir()  # リネーム対象以外の既存名（ここでは同名のフォルダー）
    conflict_ok = False
    try:
        rename_core.check_rename_plan(rename_core.build_rename_plan(rn, rn_excel, "Sheet", "B1", modes.PER_CELL), rn)
    except FileExistsError as exc:
        conflict_ok = "衝突します" in str(exc) and "一.pdf" in str(exc)
    conflict_rows = preview.build_rename_preview(rn, rn_excel, "Sheet", "B1", modes.PER_CELL)
    check(
        "39. FR-R7 リネーム対象以外の既存名と衝突すると実行せず、プレビューは [衝突]",
        conflict_ok
        and any("[衝突]" in r.status for r in conflict_rows)
        and sorted(p.name for p in rn.iterdir()) == ["10_c.pdf", "2_a.pdf", "8_b.pdf", "一.pdf"],
    )
    decoy.rmdir()

    before_rows = preview.build_rename_preview(rn, rn_excel, "Sheet", "B1", modes.PER_CELL)
    check(
        "40. FR-R11 プレビューは 順序/現在のファイル名/新しいファイル名/状態 を全件出し、名前変更していない",
        [r.display_cells() for r in before_rows]
        == [
            ("1", "2_a.pdf", "一.pdf", ""),
            ("2", "8_b.pdf", "二.pdf", ""),
            ("3", "10_c.pdf", "三.pdf", ""),
        ]
        and sorted(p.name for p in rn.iterdir()) == ["10_c.pdf", "2_a.pdf", "8_b.pdf"],
    )

    first = rename_core.apply_renames(rename_core.build_rename_plan(rn, rn_excel, "Sheet", "B1", modes.PER_CELL))
    check(
        "41. FR-R12 実行結果の件数（変更 3 件 / 変更不要 0 件）と実際のファイル名",
        (first.renamed, first.unchanged) == (3, 0)
        and sorted(p.name for p in rn.iterdir()) == ["一.pdf", "三.pdf", "二.pdf"],
    )
    rn4 = root / "rename_again"
    rn4.mkdir()
    for name in ("10_c.pdf", "2_a.pdf", "8_b.pdf"):
        make_pdf(rn4 / name, 1)
    first_a = rename_core.apply_renames(rename_core.build_rename_plan(rn4, rn_excel, "Sheet", "A1", modes.BASE_SEQ))
    again_rows = preview.build_rename_preview(rn4, rn_excel, "Sheet", "A1", modes.BASE_SEQ)
    second = rename_core.apply_renames(rename_core.build_rename_plan(rn4, rn_excel, "Sheet", "A1", modes.BASE_SEQ))
    check(
        "42. FR-R8 現在名と新名が同じファイルは [変更不要] でスキップ（同じ名前で再実行できる）",
        (first_a.renamed, first_a.unchanged) == (3, 0)
        and all("[変更不要]" in r.status for r in again_rows)
        and (second.renamed, second.unchanged) == (0, 3)
        and sorted(p.name for p in rn4.iterdir())
        == ["報告書_001.pdf", "報告書_002.pdf", "報告書_003.pdf"],
    )
    check(
        "43. FR-R13 PDF・Excelの内容は一切変更しない（ハッシュ一致・名前のみ変更）",
        sorted(digest(rn).values()) == contents_before
        and hashlib.sha256(rn_excel.read_bytes()).hexdigest() == excel_before
        and temp_names(rn) == [],
    )

    rn2 = root / "rename_fail"
    rn2.mkdir()
    for name in ("1_a.pdf", "2_b.pdf", "3_c.pdf"):
        make_pdf(rn2 / name, 1)
    real_rename = os.rename
    calls = {"n": 0}

    def flaky_rename(source, target):
        calls["n"] += 1
        if calls["n"] == 4:  # ②最終名へ変える段階で失敗させる
            raise OSError("検証による意図的な失敗")
        return real_rename(source, target)

    os.rename = flaky_rename  # 名前変更は pathops 経由でも os.rename を呼ぶ
    fail_message = ""
    try:
        rename_core.apply_renames(rename_core.build_rename_plan(rn2, rn_excel, "Sheet", "A1", modes.BASE_SEQ))
    except ValueError as exc:
        fail_message = str(exc)
    finally:
        os.rename = real_rename
    check(
        "44. FR-R9 2段階リネームの②で失敗しても全ファイルを元名へ戻す（ロールバック）",
        "変更済みの名前は元に戻しました" in fail_message
        and sorted(p.name for p in rn2.iterdir()) == ["1_a.pdf", "2_b.pdf", "3_c.pdf"]
        and temp_names(rn2) == [],
    )

    rn3 = root / "rename_swap"
    rn3.mkdir()
    make_pdf(rn3 / "a.pdf", 1)
    make_pdf(rn3 / "b.pdf", 1)
    swap_excel = root / "rename_swap.xlsx"
    make_excel(swap_excel, {"A1": "報告書", "B1": "b", "B2": "a"})
    swap = rename_core.apply_renames(rename_core.build_rename_plan(rn3, swap_excel, "Sheet", "B1", modes.PER_CELL))
    check(
        "45. FR-R9 名前を入れ替えるリネームも2段階で成功する（変更 2 件）",
        (swap.renamed, swap.unchanged) == (2, 0)
        and sorted(p.name for p in rn3.iterdir()) == ["a.pdf", "b.pdf"]
        and temp_names(rn3) == [],
    )


def verify_copy(root: Path) -> None:
    """FR-R16〜FR-R18: 出力先へコピーしながら名前を変える（入力のPDFは無傷）."""
    src = root / "copy_in"
    src.mkdir()
    make_pdf(src / "2_a.pdf", 2)
    make_pdf(src / "8_b.pdf", 1)
    excel = root / "copy_names.xlsx"
    make_excel(excel, {"A1": "報告書"})
    out = root / "copy_out"
    out.mkdir()
    before = digest(src)

    def plan(target: Path):
        items = rename_core.build_rename_plan(src, excel, "Sheet", "A1", modes.BASE_SEQ)
        return rename_core.build_copy_plan(items, target)

    result = rename_core.apply_copies(plan(out))
    check(
        "46. FR-R16 コピー実行で入力のPDFは無傷、出力先に新しい名前が揃う",
        (result.copied, result.unchanged) == (2, 0)
        and digest(src) == before
        and sorted(p.name for p in out.iterdir()) == ["報告書_001.pdf", "報告書_002.pdf"]
        and temp_names(out) == [],
    )

    same = root / "copy_same"
    same.mkdir()
    make_pdf(same / "a.pdf", 1)
    make_excel(root / "copy_same_names.xlsx", {"A1": "報告書"})
    same_result = rename_core.apply_copies(
        rename_core.build_copy_plan(
            rename_core.build_rename_plan(same, root / "copy_same_names.xlsx", "Sheet", "A1", modes.BASE_SEQ), same,
        )
    )
    check(
        "47. FR-R16 入力と同じフォルダーへコピーすると元の名前も残る（名前だけ増える）",
        same_result.copied == 1
        and sorted(p.name for p in same.iterdir() if p.suffix == ".pdf") == ["a.pdf", "報告書_001.pdf"],
    )

    blocked = root / "copy_blocked"
    blocked.mkdir()
    make_pdf(blocked / "報告書_001.pdf", 1)
    blocked_message = ""
    try:
        rename_core.check_copy_plan(plan(blocked), blocked)
    except FileExistsError as exc:
        blocked_message = str(exc)
    check(
        "48. FR-R17 出力先に同じ名前が既にあると実行前に中止し、上書きしない",
        "出力先に同じ名前のファイルが既にあります" in blocked_message
        and sorted(p.name for p in blocked.iterdir()) == ["報告書_001.pdf"],
    )

    missing_folder = root / "copy_無いフォルダー"
    missing_message = ""
    try:
        rename_core.check_copy_plan(plan(missing_folder), missing_folder)
    except ValueError as exc:
        missing_message = str(exc)
    check(
        "49. FR-R17 出力先フォルダーが見つからない場合は実行前に案内する",
        "出力先フォルダーが見つかりません" in missing_message and not missing_folder.exists(),
    )

    check(
        "50. FR-R18 コピー後のファイルは元ファイルと同じサイズ・ページ数で読める",
        all(
            (out / name).stat().st_size == (src / original).stat().st_size
            and page_count(out / name) == page_count(src / original)
            for name, original in (("報告書_001.pdf", "2_a.pdf"), ("報告書_002.pdf", "8_b.pdf"))
        ),
    )

    vanishing_out = root / "copy_vanish"
    vanishing_out.mkdir()
    real_copy_into = pathops.copy_into

    def vanishing_copy_into(source: Path, destination: Path) -> None:
        real_copy_into(source, destination)
        os.remove(destination)  # 共有で「書いたのに残らない」真似

    pathops.copy_into = vanishing_copy_into
    vanish_message = ""
    try:
        rename_core.apply_copies(plan(vanishing_out))
    except ValueError as exc:
        vanish_message = str(exc)
    finally:
        pathops.copy_into = real_copy_into
    check(
        "51. FR-R18 コピーしたファイルが出力先に残らなければエラーで案内する（入力は無傷）",
        "出力先に正しく残っていません" in vanish_message
        and digest(src) == before
        and list(vanishing_out.iterdir()) == [],
    )


def verify_no_split(root: Path) -> None:
    """FR-R19: リネームタブはPDFを分割しない（PDF 1件 = 1ファイル、ページ数そのまま）."""
    inbox = root / "nosplit_in"
    inbox.mkdir()
    make_pdf(inbox / "multi.pdf", 3)
    make_pdf(inbox / "single.pdf", 1)
    excel = root / "nosplit_names.xlsx"
    make_excel(excel, {"A1": "報告書"})
    items = rename_core.build_rename_plan(inbox, excel, "Sheet", "A1", modes.BASE_SEQ)
    check(
        "52. FR-R19 リネーム計画はPDF 1件に1ファイル（ページ数ではなく件数で名前を作る）",
        [item.source.name for item in items] == ["multi.pdf", "single.pdf"],
    )
    rename_core.check_rename_plan(items, inbox)
    rename_core.apply_renames(items)
    check(
        "53. FR-R19 その場で改名でもファイル数と各PDFのページ数は変わらない（3ページ/1ページ）",
        sorted(p.name for p in inbox.glob("*.pdf")) == ["報告書_001.pdf", "報告書_002.pdf"]
        and page_count(inbox / "報告書_001.pdf") == 3
        and page_count(inbox / "報告書_002.pdf") == 1,
    )
    out = root / "nosplit_out"
    out.mkdir()
    copy_plan = rename_core.build_copy_plan(
        rename_core.build_rename_plan(inbox, excel, "Sheet", "A1", modes.BASE_SEQ), out
    )
    rename_core.check_copy_plan(copy_plan, out)
    rename_core.apply_copies(copy_plan)
    check(
        "54. FR-R19 出力先へコピーでもページ数はそのまま（3/1）で入力のPDFも無傷",
        page_count(out / "報告書_001.pdf") == 3
        and page_count(out / "報告書_002.pdf") == 1
        and page_count(inbox / "報告書_001.pdf") == 3,
    )
    rows = preview.build_rename_preview(inbox, excel, "Sheet", "A1", modes.BASE_SEQ)
    check(
        "55. FR-R19 プレビューはPDF 1件=1行（複数ページでも行数はページ数にならない）",
        len(rows) == 2 and [row.old_name for row in rows] == ["報告書_001.pdf", "報告書_002.pdf"],
    )


def main() -> int:
    root = Path(tempfile.mkdtemp(prefix=".pdf_split_verify_", dir=str(PROJECT_ROOT)))
    try:
        verify(root)
        verify_merge(root)
        verify_rename(root)
        verify_copy(root)
        verify_no_split(root)
    except Exception as exc:  # 想定外例外も合否として計上する
        check(f"検証スクリプトが想定外のエラーで終了した: {exc}", False)
    finally:
        shutil.rmtree(root, ignore_errors=True)

    succeeded = sum(1 for _label, ok in RESULTS if ok)
    for label, ok in RESULTS:
        print(("PASS: " if ok else "FAIL: ") + label)
    print(f"TOTAL: {succeeded}/{len(RESULTS)}")
    return 0 if succeeded == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())


