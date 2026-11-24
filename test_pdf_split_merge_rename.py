"""pdf-split-merge-rename の統合テスト（分割 FR-12 / 結合 FR-M / リネーム FR-R）.

参考コード2本（refs/PDF-Split-Rename, refs/PDF-Split-Rename-By-Cell）
の単体テストを移植し、共通コア・分割2モード・自然順・結合・リネーム・プレビューのテストを追加する。
実行: `py -3 -m unittest -v`
"""

from __future__ import annotations

import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import call, patch

from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

import merge_core
import modes
import natural_sort
import preview
import rename_core
import split_core


class FakeSheet:
    """openpyxlのシートを模す。既知の番地だけ値を返す."""

    def __init__(self, cells):
        self.cells = cells

    def __getitem__(self, address):
        return self.cells.get(address, SimpleNamespace(value=None, data_type="n"))


class FakeWorkbook:
    """load_workbook の戻り値を模すダブリング."""

    sheetnames = ["Sheet1"]

    def __init__(self, cells):
        self.sheet = FakeSheet(cells)
        self.closed = False

    def __getitem__(self, sheet_name):
        return self.sheet

    def close(self):
        self.closed = True


class RealDataMixin:
    """実データ（ページ数分の空白PDF・A1から並ぶExcel・空の出力先）を作る."""

    @staticmethod
    def make_data(root: Path, values=("あいう", "かきく", "さしす")):
        pdf_path = root / "input.pdf"
        writer = PdfWriter()
        for _ in values:
            writer.add_blank_page(width=72, height=72)
        with pdf_path.open("wb") as stream:
            writer.write(stream)

        excel_path = root / "names.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        for offset, value in enumerate(values):
            if value is not None:  # None はセルを作らない（空欄扱い）
                sheet[f"A{offset + 1}"] = value
        workbook.save(excel_path)
        workbook.close()

        output_dir = root / "output"
        output_dir.mkdir()
        return pdf_path, excel_path, output_dir


class ValidateFilenameBaseTests(unittest.TestCase):
    def test_ordinary_value_is_kept(self):
        self.assertEqual(split_core.validate_filename_base("ABC123", "A1"), "ABC123")

    def test_invalid_windows_characters_rejected(self):
        for bad in ("bad/name", 'bad:name', "bad*name", "bad?name", 'bad"name',
                    "bad<name", "bad>name", "bad|name", "bad\\name"):
            with self.subTest(bad=bad):
                with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
                    split_core.validate_filename_base(bad, "B2")

    def test_control_characters_rejected(self):
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            split_core.validate_filename_base("bad\x07name", "B2")

    def test_trailing_space_or_dot_rejected(self):
        with self.assertRaisesRegex(ValueError, "末尾に空白またはピリオド"):
            split_core.validate_filename_base("name. ", "A1")
        with self.assertRaisesRegex(ValueError, "末尾に空白またはピリオド"):
            split_core.validate_filename_base("name.", "A1")

    def test_windows_reserved_names_rejected_case_insensitively(self):
        for reserved in ("CON", "con", "PRN", "aux", "NUL", "com1", "COM9", "lpt1"):
            with self.subTest(reserved=reserved):
                with self.assertRaisesRegex(ValueError, "予約名"):
                    split_core.validate_filename_base(reserved, "A1")

    def test_names_starting_like_reserved_are_allowed(self):
        self.assertEqual(split_core.validate_filename_base("CONSOLE1", "A1"), "CONSOLE1")
        self.assertEqual(split_core.validate_filename_base("COM10", "A1"), "COM10")
class ParseCellAddressTests(unittest.TestCase):
    def test_plain_and_absolute_forms(self):
        self.assertEqual(split_core.parse_cell_address("a1"), ("A", 1))
        self.assertEqual(split_core.parse_cell_address("$B$3"), ("B", 3))
        self.assertEqual(split_core.parse_cell_address(" AA123 "), ("AA", 123))

    def test_invalid_forms_use_japanese_label(self):
        for bad in ("1A", "A0", "A-1", "", "A1B", "ABBBC1"):
            with self.subTest(bad=bad):
                with self.assertRaisesRegex(ValueError, "開始セルを正しく入力してください"):
                    split_core.parse_cell_address(bad, "開始セル")

    def test_label_follows_mode_wording(self):
        with self.assertRaisesRegex(ValueError, "セル番地を正しく入力してください"):
            split_core.parse_cell_address("!!", "セル番地")


class ReadCellValuesTests(unittest.TestCase):
    """参考コード2本の数式キャッシュ検査を移植（load_workbook は2回読みの順も確認）."""

    def read(self, formula_cells, value_cells, addresses=("A1",), messages=None):
        formula_book = FakeWorkbook(formula_cells)
        value_book = FakeWorkbook(value_cells)
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            kwargs = {} if messages is None else {"messages": messages}
            values = split_core.read_cell_values(Path("data.xlsx"), "Sheet1", addresses, **kwargs)
        self.assertEqual(loader.call_args_list, [
            call(Path("data.xlsx"), read_only=True, data_only=False),
            call(Path("data.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return values

    def test_ordinary_value_cells(self):
        cells = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.read(cells, cells, ("A1", "A2")), {"A1": "あいう", "A2": "かきく"})

    def test_cached_formula_results_are_used_without_calculation(self):
        formulas = {
            "A1": SimpleNamespace(value='="あいう"', data_type="f"),
            "A2": SimpleNamespace(value='="かきく"', data_type="f"),
        }
        values = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.read(formulas, values, ("A1", "A2")), {"A1": "あいう", "A2": "かきく"})

    def test_missing_formula_cache_prompts_excel_recalculation(self):
        formulas = {"A1": SimpleNamespace(value="=B1", data_type="f")}
        values = {"A1": SimpleNamespace(value=None, data_type="n")}
        with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
            self.read(formulas, values)

    def test_error_value_cell_rejected(self):
        formulas = {"A1": SimpleNamespace(value="=B1/0", data_type="f")}
        values = {"A1": SimpleNamespace(value="#DIV/0!", data_type="e")}
        with self.assertRaisesRegex(ValueError, "数式でエラーが発生しています"):
            self.read(formulas, values)

    def test_empty_cell_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "が空です"):
            self.read({}, {})

    def test_empty_cell_message_follows_style(self):
        with self.assertRaisesRegex(ValueError, "指定したセル（Sheet1!A1）が空です。ベースファイル名を入力してください。"):
            self.read({}, {}, messages=split_core.BASE_CELL_MESSAGES)

    def test_missing_sheet_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="x", data_type="s")}
        formula_book = FakeWorkbook(cells)
        value_book = FakeWorkbook(cells)
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]):
            with self.assertRaisesRegex(ValueError, "指定したシート「Other」がExcelブックにありません。"):
                split_core.read_cell_values(Path("data.xlsx"), "Other", ["A1"])

    def test_invalid_filename_value_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="bad/name", data_type="s")}
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            self.read(cells, cells)

    def test_unreadable_excel_raises_japanese_error(self):
        with patch.object(split_core, "load_workbook", side_effect=OSError("broken")):
            with self.assertRaisesRegex(ValueError, "Excelファイルを読み込めませんでした"):
                split_core.read_cell_values(Path("data.xlsx"), "Sheet1", ["A1"])
class PerCellModeTests(unittest.TestCase):
    """参考コード test_pdf_splitter_by_cell.py の読取テストを移植."""

    def build(self, formula_cells, value_cells, start="A1", page_count=2):
        formula_book = FakeWorkbook(formula_cells)
        value_book = FakeWorkbook(value_cells)
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            names = modes.build_filenames_per_cell(Path("data.xlsx"), "Sheet1", start, page_count)
        self.assertEqual(loader.call_args_list, [
            call(Path("data.xlsx"), read_only=True, data_only=False),
            call(Path("data.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return names

    def test_reads_downward_from_selected_start_cell_without_sequence(self):
        cells = {
            "B3": SimpleNamespace(value="あいう", data_type="s"),
            "B4": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.build(cells, cells, start="B3"), ["あいう.pdf", "かきく.pdf"])

    def test_uses_cached_formula_results(self):
        formulas = {
            "A1": SimpleNamespace(value='="あいう"', data_type="f"),
            "A2": SimpleNamespace(value='="かきく"', data_type="f"),
        }
        values = {
            "A1": SimpleNamespace(value="あいう", data_type="s"),
            "A2": SimpleNamespace(value="かきく", data_type="s"),
        }
        self.assertEqual(self.build(formulas, values), ["あいう.pdf", "かきく.pdf"])

    def test_missing_formula_cache_prompts_excel_recalculation(self):
        formulas = {"A1": SimpleNamespace(value="=B1", data_type="f")}
        values = {"A1": SimpleNamespace(value=None, data_type="n")}
        with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
            self.build(formulas, values, page_count=1)

    def test_empty_cell_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "が空です"):
            self.build({}, {}, page_count=1)

    def test_duplicate_names_are_rejected_case_insensitively(self):
        cells = {
            "A1": SimpleNamespace(value="Report", data_type="s"),
            "A2": SimpleNamespace(value="report", data_type="s"),
        }
        with self.assertRaisesRegex(ValueError, "同じファイル名"):
            self.build(cells, cells)

    def test_invalid_windows_filename_is_rejected(self):
        cells = {"A1": SimpleNamespace(value="bad/name", data_type="s")}
        with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
            self.build(cells, cells, page_count=1)

    def test_reading_past_excel_last_row_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Excelの最終行を超えます"):
            modes.build_filenames_per_cell(Path("data.xlsx"), "Sheet1", "A1048576", 2)

    def test_page_count_below_one_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "PDFにページがありません"):
            modes.build_filenames_per_cell(Path("data.xlsx"), "Sheet1", "A1", 0)


class BaseSeqModeTests(unittest.TestCase):
    """参考コード test_pdf_splitter.py の数式キャッシュ検査を移植."""

    def build(self, formula_cell, cached_cell, page_count=2, cell="A1"):
        formula_book = FakeWorkbook({"A1": formula_cell})
        value_book = FakeWorkbook({"A1": cached_cell})
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]) as loader:
            names = modes.build_filenames_base_seq(Path("book.xlsx"), "Sheet1", cell, page_count)
        self.assertEqual(loader.call_args_list, [
            call(Path("book.xlsx"), read_only=True, data_only=False),
            call(Path("book.xlsx"), read_only=True, data_only=True),
        ])
        self.assertTrue(formula_book.closed)
        self.assertTrue(value_book.closed)
        return names

    def test_ordinary_value_cell_keeps_using_its_value(self):
        names = self.build(
            SimpleNamespace(value="ordinary-name", data_type="s"),
            SimpleNamespace(value="ordinary-name", data_type="s"),
            page_count=3,
        )
        self.assertEqual(names, [
            "ordinary-name_001.pdf", "ordinary-name_002.pdf", "ordinary-name_003.pdf",
        ])

    def test_formula_cell_uses_saved_cached_value(self):
        names = self.build(
            SimpleNamespace(value='=A2&"-name"', data_type="f"),
            SimpleNamespace(value="cached-name", data_type="s"),
            page_count=1,
        )
        self.assertEqual(names, ["cached-name_001.pdf"])

    def test_formula_without_cached_value_explains_recalculation(self):
        with self.assertRaisesRegex(ValueError, "Excelでブックを開き、数式を再計算してから保存"):
            self.build(SimpleNamespace(value="=A2", data_type="f"), SimpleNamespace(value=None), page_count=1)

    def test_sequence_digit_grows_beyond_999(self):
        names = self.build(
            SimpleNamespace(value="x", data_type="s"),
            SimpleNamespace(value="x", data_type="s"),
            page_count=1001,
        )
        self.assertEqual(names[999], "x_1000.pdf")
        self.assertEqual(names[1000], "x_1001.pdf")

    def test_address_is_normalized_before_reading(self):
        names = self.build(
            SimpleNamespace(value="x", data_type="s"),
            SimpleNamespace(value="x", data_type="s"),
            page_count=1, cell="$a$1",
        )
        self.assertEqual(names, ["x_001.pdf"])

    def test_empty_cell_message_keeps_base_mode_wording(self):
        with self.assertRaisesRegex(ValueError, "指定したセル（Sheet1!A1）が空です。ベースファイル名を入力してください。"):
            self.build(
                SimpleNamespace(value=None, data_type="n"),
                SimpleNamespace(value=None, data_type="n"),
                page_count=1,
            )

    def test_missing_sheet_keeps_base_mode_wording(self):
        formula_book = FakeWorkbook({})
        value_book = FakeWorkbook({})
        with patch.object(split_core, "load_workbook", side_effect=[formula_book, value_book]):
            with self.assertRaisesRegex(ValueError, "指定したシート「Other」がExcelにありません。"):
                modes.build_filenames_base_seq(Path("book.xlsx"), "Other", "A1", 1)

    def test_page_count_below_one_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "PDFにページがありません"):
            modes.build_filenames_base_seq(Path("book.xlsx"), "Sheet1", "A1", 0)

    def test_dispatcher_unknown_mode(self):
        with self.assertRaisesRegex(ValueError, "未対応のモードです"):
            modes.build_filenames("no-such-mode", Path("book.xlsx"), "Sheet1", "A1", 1)
class LoadPdfTests(RealDataMixin, unittest.TestCase):
    def make_pdf(self, root: Path) -> Path:
        pdf_path, _, _ = self.make_data(root)
        return pdf_path

    def test_encrypted_pdf_requires_unlocked_file(self):
        with TemporaryDirectory() as temporary:
            path = Path(temporary) / "secret.pdf"
            writer = PdfWriter()
            writer.add_blank_page(width=72, height=72)
            writer.encrypt("hunter2")
            with path.open("wb") as stream:
                writer.write(stream)
            with self.assertRaisesRegex(ValueError, "パスワードで保護されています"):
                split_core.load_pdf(path)

    def test_broken_pdf_raises_japanese_error(self):
        with TemporaryDirectory() as temporary:
            path = Path(temporary) / "broken.pdf"
            path.write_bytes(b"this is not a pdf")
            with self.assertRaisesRegex(ValueError, "PDFを読み込めませんでした"):
                split_core.load_pdf(path)

    def test_zero_page_pdf_is_rejected(self):
        with TemporaryDirectory() as temporary:
            path = Path(temporary) / "empty.pdf"
            with path.open("wb") as stream:
                PdfWriter().write(stream)
            with self.assertRaisesRegex(ValueError, "PDFにページがありません"):
                split_core.load_pdf(path)

    def test_input_path_checks(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            missing = root / "nope.pdf"
            notes = root / "notes.txt"
            notes.write_text("plain text", encoding="utf-8")
            cases = [
                (missing, excel_path, output_dir, "指定したPDFファイルが見つかりません"),
                (pdf_path, root / "nope.xlsx", output_dir, "指定したExcelファイルが見つかりません"),
                (excel_path, excel_path, output_dir, "PDFファイルを選択してください"),
                (pdf_path, notes, output_dir, "Excelファイル（.xlsx または .xlsm）を選択してください"),
                (pdf_path, excel_path, root / "no-dir", "出力先フォルダーが見つかりません"),
            ]
            for pdf, excel, out, message in cases:
                with self.subTest(pdf=pdf, excel=excel, out=out):
                    with self.assertRaisesRegex(ValueError, message):
                        split_core.check_input_paths(pdf, excel, out)


class WritePagesTests(RealDataMixin, unittest.TestCase):
    def make_reader_and_names(self, root: Path, names: list[str]):
        pdf_path, _, _ = self.make_data(root)
        reader = split_core.load_pdf(pdf_path)
        return reader, names

    def test_existing_output_blocks_without_overwrite(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            output_dir = root / "output"
            guard = output_dir / "かきく.pdf"
            guard.write_bytes(b"ORIGINAL")
            with self.assertRaisesRegex(FileExistsError, "同名の出力ファイルが既にあります。上書きは行いません"):
                split_core.write_pages(reader, names, output_dir)
            self.assertEqual(guard.read_bytes(), b"ORIGINAL")
            self.assertFalse((output_dir / "あいう.pdf").exists())
            self.assertFalse((output_dir / "さしす.pdf").exists())
            self.assertEqual([p.name for p in output_dir.iterdir()], ["かきく.pdf"])  # 一時フォルダーも残さない

    def test_existing_list_is_capped_with_remainder_count(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            names = [f"file{i}.pdf" for i in range(1, 11)]
            pdf_path = root / "input.pdf"
            writer = PdfWriter()
            for _ in names:
                writer.add_blank_page(width=72, height=72)
            with pdf_path.open("wb") as stream:
                writer.write(stream)
            reader = split_core.load_pdf(pdf_path)
            output_dir = root / "output"
            output_dir.mkdir()
            for name in names[:9]:
                (output_dir / name).write_bytes(b"x")
            with self.assertRaisesRegex(FileExistsError, "ほか 1 件"):
                split_core.write_pages(reader, names, output_dir)

    def test_page_name_count_mismatch_is_rejected(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["only.pdf"])
            output_dir = root / "output"
            with self.assertRaisesRegex(ValueError, "ページ数とファイル名リストの数が一致しません"):
                split_core.write_pages(reader, names, output_dir)

    def test_staging_failure_leaves_nothing_behind(self):
        class FailingWriter:
            def __init__(self):
                pass

            def add_page(self, page):
                pass

            def write(self, stream):
                raise OSError("simulated disk failure")

        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["a.pdf", "b.pdf", "c.pdf"])
            output_dir = root / "output"
            with patch.object(split_core, "PdfWriter", FailingWriter):
                with self.assertRaisesRegex(ValueError, "分割ファイルを保存できませんでした"):
                    split_core.write_pages(reader, names, output_dir)
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_move_failure_rolls_back_created_outputs(self):
        real_rename = os.rename
        state = {"calls": 0}

        def flaky_rename(src, dst):
            state["calls"] += 1
            if state["calls"] == 1:
                real_rename(src, dst)
            else:
                raise OSError("simulated move failure")

        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            reader, names = self.make_reader_and_names(root, ["a.pdf", "b.pdf", "c.pdf"])
            output_dir = root / "output"
            with patch("os.rename", side_effect=flaky_rename):
                with self.assertRaisesRegex(ValueError, "分割ファイルを保存できませんでした"):
                    split_core.write_pages(reader, names, output_dir)
            self.assertGreaterEqual(state["calls"], 2)
            self.assertEqual(list(output_dir.iterdir()), [])  # 移動済み1件も消えて片付いている
class SplitEndToEndTests(RealDataMixin, unittest.TestCase):
    def check_one_page_each(self, results):
        self.assertEqual([len(PdfReader(str(path)).pages) for path in results], [1, 1, 1])

    def test_base_seq_split_end_to_end(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            before = (pdf_path.read_bytes(), excel_path.read_bytes())
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_base_seq,
            )
            self.assertEqual([p.name for p in results], ["あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf"])
            self.check_one_page_each(results)
            self.assertEqual(sorted(p.name for p in output_dir.iterdir()), [
                "あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf",
            ])
            self.assertEqual((pdf_path.read_bytes(), excel_path.read_bytes()), before)  # 元ファイルは無変更

    def test_per_cell_split_end_to_end(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_per_cell,
            )
            self.assertEqual([p.name for p in results], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            self.check_one_page_each(results)

    def test_dispatcher_mode_selection_same_names(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, "Sheet", "A1",
                lambda excel, sheet, cell, page_count: modes.build_filenames(
                    modes.PER_CELL, excel, sheet, cell, page_count,
                ),
            )
            self.assertEqual([p.name for p in results], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])

    def test_blank_cell_stops_before_any_output(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("あいう", None, "さしす"))
            with self.assertRaisesRegex(ValueError, "が空です"):
                split_core.split_pdf(
                    pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_per_cell,
                )
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_duplicate_names_stop_before_any_output(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("Report", "report", "third"))
            with self.assertRaisesRegex(ValueError, "同じファイル名"):
                split_core.split_pdf(
                    pdf_path, excel_path, output_dir, "Sheet", "A1", modes.build_filenames_per_cell,
                )
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_mode_labels_match_specification(self):
        self.assertEqual(modes.MODE_LABELS[modes.BASE_SEQ], "セル値＋連番")
        self.assertEqual(modes.MODE_LABELS[modes.PER_CELL], "セル別（連番なし）")


class PreviewTests(RealDataMixin, unittest.TestCase):
    def test_existing_files_marked_and_nothing_written(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            guard = output_dir / "かきく.pdf"
            guard.write_bytes(b"KEEP")

            rows = preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, output_dir, pdf_path)

            self.assertEqual([row.page_no for row in rows], [1, 2, 3])
            self.assertEqual([row.filename for row in rows], ["あいう.pdf", "かきく.pdf", "さしす.pdf"])
            self.assertEqual([row.status for row in rows], ["", "[既存]", ""])
            self.assertEqual(guard.read_bytes(), b"KEEP")  # 既存ファイルは無変更
            self.assertEqual(sorted(p.name for p in output_dir.iterdir()), ["かきく.pdf"])  # 書き出しなし

    def test_duplicate_rows_marked_without_error(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("Report", "report", "third"))
            rows = preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, output_dir, pdf_path)
            self.assertEqual([row.status for row in rows], ["[重複]", "[重複]", ""])
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_base_seq_preview_names_and_no_duplicate_marks(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root)
            rows = preview.build_preview(excel_path, "Sheet", "A1", modes.BASE_SEQ, output_dir, pdf_path)
            self.assertEqual([row.filename for row in rows], [
                "あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf",
            ])
            self.assertTrue(all(not row.duplicate and not row.exists for row in rows))
            self.assertEqual(list(output_dir.iterdir()), [])

    def test_invalid_cells_propagate_japanese_errors(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf_path, excel_path, output_dir = self.make_data(root, values=("bad/name", "かきく", "さしす"))
            with self.assertRaisesRegex(ValueError, "ファイル名に使えない文字"):
                preview.build_preview(excel_path, "Sheet", "A1", modes.PER_CELL, output_dir, pdf_path)

    def test_truncate_display_keeps_all_data_and_announces_remainder(self):
        rows = [preview.PreviewRow(page_no, f"f{page_no}.pdf", False, False) for page_no in range(1, 151)]
        shown, extra = preview.truncate_display(rows)
        self.assertEqual(len(shown), preview.MAX_ANNOUNCED)
        self.assertEqual(extra, 50)
        self.assertEqual(rows[149].display_cells(), ("150", "f150.pdf", ""))
        short, extra2 = preview.truncate_display(rows[:preview.MAX_ANNOUNCED])
        self.assertEqual(extra2, 0)

    def test_status_joins_both_marks(self):
        row = preview.PreviewRow(1, "a.pdf", exists=True, duplicate=True)
        self.assertEqual(row.status, "[既存] [重複]")


def make_blank_pdf(path: Path, pages: int = 1) -> None:
    """ページ数分の空白ページだけを持つPDFを作る（結合・リネーム検証用）."""
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=72, height=72)
    with path.open("wb") as stream:
        writer.write(stream)


def make_marked_pdf(path: Path, marks) -> None:
    """各ページに英字マーカーを描いたPDFを作る（結合順の検証用）."""
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


def make_excel(path: Path, cells, sheet_title: str = "Sheet") -> None:
    """番地→値の辞書から実Excelを作る（結合・リネーム検証用）."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_title
    for address, value in cells.items():
        sheet[address] = value
    workbook.save(path)


class NaturalSortTests(unittest.TestCase):
    """FR-M2/FR-R3: 自然順（数字の連なりは数値として比べる）."""

    def test_digit_runs_compare_as_numbers(self):
        self.assertEqual(
            natural_sort.sort_pdf_names(["10.pdf", "9.pdf", "008.pdf", "003.pdf", "001.pdf"]),
            ["001.pdf", "003.pdf", "008.pdf", "9.pdf", "10.pdf"],
        )

    def test_text_parts_ignore_case(self):
        self.assertEqual(
            natural_sort.sort_pdf_names(["scan010.pdf", "Scan01.pdf", "scan002.pdf"]),
            ["Scan01.pdf", "scan002.pdf", "scan010.pdf"],
        )

    def test_non_pdf_names_are_excluded(self):
        self.assertEqual(
            natural_sort.sort_pdf_names(["notes.txt", "b.pdf", "report.PDF"]),
            ["b.pdf", "report.PDF"],
        )

    def test_text_blocks_come_before_digit_blocks(self):
        # 3.4 の実装（数字連なり=(1,数値,"")、それ以外=(0,0,テキスト)）どおりの順序
        self.assertEqual(natural_sort.sort_pdf_names(["9.pdf", "a.pdf"]), ["a.pdf", "9.pdf"])

    def test_same_rank_is_decided_deterministically(self):
        self.assertEqual(natural_sort.sort_pdf_names(["a1.pdf", "A1.pdf"]), ["A1.pdf", "a1.pdf"])

    def test_key_shape_is_uniform(self):
        self.assertEqual(
            natural_sort.natural_sort_key("A2"),
            ((0, 0, "a"), (1, 2, ""), (2, 0, "a2"), (3, 0, "A2")),
        )

    def test_extremely_long_digit_run_falls_back_to_text(self):
        long_nine = "9" * 40
        one_then_zeros = "1" + "0" * 40
        self.assertEqual(
            natural_sort.sort_pdf_names([f"{long_nine}.pdf", f"{one_then_zeros}.pdf"]),
            [f"{one_then_zeros}.pdf", f"{long_nine}.pdf"],
        )
        self.assertIn((0, 0, long_nine), natural_sort.natural_sort_key(long_nine))


class MergeDataMixin:
    """結合検証用のフォルダー構成（入力 inbox・出力先 outbox）を作る."""

    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.root = Path(self._tmp.name)
        self.inbox = self.root / "inbox"
        self.inbox.mkdir()
        self.outbox = self.root / "outbox"
        self.outbox.mkdir()
        self.excel = self.root / "names.xlsx"

    def add_pdf(self, name: str, pages: int = 1) -> Path:
        """入力フォルダーへ空白PDFを追加する."""
        path = self.inbox / name
        make_blank_pdf(path, pages)
        return path

    def leftover_temp_dirs(self):
        """作成後に残った一時フォルダーの一覧（結合・分割・リネーム共通）."""
        return [p.name for p in self.root.iterdir() if merge_core.is_temp_name(p.name)]


class MergeListingTests(MergeDataMixin, unittest.TestCase):
    """FR-M1〜FR-M4: 入力PDFの収集と結合後ファイル名の既定値."""

    def test_sources_are_listed_in_natural_order(self):
        for name in ("10.pdf", "9.pdf", "008.pdf", "001.pdf", "003.pdf"):
            self.add_pdf(name)
        (self.inbox / "notes.txt").write_text("対象外", encoding="utf-8")
        (self.inbox / "sub").mkdir()
        make_blank_pdf(self.inbox / "sub" / "000.pdf")
        (self.inbox / ".pdf_split_staging").mkdir()
        make_blank_pdf(self.inbox / ".pdf_split_staging" / "000.pdf")
        self.assertEqual(
            [p.name for p in merge_core.list_merge_sources(self.inbox)],
            ["001.pdf", "003.pdf", "008.pdf", "9.pdf", "10.pdf"],
        )

    def test_pdf_extension_is_case_insensitive(self):
        self.add_pdf("scan.PDF")
        self.add_pdf("memo.pdf")
        self.assertEqual([p.name for p in merge_core.list_merge_sources(self.inbox)], ["memo.pdf", "scan.PDF"])

    def test_no_pdf_is_reported(self):
        with self.assertRaises(ValueError) as caught:
            merge_core.list_merge_sources(self.inbox)
        self.assertIn("PDFが見つかりません", str(caught.exception))

    def test_single_pdf_is_reported(self):
        self.add_pdf("only.pdf")
        with self.assertRaises(ValueError) as caught:
            merge_core.list_merge_sources(self.inbox)
        self.assertEqual(str(caught.exception), "結合には2つ以上のPDFが必要です。")

    def test_missing_folder_is_reported(self):
        with self.assertRaises(ValueError) as caught:
            merge_core.list_merge_sources(self.root / "nowhere")
        self.assertIn("入力フォルダーが見つかりません", str(caught.exception))

    def test_default_output_name_is_first_source_stem(self):
        for name in ("10.pdf", "001.pdf"):
            self.add_pdf(name)
        sources = merge_core.list_merge_sources(self.inbox)
        self.assertEqual(merge_core.default_output_name(sources), "001")

    def test_output_name_gets_pdf_extension(self):
        self.assertEqual(merge_core.validate_output_name(" 結合結果 "), "結合結果.pdf")
        self.assertEqual(merge_core.validate_output_name("結合結果.pdf"), "結合結果.pdf")

    def test_output_name_rejects_other_extension_and_bad_characters(self):
        with self.assertRaises(ValueError) as caught:
            merge_core.validate_output_name("結果.txt")
        self.assertIn("拡張子", str(caught.exception))
        with self.assertRaises(ValueError):
            merge_core.validate_output_name("bad/name")
        with self.assertRaises(ValueError):
            merge_core.validate_output_name("   ")


class MergePlanTests(MergeDataMixin, unittest.TestCase):
    """FR-M7〜FR-M10: 書き出し前の検査（全件開く・同名・既存）."""

    def test_plan_counts_pages_and_target_path(self):
        self.add_pdf("001.pdf", 2)
        self.add_pdf("003.pdf", 3)
        plan = merge_core.build_merge_plan(self.inbox, "結合結果", self.outbox)
        self.assertEqual([p.name for p in plan.sources], ["001.pdf", "003.pdf"])
        self.assertEqual(plan.page_counts, (2, 3))
        self.assertEqual(plan.total_pages, 5)
        self.assertEqual(plan.output_path, self.outbox / "結合結果.pdf")

    def test_output_name_equal_to_input_pdf_is_rejected(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        with self.assertRaises(ValueError) as caught:
            merge_core.build_merge_plan(self.inbox, "001.PDF", self.outbox)
        self.assertIn("入力PDFを出力名にできません", str(caught.exception))

    def test_existing_output_is_rejected_without_overwrite(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        target = self.outbox / "merged.pdf"
        target.write_bytes(b"OLD")
        with self.assertRaises(FileExistsError) as caught:
            merge_core.build_merge_plan(self.inbox, "merged", self.outbox)
        self.assertIn("上書きは行いません", str(caught.exception))
        self.assertEqual(target.read_bytes(), b"OLD")

    def test_missing_output_folder_is_rejected(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        with self.assertRaises(ValueError) as caught:
            merge_core.build_merge_plan(self.inbox, "out", self.root / "nowhere")
        self.assertIn("出力先フォルダーが見つかりません", str(caught.exception))

    def test_unreadable_pdfs_are_collected_before_writing(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        (self.inbox / "bad.pdf").write_bytes(b"not a pdf at all")
        encrypted = self.inbox / "enc.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=72, height=72)
        writer.encrypt("secret")
        with encrypted.open("wb") as stream:
            writer.write(stream)
        with self.assertRaises(ValueError) as caught:
            merge_core.build_merge_plan(self.inbox, "out", self.outbox)
        message = str(caught.exception)
        self.assertIn("読み込めないPDFがあります", message)
        self.assertIn("bad.pdf", message)
        self.assertIn("enc.pdf", message)
        self.assertEqual(list(self.outbox.iterdir()), [])

    def test_unreadable_pdfs_are_summarised_after_eight(self):
        self.add_pdf("001.pdf")
        for index in range(9):
            (self.inbox / f"bad{index}.pdf").write_bytes(b"broken")
        with self.assertRaises(ValueError) as caught:
            merge_core.build_merge_plan(self.inbox, "out", self.outbox)
        self.assertIn("ほか 1 件", str(caught.exception))


class MergeExecutionTests(MergeDataMixin, unittest.TestCase):
    """FR-M11〜FR-M12: 一時フォルダー経由の書き出しとロールバック."""

    def test_pages_are_merged_in_natural_order(self):
        make_marked_pdf(self.inbox / "9.pdf", ["NINE"])
        make_marked_pdf(self.inbox / "10.pdf", ["TEN"])
        make_marked_pdf(self.inbox / "001.pdf", ["ONE", "ONE2"])
        plan = merge_core.build_merge_plan(self.inbox, "結合結果", self.outbox)
        saved = merge_core.merge_pdfs(plan)
        self.assertEqual(saved.name, "結合結果.pdf")
        reader = PdfReader(str(saved))
        self.assertEqual(len(reader.pages), 4)
        texts = [page.extract_text() for page in reader.pages]
        self.assertEqual(texts, ["ONE", "ONE2", "NINE", "TEN"])

    def test_no_temp_folder_is_left_behind(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        merge_core.merge_pdfs(merge_core.build_merge_plan(self.inbox, "out", self.outbox))
        self.assertEqual(sorted(p.name for p in self.outbox.iterdir()), ["out.pdf"])
        self.assertEqual(self.leftover_temp_dirs(), [])

    def test_output_is_not_created_when_writing_fails(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")

        class FailingWriter:
            def __init__(self):
                self.pages = []

            def add_page(self, page):
                self.pages.append(page)

            def write(self, stream):
                raise OSError("ディスク書き込み失敗（模擬）")

        with patch.object(merge_core, "PdfWriter", FailingWriter):
            with self.assertRaises(ValueError) as caught:
                merge_core.merge_pdfs(merge_core.build_merge_plan(self.inbox, "out", self.outbox))
        self.assertIn("結合ファイルを保存できませんでした", str(caught.exception))
        self.assertEqual(list(self.outbox.iterdir()), [])
        self.assertEqual(self.leftover_temp_dirs(), [])

    def test_input_pdfs_are_not_modified(self):
        self.add_pdf("001.pdf", 2)
        self.add_pdf("003.pdf", 3)
        before = {p.name: p.read_bytes() for p in self.inbox.iterdir()}
        merge_core.merge_pdfs(merge_core.build_merge_plan(self.inbox, "out", self.outbox))
        self.assertEqual({p.name: p.read_bytes() for p in self.inbox.iterdir()}, before)


class RenameDataMixin:
    """リネーム検証用のフォルダー・Excel構成を作る."""

    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.root = Path(self._tmp.name)
        self.inbox = self.root / "inbox"
        self.inbox.mkdir()
        self.excel = self.root / "names.xlsx"

    def add_pdf(self, name: str, pages: int = 1) -> Path:
        """リネーム対象のPDFを追加する."""
        path = self.inbox / name
        make_blank_pdf(path, pages)
        return path


class RenamePlanTests(RenameDataMixin, unittest.TestCase):
    """FR-R1〜FR-R8: 命名計画（自然順・2ルール・衝突検査）."""

    SCATTERED = ("10_末尾.pdf", "9_中.pdf", "008.pdf", "001.pdf", "003.pdf")  # 作成順（わざとばらばら）
    NATURAL = ("001.pdf", "003.pdf", "008.pdf", "9_中.pdf", "10_末尾.pdf")  # 期待する自然順

    def test_targets_are_listed_in_natural_order(self):
        for name in self.SCATTERED:
            self.add_pdf(name)
        (self.inbox / "notes.txt").write_text("対象外", encoding="utf-8")
        self.assertEqual(
            [p.name for p in rename_core.list_rename_targets(self.inbox)],
            list(self.NATURAL),
        )

    def test_no_pdf_is_reported(self):
        with self.assertRaises(ValueError) as caught:
            rename_core.list_rename_targets(self.inbox)
        self.assertIn("PDFが見つかりません", str(caught.exception))

    def test_base_seq_rule_numbers_in_natural_order(self):
        for name in self.SCATTERED:
            self.add_pdf(name)
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        self.assertEqual([it.source.name for it in items], list(self.NATURAL))
        self.assertEqual(
            [it.new_name for it in items],
            ["報告書_001.pdf", "報告書_002.pdf", "報告書_003.pdf", "報告書_004.pdf", "報告書_005.pdf"],
        )
        self.assertEqual([it.unchanged for it in items], [False] * 5)

    def test_per_cell_rule_uses_cells_below_start(self):
        for name in self.SCATTERED:
            self.add_pdf(name)
        make_excel(self.excel, {"B1": "表紙", "B2": "本文", "B3": "付録", "B4": "索引", "B5": "奥付"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "B1", modes.PER_CELL)
        self.assertEqual(
            [it.new_name for it in items],
            ["表紙.pdf", "本文.pdf", "付録.pdf", "索引.pdf", "奥付.pdf"],
        )

    def test_duplicate_new_names_are_aborted(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        make_excel(self.excel, {"B1": "同じ", "B2": "同じ"})
        with self.assertRaises(ValueError) as caught:
            rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "B1", modes.PER_CELL)
        self.assertIn("同じファイル名", str(caught.exception))

    def test_unchanged_items_are_flagged(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        rename_core.apply_renames(items)
        again = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        self.assertEqual([it.new_name for it in again], ["報告書_001.pdf", "報告書_002.pdf"])
        self.assertEqual([it.unchanged for it in again], [True, True])

    def test_conflict_with_existing_non_target_is_reported(self):
        for name in self.SCATTERED:
            self.add_pdf(name)
        make_excel(self.excel, {"A1": "報告書"})
        # 同名の「フォルダー」はリネーム対象ではないため衝突になる
        (self.inbox / "報告書_003.pdf").mkdir()
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        with self.assertRaises(FileExistsError) as caught:
            rename_core.check_rename_plan(items, self.inbox)
        self.assertIn("衝突します", str(caught.exception))
        self.assertIn("報告書_003.pdf", str(caught.exception))
        self.assertEqual({p.name for p in self.inbox.iterdir() if p.is_file()}, set(self.NATURAL))

    def test_conflicts_are_summarised_after_eight(self):
        items = []
        for index in range(10):
            source = self.add_pdf(f"src_{index:03d}.pdf")
            (self.inbox / f"既存_{index:03d}.pdf").write_bytes(b"x")
            items.append(
                rename_core.RenameItem(
                    order=index + 1, source=source, new_name=f"既存_{index:03d}.pdf", unchanged=False
                )
            )
        with self.assertRaises(FileExistsError) as caught:
            rename_core.check_rename_plan(items, self.inbox)
        self.assertIn("ほか 2 件", str(caught.exception))

    def test_preview_style_plan_keeps_duplicates(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        make_excel(self.excel, {"B1": "同じ", "B2": "同じ"})
        items = rename_core.build_rename_plan(
            self.inbox, self.excel, "Sheet", "B1", modes.PER_CELL, abort_on_duplicates=False
        )
        self.assertEqual([it.new_name for it in items], ["同じ.pdf", "同じ.pdf"])


class RenameExecutionTests(RenameDataMixin, unittest.TestCase):
    """FR-R9〜FR-R11: 2段階リネーム・ロールバック・残骸掃除."""

    def test_names_are_changed_and_no_temp_remains(self):
        for name in ("10.pdf", "9.pdf", "001.pdf"):
            self.add_pdf(name)
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        result = rename_core.apply_renames(items)
        self.assertEqual((result.renamed, result.unchanged), (3, 0))
        self.assertEqual(
            sorted(p.name for p in self.inbox.iterdir()),
            ["報告書_001.pdf", "報告書_002.pdf", "報告書_003.pdf"],
        )

    def test_pdf_content_is_preserved(self):
        first = self.add_pdf("001.pdf", 2)
        self.add_pdf("003.pdf", 3)
        digest = first.read_bytes()
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        rename_core.apply_renames(items)
        self.assertEqual((self.inbox / "報告書_001.pdf").read_bytes(), digest)

    def test_swapping_two_names_works(self):
        self.add_pdf("a.pdf")
        self.add_pdf("b.pdf")
        make_excel(self.excel, {"B1": "b", "B2": "a"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "B1", modes.PER_CELL)
        result = rename_core.apply_renames(items)
        self.assertEqual((result.renamed, result.unchanged), (2, 0))
        self.assertEqual(sorted(p.name for p in self.inbox.iterdir()), ["a.pdf", "b.pdf"])
        self.assertEqual(len(PdfReader(str(self.inbox / "a.pdf")).pages), 1)

    def test_rerun_reports_all_unchanged(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        make_excel(self.excel, {"A1": "報告書"})
        rename_core.apply_renames(rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ))
        result = rename_core.apply_renames(
            rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        )
        self.assertEqual((result.renamed, result.unchanged), (0, 2))

    def _fail_rename_at(self, fail_at: int):
        """n回目の os.rename を失敗させる（ロールバック検証用）."""
        real_rename = os.rename
        state = {"calls": 0}

        def fake_rename(src, dst):
            state["calls"] += 1
            if state["calls"] == fail_at:
                raise OSError("名前変更失敗（模擬）")
            return real_rename(src, dst)

        return patch("rename_core.os.rename", side_effect=fake_rename), state

    def test_failure_in_stage_one_rolls_back(self):
        for name in ("001.pdf", "003.pdf", "008.pdf"):
            self.add_pdf(name)
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        patcher, _state = self._fail_rename_at(2)  # 2段階目の一時リネームで失敗
        with patcher:
            with self.assertRaises(ValueError) as caught:
                rename_core.apply_renames(items)
        self.assertIn("変更済みの名前は元に戻しました", str(caught.exception))
        self.assertEqual(sorted(p.name for p in self.inbox.iterdir()), ["001.pdf", "003.pdf", "008.pdf"])

    def test_failure_in_stage_two_rolls_back(self):
        for name in ("001.pdf", "003.pdf", "008.pdf"):
            self.add_pdf(name)
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        patcher, _state = self._fail_rename_at(5)  # 3件中2件目の最終リネームで失敗
        with patcher:
            with self.assertRaises(ValueError) as caught:
                rename_core.apply_renames(items)
        self.assertIn("変更済みの名前は元に戻しました", str(caught.exception))
        self.assertEqual(sorted(p.name for p in self.inbox.iterdir()), ["001.pdf", "003.pdf", "008.pdf"])

    def test_no_temp_names_remain_after_failure(self):
        for name in ("001.pdf", "003.pdf", "008.pdf"):
            self.add_pdf(name)
        make_excel(self.excel, {"A1": "報告書"})
        items = rename_core.build_rename_plan(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        patcher, _state = self._fail_rename_at(4)
        with patcher:
            with self.assertRaises(ValueError):
                rename_core.apply_renames(items)
        self.assertEqual([p.name for p in self.inbox.iterdir() if merge_core.is_temp_name(p.name)], [])


class MergeRenamePreviewTests(MergeDataMixin, unittest.TestCase):
    """FR-M13/FR-R11: プレビュー一覧（書き出し・名前変更を行わない）."""

    def test_merge_preview_marks_unreadable_and_existing(self):
        self.add_pdf("001.pdf", 2)
        self.add_pdf("003.pdf", 3)
        (self.inbox / "999_bad.pdf").write_bytes(b"broken")
        rows = preview.build_merge_preview(self.inbox, "out", self.outbox)
        self.assertEqual([r.filename for r in rows], ["001.pdf", "003.pdf", "999_bad.pdf"])
        self.assertEqual([r.page_count for r in rows], [2, 3, 0])
        self.assertEqual(rows[2].status, "[読込不可]")
        self.assertEqual([r.status for r in rows[:2]], ["", ""])
        self.assertEqual(rows[0].display_cells(), ("1", "001.pdf", "2", ""))
        self.assertEqual(rows[2].display_cells(), ("3", "999_bad.pdf", "-", "[読込不可]"))
        self.assertEqual(list(self.outbox.iterdir()), [])
        existing = preview.build_merge_preview(self.inbox, "001", self.inbox)
        self.assertTrue(all(r.exists for r in existing))

    def test_merge_preview_rejects_unreadable_plan(self):
        self.add_pdf("001.pdf")
        (self.inbox / "bad.pdf").write_bytes(b"broken")
        rows = preview.build_merge_preview(self.inbox, "out", self.outbox)
        self.assertEqual(sum(1 for r in rows if r.unreadable), 1)

    def test_rename_preview_marks_unchanged_conflict_and_duplicate(self):
        for name in ("001.pdf", "003.pdf", "008.pdf"):
            self.add_pdf(name)
        make_excel(self.excel, {"B1": "001", "B2": "同じ", "B3": "同じ"})
        rows = preview.build_rename_preview(self.inbox, self.excel, "Sheet", "B1", modes.PER_CELL)
        self.assertEqual([r.old_name for r in rows], ["001.pdf", "003.pdf", "008.pdf"])
        self.assertEqual([r.new_name for r in rows], ["001.pdf", "同じ.pdf", "同じ.pdf"])
        self.assertEqual([r.status for r in rows], ["[変更不要]", "[重複]", "[重複]"])
        self.assertEqual(rows[0].display_cells(), ("1", "001.pdf", "001.pdf", "[変更不要]"))
        self.assertEqual(sorted(p.name for p in self.inbox.iterdir()), ["001.pdf", "003.pdf", "008.pdf"])

    def test_rename_preview_flags_conflict(self):
        self.add_pdf("001.pdf")
        self.add_pdf("003.pdf")
        make_excel(self.excel, {"A1": "報告書"})
        (self.inbox / "報告書_001.pdf").mkdir()
        rows = preview.build_rename_preview(self.inbox, self.excel, "Sheet", "A1", modes.BASE_SEQ)
        self.assertEqual([r.status for r in rows], ["[衝突]", ""])

    def test_more_than_100_rows_are_summarised(self):
        for index in range(105):
            self.add_pdf(f"{index:03d}.pdf")
        rows = preview.build_merge_preview(self.inbox, "out", self.outbox)
        shown, extra = preview.truncate_display(rows)
        self.assertEqual(len(shown), preview.MAX_ANNOUNCED)
        self.assertEqual(extra, 5)
        rename_rows = [preview.RenamePreviewRow(order=i, old_name=f"{i}.pdf", new_name=f"n{i}.pdf") for i in range(120)]
        shown, extra = preview.truncate_display(rename_rows)
        self.assertEqual(len(shown), preview.MAX_ANNOUNCED)
        self.assertEqual(extra, 20)


class GuiWordingTests(unittest.TestCase):
    """GUI は表示のみ。ウィンドウを開かずに、文言とインポート確認だけ行う（FR-15）."""

    def test_gui_module_imports_without_opening_window(self):
        import pdf_split_merge_rename as gui

        self.assertTrue(hasattr(gui, "UnifiedPdfSplitApp"))
        self.assertTrue(hasattr(gui, "main"))
        self.assertEqual(gui.PREVIEW_VISIBLE_ROWS, 12)
        self.assertEqual(gui.STATE_STALE, "入力変更あり: プレビューを更新してください")


if __name__ == "__main__":
    unittest.main()




