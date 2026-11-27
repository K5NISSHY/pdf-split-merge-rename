r"""アシスタント側（実装者とは別系統）の独立検証スクリプト.

OpenCode が書いた実装に対し、要件定義（docs/PROJECT_DOCUMENT.md 第1〜3章）の
項目が実際に満たされているかを、こちらの作る新規PDF・Excelで機械的に確かめる。
実装者作成の tools/verify_independent.py とは検査項目を重複させない。
分割（FR-1〜FR-16）に加え、タブ4枚化（FR-C）と結合（FR-M）・リネーム（FR-R）も
非表示ウィンドウのGUI操作まで含めて検査する。

実行: py -3 tools\verify_assistant.py
出力: 合否行と最後に TOTAL: <成功>/<総数>。全件成功なら終了コード0。
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import split_core  # noqa: E402
import modes  # noqa: E402
import preview as preview_mod  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from pypdf import PdfReader, PdfWriter  # noqa: E402

RESULTS: list[tuple[bool, str]] = []


def check(ok: bool, message: str) -> None:
    RESULTS.append((bool(ok), message))
    print(("PASS: " if ok else "FAIL: ") + message)


def make_pdf(path: Path, pages: int) -> None:
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=200, height=120)
    with path.open("wb") as stream:
        writer.write(stream)


def make_excel(path: Path, sheet: str, values: list[str], formula_cell: str | None = None) -> None:
    book = Workbook()
    ws = book.active
    ws.title = sheet
    for offset, value in enumerate(values, start=1):
        ws[f"A{offset}"] = value
    if formula_cell:
        ws[formula_cell] = "=LEN(A1)"
    book.save(path)
    book.close()


def digest(paths) -> list[str]:
    if isinstance(paths, Path):
        paths = [paths]
    return [hashlib.sha256(p.read_bytes()).hexdigest() for p in paths]


def staging_dirs(folder: Path) -> list[str]:
    return [p.name for p in folder.iterdir() if p.name.startswith(".pdf_split_")]


def main() -> int:
    work = Path(tempfile.mkdtemp(prefix="pdfsplit_assistant_"))
    try:
        # --- 準備: 4ページPDF + A1〜A4 に名前を持つExcel
        src_pdf = work / "input.pdf"
        src_xlsx = work / "names.xlsx"
        make_pdf(src_pdf, 4)
        make_excel(src_xlsx, "名簿", ["あいう", "かきく", "さしす", "せあそ"])
        out1 = work / "out_base"
        out1.mkdir()
        src_digests = digest([src_pdf, src_xlsx])

        names = modes.build_filenames(modes.BASE_SEQ, src_xlsx, "名簿", "A1", 4)
        check(names == ["あいう_001.pdf", "あいう_002.pdf", "あいう_003.pdf", "あいう_004.pdf"],
              f"FR-3 base_seq のファイル名（連番001〜）: {names}")
        files = split_core.split_pdf(src_pdf, src_xlsx, out1, "名簿", "A1",
                                     lambda e, s, c, n: modes.build_filenames(modes.BASE_SEQ, e, s, c, n))
        check([p.name for p in files] == names and all(p.is_file() for p in files),
              "FR-7 base_seq で4つのPDFが実際に出力される")
        counts = [len(PdfReader(str(p)).pages) for p in files]
        check(counts == [1, 1, 1, 1], f"FR-8 各出力PDFはちょうど1ページ: {counts}")
        check(digest([src_pdf, src_xlsx]) == src_digests,
              "FR-9 元のPDF・Excelは読み取りのみで無変更")
        check(staging_dirs(out1) == [], "FR-12 一時フォルダー（.pdf_split_*）は成功後に残らない")

        # --- FR-4: 1000ページ超で連番の桁が増える
        big = work / "big.pdf"
        make_pdf(big, 1005)
        big_names = modes.build_filenames(modes.BASE_SEQ, src_xlsx, "名簿", "A1", 1005)
        check(len(big_names) == 1005 and big_names[0] == "あいう_001.pdf"
              and big_names[999] == "あいう_1000.pdf" and big_names[-1] == "あいう_1005.pdf",
              "FR-4 1000ページ超で連番が4桁に伸びる（001→1000→1005）")

        # --- FR-5: per_cell は連番なしで開始セルから下行き
        out2 = work / "out_cell"
        out2.mkdir()
        three_pdf = work / "three.pdf"
        make_pdf(three_pdf, 3)
        cell_names = modes.build_filenames(modes.PER_CELL, src_xlsx, "名簿", "A2", 3)
        check(cell_names == ["かきく.pdf", "さしす.pdf", "せあそ.pdf"],
              f"FR-5 per_cell は開始セルA2から下行き・連番なし: {cell_names}")
        cell_files = split_core.split_pdf(three_pdf, src_xlsx, out2, "名簿", "A2",
                                         lambda e, s, c, n: modes.build_filenames(modes.PER_CELL, e, s, c, n))
        check([p.name for p in cell_files] == cell_names, "FR-7 per_cell で3つのPDFが実際に出力される")

        # --- FR-10: 既存ファイルがある場合は上書きせず中止（最大8件提示）
        before = digest(out2 / "かきく.pdf")
        try:
            split_core.split_pdf(three_pdf, src_xlsx, out2, "名簿", "A2",
                                 lambda e, s, c, n: modes.build_filenames(modes.PER_CELL, e, s, c, n))
            check(False, "FR-10 既存ファイル時に中止していない")
        except FileExistsError as exc:
            check("上書きは行いません" in str(exc), "FR-10 既存ファイル時に FileExistsError で中止（文言に上書きしない旨）")
        check(digest([out2 / "かきく.pdf"]) == before, "FR-10 中止後も既存出力の中身は無変更")
        check(staging_dirs(out2) == [], "FR-12 中止時にも一時フォルダーが残らない")

        # --- FR-6: 不正なセル値の検査（参考コードと同じ文言）
        bad_cases = [
            (["A B", "不正:名前", "c"], "ファイル名に使えない文字"),
            (["A B", "CON", "c"], "予約名"),
            (["A B", "trail.", "c"], "末尾に空白またはピリオド"),
            (["A B", "   ", "c"], "が空です"),
        ]
        for idx, (values, expect) in enumerate(bad_cases, start=1):
            bad_xlsx = work / f"bad{idx}.xlsx"
            make_excel(bad_xlsx, "Sheet1", values)
            try:
                modes.build_filenames(modes.PER_CELL, bad_xlsx, "Sheet1", "A1", 3)
                check(False, f"FR-6 不正値 {values[1]!r} が通ってしまった")
            except ValueError as exc:
                check(expect in str(exc), f"FR-6 不正値 {values[1]!r} を日本語エラーで拒否（{str(exc).splitlines()[0][:40]}）")

        # --- 数式セルに保存済み計算結果が無い場合（openpyxlは計算しない）
        fx = work / "formula.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # openpyxl は数式を書くが計算結果（キャッシュ）は書かないため、この状態が「再計算が必要」ケースになる
        ws["A1"] = '=LEN("abc")'
        wb.save(fx)
        wb.close()
        try:
            modes.build_filenames(modes.PER_CELL, fx, "Sheet1", "A1", 1)
            check(False, "FR-6 計算結果の無い数式セルが通ってしまった")
        except ValueError as exc:
            check("再計算" in str(exc), f"FR-6 数式セルに計算結果が無い場合は再計算を案内（{str(exc).splitlines()[-1][:30]}）")

        # --- 暗号化PDF（参考コードと同じ decrypt("") 方式）
        enc = work / "enc.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.encrypt("secret")
        with enc.open("wb") as stream:
            writer.write(stream)
        try:
            split_core.load_pdf(enc)
            check(False, "FR-9 暗号化PDFを拒否していない")
        except ValueError as exc:
            check("パスワードで保護" in str(exc), "FR-9 暗号化PDFは日本語エラーで拒否")

        # --- FR-13: フォルダー選択・シート選択の共通欄（入力検証の文言）
        try:
            split_core.check_input_paths(work / "no.pdf", src_xlsx, out1)
            check(False, "FR-13 存在しないPDFを拒否していない")
        except ValueError as exc:
            check(str(exc) == "指定したPDFファイルが見つかりません。", f"FR-13 PDF不在の文言: {exc}")
        (work / "notes.txt").write_text("not excel", encoding="utf-8")
        try:
            split_core.check_input_paths(src_pdf, work / "notes.txt", out1)
            check(False, "FR-13 Excel拡張子検査を通過した")
        except ValueError as exc:
            check("Excelファイル（.xlsx または .xlsm）" in str(exc), f"FR-13 Excel拡張子検査: {exc}")

        # --- FR-14: プレビューは [既存] / [重複] を付け、書き出ししない
        prev_out = work / "out_prev"
        prev_out.mkdir()
        dup_xlsx = work / "dup.xlsx"
        make_excel(dup_xlsx, "Sheet1", ["Report", "report", "only"])
        (prev_out / "Report.pdf").write_bytes(b"EXISTING")
        keep = digest([prev_out / "Report.pdf"])
        rows = preview_mod.build_preview(dup_xlsx, "Sheet1", "A1", modes.PER_CELL, prev_out, three_pdf)
        marks = [(r.page_no, r.filename, r.status) for r in rows]
        check(marks[0][2].startswith("[既存]") and "[重複]" in marks[0][2],
              f"FR-14 1行目に [既存] と [重複]: {marks}")
        check("[重複]" in marks[1][2], f"FR-14 2行目（report）にも [重複]（大文字小文字無視）: {marks[1][2]}")
        check(marks[2][2] == "", "FR-14 重複・既存の無い3行目の状態は空")
        check(sorted(p.name for p in prev_out.iterdir()) == ["Report.pdf"],
              "FR-14 プレビューでファイルは一切書き出していない")
        check(digest([prev_out / "Report.pdf"]) == keep, "FR-14 プレビューで既存ファイルの中身は無変更")
        many = [preview_mod.PreviewRow(i, f"n{i}.pdf", False, False) for i in range(1, 106)]
        shown, extra = preview_mod.truncate_display(many)
        check(len(shown) == 100 and extra == 5 and len(many) == 105,
              f"FR-14 105件中100件を表示＋「ほか 5 件」（データは全件保持）: {len(shown)}/{extra}")

        # --- FR-12: 書き出し途中で失敗した場合、作成済み出力を削除してロールバック
        fail_out = work / "out_fail"
        fail_out.mkdir()
        real_rename = os.rename
        calls = {"n": 0}

        def flaky_rename(src, dst):
            calls["n"] += 1
            if calls["n"] == 3:
                raise OSError("検証による意図的な失敗")
            return real_rename(src, dst)

        os.rename = flaky_rename  # 名前変更は pathops 経由でも os.rename を呼ぶ
        try:
            split_core.split_pdf(src_pdf, src_xlsx, fail_out, "名簿", "A1",
                                 lambda e, s, c, n: modes.build_filenames(modes.BASE_SEQ, e, s, c, n))
            check(False, "FR-12 書き出し失敗を検知していない")
        except ValueError as exc:
            check("分割ファイルを保存できませんでした" in str(exc), "FR-12 書き出し失敗は日本語エラーに変換")
        finally:
            os.rename = real_rename
        check(list(fail_out.iterdir()) == [], "FR-12 失敗後に作成済み出力が残らない（ロールバック）")
        check(staging_dirs(fail_out) == [], "FR-12 失敗後に一時フォルダーも残らない")

        # --- FR-1/FR-2/FR-11/FR-14/FR-15: 統合GUIの構造（非表示ウィンドウで検証）
        import tkinter as tk
        import pdf_split_merge_rename as gui

        # 検証中のダイアログは人の操作を待たないよう差し替える（記録だけ残す）
        dialogs = {"error": [], "warning": [], "info": []}
        gui.messagebox.showerror = lambda *a: dialogs["error"].append(a)
        gui.messagebox.showwarning = lambda *a: dialogs["warning"].append(a)
        gui.messagebox.showinfo = lambda *a: dialogs["info"].append(a)

        root = tk.Tk()
        root.withdraw()
        try:
            app = gui.UnifiedPdfSplitApp(root)
            tabs = [app.notebook.tab(tab, "text") for tab in app.notebook.tabs()]
            check(tabs == ["セル値＋連番", "セル別（連番なし）", "PDF結合", "Excelでリネーム"],
                  f"FR-1/FR-C1 ttk.Notebook のタブ4枚（順序含む）: {tabs}")
            check(gui.TAB_ORDER == (modes.BASE_SEQ, modes.PER_CELL, gui.TAB_MERGE, gui.TAB_RENAME)
                  and [gui.TAB_LABELS[tab] for tab in gui.TAB_ORDER] == tabs,
                  f"FR-C1 タブIDと表示ラベルの対応: {gui.TAB_ORDER}")
            check(app.current_mode() == modes.BASE_SEQ, "FR-2 アクティブタブからモードが決まる（既定はセル値＋連番）")
            app.notebook.select(1)
            check(app.current_mode() == modes.PER_CELL, "FR-2 タブ2枚目で per_cell に切り替わる")
            app.notebook.select(0)
            labeled = top_level_buttons(root)
            check(labeled.count("プレビュー") == 1 and labeled.count("分割を実行") == 2
                  and labeled.count("結合を実行") == 1 and labeled.count("リネームを実行") == 1,
                  f"FR-11/FR-C8 実行ボタンは各タブ内（プレビュー1・分割を実行2・結合を実行1・リネームを実行1）: {labeled}")
            check(root.title() == "PDF 分割・結合・ファイル名自動生成ツール（統合版）", f"FR-11 タイトル: {root.title()}")
            headers = [app.preview_tree.heading(col, "text") for col in ("page", "filename", "status")]
            check(headers == ["ページ番号", "出力ファイル名", "状態"], f"FR-14 一覧の列構成: {headers}")
            # 102ページ分のプレビュー表示（ほか 2 件 検証）
            app.pdf_var.set(str(work / "big.pdf"))
            app.excel_var.set(str(src_xlsx))
            app.sheet_var.set("名簿")
            app.output_var.set(str(out1))
            app.base_cell_var.set("A1")
            app.run_preview()
            tree_rows = app.preview_tree.get_children()
            last = app.preview_tree.item(tree_rows[-1], "values")[2]
            check(len(app.preview_rows) == 1005 and len(tree_rows) == 101 and last == "ほか 905 件",
                  f"FR-14 1005件→100行＋「ほか 905 件」: 一覧 {len(tree_rows)} 行 / 最終 {last}")
            check("書き出しは行っていません" in app.status_label.cget("text"),
                  f"FR-14 ステータスに書き出ししない旨: {app.status_label.cget('text')[:60]}")
            app.base_cell_var.set("A2")
            check(app.status_label.cget("text") == gui.STATE_STALE and len(app.preview_tree.get_children()) == 0,
                  f"FR-15 入力変更で一覧が消え「要更新」表示: {app.status_label.cget('text')}")
            # 実行はプレビューとは別の新しい出力先で行い、その時点のセル値（A2=かきく）が使われることを実測で確認
            run_out = work / "out_gui_run"
            run_out.mkdir()
            app.pdf_var.set(str(src_pdf))
            app.output_var.set(str(run_out))
            app.base_cell_var.set("A2")
            app.run()
            produced = sorted(p.name for p in run_out.iterdir())
            check(produced == ["かきく_001.pdf", "かきく_002.pdf", "かきく_003.pdf", "かきく_004.pdf"],
                  f"FR-15 実行時に最新の入力（A2）で再計算して実行: {produced}")
            check(dialogs["info"] and "4ページを分割しました" in dialogs["info"][-1][1],
                  f"FR-11 成功時は完了ダイアログ（4ページ）: {dialogs['info'][-1][1].splitlines()[0]}")
            check(app.status_label.cget("text") == gui.STATE_STALE,
                  "FR-15 実行後も一覧は「要更新」表示（古いプレビューを残さない）")
            # 未入力時は警告のみで実行しない
            dialogs["warning"].clear()
            app.pdf_var.set("")
            app.run()
            check(bool(dialogs["warning"]), "FR-13 未入力欄があれば警告出して実行しない")
            # --- FR-C2/FR-C5/FR-C9/FR-M: 結合タブ（タブ3枚目）
            merge_in = work / "merge_in"
            merge_in.mkdir()
            make_pdf(merge_in / "002.pdf", 2)
            make_pdf(merge_in / "9.pdf", 1)
            make_pdf(merge_in / "10.pdf", 1)
            merge_files = [merge_in / name for name in ("002.pdf", "9.pdf", "10.pdf")]
            merge_before = digest(merge_files)
            app.notebook.select(2)
            root.update()  # タブ切替イベント（列構成の作り直し）を処理させる
            check(app.current_tab() == gui.TAB_MERGE and app.base_cell_var.get() == "A2",
                  "FR-C1/FR-C2 タブ3枚目は結合で、分割タブで入力した値は保持されている")
            merge_columns = list(app.preview_tree["columns"])
            merge_headers = [app.preview_tree.heading(col, "text") for col in merge_columns]
            check(merge_columns == ["order", "source", "pages", "status"]
                  and merge_headers == ["順序", "入力PDF", "ページ数", "状態"],
                  f"FR-C5 結合タブの列構成: {merge_headers}")
            app.folder_var.set(str(merge_in))
            app.run_preview()
            check(app.merge_name_var.get() == "002",
                  f"FR-M4 結合後ファイル名の初期値は自然順で最初のPDF名: {app.merge_name_var.get()}")
            merge_rows = [app.preview_tree.item(row, "values") for row in app.preview_tree.get_children()]
            check([row[1] for row in merge_rows] == ["002.pdf", "9.pdf", "10.pdf"],
                  f"FR-M2 一覧は自然順（Explorer順 002/010/009 ではない）: {[row[1] for row in merge_rows]}")
            check("合計 3 ファイル / 4 ページ" in app.status_label.cget("text")
                  and sorted(p.name for p in merge_in.iterdir()) == ["002.pdf", "10.pdf", "9.pdf"],
                  f"FR-M13 合計件数とページ数を表示し、書き出しはしていない: {app.status_label.cget('text')}")
            check(app.check_inputs(gui.TAB_MERGE) is True and app.pdf_var.get() == "",
                  "FR-C9 結合タブでは「分割するPDF」が空でも処理を開始できる")
            app.merge_name_var.set("手入力名")
            app.run_preview()
            check(app.merge_name_var.get() == "手入力名", "FR-M5 手入力した結合後ファイル名は自動補完で書き換えない")
            app.output_var.set("")
            check(str(app.merge_inputs()[1]) == str(merge_in), "FR-M6 出力先が空なら入力フォルダーを代用する")
            app.run_merge()
            check(bool(dialogs["info"]) and "3 ファイル / 4 ページを結合しました" in dialogs["info"][-1][1],
                  f"FR-M14 完了報告に N ファイル / M ページ: {dialogs['info'][-1][1].splitlines()[0]}")
            merged = merge_in / "手入力名.pdf"
            check(merged.exists() and len(PdfReader(str(merged)).pages) == 4, "FR-M14 結合結果は4ページのPDF")
            check(digest(merge_files) == merge_before, "FR-M11 結合で入力PDFは一切変更されない")
            check([p.name for p in merge_in.iterdir() if p.name.startswith(".pdf_")] == [],
                  "FR-M12 結合後に一時フォルダー（.pdf_*）が残らない")
            one_dir = work / "merge_one"
            one_dir.mkdir()
            make_pdf(one_dir / "only.pdf", 1)
            app.folder_var.set(str(one_dir))
            app.merge_name_var.set("only")
            dialogs["error"].clear()
            app.run_merge()
            check(bool(dialogs["error"]) and "結合には2つ以上のPDFが必要です。" in dialogs["error"][-1][1]
                  and sorted(p.name for p in one_dir.iterdir()) == ["only.pdf"],
                  "FR-M3 PDFが1件だけなら実行せず、エラーのみで出力も作らない")

            # --- FR-C5/FR-R: リネームタブ（タブ4枚目）
            rename_in = work / "rename_in"
            rename_in.mkdir()
            for name in ("2_a.pdf", "8_b.pdf", "10_c.pdf"):
                make_pdf(rename_in / name, 1)
            rename_xlsx = work / "rename_names.xlsx"
            make_excel(rename_xlsx, "Sheet1", ["報告書"])
            app.notebook.select(3)
            root.update()  # タブ切替イベント（列構成の作り直し）を処理させる
            rename_columns = list(app.preview_tree["columns"])
            rename_headers = [app.preview_tree.heading(col, "text") for col in rename_columns]
            check(app.current_tab() == gui.TAB_RENAME
                  and rename_columns == ["order", "old", "new", "status"]
                  and rename_headers == ["順序", "現在のファイル名", "新しいファイル名", "状態"],
                  f"FR-C1/FR-C5 リネームタブの列構成: {rename_headers}")
            app.folder_var.set(str(rename_in))
            app.excel_var.set(str(rename_xlsx))
            app.sheet_var.set("Sheet1")
            app.rename_cell_var.set("A1")
            app.rename_rule_var.set(modes.RENAME_RULE_BASE_SEQ)
            check(app.rename_rule_var.get() == modes.RENAME_RULE_BASE_SEQ
                  and app.rename_inputs()[4] == modes.BASE_SEQ,
                  f"FR-R2 ルールradio（固定セル値＋連番）→ 命名モードID: {app.rename_inputs()[4]}")
            app.run_preview()
            rename_rows = [app.preview_tree.item(row, "values") for row in app.preview_tree.get_children()]
            check([row[1] for row in rename_rows] == ["2_a.pdf", "8_b.pdf", "10_c.pdf"]
                  and [row[2] for row in rename_rows] == ["報告書_001.pdf", "報告書_002.pdf", "報告書_003.pdf"],
                  f"FR-R1/FR-R11 プレビューは自然順と新しい名前（まだ変更していない）: {rename_rows}")
            check("変更 3 / 変更不要 0" in app.status_label.cget("text")
                  and sorted(p.name for p in rename_in.iterdir()) == ["10_c.pdf", "2_a.pdf", "8_b.pdf"],
                  f"FR-R11/FR-R12 プレビューに件数表示、実行前は名前変更なし: {app.status_label.cget('text')}")
            app.rename_rule_var.set(modes.RENAME_RULE_PER_CELL)
            check(app.rename_rule_var.get() == modes.RENAME_RULE_PER_CELL
                  and app.rename_inputs()[4] == modes.PER_CELL, "FR-R2 ルールradioでルールB（開始セルから下の行）へ切り替わる")
            app.rename_rule_var.set(modes.RENAME_RULE_BASE_SEQ)
            decoy = rename_in / "報告書_002.pdf"
            decoy.mkdir()  # リネーム対象以外の既存名（ここでは同名のフォルダー）
            dialogs["error"].clear()
            app.run_rename()
            check(bool(dialogs["error"]) and "衝突します" in dialogs["error"][-1][1]
                  and sorted(p.name for p in rename_in.iterdir()) == ["10_c.pdf", "2_a.pdf", "8_b.pdf", "報告書_002.pdf"],
                  "FR-R7/FR-R10 既存名と衝突すると実行前に中止し、名前変更は1つも起きていない")
            decoy.rmdir()
            app.run_rename()
            check(bool(dialogs["info"]) and "変更 3 件 / 変更不要 0 件" in dialogs["info"][-1][1]
                  and sorted(p.name for p in rename_in.iterdir())
                  == ["報告書_001.pdf", "報告書_002.pdf", "報告書_003.pdf"],
                  "FR-R12 完了報告に変更/変更不要の件数、実際にリネーム済み")
            app.run_rename()
            check("変更 0 件 / 変更不要 3 件" in dialogs["info"][-1][1],
                  "FR-R8 同じ条件で再実行すると全て変更不要（上書きしない）")

            # --- FR-R16〜FR-R18: 出力先へコピーしながら改名する ---
            copy_in = work / "copy_in"
            copy_in.mkdir()
            for name in ("2_a.pdf", "8_b.pdf"):
                make_pdf(copy_in / name, 1)
            copy_out = work / "copy_out"
            copy_out.mkdir()
            output_before = app.output_var.get()
            app.folder_var.set(str(copy_in))
            app.excel_var.set(str(rename_xlsx))
            app.output_var.set(str(copy_out))
            app.rename_target_var.set(gui.RENAME_TARGET_COPY)
            check(app.rename_target() == gui.RENAME_TARGET_COPY and app.rename_copy_folder() == copy_out,
                  "FR-R16 適用先radio（出力先へコピー）→ 共通欄の出力先フォルダーを使う")
            dialogs["warning"].clear()
            app.output_var.set("")
            app.run_rename()
            check(bool(dialogs["warning"]) and "出力先フォルダー" in dialogs["warning"][-1][1]
                  and list(copy_out.iterdir()) == [],
                  "FR-R16 出力先未指定なら警告して実行しない（コピーは始まらない）")
            app.output_var.set(str(copy_out))
            app.run_preview()
            copy_rows = [app.preview_tree.item(row, "values") for row in app.preview_tree.get_children()]
            check([row[2] for row in copy_rows] == ["報告書_001.pdf", "報告書_002.pdf"]
                  and list(copy_out.iterdir()) == []
                  and "コピー 2 / 変更不要 0" in app.status_label.cget("text"),
                  f"FR-R11/FR-R16 コピーのプレビューでも書き出しはしない: {app.status_label.cget('text')}")
            dialogs["info"].clear()
            app.run_rename()
            check(bool(dialogs["info"]) and "コピー 2 件 / 変更不要 0 件" in dialogs["info"][-1][1]
                  and sorted(p.name for p in copy_in.iterdir()) == ["2_a.pdf", "8_b.pdf"]
                  and sorted(p.name for p in copy_out.iterdir()) == ["報告書_001.pdf", "報告書_002.pdf"],
                  "FR-R16 完了報告にコピー件数、入力のPDFは無傷で出力先に新しい名前が揃う")
            check(all((copy_out / name).stat().st_size == (copy_in / src).stat().st_size
                      and len(PdfReader(copy_out / name).pages) == 1
                      for name, src in (("報告書_001.pdf", "2_a.pdf"), ("報告書_002.pdf", "8_b.pdf"))),
                  "FR-R18 出力先の各ファイルは元ファイルと同じサイズで読み取れる")
            dialogs["error"].clear()
            app.run_rename()
            check(bool(dialogs["error"]) and "出力先に同じ名前のファイルが既にあります" in dialogs["error"][-1][1]
                  and sorted(p.name for p in copy_out.iterdir()) == ["報告書_001.pdf", "報告書_002.pdf"],
                  "FR-R17 出力先に既存名があると実行前に中止し、上書きしていない")
            app.rename_target_var.set(gui.RENAME_TARGET_INPLACE)
            app.folder_var.set(str(rename_in))
            app.output_var.set(output_before)

            app.run_preview()  # 一覧に行を出した状態でタブを切り、消えることを確かめる
            app.notebook.select(2)
            root.update()  # タブ切替イベント（一覧を消す処理）を処理させる
            check(app.preview_tree.get_children() == () and app.status_label.cget("text") == gui.STATE_STALE,
                  "FR-C6 タブを切り替えると一覧が消え「要更新」表示になる")

            check(root.state() == "withdrawn", "GUI検証は非表示ウィンドウで実施（目視表示していない）")
        finally:
            root.destroy()

    except Exception as exc:  # 想定外エラーは全FAIL扱い
        import traceback

        traceback.print_exc()
        RESULTS.append((False, f"検証スクリプトが想定外のエラーで終了した: {exc}"))
    finally:
        shutil.rmtree(work, ignore_errors=True)

    succeeded = sum(1 for ok, _ in RESULTS if ok)
    print(f"TOTAL: {succeeded}/{len(RESULTS)}")
    return 0 if succeeded == len(RESULTS) else 1


def top_level_buttons(root):
    """フォーム内の操作ボタン（プレビュー / 分割・結合・リネームの実行）をウィジェット探索で集める."""
    import tkinter as tk

    found = []

    def walk(widget):
        for child in widget.winfo_children():
            if isinstance(child, tk.Button):
                found.append(child.cget("text"))
            else:
                walk(child)

    walk(root)
    return [name for name in found if name in ("プレビュー", "分割を実行", "結合を実行", "リネームを実行")]


if __name__ == "__main__":
    raise SystemExit(main())
