r"""GUI実表示スモーク（一瞬だけ実ウィンドウを作り、表示確認後に閉じる）.

表示確認: root.winfo_viewable() / 実ウィンドウサイズ / プレビュー一覧の実測表示件数 /
タブ4枚（FR-C1）とタブ別の列構成切替（FR-C5）、タブ3・4のプレビュー。
実行: py -3 tools\verify_gui_display.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import tkinter as tk  # noqa: E402

import pdf_split_merge_rename as gui  # noqa: E402
import modes  # noqa: E402


def main() -> int:
    work = Path(tempfile.mkdtemp(prefix="pdfsplit_gui_display_"))
    try:
        import verify_assistant as va  # 検証用のPDF・Excel生成関数を再利用

        va.make_pdf(work / "input.pdf", 4)
        va.make_pdf(work / "2.pdf", 2)  # 結合タブは2つ以上必要
        va.make_excel(work / "names.xlsx", "名簿", ["あいう", "かきく:不正", "さしす", "CON"])
        (work / "out").mkdir()

        root = tk.Tk()
        app = gui.UnifiedPdfSplitApp(root)
        root.update_idletasks()
        root.update()
        viewable = bool(root.winfo_viewable())
        width, height = root.winfo_width(), root.winfo_height()
        tabs = [app.notebook.tab(tab_id, "text") for tab_id in app.notebook.tabs()]
        app.pdf_var.set(str(work / "input.pdf"))
        app.excel_var.set(str(work / "names.xlsx"))
        app.sheet_var.set("名簿")
        app.output_var.set(str(work / "out"))
        app.base_cell_var.set("A1")
        root.update()
        app.run_preview()
        root.update()
        rows = len(app.preview_tree.get_children())
        bbox = app.preview_tree.bbox(app.preview_tree.get_children()[0]) if rows else None
        split_columns = tuple(app.preview_tree["columns"])
        # タブ2枚目へ切り替えても共通入力が保持されることを実表示で確認
        app.notebook.select(1)
        root.update()
        kept = (app.pdf_var.get(), app.excel_var.get(), app.sheet_var.get(), app.output_var.get())
        mode = app.current_mode()
        # 不正文字を含むA2のため、プレビューはエラー欄に日本語文言を出す（一覧は空）
        app.per_cell_var.set("A1")
        app.run_preview()
        root.update()
        err_text = app.error_label.cget("text")
        # タブ3枚目 PDF結合: 結合後ファイル名の自動補完と4列構成
        app.notebook.select(2)
        root.update()
        app.folder_var.set(str(work))
        app.run_preview()
        root.update()
        merge_name = app.merge_name_var.get()
        merge_columns = tuple(app.preview_tree["columns"])
        merge_rows = [app.preview_tree.item(item, "values") for item in app.preview_tree.get_children()]
        merge_summary = app.status_label.cget("text")
        # タブ4枚目 Excelでリネーム: 現在の名前 → 新しい名前の4列構成
        app.notebook.select(3)
        root.update()
        rename_columns = tuple(app.preview_tree["columns"])
        app.rename_cell_var.set("A1")
        app.run_preview()
        root.update()
        rename_rows = [app.preview_tree.item(item, "values") for item in app.preview_tree.get_children()]
        rename_summary = app.status_label.cget("text")
        root.destroy()
    finally:
        import shutil

        shutil.rmtree(work, ignore_errors=True)

    print(f"表示可能ウィンドウ: {viewable} / 実サイズ: {width}x{height}")
    print(f"タブ（FR-C1）: {' / '.join(tabs)}")
    print(f"プレビュー表示行数（A1=あいう・4ページ）: {rows} / 先頭行の描画領域: {bbox}")
    print(f"タブ切替後の共通入力保持: {kept[1].split('/')[-1]} / mode={mode}")
    print(f"プレビュー時のエラー欄: {err_text[:60]}")
    print(f"タブ3の列構成（FR-C5）: {' / '.join(merge_columns)} / 結合後ファイル名: {merge_name}")
    print(f"タブ3のプレビュー: {merge_rows}")
    print(f"タブ3の要約欄: {merge_summary}")
    print(f"タブ4の列構成（FR-C5）: {' / '.join(rename_columns)}")
    print(f"タブ4のプレビュー: {rename_rows}")
    print(f"タブ4の要約欄: {rename_summary}")
    ok = (
        viewable and width > 600 and height > 400 and rows == 4
        and mode == modes.PER_CELL and "ファイル名に使えない文字" in err_text
        and tabs == [gui.TAB_LABELS[tab] for tab in gui.TAB_ORDER]
        and split_columns == ("page", "filename", "status")
        and merge_columns == ("order", "source", "pages", "status")
        and rename_columns == ("order", "old", "new", "status")
        and merge_name == "input" and len(merge_rows) == 2
        and "6 ページ" in merge_summary
        and len(rename_rows) == 2 and rename_rows[0][2].endswith(".pdf")
    )
    print("TOTAL:", "1/1" if ok else "0/1")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
