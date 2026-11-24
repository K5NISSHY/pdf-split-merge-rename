"""リリース前検査（アシスタント設計の独立検証・補助役の検証とは別系統）.

補助役が書いた `verify_independent.py` / `verify_assistant.py` とは**別の観点**で、
実PDF・実Excelを `output\\release_check` に作り直して次を機械判定する。

- 自然順の境界（001/003/008/9/10、20/100、a9/a10）
- 結合順序が実ページ順序として結果に反映されていること（ページ幅で識別）
- 結合の安全規則（既存出力・入力と同名・暗号化・破損・一時残骸・入力PDF不変）
- 結合の同一フォルダー出力（初回は成功、2回目は中止）
- リネームの割り当て順・両ルール・変更不要・衝突中止・ロールバック・一時残骸
- GUI の結合後ファイル名自動補完がユーザー編集を守ること（FR-M5）

実行: `py -3 tools\\verify_release_check.py`（全件成功で `TOTAL: n/n` / 終了コード0）
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter

import merge_core
import rename_core
from natural_sort import natural_sort_key, sort_pdf_names

WORK = ROOT / "output" / "release_check"

RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, condition: bool, detail: str = "") -> None:
    RESULTS.append((name, bool(condition), detail))


def make_pdf(path: Path, pages: int, width: int) -> None:
    """ページ幅を width にして作例。結合後の順序を幅で識別できるようにする."""
    writer = PdfWriter()
    for _ in range(pages):
        page = writer.add_blank_page(width=width, height=100)
        page.mediabox  # 参照だけ触る（属性確定のため）
    with path.open("wb") as stream:
        writer.write(stream)


def page_widths(path: Path) -> list[int]:
    reader = PdfReader(str(path))
    try:
        return [int(page.mediabox.width) for page in reader.pages]
    finally:
        reader.close()


def make_excel(path: Path, base: str, names: list[str]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "検査"
    sheet["A1"] = base
    for offset, name in enumerate(names):
        sheet.cell(row=1 + offset, column=3, value=name)
    book.save(path)
    book.close()


def snapshot(folder: Path) -> dict[str, tuple[int, int]]:
    return {p.name: (p.stat().st_size, p.stat().st_mtime_ns) for p in folder.iterdir() if p.is_file()}


def fresh(name: str) -> Path:
    folder = WORK / name
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True)
    return folder


# ---------------------------------------------------------------- 自然順
def check_natural_sort() -> None:
    check("自然順: 001<003<008<9<10", sort_pdf_names(["10.pdf", "9.pdf", "008.pdf", "003.pdf", "001.pdf"])
          == ["001.pdf", "003.pdf", "008.pdf", "9.pdf", "10.pdf"])
    check("自然順: 20<100", sort_pdf_names(["100.pdf", "20.pdf"]) == ["20.pdf", "100.pdf"])
    check("自然順: 英字混在 a9<a10", sort_pdf_names(["a10.pdf", "a9.pdf"]) == ["a9.pdf", "a10.pdf"])
    check("自然順: .pdf 以外を除外", sort_pdf_names(["a.pdf", "b.txt", "c.PDF"]) == ["a.pdf", "c.PDF"])
    check("自然順: 同順位が決定的", sorted(["x.pdf", "X.pdf"], key=natural_sort_key) == ["X.pdf", "x.pdf"])


# ---------------------------------------------------------------- 結合
def check_merge_order_and_safety() -> None:
    src = fresh("結合_順序")
    # 自然順で 001(幅101) 003(幅103) 008(幅108) 9(幅109) 10(幅110)
    for name, width in [("10.pdf", 110), ("9.pdf", 109), ("008.pdf", 108), ("003.pdf", 103), ("001.pdf", 101)]:
        make_pdf(src / name, 2, width)
    out = fresh("結合_順序_出力")

    plan = merge_core.build_merge_plan(src, "結合結果", out)
    check("結合: 順序が自然順5件", [p.name for p in plan.sources]
          == ["001.pdf", "003.pdf", "008.pdf", "9.pdf", "10.pdf"], str([p.name for p in plan.sources]))
    check("結合: 合計ページ数10", plan.total_pages == 10, str(plan.total_pages))
    result = merge_core.merge_pdfs(plan)
    check("結合: 結果は1ファイル", sorted(p.name for p in out.iterdir()) == ["結合結果.pdf"], str(list(out.iterdir())))
    check("結合: ページ順序が結果に反映", page_widths(result) == [101, 101, 103, 103, 108, 108, 109, 109, 110, 110],
          str(page_widths(result)))
    check("結合: 一時フォルダー残骸なし", not any(p.name.startswith(".pdf_merge_") for p in out.iterdir()))
    check("結合: 入力PDFは不変", {p.name for p in src.iterdir()}
          == {"001.pdf", "003.pdf", "008.pdf", "9.pdf", "10.pdf"}, str(sorted(p.name for p in src.iterdir())))

    # 既存出力 → 中止（上書きしない）
    try:
        merge_core.build_merge_plan(src, "結合結果", out)
        check("結合: 既存出力で中止", False)
    except FileExistsError as exc:
        check("結合: 既存出力で中止", "上書きは行いません" in str(exc), str(exc).splitlines()[0])

    # 入力PDFと同名 → 中止（デフォルト名はここに入る）
    try:
        merge_core.build_merge_plan(src, "001", out)
        check("結合: 入力と同名を拒否", False)
    except ValueError as exc:
        check("結合: 入力と同名を拒否", "入力PDFを出力名にできません" in str(exc), str(exc).splitlines()[0])

    # 不正名
    for bad, label in [("a/b", "不正文字"), ("CON", "予約名"), ("", "空"), ("x.docx", "別拡張子")]:
        try:
            merge_core.validate_output_name(bad)
            check(f"結合: {label}を拒否", False)
        except ValueError:
            check(f"結合: {label}を拒否", True)

    # 0件・1件
    empty = fresh("結合_0件")
    try:
        merge_core.list_merge_sources(empty)
        check("結合: 0件でエラー", False)
    except ValueError as exc:
        check("結合: 0件でエラー", "PDFが見つかりません" in str(exc))
    one = fresh("結合_1件")
    make_pdf(one / "only.pdf", 1, 100)
    try:
        merge_core.list_merge_sources(one)
        check("結合: 1件でエラー", False)
    except ValueError as exc:
        check("結合: 1件でエラー", "2つ以上" in str(exc))


def check_merge_same_folder_and_bad_inputs() -> None:
    folder = fresh("結合_同一フォルダー")
    for name, width in [("001.pdf", 101), ("002.pdf", 102)]:
        make_pdf(folder / name, 1, width)
    before = snapshot(folder)

    plan = merge_core.build_merge_plan(folder, "まとめ", folder)
    merge_core.merge_pdfs(plan)
    check("結合: 同一フォルダー出力の初回は成功", (folder / "まとめ.pdf").is_file())
    check("結合: 2回目は既存として中止", True)
    try:
        merge_core.build_merge_plan(folder, "まとめ", folder)
        check("結合: 2回目は既存として中止", False)
    except (FileExistsError, ValueError) as exc:
        # 2回目は まとめ.pdf 自体が入力PDFになる（FR-M8/M9 のどちらかで止まる）
        check("結合: 2回目は既存として中止", "上書きは行いません" in str(exc) or "入力PDFを出力名にできません" in str(exc),
              str(exc).splitlines()[0])
    check("結合: 一時残骸なし（同一フォルダー）",
          not any(p.name.startswith(".pdf_merge_") for p in folder.iterdir()))
    check("結合: 既存PDFの内容は不変",
          all((folder / n).stat().st_size == size for n, (size, _) in before.items()))

    # 暗号化・破損は全件検査してからまとめてエラー
    bad = fresh("結合_bad")
    make_pdf(bad / "001.pdf", 1, 101)
    writer = PdfWriter()
    writer.add_blank_page(width=102, height=100)
    writer.encrypt("secret")
    with (bad / "002.pdf").open("wb") as stream:
        writer.write(stream)
    (bad / "003.pdf").write_bytes(b"%PDF-1.7 garbage not a real pdf")
    out = fresh("結合_bad_出力")
    try:
        merge_core.build_merge_plan(bad, "まとめ", out)
        check("結合: 暗号化・破損で中止", False)
    except ValueError as exc:
        text = str(exc)
        check("結合: 暗号化・破損で中止", "読み込めないPDFがあります" in text
              and "002.pdf" in text and "003.pdf" in text, text.splitlines()[0])
    check("結合: 失敗時に出力なし", not any(p.suffix.lower() == ".pdf" for p in out.iterdir()), str(list(out.iterdir())))


# ---------------------------------------------------------------- リネーム
def check_rename() -> None:
    folder = fresh("リネーム_連番")
    for name in ("b3.pdf", "a10.pdf", "a9.pdf", "z1.pdf", "m2.pdf"):
        make_pdf(folder / name, 1, 100)
    excel = WORK / "names.xlsx"
    excel.parent.mkdir(parents=True, exist_ok=True)
    make_excel(excel, "請求書", ["一", "二", "三", "四", "五"])

    items = rename_core.build_rename_plan(folder, excel, "検査", "A1", "base_seq")
    check("リネーム: 自然順に連番を割り当て",
          [(i.source.name, i.new_name) for i in items]
          == [("a9.pdf", "請求書_001.pdf"), ("a10.pdf", "請求書_002.pdf"), ("b3.pdf", "請求書_003.pdf"),
              ("m2.pdf", "請求書_004.pdf"), ("z1.pdf", "請求書_005.pdf")],
          str([(i.source.name, i.new_name) for i in items]))
    result = rename_core.apply_renames(items)
    check("リネーム: 5件変更", (result.renamed, result.unchanged) == (5, 0), str(result))
    check("リネーム: 結果ファイル名", sorted(p.name for p in folder.iterdir())
          == ["請求書_001.pdf", "請求書_002.pdf", "請求書_003.pdf", "請求書_004.pdf", "請求書_005.pdf"],
          str(sorted(p.name for p in folder.iterdir())))
    check("リネーム: 一時名残骸なし", not any(p.name.startswith(".pdf_rename_") for p in folder.iterdir()))

    # 変更不要（同じ名前で再実行できる）
    again = rename_core.build_rename_plan(folder, excel, "検査", "A1", "base_seq")
    check("リネーム: 再実行は全て変更不要", all(i.unchanged for i in again))
    result2 = rename_core.apply_renames(again)
    check("リネーム: 再実行で0件変更", (result2.renamed, result2.unchanged) == (0, 5), str(result2))

    # セル別ルール
    folder2 = fresh("リネーム_セル別")
    for name in ("scan10.pdf", "scan9.pdf", "scan1.pdf"):
        make_pdf(folder2 / name, 1, 100)
    items2 = rename_core.build_rename_plan(folder2, excel, "検査", "C1", "per_cell")
    check("リネーム: セル別は自然順にセル値",
          [(i.source.name, i.new_name) for i in items2]
          == [("scan1.pdf", "一.pdf"), ("scan9.pdf", "二.pdf"), ("scan10.pdf", "三.pdf")],
          str([(i.source.name, i.new_name) for i in items2]))
    rename_core.apply_renames(items2)
    check("リネーム: セル別の結果", {p.name for p in folder2.iterdir()} == {"一.pdf", "二.pdf", "三.pdf"},
          str(sorted(p.name for p in folder2.iterdir())))

    # 名前を入れ替えるケース（ステージングがないと衝突する）
    folder3 = fresh("リネーム_入替")
    make_pdf(folder3 / "002.pdf", 1, 100)
    make_pdf(folder3 / "001.pdf", 1, 100)
    excel3 = WORK / "swap.xlsx"
    make_excel(excel3, "x", ["002", "001"])  # 名前を入れ替える（ステージング無しでは衝突する）
    items3 = rename_core.build_rename_plan(folder3, excel3, "検査", "C1", "per_cell")
    rename_core.apply_renames(items3)
    check("リネーム: 名前入替が成立", sorted(p.name for p in folder3.iterdir()) == ["001.pdf", "002.pdf"],
          str(sorted(p.name for p in folder3.iterdir())))

    # 衝突（対象以外の既存ファイルと同名）→ 名前変更前に中止
    folder4 = fresh("リネーム_衝突")
    make_pdf(folder4 / "001.pdf", 1, 100)
    make_pdf(folder4 / "002.pdf", 1, 100)
    # 衝突させるには、同名だがリネーム対象でないものが必要（.pdf ファイルは全て対象になる）
    # → 同じ名前の「フォルダー」を置いて検証する
    (folder4 / "請求書_002.pdf").mkdir()
    excel4 = WORK / "conflict.xlsx"
    make_excel(excel4, "請求書", [])
    items4 = rename_core.build_rename_plan(folder4, excel4, "検査", "A1", "base_seq")
    try:
        rename_core.check_rename_plan(items4, folder4)
        check("リネーム: 衝突で中止", False)
    except FileExistsError as exc:
        check("リネーム: 衝突で中止", "衝突します" in str(exc) and "請求書_002.pdf" in str(exc), str(exc).splitlines()[0])
    check("リネーム: 衝突時に名前が変わっていない",
          sorted(p.name for p in folder4.iterdir()) == ["001.pdf", "002.pdf", "請求書_002.pdf"],
          str(sorted(p.name for p in folder4.iterdir())))
    check("リネーム: 衝突させた同名フォルダーは不変", (folder4 / "請求書_002.pdf").is_dir())

    # ロールバック: 2段階目の途中失敗を再現して元名に戻ることを確認
    folder5 = fresh("リネーム_ロールバック")
    make_pdf(folder5 / "001.pdf", 1, 100)
    make_pdf(folder5 / "002.pdf", 1, 100)
    items5 = rename_core.build_rename_plan(folder5, excel4, "検査", "A1", "base_seq")
    import os as _os

    real_rename = _os.rename
    calls = {"n": 0}

    def failing_rename(src, dst):
        calls["n"] += 1
        if calls["n"] == 3:  # ① 2件完了 → ② 1件目で失敗
            raise OSError("模擬的な失敗")
        real_rename(src, dst)

    _os.rename = failing_rename
    try:
        rename_core.apply_renames(items5)
        check("リネーム: 失敗時にロールバック", False)
    except ValueError as exc:
        check("リネーム: 失敗時にロールバック", "元に戻しました" in str(exc), str(exc).splitlines()[0])
    finally:
        _os.rename = real_rename
    check("リネーム: ロールバック後に全ファイルが元名",
          sorted(p.name for p in folder5.iterdir()) == ["001.pdf", "002.pdf"],
          str(sorted(p.name for p in folder5.iterdir())))
    check("リネーム: ロールバック後に一時名残骸なし",
          not any(p.name.startswith(".pdf_rename_") for p in folder5.iterdir()))

    # 重複名（セル別に同じ値）→ 実行時はエラー、プレビューは通る
    folder6 = fresh("リネーム_重複")
    make_pdf(folder6 / "001.pdf", 1, 100)
    make_pdf(folder6 / "002.pdf", 1, 100)
    excel6 = WORK / "dup.xlsx"
    make_excel(excel6, "x", ["同じ", "同じ"])
    try:
        rename_core.build_rename_plan(folder6, excel6, "検査", "C1", "per_cell")
        check("リネーム: 重複名で中止", False)
    except ValueError as exc:
        check("リネーム: 重複名で中止", "重複" in str(exc), str(exc).splitlines()[0])
    preview_items = rename_core.build_rename_plan(folder6, excel6, "検査", "C1", "per_cell", abort_on_duplicates=False)
    check("リネーム: プレビューは重複を通す", len(preview_items) == 2)

    # 0件
    try:
        rename_core.list_rename_targets(fresh("リネーム_0件"))
        check("リネーム: 0件でエラー", False)
    except ValueError as exc:
        check("リネーム: 0件でエラー", "PDFが見つかりません" in str(exc))


# ---------------------------------------------------------------- GUI（FR-M5）
def check_gui_name_autofill() -> None:
    import tkinter as tk

    import pdf_split_merge_rename as gui

    root = tk.Tk()
    root.withdraw()
    try:
        app = gui.UnifiedPdfSplitApp(root)
        root.update_idletasks()
        tab_count = len(app.notebook.tabs())
        check("GUI: タブ4枚", tab_count == 4, str(tab_count))

        src = fresh("GUI_結合")
        make_pdf(src / "001.pdf", 1, 101)
        make_pdf(src / "002.pdf", 1, 102)
        # 参照ボタン（choose_folder）と同じ動作を再現する
        app.folder_var.set(str(src))
        if not app.output_var.get():
            app.output_var.set(str(src))
        app._autofill_merge_name()
        root.update()
        auto = app.merge_name_var.get()
        check("GUI: 結合後名は最初のPDF名", auto == "001", auto)
        check("GUI: 出力先は入力フォルダーが既定", Path(app.output_var.get()).samefile(src), app.output_var.get())
        app.merge_name_var.set("私の名前")
        app.folder_var.set("")
        root.update()
        app.folder_var.set(str(src))
        app._autofill_merge_name()
        root.update()
        check("GUI: ユーザー編集した名は変えない", app.merge_name_var.get() == "私の名前", app.merge_name_var.get())
    finally:
        root.destroy()


def main() -> int:
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    for step in (check_natural_sort, check_merge_order_and_safety, check_merge_same_folder_and_bad_inputs,
                 check_rename, check_gui_name_autofill):
        try:
            step()
        except Exception as exc:  # 予期しない例外は1項目の失敗として記録
            check(f"{step.__name__} が例外", False, f"{type(exc).__name__}: {exc}")
    failed = [name for name, ok, _ in RESULTS if not ok]
    for name, ok, detail in RESULTS:
        if not ok:
            print(f"FAIL: {name} {detail}")
    print(f"TOTAL: {len(RESULTS) - len(failed)}/{len(RESULTS)}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
