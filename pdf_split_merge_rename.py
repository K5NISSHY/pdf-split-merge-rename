"""PDF分割・結合・リネームツールのGUI・エントリポイント.

上部に共通入力欄5行（分割するPDF / 入力フォルダー / Excelファイル / シート / 出力先フォルダー）、
中央に ttk.Notebook によるタブ4枚（セル値＋連番 / セル別（連番なし） / PDF結合 / Excelでリネーム）、
その下にタブごとに列構成の変わるプレビュー一覧欄を置く単一ウィンドウ。
タブを切り替えても各欄の入力は保持され、実行は常にその時点の入力で組み直す。
ロジックは split_core / modes / merge_core / rename_core / preview に置く。
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

import merge_core
import modes
import preview
import rename_core
import split_core

PREVIEW_VISIBLE_ROWS = 12  # 一覧欄の表示行数（データは全件保持）
APP_VERSION = "2.1.1"  # タイトル欄に表示する版（どの版を起動しているか判別できるようにする）
STATE_INITIAL = "プレビューを押すと、出力予定のファイル名を一覧に表示します。"
STATE_STALE = "入力変更あり: プレビューを更新してください"

# FR-C1: タブは 分割2枚 → PDF結合 → Excelでリネーム の順で並べる
TAB_MERGE = "merge"
TAB_RENAME = "rename"
SPLIT_TABS = (modes.BASE_SEQ, modes.PER_CELL)
TAB_ORDER = (modes.BASE_SEQ, modes.PER_CELL, TAB_MERGE, TAB_RENAME)
TAB_LABELS = {
    modes.BASE_SEQ: modes.MODE_LABELS[modes.BASE_SEQ],
    modes.PER_CELL: modes.MODE_LABELS[modes.PER_CELL],
    TAB_MERGE: "PDF結合",
    TAB_RENAME: "Excelでリネーム",
}

# FR-R16: リネームの適用先（その場で改名 / 出力先へコピーして改名）
RENAME_TARGET_INPLACE = "inplace"
RENAME_TARGET_COPY = "copy"
RENAME_TARGET_LABELS = {
    RENAME_TARGET_INPLACE: "入力フォルダー内で名前を変える",
    RENAME_TARGET_COPY: "出力先フォルダーへコピーして名前を変える",
}

# FR-C5: タブ別の列構成 (列ID, 見出し, 幅, 位置, 伸縮)
TAB_COLUMNS = {
    modes.BASE_SEQ: (
        ("page", "ページ番号", 90, "e", False),
        ("filename", "出力ファイル名", 470, "w", True),
        ("status", "状態", 130, "w", False),
    ),
    modes.PER_CELL: (
        ("page", "ページ番号", 90, "e", False),
        ("filename", "出力ファイル名", 470, "w", True),
        ("status", "状態", 130, "w", False),
    ),
    TAB_MERGE: (
        ("order", "順序", 60, "e", False),
        ("source", "入力PDF", 400, "w", True),
        ("pages", "ページ数", 90, "e", False),
        ("status", "状態", 140, "w", False),
    ),
    TAB_RENAME: (
        ("order", "順序", 60, "e", False),
        ("old", "現在のファイル名", 250, "w", True),
        ("new", "新しいファイル名", 250, "w", True),
        ("status", "状態", 140, "w", False),
    ),
}


class UnifiedPdfSplitApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(f"PDF 分割・結合・ファイル名自動生成ツール（統合版） v{APP_VERSION}")
        root.resizable(False, False)
        # FR-C2: 共通欄・タブ内の欄はすべて独立した変数（タブを切り替えても保持される）
        self.pdf_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.base_cell_var = tk.StringVar(value="A1")
        self.per_cell_var = tk.StringVar(value="A1")
        self.merge_name_var = tk.StringVar()
        self.rename_rule_var = tk.StringVar(value=modes.RENAME_RULE_BASE_SEQ)
        self.rename_cell_var = tk.StringVar(value="A1")
        self.rename_target_var = tk.StringVar(value=RENAME_TARGET_INPLACE)
        self.preview_rows: list[preview.PreviewRow | preview.MergePreviewRow | preview.RenamePreviewRow] = []
        self.preview_active = False
        self.merge_name_autofilled = False  # FR-M5: 自動補完した値かどうか（ユーザー編集名は変えない）

        # 注: tk.Frame/Notebookのpackで渡すpadx/padyは単一値のみ指定可（Python 3.14以降のtkinter）。
        # grid側のpady=(上,下)タプル形式は従来どおり使えるため、詰め幅はgrid側で調整している。
        top = tk.Frame(root, padx=16, pady=12)
        top.pack(fill="x")
        # FR-C3: 共通欄は5行（分割するPDF / 入力フォルダー / Excel / シート / 出力先）
        self._file_row(top, 0, "分割するPDF", self.pdf_var, self.choose_pdf)
        self._file_row(top, 1, "入力フォルダー", self.folder_var, self.choose_folder)
        self._file_row(top, 2, "Excelファイル", self.excel_var, self.choose_excel)
        tk.Label(top, text="シート", width=20, anchor="w").grid(row=3, column=0, sticky="w", pady=5)
        self.sheet_box = ttk.Combobox(top, textvariable=self.sheet_var, state="readonly", width=56)
        self.sheet_box.grid(row=3, column=1, columnspan=2, sticky="w", padx=6, pady=5)
        self._file_row(top, 4, "出力先フォルダー", self.output_var, self.choose_output)

        buttons = tk.Frame(top)
        buttons.grid(row=5, column=0, columnspan=3, sticky="w")
        tk.Button(buttons, text="プレビュー", width=14, command=self.run_preview).pack(side="left")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="x", padx=16, pady=(8, 4))
        tab_seq = tk.Frame(self.notebook, padx=12, pady=10)
        tab_cell = tk.Frame(self.notebook, padx=12, pady=10)
        tab_merge = tk.Frame(self.notebook, padx=12, pady=10)
        tab_rename = tk.Frame(self.notebook, padx=12, pady=10)
        for tab, tab_id in zip((tab_seq, tab_cell, tab_merge, tab_rename), TAB_ORDER):
            self.notebook.add(tab, text=TAB_LABELS[tab_id])

        # タブ1: セル値＋連番（分割）
        tk.Label(
            tab_seq, text="選択したセルの値をファイル名に使用します（例: ABC123_001.pdf）。", anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        tk.Label(tab_seq, text="セル番地", width=20, anchor="w").grid(row=1, column=0, sticky="w")
        tk.Entry(tab_seq, textvariable=self.base_cell_var, width=20).grid(row=1, column=1, sticky="w")
        # FR-C8: 実行ボタンはタブ内に置く（他タブの誤実行ができない）
        tk.Button(tab_seq, text="分割を実行", width=14, command=self.run).grid(row=2, column=0, sticky="w", pady=(8, 0))

        # タブ2: セル別（連番なし）（分割）
        tk.Label(
            tab_cell,
            text="開始セルから下方向へ、PDFのページ数分のセル値をファイル名に使います（連番なし）。",
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        tk.Label(tab_cell, text="開始セル", width=20, anchor="w").grid(row=1, column=0, sticky="w")
        tk.Entry(tab_cell, textvariable=self.per_cell_var, width=20).grid(row=1, column=1, sticky="w")
        tk.Button(tab_cell, text="分割を実行", width=14, command=self.run).grid(row=2, column=0, sticky="w", pady=(8, 0))

        # タブ3: PDF結合
        tk.Label(
            tab_merge,
            text="入力フォルダーの直下にあるPDFを自然順（001 → 003 → 008 → 9 → 10）で1つに結合します。",
            anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        tk.Label(tab_merge, text="結合後のファイル名", width=20, anchor="w").grid(row=1, column=0, sticky="w")
        tk.Entry(tab_merge, textvariable=self.merge_name_var, width=40).grid(row=1, column=1, sticky="w")
        tk.Label(
            tab_merge, text="初期値は自然順で最初のPDFのファイル名です。書き換えるとその名前が使われます（.pdf を補います）。",
            anchor="w",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))
        tk.Button(tab_merge, text="結合を実行", width=14, command=self.run_merge).grid(row=3, column=0, sticky="w", pady=(8, 0))

        # タブ4: Excelでリネーム
        tk.Label(
            tab_rename,
            text="入力フォルダーの直下のPDFを自然順に並べ、Excelの命名ルールで名前を変えます。上書きはしません。\n"
                 "このタブはPDFの分割を行いません。PDF 1件には1ファイル（ページ数はそのまま）を対応させます。\n"
                 "その場で改名: 同じフォルダー内で名前を変えます（コピーは作りません）。\n"
                 "出力先へコピー: 入力フォルダーのPDFを残したまま、共通欄の「出力先フォルダー」へ新しい名前で書き出します。",
            anchor="w",
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        tk.Label(tab_rename, text="リネームのルール", width=20, anchor="w").grid(row=1, column=0, sticky="w")
        rule_box = tk.Frame(tab_rename)
        rule_box.grid(row=1, column=1, columnspan=2, sticky="w")
        for rule_id in (modes.RENAME_RULE_BASE_SEQ, modes.RENAME_RULE_PER_CELL):
            tk.Radiobutton(
                rule_box, text=modes.RENAME_RULE_LABELS[rule_id],
                variable=self.rename_rule_var, value=rule_id,
            ).pack(side="left", padx=(0, 14))
        tk.Label(tab_rename, text="適用先", width=20, anchor="w").grid(row=2, column=0, sticky="w")
        target_box = tk.Frame(tab_rename)
        target_box.grid(row=2, column=1, columnspan=2, sticky="w")
        for target_id in (RENAME_TARGET_INPLACE, RENAME_TARGET_COPY):
            tk.Radiobutton(
                target_box, text=RENAME_TARGET_LABELS[target_id],
                variable=self.rename_target_var, value=target_id,
            ).pack(side="left", padx=(0, 14))
        tk.Label(tab_rename, text="セル番地", width=20, anchor="w").grid(row=3, column=0, sticky="w")
        tk.Entry(tab_rename, textvariable=self.rename_cell_var, width=20).grid(row=3, column=1, sticky="w")
        tk.Label(
            tab_rename, text="連番を使うルールは A1 のような1セル、行のセル値を使うルールは開始セルを指定します。",
            anchor="w",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))
        tk.Button(tab_rename, text="リネームを実行", width=14, command=self.run_rename).grid(row=5, column=0, sticky="w", pady=(8, 0))

        list_frame = tk.Frame(root, padx=16, pady=8)
        list_frame.pack(fill="both", expand=True)
        self.preview_tree = ttk.Treeview(list_frame, columns=(), show="headings", height=PREVIEW_VISIBLE_ROWS)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        self._apply_columns(TAB_ORDER[0])
        self.status_label = tk.Label(root, text=STATE_INITIAL, anchor="w", fg="#555555")
        self.status_label.pack(fill="x", padx=16)
        self.error_label = tk.Label(root, text="", anchor="w", fg="#a00000", justify="left")
        self.error_label.pack(fill="x", padx=16, pady=(0, 8))

        self._trace_inputs()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _file_row(self, parent: tk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        """共通欄1行（ラベル / 入力欄 / 参照...）を作る."""
        tk.Label(parent, text=label, width=20, anchor="w").grid(row=row, column=0, sticky="w", pady=5)
        tk.Entry(parent, textvariable=variable, width=56).grid(row=row, column=1, sticky="w", padx=6, pady=5)
        tk.Button(parent, text="参照...", width=10, command=command).grid(row=row, column=2, pady=5)

    def _trace_inputs(self) -> None:
        """FR-C6: 入力欄・セル欄・ルールが変わったら一覧を消して「要更新」にする."""
        for variable in (
            self.pdf_var, self.folder_var, self.excel_var, self.sheet_var, self.output_var,
            self.base_cell_var, self.per_cell_var, self.rename_cell_var, self.rename_rule_var,
            self.rename_target_var,
        ):
            variable.trace_add("write", self.invalidate_preview)
        # 結合後ファイル名はユーザーが編集した瞬間に「自動補完値」ではなくなる（FR-M5）
        self.merge_name_var.trace_add("write", self._on_merge_name_edited)

    def _on_merge_name_edited(self, *_args) -> None:
        """結合後ファイル名が編集されたら自動補完フラグを下ろす（以後、自動では書き換えない）."""
        self.merge_name_autofilled = False
        self.invalidate_preview()

    def _set_merge_name(self, name: str) -> None:
        """結合後ファイル名を自動補完する（traceで消されるフラグを後から立て直す）."""
        self.merge_name_var.set(name)
        self.merge_name_autofilled = True
        self.invalidate_preview()

    def _autofill_merge_name(self) -> None:
        """FR-M4/FR-M5: 空欄か自動補完値のままのときだけ、最初のPDF名を結合後ファイル名に補う."""
        folder_text = self.folder_var.get().strip()
        if not folder_text:
            return
        if self.merge_name_var.get().strip() and not self.merge_name_autofilled:
            return  # ユーザーが編集した名前は絶対に替えない
        try:
            sources = merge_core.list_merge_sources(Path(folder_text))
        except ValueError:
            return  # 0件・1件などの案内はプレビュー・実行側で出す
        self._set_merge_name(merge_core.default_output_name(sources))

    # --- 参照ボタン -------------------------------------------------------
    def choose_pdf(self) -> None:
        """分割するPDFを選び、出力先が未設定ならPDFのあるフォルダーを初期値にする."""
        path = filedialog.askopenfilename(title="分割するPDFを選択", filetypes=[("PDFファイル", "*.pdf")])
        if not path:
            return
        self.pdf_var.set(str(path))
        if not self.output_var.get():
            self.output_var.set(str(Path(path).parent))

    def choose_folder(self) -> None:
        """FR-M6: 入力フォルダーを選び、出力先が未設定なら同じフォルダーを初期値にする."""
        path = filedialog.askdirectory(title="入力フォルダーを選択")
        if not path:
            return
        self.folder_var.set(str(path))
        if not self.output_var.get():
            self.output_var.set(str(path))
        self._autofill_merge_name()

    def choose_excel(self) -> None:
        """Excelを選び、シート名をコンボボックスに出してアクティブシートを初期選択する（FR-2）."""
        path = filedialog.askopenfilename(
            title="Excelファイルを選択", filetypes=[("Excelブック", "*.xlsx *.xlsm")]
        )
        if not path:
            return
        self.excel_var.set(str(path))
        self._populate_sheets(Path(path))

    def choose_output(self) -> None:
        """出力先フォルダーを選ぶ."""
        path = filedialog.askdirectory(title="出力先フォルダーを選択")
        if path:
            self.output_var.set(str(path))

    def _populate_sheets(self, excel_path: Path) -> None:
        """Excelのシート名をコンボボックスに出し、アクティブシートを初期選択する."""
        try:
            workbook = load_workbook(excel_path, read_only=True)
        except Exception as exc:
            self.error_label.configure(text=f"Excelを読み込めませんでした。{exc}")
            return
        try:
            names = list(workbook.sheetnames)
            active = workbook.active.title if workbook.active is not None else ""
        finally:
            workbook.close()
        self.sheet_box.configure(values=names)
        self.sheet_var.set(active if active in names else (names[0] if names else ""))

    # --- 現在のタブと入力 -------------------------------------------------
    def current_tab(self) -> str:
        """選択中のタブID（base_seq / per_cell / merge / rename）を返す."""
        index = self.notebook.index(self.notebook.select())
        return TAB_ORDER[index]

    def current_mode(self) -> str:
        """後方互換: 選択中のタブIDを返す（分割タブでは命名モードIDと同じ）."""
        return self.current_tab()

    def check_inputs(self, tab: str) -> bool:
        """FR-C9: タブごとに必要な欄だけ検査し、不足なら警告して False を返す."""
        if tab in SPLIT_TABS:
            cell_label = "セル番地" if tab == modes.BASE_SEQ else "開始セル"
            missing = [
                label
                for label, value in (
                    ("分割するPDF", self.pdf_var.get()),
                    ("Excelファイル", self.excel_var.get()),
                    ("シート", self.sheet_var.get()),
                    (cell_label, self._cell_for(tab)),
                    ("出力先フォルダー", self.output_var.get()),
                )
                if not value.strip()
            ]
            if missing:
                messagebox.showwarning("入力が足りません", f"{'、'.join(missing)} を指定してください。")
                return False
            return True
        if tab == TAB_MERGE:
            # FR-C9: 結合タブでは Excel・シート・分割するPDFは不要（出力先は入力フォルダーを代用できる）
            missing = [
                label
                for label, value in (
                    ("入力フォルダー", self.folder_var.get()),
                    ("結合後のファイル名", self.merge_name_var.get()),
                )
                if not value.strip()
            ]
            if missing:
                messagebox.showwarning("入力が足りません", f"{'、'.join(missing)} を指定してください。")
                return False
            return True
        copy_mode = self.rename_target_var.get() == RENAME_TARGET_COPY
        required = [
            ("入力フォルダー", self.folder_var.get()),
            ("Excelファイル", self.excel_var.get()),
            ("シート", self.sheet_var.get()),
            ("セル番地", self.rename_cell_var.get()),
        ]
        if copy_mode:  # FR-R16: 出力先へコピーする場合は出力先フォルダーが必須
            required.append(("出力先フォルダー", self.output_var.get()))
        missing = [label for label, value in required if not value.strip()]
        if missing:
            messagebox.showwarning("入力が足りません", f"{'、'.join(missing)} を指定してください。")
            return False
        return True

    def _cell_for(self, tab: str) -> str:
        """タブに応じたセル欄の値を返す."""
        if tab == modes.BASE_SEQ:
            return self.base_cell_var.get()
        if tab == modes.PER_CELL:
            return self.per_cell_var.get()
        return self.rename_cell_var.get()

    def split_inputs(self, tab: str) -> tuple[Path, Path, Path, str, str]:
        """分割タブの現在入力（実行は常に現在の入力で組み直す: FR-16）."""
        return (
            Path(self.pdf_var.get().strip()),
            Path(self.excel_var.get().strip()),
            Path(self.output_var.get().strip()),
            self.sheet_var.get(),
            self._cell_for(tab).strip(),
        )

    def merge_inputs(self) -> tuple[Path, Path, str]:
        """結合タブの現在入力（出力先が空なら入力フォルダーを代用: FR-M6）."""
        folder_text = self.folder_var.get().strip()
        output_text = self.output_var.get().strip() or folder_text
        return Path(folder_text), Path(output_text), self.merge_name_var.get()

    def rename_inputs(self) -> tuple[Path, Path, str, str, str]:
        """リネームタブの現在入力（ルールradio → 命名モードID: FR-R2）."""
        return (
            Path(self.folder_var.get().strip()),
            Path(self.excel_var.get().strip()),
            self.sheet_var.get(),
            self.rename_cell_var.get().strip(),
            modes.RENAME_RULES[self.rename_rule_var.get()],
        )

    def rename_target(self) -> str:
        """リネームの適用先（その場で改名 / 出力先へコピー）を返す（FR-R16）."""
        return self.rename_target_var.get()

    def rename_copy_folder(self) -> Path:
        """コピー先フォルダー（FR-R16: 共通欄の「出力先フォルダー」を使う）."""
        return Path(self.output_var.get().strip())

    # --- プレビュー一覧 ---------------------------------------------------
    def invalidate_preview(self, *_args) -> None:
        """FR-C6: 一覧を消して「要更新」表示に戻す（実行は常に現在の入力で組み直す）。"""
        self.preview_rows = []
        self.preview_active = False
        if hasattr(self, "preview_tree"):
            self.preview_tree.delete(*self.preview_tree.get_children())
        if hasattr(self, "status_label"):
            self.status_label.configure(text=STATE_STALE)
        self.error_label.configure(text="")

    def _apply_columns(self, tab: str) -> None:
        """FR-C5: タブに応じてプレビュー一覧の列構成（columns）を作り直す."""
        spec = TAB_COLUMNS[tab]
        self.preview_tree["columns"] = tuple(column[0] for column in spec)
        for column_id, heading, width, anchor, stretch in spec:
            self.preview_tree.heading(column_id, text=heading)
            self.preview_tree.column(column_id, width=width, anchor=anchor, stretch=stretch)

    def _on_tab_changed(self, *_args) -> None:
        """FR-C1/FR-C5/FR-C6: タブ切替で列構成を作り直し、一覧を消して「要更新」にする."""
        if not hasattr(self, "preview_tree"):
            return
        self._apply_columns(self.current_tab())
        self.invalidate_preview()

    def _show_rows(self, rows: list, summary: str) -> None:
        """FR-14/FR-M13/FR-R11: 行を一覧へ並べ、100件超は「ほか N 件」にまとめる."""
        self.preview_rows = rows
        self.preview_active = True
        self.preview_tree.delete(*self.preview_tree.get_children())
        shown, extra = preview.truncate_display(rows)
        filler = ("",) * (len(TAB_COLUMNS[self.current_tab()]) - 1)
        for row in shown:
            self.preview_tree.insert("", "end", values=row.display_cells())
        if extra:
            self.preview_tree.insert("", "end", values=filler + (f"ほか {extra} 件",))
            summary = f"{summary}・一覧表示 {preview.MAX_ANNOUNCED} 件・ほか {extra} 件"
        self.status_label.configure(text=summary)

    def run_preview(self) -> None:
        """現在のタブのプレビューを再構築する（書き出し・名前変更は一切行わない）."""
        tab = self.current_tab()
        if tab == TAB_MERGE:
            self._autofill_merge_name()
        if not self.check_inputs(tab):
            return
        self.error_label.configure(text="")
        try:
            if tab in SPLIT_TABS:
                pdf_path, excel_path, output_dir, sheet_name, cell = self.split_inputs(tab)
                rows = preview.build_preview(excel_path, sheet_name, cell, tab, output_dir, pdf_path)
                self._show_rows(rows, f"プレビュー: {len(rows)} 件（書き出しは行っていません）")
            elif tab == TAB_MERGE:
                folder, output_dir, out_name = self.merge_inputs()
                rows = preview.build_merge_preview(folder, out_name, output_dir)
                total_pages = sum(row.page_count for row in rows)
                self._show_rows(
                    rows,
                    f"プレビュー: 合計 {len(rows)} ファイル / {total_pages} ページ → {out_name.strip()}（書き出しは行っていません）",
                )
            else:
                folder, excel_path, sheet_name, cell, mode = self.rename_inputs()
                copy_mode = self.rename_target() == RENAME_TARGET_COPY
                out_folder = self.rename_copy_folder() if copy_mode else None
                rows = preview.build_rename_preview(folder, excel_path, sheet_name, cell, mode, out_folder, copy_mode)
                unchanged = sum(1 for row in rows if row.unchanged)
                action = "コピー" if copy_mode else "名前変更"
                self._show_rows(
                    rows,
                    f"プレビュー: PDF {len(rows)} 件 → {len(rows)} ファイル（分割なし・{action} {len(rows) - unchanged} / "
                    f"変更不要 {unchanged}、{action}は行っていません）",
                )
        except Exception as exc:
            self.invalidate_preview()
            self.error_label.configure(text=str(exc))

    # --- 実行 -------------------------------------------------------------
    def run(self) -> None:
        """分割タブの実行（v1.0.0 と同じ流れ・同じエラー文言）."""
        tab = self.current_tab()
        if tab not in SPLIT_TABS:
            return
        if not self.check_inputs(tab):
            return
        pdf_path, excel_path, output_dir, sheet_name, cell = self.split_inputs(tab)
        try:
            results = split_core.split_pdf(
                pdf_path, excel_path, output_dir, sheet_name, cell,
                lambda excel, sheet, target_cell, page_count: modes.build_filenames(
                    tab, excel, sheet, target_cell, page_count,
                ),
            )
        except Exception as exc:
            self.invalidate_preview()
            messagebox.showerror("分割できませんでした", str(exc))
            return
        self.invalidate_preview()
        messagebox.showinfo("完了", f"{len(results)}ページを分割しました。\n保存先: {results[0].parent}")

    def run_merge(self) -> None:
        """結合タブの実行（FR-M14: N ファイル / M ページ を報告）."""
        self._autofill_merge_name()
        if not self.check_inputs(TAB_MERGE):
            return
        folder, output_dir, out_name = self.merge_inputs()
        try:
            plan = merge_core.build_merge_plan(folder, out_name, output_dir)
            saved = merge_core.merge_pdfs(plan)
        except Exception as exc:
            self.invalidate_preview()
            messagebox.showerror("結合できませんでした", str(exc))
            return
        self.invalidate_preview()
        messagebox.showinfo(
            "完了",
            f"{len(plan.sources)} ファイル / {plan.total_pages} ページを結合しました。\n保存先: {saved}",
        )

    def run_rename(self) -> None:
        """リネームタブの実行（FR-R12: 件数報告 / FR-R16: 適用先で分岐）."""
        if not self.check_inputs(TAB_RENAME):
            return
        folder, excel_path, sheet_name, cell, mode = self.rename_inputs()
        if self.rename_target() == RENAME_TARGET_COPY:
            self.run_rename_copy(folder, excel_path, sheet_name, cell, mode)
            return
        try:
            items = rename_core.build_rename_plan(folder, excel_path, sheet_name, cell, mode)
            rename_core.check_rename_plan(items, folder)
            result = rename_core.apply_renames(items)
        except Exception as exc:
            self.invalidate_preview()
            messagebox.showerror("リネームできませんでした", str(exc))
            return
        self.invalidate_preview()
        messagebox.showinfo(
            "完了",
            f"変更 {result.renamed} 件 / 変更不要 {result.unchanged} 件（PDFの分割は行っていません）\nフォルダー: {folder}",
        )

    def run_rename_copy(self, folder: Path, excel_path: Path, sheet_name: str, cell: str, mode: str) -> None:
        """出力先フォルダーへコピーしながら名前を変える（FR-R16: 入力フォルダーのPDFは無傷のまま）."""
        out_folder = self.rename_copy_folder()
        try:
            items = rename_core.build_rename_plan(folder, excel_path, sheet_name, cell, mode)
            plan = rename_core.build_copy_plan(items, out_folder)
            rename_core.check_copy_plan(plan, out_folder)
            result = rename_core.apply_copies(plan)
        except Exception as exc:
            self.invalidate_preview()
            messagebox.showerror("コピーできませんでした", str(exc))
            return
        self.invalidate_preview()
        messagebox.showinfo(
            "完了",
            f"コピー {result.copied} 件 / 変更不要 {result.unchanged} 件（PDFの分割は行っていません）\n入力: {folder}\n出力: {out_folder}",
        )


def main() -> None:
    root = tk.Tk()
    UnifiedPdfSplitApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
