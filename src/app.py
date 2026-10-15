"""Responsive Tkinter desktop interface for the recon workbook generator."""

from __future__ import annotations

import subprocess
import threading
import datetime as dt
from functools import partial
from pathlib import Path
import platform
from typing import Any
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import polars as pl

from src.recon_exporter import (
    LB_COL_MAP,
    RB_COL_MAP,
    build_recon_workbook,
    read_raw_excel,
)
from src.version import __version__

PREVIEW_PAGE_SIZE = 500
PREVIEW_PAGE_SIZES = ("100", "250", "500", "1000")
SOURCE_HEADERS = {
    "LB": {value: key for key, value in LB_COL_MAP.items()},
    "RB": {value: key for key, value in RB_COL_MAP.items()},
}


class ReconApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Recon Builder | Finance v{__version__}")
        self.geometry("1180x760")
        self.minsize(850, 580)

        self.lb_path = tk.StringVar()
        self.rb_path = tk.StringVar()
        self.out_path = tk.StringVar(
            value=str(Path.home() / "Desktop" / "recon-output.xlsx")
        )
        self.status = tk.StringVar(
            value="Pilih file Excel LB dan RB untuk memulai."
        )
        self._frames: dict[str, pl.DataFrame | None] = {"LB": None, "RB": None}
        self._preview_widgets: dict[str, ttk.Treeview] = {}
        self._row_labels: dict[str, tk.StringVar] = {}
        self._search_vars: dict[str, tk.StringVar] = {}
        self._page_vars: dict[str, tk.StringVar] = {}
        self._page_size_vars: dict[str, tk.StringVar] = {}
        self._filtered_indices: dict[str, list[int]] = {"LB": [], "RB": []}
        self._page_index = {"LB": 0, "RB": 0}
        self._edit_counts: dict[str, set[tuple[int, str]]] = {
            "LB": set(),
            "RB": set(),
        }
        self._load_tokens = {"LB": 0, "RB": 0}
        self._loading: set[str] = set()
        self._invalid_edits: set[tuple[str, int, str]] = set()
        self._busy = False
        self._editor: tk.Entry | None = None
        self._editor_target: tuple[str, int, str] | None = None
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#F4F7FB")
        style.configure("TLabelframe", background="#F4F7FB")
        style.configure("TLabelframe.Label", background="#F4F7FB",
                        foreground="#12233F", font=("Helvetica", 11, "bold"))
        style.configure("TLabel", background="#F4F7FB")
        style.configure("TNotebook", background="#F4F7FB", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(14, 8))
        style.configure("Treeview", rowheight=27, font=("Helvetica", 10))
        style.configure(
            "Treeview.Heading", font=("Helvetica", 10, "bold"),
            foreground="#12233F",
        )
        style.configure(
            "Accent.TButton", font=("Helvetica", 10, "bold"),
            foreground="#FFFFFF", background="#2878C7", padding=(12, 8),
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#1D5E9E"), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#E2E8F0")],
        )
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(20, 16, 20, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header, text="Recon Builder", font=("Helvetica", 22, "bold")
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Pilih, cari, tinjau, dan koreksi seluruh data sebelum "
                "membuat workbook recon."
            ),
            foreground="#536273",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Label(
            header, text=f"Finance · v{__version__}", foreground="#2878C7"
        ).grid(row=0, column=1, sticky="e", padx=(10, 0))

        content = ttk.Frame(self, padding=(20, 8, 20, 8))
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        files = ttk.LabelFrame(content, text="1. File Excel", padding=12)
        files.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        files.columnconfigure(1, weight=1)
        self._add_file_row(
            files, 0, "LB · Loket Bayar", self.lb_path, self._browse_lb
        )
        self._add_file_row(
            files, 1, "RB · Rajabiller", self.rb_path, self._browse_rb
        )
        self._add_file_row(
            files, 2, "File output", self.out_path, self._browse_out
        )

        preview = ttk.LabelFrame(
            content,
            text="2. Preview & edit — seluruh baris dapat ditinjau",
            padding=10,
        )
        preview.grid(row=1, column=0, sticky="nsew")
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)
        self.notebook = ttk.Notebook(preview)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        self._add_preview_tab("LB", "Preview LB")
        self._add_preview_tab("RB", "Preview RB")
        ttk.Label(
            preview,
            text="Cari mencakup seluruh baris transaksi. Klik ganda sel untuk "
            "mengedit; Enter/pindah sel menerapkan, Esc membatalkan. "
            "Perubahan ikut diekspor.",
            foreground="#536273",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        footer = ttk.Frame(self, padding=(20, 8, 20, 16))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(footer, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        actions = ttk.Frame(footer)
        actions.grid(row=0, column=1, sticky="e")
        self.generate_button = ttk.Button(
            actions,
            text="Generate Recon",
            command=self._on_generate,
            style="Accent.TButton",
        )
        self.generate_button.pack(side="left", padx=(0, 8))
        ttk.Button(
            actions, text="Buka Folder Output", command=self._open_output_folder
        ).pack(side="left")
        ttk.Label(
            footer, textvariable=self.status, foreground="#35465a"
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

    def _add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: Any,
    ) -> None:
        ttk.Label(parent, text=label, width=20).grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=5
        )
        ttk.Button(parent, text="Pilih .xlsx…", command=command).grid(
            row=row, column=2, padx=(10, 0), pady=5
        )

    def _add_preview_tab(self, side: str, title: str) -> None:
        frame = ttk.Frame(self.notebook, padding=8)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        row_label = tk.StringVar(value="Belum ada file dipilih.")
        search_var = tk.StringVar()
        page_var = tk.StringVar(value="Halaman 0 / 0")
        page_size_var = tk.StringVar(value=str(PREVIEW_PAGE_SIZE))
        self._row_labels[side] = row_label
        self._search_vars[side] = search_var
        self._page_vars[side] = page_var
        self._page_size_vars[side] = page_size_var

        controls = ttk.Frame(frame)
        controls.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        controls.columnconfigure(0, weight=1)
        search_entry = ttk.Entry(controls, textvariable=search_var)
        search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        search_entry.bind("<Return>", partial(self._search_from_key, side))
        ttk.Button(
            controls, text="Cari seluruh data",
            command=lambda current=side: self._apply_search(current),
        ).grid(row=0, column=1, padx=3)
        ttk.Button(
            controls, text="Hapus",
            command=lambda current=side: self._clear_search(current),
        ).grid(row=0, column=2, padx=3)
        ttk.Label(controls, text="Baris/halaman").grid(
            row=0, column=3, padx=(10, 5)
        )
        page_size = ttk.Combobox(
            controls,
            textvariable=page_size_var,
            values=PREVIEW_PAGE_SIZES,
            state="readonly",
            width=7,
        )
        page_size.grid(row=0, column=4, padx=(0, 8))
        page_size.bind(
            "<<ComboboxSelected>>", partial(self._page_size_selected, side)
        )

        tree = ttk.Treeview(frame, show="headings", selectmode="browse")
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(
            frame, orient="horizontal", command=tree.xview
        )
        tree.configure(
            yscrollcommand=vertical.set, xscrollcommand=horizontal.set
        )
        tree.grid(row=2, column=0, sticky="nsew")
        vertical.grid(row=2, column=1, sticky="ns")
        horizontal.grid(row=3, column=0, sticky="ew")
        tree.bind("<Double-1>", lambda event, current=side: self._begin_edit(
            current, event
        ))
        self._preview_widgets[side] = tree
        navigation = ttk.Frame(frame)
        navigation.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        navigation.columnconfigure(1, weight=1)
        ttk.Button(
            navigation, text="← Sebelumnya",
            command=lambda current=side: self._change_page(current, -1),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(navigation, textvariable=page_var).grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Button(
            navigation, text="Berikutnya →",
            command=lambda current=side: self._change_page(current, 1),
        ).grid(row=0, column=2, sticky="e")
        self.notebook.add(frame, text=title)

    def _pick_xlsx(self, title: str) -> str:
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
        )

    def _browse_lb(self) -> None:
        selected = self._pick_xlsx("Pilih file LB (Loket Bayar)")
        if selected:
            self._load_source("LB", Path(selected))

    def _browse_rb(self) -> None:
        selected = self._pick_xlsx("Pilih file RB (Rajabiller)")
        if selected:
            self._load_source("RB", Path(selected))

    def _browse_out(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="Simpan file recon sebagai",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
        )
        if selected:
            self.out_path.set(selected)

    def _load_source(self, side: str, path: Path) -> None:
        if path.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format tidak didukung", "Pilih file Excel berformat .xlsx."
            )
            return
        has_edits = bool(self._edit_counts[side]) or any(
            edit[0] == side for edit in self._invalid_edits
        )
        if has_edits and not messagebox.askyesno(
            "Ganti file sumber?",
            f"Memilih file LB/RB baru akan membuang perubahan preview {side}. "
            "Lanjutkan?",
        ):
            return
        path_variable = self.lb_path if side == "LB" else self.rb_path
        path_variable.set(str(path))
        self._load_tokens[side] += 1
        token = self._load_tokens[side]
        self._loading.add(side)
        self._frames[side] = None
        self._filtered_indices[side] = []
        self._page_index[side] = 0
        self._search_vars[side].set("")
        self._edit_counts[side].clear()
        self._invalid_edits = {
            edit for edit in self._invalid_edits if edit[0] != side
        }
        self._clear_preview(side)
        self._row_labels[side].set(f"Membaca {path.name}…")
        self.status.set(f"Membaca file {side}…")
        threading.Thread(
            target=self._load_worker,
            args=(side, path, token),
            daemon=True,
        ).start()

    def _load_worker(self, side: str, path: Path, token: int) -> None:
        try:
            mapping = LB_COL_MAP if side == "LB" else RB_COL_MAP
            frame = read_raw_excel(
                path,
                col_map=mapping,
                drop_col_non_numeric="no" if side == "RB" else None,
            )
            error = None
        except Exception as exc:
            frame = None
            error = str(exc)
        self.after(0, self._source_loaded, side, token, frame, error)

    def _source_loaded(
        self,
        side: str,
        token: int,
        frame: pl.DataFrame | None,
        error: str | None,
    ) -> None:
        if token != self._load_tokens[side]:
            return
        self._loading.discard(side)
        if error is not None or frame is None:
            self._frames[side] = None
            self._row_labels[side].set("File gagal dibaca.")
            self.status.set(f"Gagal membaca file {side}.")
            messagebox.showerror("Gagal membaca file", error or "Data kosong.")
            return

        self._frames[side] = frame
        self._filtered_indices[side] = list(range(frame.height))
        self._page_index[side] = 0
        self._render_preview(side)
        self.status.set(f"File {side} siap ditinjau dan diedit.")

    def _clear_preview(self, side: str) -> None:
        tree = self._preview_widgets[side]
        tree.delete(*tree.get_children())
        tree.configure(columns=())
        self._page_vars[side].set("Halaman 0 / 0")

    @staticmethod
    def _display_value(value: Any) -> str:
        if value is None:
            return ""
        return value.isoformat(sep=" ") if hasattr(value, "isoformat") else str(value)

    def _render_preview(self, side: str) -> None:
        frame = self._frames[side]
        if frame is None:
            return
        tree = self._preview_widgets[side]
        columns = [column for column in frame.columns if column != "_row_idx"]
        tree.delete(*tree.get_children())
        tree.configure(columns=columns)
        for column in columns:
            tree.heading(
                column,
                text=SOURCE_HEADERS[side].get(column, column),
                anchor="w",
            )
            tree.column(column, width=145, minwidth=90, stretch=False)

        page_size = int(self._page_size_vars[side].get())
        indices = self._filtered_indices[side]
        page_count = (len(indices) + page_size - 1) // page_size
        self._page_index[side] = min(
            self._page_index[side], max(0, page_count - 1)
        )
        page_start = self._page_index[side] * page_size
        page_indices = indices[page_start:page_start + page_size]
        page = frame[page_indices].iter_rows(named=True)
        for index, row in zip(page_indices, page):
            values = [self._display_value(row.get(column)) for column in columns]
            tree.insert("", "end", iid=str(index), values=values)

        edited = len(self._edit_counts[side])
        suffix = f" · {edited} sel diedit" if edited else ""
        if self._search_vars[side].get().strip():
            description = (
                f"{side}: {len(indices):,} hasil pencarian dari "
                f"{frame.height:,} baris"
            )
        else:
            description = f"{side}: {frame.height:,} baris"
        self._row_labels[side].set(f"{description}{suffix}")
        first = page_start + 1 if page_indices else 0
        last = page_start + len(page_indices)
        self._page_vars[side].set(
            f"Baris {first:,}–{last:,} dari {len(indices):,} · "
            f"Halaman {self._page_index[side] + 1 if page_count else 0} / "
            f"{page_count:,}"
        )

    def _change_page(self, side: str, offset: int) -> None:
        if self._editor is not None and self._editor_target is not None:
            if not self._commit_edit(*self._editor_target, self._editor):
                return
        page_count = (
            len(self._filtered_indices[side])
            + int(self._page_size_vars[side].get()) - 1
        ) // int(self._page_size_vars[side].get())
        self._page_index[side] = max(
            0, min(page_count - 1, self._page_index[side] + offset)
        )
        self._render_preview(side)

    def _page_size_changed(self, side: str) -> None:
        self._page_index[side] = 0
        self._render_preview(side)

    def _page_size_selected(
        self, side: str, event: tk.Event[Any]
    ) -> None:
        event.widget.get()
        self._page_size_changed(side)

    def _search_from_key(self, side: str, event: tk.Event[Any]) -> str:
        if event.keysym == "Return":
            self._apply_search(side)
        return "break"

    def _apply_search(self, side: str) -> None:
        if self._editor is not None and self._editor_target is not None:
            if not self._commit_edit(*self._editor_target, self._editor):
                return
        frame = self._frames[side]
        if frame is None:
            return
        query = self._search_vars[side].get().strip().casefold()
        if not query:
            matches = list(range(frame.height))
        else:
            matches = self._search_indices(frame, query)
        self._filtered_indices[side] = matches
        self._page_index[side] = 0
        self._render_preview(side)

    def _clear_search(self, side: str) -> None:
        self._search_vars[side].set("")
        self._apply_search(side)

    def _begin_edit(self, side: str, event: tk.Event) -> None:
        if self._busy or side in self._loading:
            return
        if self._editor is not None and self._editor_target is not None:
            if not self._commit_edit(*self._editor_target, self._editor):
                return
        tree = self._preview_widgets[side]
        item = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        if not item or not column_id:
            return
        index = int(item)
        column_index = int(column_id[1:]) - 1
        columns = list(tree["columns"])
        if index >= (self._frames[side].height if self._frames[side] else 0):
            return
        if column_index >= len(columns):
            return
        box = tree.bbox(item, column_id)
        if not box:
            return

        if self._editor is not None:
            self._editor.destroy()
        editor = tk.Entry(tree)
        editor.insert(0, tree.set(item, columns[column_index]))
        editor.select_range(0, "end")
        editor.place(x=box[0], y=box[1], width=box[2], height=box[3])
        editor.focus_set()
        self._editor = editor
        self._editor_target = (side, index, columns[column_index])
        commit = partial(
            self._commit_edit, side, index, columns[column_index], editor
        )
        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", partial(self._cancel_edit_event, editor))

    def _commit_edit(
        self,
        side: str,
        row_index: int,
        column: str,
        editor: tk.Entry,
        event: tk.Event[Any] | None = None,
    ) -> bool | str:
        if self._editor is not editor:
            return "break" if event is not None else False
        raw_value = editor.get()
        editor.destroy()
        self._editor = None
        self._editor_target = None
        frame = self._frames[side]
        if frame is None:
            return "break" if event is not None else False

        try:
            if raw_value == "":
                value = None
            elif frame.schema[column] == pl.Null:
                value = raw_value
            else:
                value = self._parse_edit_value(raw_value, frame.schema[column])
            values = frame.get_column(column).to_list()
            values[row_index] = value
            if frame.schema[column] == pl.Null:
                updated = pl.Series(column, values, strict=True)
            else:
                updated = pl.Series(
                    column, values, dtype=frame.schema[column], strict=True
                )
            self._frames[side] = frame.with_columns(updated)
        except (TypeError, ValueError, pl.exceptions.PolarsError) as exc:
            self._invalid_edits.add((side, row_index, column))
            messagebox.showerror(
                "Nilai tidak valid",
                f"Nilai untuk kolom {SOURCE_HEADERS[side].get(column, column)} "
                f"tidak sesuai tipe data.\n\n{exc}",
            )
            return "break" if event is not None else False

        self._invalid_edits.discard((side, row_index, column))
        self._edit_counts[side].add((row_index, column))
        query = self._search_vars[side].get().strip()
        if query:
            self._filtered_indices[side] = self._search_indices(
                self._frames[side], query
            )
        self._render_preview(side)
        self.status.set(
            f"Perubahan pada file {side} diterapkan dan akan ikut diekspor."
        )
        return "break" if event is not None else True

    def _cancel_edit(self, editor: tk.Entry) -> None:
        if self._editor is editor:
            editor.destroy()
            self._editor = None
            if self._editor_target is not None:
                self._invalid_edits.discard(self._editor_target)
            self._editor_target = None

    def _cancel_edit_event(
        self, editor: tk.Entry, event: tk.Event[Any]
    ) -> str:
        if event.keysym == "Escape":
            self._cancel_edit(editor)
        return "break"

    @staticmethod
    def _parse_edit_value(raw_value: str, dtype: pl.DataType) -> Any:
        base_type = dtype.base_type()
        if base_type == pl.Datetime:
            parsed = dt.datetime.fromisoformat(raw_value.replace("Z", "+00:00"))
            return pl.Series("_edit", [parsed], dtype=dtype, strict=True).item()
        if base_type == pl.Date:
            parsed_date = dt.date.fromisoformat(raw_value)
            return pl.Series("_edit", [parsed_date], dtype=dtype, strict=True).item()
        if base_type == pl.Boolean:
            normalized = raw_value.strip().lower()
            if normalized not in {"true", "false"}:
                raise ValueError("Nilai boolean harus true atau false.")
            return normalized == "true"
        return pl.Series(
            "_edit", [raw_value], dtype=pl.String
        ).cast(dtype, strict=True).item()

    @staticmethod
    def _search_indices(frame: pl.DataFrame, query: str) -> list[int]:
        columns = [name for name in frame.columns if name != "_row_idx"]
        contains_query = pl.any_horizontal(
            pl.col(column)
            .cast(pl.String)
            .str.to_lowercase()
            .str.contains(query.casefold(), literal=True)
            .fill_null(False)
            for column in columns
        )
        return (
            frame.select(pl.int_range(pl.len()).alias("_index"))
            .with_columns(
                frame.select(contains_query).to_series().alias("_match")
            )
            .filter(pl.col("_match"))
            .get_column("_index")
            .to_list()
        )

    def _open_output_folder(self) -> None:
        folder = Path(self.out_path.get()).expanduser().parent
        if not folder.is_dir():
            messagebox.showwarning(
                "Folder tidak ada", f"Folder output belum ada:\n{folder}"
            )
            return
        current_platform = platform.system()
        if current_platform == "Darwin":
            command = ["open", str(folder)]
        elif current_platform == "Windows":
            command = ["explorer", str(folder)]
        else:
            command = ["xdg-open", str(folder)]
        try:
            subprocess.Popen(command)
        except OSError as exc:
            messagebox.showerror("Folder tidak bisa dibuka", str(exc))

    def _on_close(self) -> None:
        if self._busy or self._loading:
            messagebox.showwarning(
                "Proses masih berjalan",
                "Tunggu pembacaan file/export selesai sebelum menutup aplikasi.",
            )
            return
        self.destroy()

    def _on_generate(self) -> None:
        if self._editor is not None:
            target = self._editor_target
            if target is not None:
                if not self._commit_edit(*target, self._editor):
                    return
        if self._invalid_edits:
            messagebox.showerror(
                "Periksa nilai yang diedit",
                "Ada perubahan preview yang tidak valid. Perbaiki atau batalkan "
                "edit tersebut sebelum export.",
            )
            return
        if self._loading:
            messagebox.showwarning(
                "File sedang dibaca",
                "Tunggu sampai pembacaan file LB/RB selesai.",
            )
            return
        lb_path = Path(self.lb_path.get().strip()).expanduser()
        rb_path = Path(self.rb_path.get().strip()).expanduser()
        output = Path(self.out_path.get().strip()).expanduser()

        if self._frames["LB"] is None or not lb_path.is_file():
            messagebox.showerror("File LB", "Pilih dan muat file LB terlebih dahulu.")
            return
        if self._frames["RB"] is None or not rb_path.is_file():
            messagebox.showerror("File RB", "Pilih dan muat file RB terlebih dahulu.")
            return
        if not self.out_path.get().strip():
            messagebox.showerror("File output", "Tentukan lokasi file output.")
            return
        if output.resolve() in {lb_path.resolve(), rb_path.resolve()}:
            messagebox.showerror(
                "File output", "File output tidak boleh menimpa file input."
            )
            return
        if output.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format tidak didukung", "File output harus berakhiran .xlsx."
            )
            return

        self._busy = True
        self.generate_button.configure(state="disabled")
        self.progress.start(10)
        self.status.set("Membuat file recon dari data yang sudah ditinjau…")
        threading.Thread(
            target=self._export_worker,
            args=(self._frames["LB"], self._frames["RB"], output),
            daemon=True,
        ).start()

    def _export_worker(
        self, lb_frame: pl.DataFrame | None, rb_frame: pl.DataFrame | None,
        output: Path,
    ) -> None:
        try:
            if lb_frame is None or rb_frame is None:
                raise ValueError("File LB dan RB harus dimuat sebelum export.")
            result = build_recon_workbook(lb_frame, rb_frame, output)
            self.after(0, self._export_finished, result, None)
        except Exception as exc:
            self.after(0, self._export_finished, None, str(exc))

    def _export_finished(self, result: Path | None, error: str | None) -> None:
        self.progress.stop()
        self._busy = False
        self.generate_button.configure(state="normal")
        if error is not None:
            self.status.set("Export gagal.")
            messagebox.showerror("Gagal membuat recon", error)
            return
        self.status.set(f"Export selesai: {result}")
        messagebox.showinfo("Recon selesai", f"File berhasil dibuat:\n{result}")


def main() -> None:
    ReconApp().mainloop()


if __name__ == "__main__":
    main()
