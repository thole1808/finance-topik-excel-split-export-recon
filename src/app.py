"""Responsive Tkinter desktop interface for the recon workbook generator."""

from __future__ import annotations

import subprocess
import threading
import datetime as dt
import os
import sys
import shutil
import time
import json
import getpass
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
    read_backdate_workbook,
    read_recon_reference,
    read_raw_excel,
)
from src.audit_log import AuditLog, generated_files_directory
from src.announcements import announce
from src.localization import format_local_datetime
from src.version import __version__
from src.updater import (
    UpdateError,
    UpdateRelease,
    download_update,
    latest_update,
    start_installer,
)

PREVIEW_PAGE_SIZE = 500
PREVIEW_PAGE_SIZES = ("100", "250", "500", "1000", "Semua")
SOURCE_HEADERS = {
    "LB": {value: key for key, value in LB_COL_MAP.items()},
    "RB": {value: key for key, value in RB_COL_MAP.items()},
}


class ReconApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.withdraw()
        self.title(f"Recon Builder | Finance v{__version__}")
        self.audit = AuditLog()
        self.operator = getpass.getuser().strip() or "Operator"
        self.audit.record(self.operator, "app_started")
        self.geometry("1440x900")
        self.minsize(960, 660)

        self.lb_path = tk.StringVar()
        self.rb_path = tk.StringVar()
        self.recon_path = tk.StringVar()
        self.backdate_path = tk.StringVar()
        self.out_path = tk.StringVar(
            value=str(generated_files_directory())
        )
        self.last_output_path: Path | None = self.audit.last_successful_output()
        self._pending_output_path: Path | None = None
        self.status = tk.StringVar(
            value="Upload file Excel LB, RB, dan recon untuk memulai."
        )
        self._frames: dict[str, pl.DataFrame | None] = {
            "LB": None,
            "RB": None,
            "Recon Summary": None,
            "Recon Records": None,
            "Backdate LB": None,
            "Backdate RB": None,
            "Backdate RECON": None,
        }
        self._preview_widgets: dict[str, ttk.Treeview] = {}
        self._row_labels: dict[str, tk.StringVar] = {}
        self._search_vars: dict[str, tk.StringVar] = {}
        self._page_vars: dict[str, tk.StringVar] = {}
        self._page_size_vars: dict[str, tk.StringVar] = {}
        self._selected_cells: dict[str, tuple[int, str] | None] = {
            side: None for side in self._frames
        }
        self._filtered_indices: dict[str, list[int]] = {
            side: [] for side in self._frames
        }
        self._page_index = {side: 0 for side in self._frames}
        self._edit_counts: dict[str, set[tuple[int, str]]] = {
            side: set() for side in self._frames
        }
        self._load_tokens = {"LB": 0, "RB": 0, "Recon": 0, "Backdate": 0}
        self._loading: set[str] = set()
        self._invalid_edits: set[tuple[str, int, str]] = set()
        self._busy = False
        self._editor: tk.Entry | None = None
        self._editor_target: tuple[str, int, str] | None = None
        self._terminal: tk.Text | None = None
        self._history_tree: ttk.Treeview | None = None
        self._generated_files_tree: ttk.Treeview | None = None
        self._generated_file_paths: dict[str, Path] = {}
        self._header: ttk.Frame | None = None
        self._clock_var = tk.StringVar()
        self._release_status_var = tk.StringVar(
            value=f"APP v{__version__}  ·  RELEASE v{__version__}"
        )
        self._update_button: ttk.Button | None = None
        self._voice_test_button: ttk.Button | None = None
        self._update_info: UpdateRelease | None = None
        self._update_checking = False
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.deiconify()
        self.after(100, self._maximize_window)
        self._update_clock()
        self.audit.record(self.operator, "audit_chain_check", **{
            "valid": self.audit.verify_chain()[0]
        })
        self._refresh_history()
        self.after(1500, self._check_updates_silently)

    def _maximize_window(self) -> None:
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")

    def _update_clock(self) -> None:
        self._clock_var.set(format_local_datetime(dt.datetime.now()))
        self.after(1000, self._update_clock)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TFrame", background="#F1F5F9")
        style.configure("Header.TFrame", background="#102A43")
        style.configure(
            "TLabelframe", background="#FFFFFF", borderwidth=0,
            relief="flat", padding=2,
        )
        style.configure("TLabelframe.Label", background="#EDF3F8",
                        foreground="#123B5D", font=("Helvetica", 11, "bold"),
                        padding=(2, 2))
        style.configure("TLabel", background="#F1F5F9", foreground="#263D56")
        style.configure(
            "Header.TLabel", background="#102A43", foreground="#FFFFFF",
            font=("Helvetica", 23, "bold"),
        )
        style.configure(
            "Subtitle.TLabel", background="#102A43", foreground="#C7D8E7",
            font=("Helvetica", 10),
        )
        style.configure("TNotebook", background="#EDF3F8", borderwidth=0)
        style.configure(
            "TNotebook.Tab", padding=(18, 11),
            font=("Helvetica", 10, "bold"), background="#DCE7F0",
            foreground="#455A72",
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#FFFFFF")],
            foreground=[("selected", "#1F6EAA")],
        )
        style.configure(
            "Treeview", rowheight=30, font=("Helvetica", 10),
            background="#FFFFFF", fieldbackground="#FFFFFF",
            borderwidth=0, relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", "#D7EAFE")],
            foreground=[("selected", "#142B49")],
        )
        style.configure(
            "Treeview.Heading", font=("Helvetica", 10, "bold"),
            foreground="#263D56", background="#EAF0F6",
            padding=(8, 7), relief="flat",
        )
        style.configure(
            "Accent.TButton", font=("Helvetica", 10, "bold"),
            foreground="#FFFFFF", background="#087E8B", padding=(14, 9),
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#185A8C"), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#E2E8F0")],
        )
        style.configure(
            "TButton", font=("Helvetica", 10), padding=(10, 7),
            background="#FFFFFF", foreground="#263D56",
        )
        style.map("TButton", background=[("active", "#E8F1FA")])
        style.configure(
            "Header.TButton", font=("Helvetica", 10, "bold"),
            foreground="#FFFFFF", background="#087E8B", padding=(14, 9),
            borderwidth=0,
        )
        style.map(
            "Header.TButton",
            background=[("active", "#0A929B"), ("disabled", "#536B80")],
            foreground=[("disabled", "#DCE7F0")],
        )
        style.configure(
            "TEntry", padding=(8, 6), fieldbackground="#FFFFFF",
        )
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(28, 19, 28, 16), style="Header.TFrame")
        self._header = header
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header, text="Recon Builder", style="Header.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text=(
                "Upload, cari, tinjau, dan koreksi data sebelum "
                "membuat workbook recon."
            ),
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._release_status_label = ttk.Label(
            header, textvariable=self._release_status_var,
            background="#102A43", foreground="#55D6BE",
            font=("Helvetica", 9, "bold"),
        )
        self._release_status_label.grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            header, text="Powered By DevOps", background="#102A43",
            foreground="#C7D8E7", font=("Helvetica", 10, "bold"),
        ).grid(row=2, column=2, sticky="e", pady=(8, 0))
        self._update_button = ttk.Button(
            header, text="Periksa Update", command=self._on_update_button,
            style="Header.TButton",
        )
        self._update_button.grid(row=0, column=2, rowspan=2, sticky="e", padx=(24, 0))
        ttk.Label(
            header, textvariable=self._clock_var, background="#102A43",
            foreground="#FFFFFF", font=("Helvetica", 10),
        ).grid(row=0, column=1, rowspan=2, sticky="e", padx=(24, 0))

        content = ttk.Frame(self, padding=(24, 12, 24, 8))
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        files = ttk.LabelFrame(content, text="1. Upload Excel", padding=12)
        files.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        files.columnconfigure(1, weight=1)
        self._add_file_row(
            files, 0, "File Backdate (opsional)",
            self.backdate_path, self._browse_backdate,
        )
        self._add_file_row(
            files, 1, "Upload LB", self.lb_path, self._browse_lb
        )
        self._add_file_row(
            files, 2, "Upload RB", self.rb_path, self._browse_rb
        )
        self._add_file_row(
            files, 3, "Upload Recon",
            self.recon_path, self._browse_recon,
        )
        self._add_file_row(
            files, 4, "Folder output", self.out_path, self._browse_out
        )
        ttk.Label(
            files,
            text="LB = Loket Bayar · RB = Rajabiller · file recon referensi "
            "wajib memiliki sheet Summary dan Records. Backdate opsional "
            "digabung saat Generate.",
            foreground="#536273",
        ).grid(row=5, column=1, columnspan=2, sticky="w", pady=(2, 0))

        preview = ttk.LabelFrame(
            content,
            text="2. Preview",
            padding=10,
        )
        preview.grid(row=1, column=0, sticky="nsew")
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)
        self.notebook = ttk.Notebook(preview)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        self.notebook.bind("<<NotebookTabChanged>>", self._tab_changed)
        self._add_preview_tab("LB", "Preview LB", editable=False)
        self._add_preview_tab("RB", "Preview RB", editable=False)
        self._add_preview_tab(
            "Recon Summary", "Recon · Summary", editable=False
        )
        self._add_preview_tab(
            "Recon Records", "Recon · Records", editable=False
        )
        self._add_preview_tab("Backdate LB", "Backdate · LB", editable=False)
        self._add_preview_tab("Backdate RB", "Backdate · RB", editable=False)
        self._add_preview_tab("Backdate RECON", "Backdate · RECON", editable=False)
        self._add_generated_files_tab()
        self._add_history_tab()
        ttk.Label(
            preview,
            text="Pilih Semua untuk melihat seluruh baris. Preview hanya "
            "untuk melihat dan mencari data (read-only); pengeditan "
            "sel dimatikan. Summary/Records dari file recon hanya "
            "sebagai referensi.",
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
        if not self.audit.has_action("voice_tested"):
            self._voice_test_button = ttk.Button(
                actions, text="Tes Suara", command=self._test_voice
            )
            self._voice_test_button.pack(side="left", padx=(8, 0))
        ttk.Label(
            footer, textvariable=self.status, foreground="#35465a"
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 5))
        ttk.Label(
            footer, text="Terminal proses Generate", foreground="#123B5D",
            font=("Helvetica", 9, "bold"),
        ).grid(row=2, column=0, columnspan=2, sticky="w")
        terminal_frame = ttk.Frame(footer)
        terminal_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(3, 0))
        terminal_frame.columnconfigure(0, weight=1)
        self._terminal = tk.Text(
            terminal_frame, height=9, wrap="word", state="disabled",
            background="#102A43", foreground="#D9E7F2",
            insertbackground="#FFFFFF", relief="flat", font=("Menlo", 10),
        )
        terminal_scroll = ttk.Scrollbar(
            terminal_frame, orient="vertical", command=self._terminal.yview
        )
        self._terminal.configure(yscrollcommand=terminal_scroll.set)
        self._terminal.grid(row=0, column=0, sticky="ew")
        terminal_scroll.grid(row=0, column=1, sticky="ns")

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
        ttk.Button(parent, text="Upload .xlsx…", command=command).grid(
            row=row, column=2, padx=(10, 0), pady=5
        )

    def _add_preview_tab(
        self, side: str, title: str, editable: bool = True
    ) -> None:
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
        tree.tag_configure("even", background="#FFFFFF")
        tree.tag_configure("odd", background="#F5F8FC")
        tree.grid(row=2, column=0, sticky="nsew")
        vertical.grid(row=2, column=1, sticky="ns")
        horizontal.grid(row=3, column=0, sticky="ew")
        tree.bind(
            "<ButtonRelease-1>",
            lambda event, current=side: self._select_preview_cell(
                current, event
            ),
        )
        tree.bind("<Double-1>", lambda event, current=side: self._begin_edit(
            current, event
        ))
        if not editable:
            tree.unbind("<Double-1>")
        self._preview_widgets[side] = tree
        if editable:
            edit_bar = ttk.Frame(frame)
            edit_bar.grid(
                row=1, column=0, columnspan=2, sticky="ew", pady=(0, 5)
            )
            ttk.Button(
                edit_bar,
                text="Edit sel terpilih",
                command=lambda current=side: self._edit_selected_cell(
                    current
                ),
            ).pack(side="left")
            ttk.Label(
                edit_bar,
                text="Pilih sel lalu klik Edit, atau klik ganda sel. "
                "Perubahan akan masuk ke hasil generate.",
                foreground="#61748B",
            ).pack(side="left", padx=(10, 0))
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

    def _add_generated_files_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=10)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        controls = ttk.Frame(frame)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            controls,
            text="File hasil tersimpan di folder data Recon Builder pada komputer ini.",
        ).pack(side="left")
        ttk.Button(
            controls, text="Muat ulang", command=self._refresh_generated_files
        ).pack(side="right")
        ttk.Button(
            controls, text="Unduh / Simpan Salinan…",
            command=self._save_generated_copy,
        ).pack(side="right", padx=(0, 6))
        ttk.Button(
            controls, text="Buka File", command=self._open_generated_file
        ).pack(side="right", padx=(0, 6))
        columns = ("time", "file", "size", "location")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for key, title, width in (
            ("time", "Waktu dibuat", 280), ("file", "Nama file", 330),
            ("size", "Ukuran", 100), ("location", "Lokasi tersimpan", 500),
        ):
            tree.heading(key, text=title, anchor="w")
            tree.column(key, width=width, minwidth=90, stretch=key == "location")
        tree.bind("<Double-1>", lambda _event: self._open_generated_file())
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=1, column=0, sticky="nsew")
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal.grid(row=2, column=0, sticky="ew")
        self._generated_files_tree = tree
        self.notebook.add(frame, text="File Hasil")

    def _refresh_generated_files(self) -> None:
        tree = self._generated_files_tree
        if tree is None:
            return
        tree.delete(*tree.get_children())
        self._generated_file_paths.clear()
        for index, record in enumerate(self.audit.generated_outputs()):
            path = Path(record["path"])
            if not path.is_file():
                continue
            item = f"generated-{index}"
            self._generated_file_paths[item] = path
            size = path.stat().st_size
            size_label = (
                f"{size / (1024 * 1024):.1f} MB"
                if size >= 1024 * 1024 else f"{size / 1024:.0f} KB"
            )
            tree.insert(
                "", "end", iid=item,
                values=(
                    format_local_datetime(record["occurred_at"]),
                    path.name, size_label, str(path.parent),
                ),
            )

    def _selected_generated_file(self) -> Path | None:
        tree = self._generated_files_tree
        if tree is None or not tree.selection():
            messagebox.showinfo("File hasil", "Pilih file hasil terlebih dahulu.")
            return None
        path = self._generated_file_paths.get(tree.selection()[0])
        if path is None or not path.is_file():
            messagebox.showwarning(
                "File tidak ditemukan",
                "File hasil ini sudah dipindah atau dihapus dari komputer.",
            )
            self._refresh_generated_files()
            return None
        return path

    def _open_generated_file(self) -> None:
        path = self._selected_generated_file()
        if path is None:
            return
        try:
            if platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            elif platform.system() == "Windows":
                os.startfile(str(path))
            else:
                subprocess.Popen(["xdg-open", str(path)])
            self.audit.record(self.operator, "generated_file_opened", file=path.name)
        except OSError as exc:
            messagebox.showerror("File tidak bisa dibuka", str(exc))

    def _save_generated_copy(self) -> None:
        source = self._selected_generated_file()
        if source is None:
            return
        destination = filedialog.asksaveasfilename(
            title="Simpan salinan file hasil",
            initialfile=source.name,
            defaultextension=".xlsx",
            filetypes=[("Workbook Excel", "*.xlsx")],
        )
        if not destination:
            return
        target = Path(destination).expanduser()
        try:
            if target.resolve() != source.resolve():
                shutil.copy2(source, target)
            self.audit.record(
                self.operator, "generated_file_saved_copy",
                source=source.name, destination=str(target),
            )
            messagebox.showinfo("File disimpan", f"Salinan file tersimpan di:\n{target}")
        except OSError as exc:
            messagebox.showerror("Gagal menyimpan salinan", str(exc))

    def _add_history_tab(self) -> None:
        frame = ttk.Frame(self.notebook, padding=10)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        controls = ttk.Frame(frame)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(
            controls,
            text=f"Operator aktif: {self.operator} · riwayat disimpan lokal",
        ).pack(side="left")
        ttk.Button(
            controls, text="Periksa integritas", command=self._verify_audit
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            controls, text="Ekspor CSV", command=self._export_audit_csv
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            controls, text="Muat ulang", command=self._refresh_history
        ).pack(side="right")
        columns = ("time", "operator", "action", "details", "host", "ip", "mac")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for key, title, width in (
            ("time", "Waktu lokal", 300), ("operator", "Akun PC", 140),
            ("action", "Aktivitas", 180), ("details", "Rincian", 620),
            ("host", "Perangkat", 170), ("ip", "IP (masked)", 135),
            ("mac", "MAC (masked)", 145),
        ):
            tree.heading(key, text=title, anchor="w")
            tree.column(key, width=width, minwidth=90, stretch=key == "details")
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=1, column=0, sticky="nsew")
        vertical.grid(row=1, column=1, sticky="ns")
        horizontal.grid(row=2, column=0, sticky="ew")
        self._history_tree = tree
        self.notebook.add(frame, text="Riwayat")

    def _refresh_history(self) -> None:
        tree = self._history_tree
        if tree is None:
            return
        tree.delete(*tree.get_children())
        for event in self.audit.recent_events(1000):
            details = json.loads(event["details_json"])
            tree.insert(
                "", "end",
                values=(
                    format_local_datetime(event["occurred_at"]),
                    event["operator"], event["action"],
                    event["details_json"], event["host"],
                    details.get("ip_masked", "—"), details.get("mac_masked", "—"),
                ),
            )

    def _verify_audit(self) -> None:
        valid, broken_id = self.audit.verify_chain()
        self.audit.record(
            self.operator, "audit_integrity_checked",
            valid=valid, first_invalid_event=broken_id,
        )
        self._refresh_history()
        self._refresh_generated_files()
        if valid:
            messagebox.showinfo("Riwayat", "Rantai hash riwayat masih valid.")
        else:
            messagebox.showerror(
                "Riwayat berubah",
                f"Integritas riwayat gagal mulai event #{broken_id}.",
            )

    def _export_audit_csv(self) -> None:
        destination = filedialog.asksaveasfilename(
            title="Ekspor riwayat pemakaian",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile=f"ReconBuilder_History_{dt.datetime.now():%Y%m%d_%H%M%S}.csv",
        )
        if not destination:
            return
        path = self.audit.export_csv(Path(destination))
        self.audit.record(self.operator, "audit_exported", file=path.name)
        self._refresh_history()
        messagebox.showinfo("Riwayat diekspor", f"File riwayat:\n{path}")

    def _pick_xlsx(self, title: str) -> str:
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")],
        )

    def _browse_lb(self) -> None:
        selected = self._pick_xlsx("Upload file LB (Loket Bayar)")
        if selected:
            self._load_source("LB", Path(selected))

    def _browse_rb(self) -> None:
        selected = self._pick_xlsx("Upload file RB (Rajabiller)")
        if selected:
            self._load_source("RB", Path(selected))

    def _browse_recon(self) -> None:
        selected = self._pick_xlsx(
            "Upload file recon (Summary dan Records)"
        )
        if selected:
            self._load_recon_source(Path(selected))

    def _load_backdate_source(self, path: Path) -> None:
        if path.suffix.lower() != ".xlsx":
            return
        self._load_tokens["Backdate"] += 1
        token = self._load_tokens["Backdate"]
        self._loading.add("Backdate")
        for side in ("Backdate LB", "Backdate RB", "Backdate RECON"):
            self._frames[side] = None
            self._filtered_indices[side] = []
            self._page_index[side] = 0
            self._search_vars[side].set("")
            self._clear_preview(side)
            self._row_labels[side].set(f"Membaca {path.name}…")
        self.status.set("Membaca dan mengelompokkan sheet Backdate…")
        self.audit.record(
            self.operator, "backdate_upload_selected", file=path.name
        )
        threading.Thread(
            target=self._backdate_load_worker, args=(path, token), daemon=True
        ).start()

    def _backdate_load_worker(self, path: Path, token: int) -> None:
        try:
            frames = read_backdate_workbook(path)
            error = None
        except Exception as exc:
            frames = None
            error = str(exc)
        self.after(0, self._backdate_loaded, token, path, frames, error)

    def _backdate_loaded(
        self,
        token: int,
        path: Path,
        frames: dict[str, pl.DataFrame] | None,
        error: str | None,
    ) -> None:
        if token != self._load_tokens["Backdate"]:
            return
        self._loading.discard("Backdate")
        if error is not None or frames is None:
            self.audit.record(
                self.operator, "backdate_upload_failed",
                file=path.name, error=error or "unknown",
            )
            for side in ("Backdate LB", "Backdate RB", "Backdate RECON"):
                self._row_labels[side].set("File gagal dibaca.")
            self.status.set("Gagal membaca workbook Backdate.")
            messagebox.showerror(
                "Gagal membaca Backdate", error or "Workbook kosong."
            )
            return
        for sheet_name, frame in frames.items():
            side = f"Backdate {sheet_name}"
            self._frames[side] = frame
            self._filtered_indices[side] = list(range(frame.height))
            self._page_index[side] = 0
            self._render_preview(side)
        self.audit.record(
            self.operator, "backdate_workbook_loaded", file=path.name,
            sheets={name: frame.height for name, frame in frames.items()},
        )
        self.status.set(
            "Backdate siap ditinjau: " + ", ".join(
                f"{name} {frame.height:,} baris" for name, frame in frames.items()
            )
        )

    def _browse_backdate(self) -> None:
        selected = self._pick_xlsx("Upload File Backdate (.xlsx)")
        if not selected:
            return
        path = Path(selected)
        if path.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format tidak didukung",
                "File Backdate harus berformat Excel .xlsx.",
            )
            return
        self.backdate_path.set(str(path))
        self._load_backdate_source(path)

    def _browse_out(self) -> None:
        initial = Path(self.out_path.get()).expanduser()
        if not initial.is_dir():
            initial = Path.home() / "Desktop"
        selected = filedialog.askdirectory(
            title="Pilih folder untuk menyimpan hasil recon",
            initialdir=str(initial),
        )
        if selected:
            self.out_path.set(selected)

    def _page_size_for(self, side: str) -> int:
        selection = self._page_size_vars[side].get()
        if selection == "Semua":
            return max(1, len(self._filtered_indices[side]))
        return int(selection)

    def _load_source(self, side: str, path: Path) -> None:
        if path.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format tidak didukung", "Upload file Excel berformat .xlsx."
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
        self.audit.record(
            self.operator, "source_upload_selected", side=side, file=path.name
        )
        self._load_tokens[side] += 1
        token = self._load_tokens[side]
        self._loading.add(side)
        self._frames[side] = None
        self._filtered_indices[side] = []
        self._page_index[side] = 0
        self._selected_cells[side] = None
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

    def _load_recon_source(self, path: Path) -> None:
        if path.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format tidak didukung", "Upload file Excel berformat .xlsx."
            )
            return
        self.recon_path.set(str(path))
        self.audit.record(self.operator, "recon_upload_selected", file=path.name)
        self._load_tokens["Recon"] += 1
        token = self._load_tokens["Recon"]
        self._loading.add("Recon")
        for side in ("Recon Summary", "Recon Records"):
            self._frames[side] = None
            self._filtered_indices[side] = []
            self._page_index[side] = 0
            self._selected_cells[side] = None
            self._search_vars[side].set("")
            self._clear_preview(side)
            self._row_labels[side].set(f"Membaca {path.name}…")
        self.status.set("Membaca sheet Summary dan Records dari file recon…")
        threading.Thread(
            target=self._recon_load_worker,
            args=(path, token),
            daemon=True,
        ).start()

    def _recon_load_worker(self, path: Path, token: int) -> None:
        try:
            frames = read_recon_reference(path)
            error = None
        except Exception as exc:
            frames = None
            error = str(exc)
        self.after(0, self._recon_loaded, token, frames, error)

    def _recon_loaded(
        self,
        token: int,
        frames: dict[str, pl.DataFrame] | None,
        error: str | None,
    ) -> None:
        if token != self._load_tokens["Recon"]:
            return
        self._loading.discard("Recon")
        if error is not None or frames is None:
            for side in ("Recon Summary", "Recon Records"):
                self._frames[side] = None
                self._row_labels[side].set("File gagal dibaca.")
            self.status.set("Gagal membaca file recon.")
            self.audit.record(
                self.operator, "recon_upload_failed", file=Path(self.recon_path.get()).name,
                error=error or "unknown",
            )
            messagebox.showerror(
                "Gagal membaca file recon", error or "Data kosong."
            )
            return

        for side, frame in frames.items():
            self._frames[side] = frame
            self._filtered_indices[side] = list(range(frame.height))
            self._page_index[side] = 0
            self._render_preview(side)
        summary_count = frames["Recon Summary"].height
        record_count = frames["Recon Records"].height
        self.status.set(
            f"File recon siap sebagai referensi: {summary_count:,} baris "
            f"Summary, {record_count:,} baris Records."
        )
        self.audit.record(
            self.operator, "recon_reference_loaded",
            file=Path(self.recon_path.get()).name,
            sheets={name.removeprefix("Recon "): frame.height for name, frame in frames.items()},
        )

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
            self.audit.record(
                self.operator, "source_upload_failed", side=side,
                file=self._path_for_side(side).name, error=error or "unknown",
            )
            messagebox.showerror("Gagal membaca file", error or "Data kosong.")
            return

        self._frames[side] = frame
        self._filtered_indices[side] = list(range(frame.height))
        self._page_index[side] = 0
        self._render_preview(side)
        self.status.set(f"File {side} siap ditinjau dan diedit.")
        self.audit.record(
            self.operator, "source_loaded", side=side,
            file=self._path_for_side(side).name, rows=frame.height,
        )

    def _path_for_side(self, side: str) -> Path:
        raw = self.lb_path.get() if side == "LB" else self.rb_path.get()
        return Path(raw) if raw else Path("unknown.xlsx")

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
                text=SOURCE_HEADERS.get(
                    side.split()[-1] if side.startswith("Backdate ") else side,
                    {},
                ).get(column, column),
                anchor="w",
            )
            tree.column(column, width=145, minwidth=90, stretch=False)

        page_size = self._page_size_for(side)
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
            row_tag = "odd" if index % 2 else "even"
            tree.insert(
                "", "end", iid=str(index), values=values, tags=(row_tag,)
            )

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
        page_size = self._page_size_for(side)
        page_count = (len(self._filtered_indices[side]) + page_size - 1) // page_size
        self._page_index[side] = max(
            0, min(page_count - 1, self._page_index[side] + offset)
        )
        self.audit.record(
            self.operator, "preview_page_changed", tab=side,
            page=self._page_index[side] + 1,
        )
        self._render_preview(side)

    def _page_size_changed(self, side: str) -> None:
        self._page_index[side] = 0
        self.audit.record(
            self.operator, "preview_page_size_changed", tab=side,
            page_size=self._page_size_vars[side].get(),
        )
        self._render_preview(side)

    def _tab_changed(self, _event: tk.Event[Any]) -> None:
        selected = self.notebook.select()
        if selected:
            tab_title = self.notebook.tab(selected, "text")
            self.audit.record(
                self.operator, "tab_opened", tab=tab_title,
            )
            if tab_title == "Riwayat":
                self._refresh_history()

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
        self.audit.record(
            self.operator, "preview_searched", tab=side,
            query=query, matches=len(matches),
        )
        self._filtered_indices[side] = matches
        self._page_index[side] = 0
        self._render_preview(side)

    def _clear_search(self, side: str) -> None:
        self._search_vars[side].set("")
        self._apply_search(side)

    def _begin_edit(self, side: str, event: tk.Event) -> None:
        # Preview read-only: pengeditan sel dimatikan untuk semua tab.
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
        self._selected_cells[side] = (index, columns[column_index])
        self._open_cell_editor(side, index, columns[column_index])

    def _select_preview_cell(self, side: str, event: tk.Event) -> None:
        if side not in {"LB", "RB"}:
            return
        tree = self._preview_widgets[side]
        item = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        if not item or not column_id:
            return
        columns = list(tree["columns"])
        column_index = int(column_id[1:]) - 1
        if column_index < len(columns):
            self._selected_cells[side] = (
                int(item), columns[column_index]
            )

    def _edit_selected_cell(self, side: str) -> None:
        # Preview read-only: pengeditan sel dimatikan untuk semua tab.
        self.status.set(
            "Preview hanya untuk melihat dan mencari data (read-only)."
        )
        return

    def _open_cell_editor(
        self, side: str, index: int, column: str
    ) -> None:
        # Preview read-only: pengeditan sel dimatikan untuk semua tab.
        return

    def _commit_edit(
        self,
        side: str,
        row_index: int,
        column: str,
        editor: tk.Entry,
        event: tk.Event[Any] | None = None,
    ) -> bool | str:
        # Preview read-only: pengeditan sel dimatikan untuk semua tab.
        # Editor dihancurkan tanpa menerapkan perubahan apa pun.
        if self._editor is editor:
            try:
                editor.destroy()
            except Exception:
                pass
            self._editor = None
            self._editor_target = None
        self.status.set(
            "Preview hanya untuk melihat dan mencari data (read-only)."
        )
        return "break" if event is not None else False

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
        output_file = self.last_output_path
        folder = (
            output_file.parent if output_file is not None
            else Path(self.out_path.get()).expanduser()
        )
        if not folder.is_dir():
            messagebox.showwarning(
                "Folder tidak ada", f"Folder output belum ada:\n{folder}"
            )
            return
        current_platform = platform.system()
        if current_platform == "Darwin":
            command = ["open", "-R", str(output_file)] if output_file and output_file.is_file() else ["open", str(folder)]
        elif current_platform == "Windows":
            command = ["explorer", "/select,", str(output_file)] if output_file and output_file.is_file() else ["explorer", str(folder)]
        else:
            command = (
                ["xdg-open", str(output_file)]
                if output_file and output_file.is_file()
                else ["xdg-open", str(folder)]
            )
        try:
            subprocess.Popen(command)
            self.audit.record(
                self.operator, "output_location_opened",
                file=output_file.name if output_file else None,
            )
        except OSError as exc:
            messagebox.showerror("Folder tidak bisa dibuka", str(exc))

    def _on_close(self) -> None:
        if self._busy or self._loading or self._update_checking:
            messagebox.showwarning(
                "Proses masih berjalan",
                "Tunggu proses file atau pemeriksaan update selesai sebelum "
                "menutup aplikasi.",
            )
            return
        self.audit.record(self.operator, "app_closed")
        self.destroy()

    def _terminal_write(self, line: str) -> None:
        if self._terminal is None:
            return
        timestamp = dt.datetime.now().astimezone().strftime("%H:%M:%S")
        self._terminal.configure(state="normal")
        self._terminal.insert("end", f"[{timestamp}] {line}\n")
        self._terminal.see("end")
        self._terminal.configure(state="disabled")

    def _terminal_from_worker(self, line: str) -> None:
        self.after(0, self._terminal_write, line)

    def _test_voice(self) -> None:
        self.audit.record(self.operator, "voice_tested")
        announce("Halo, ini tes suara Recon Builder. Speaker berfungsi.")
        self.status.set("Tes suara dikirim ke speaker perangkat.")
        if self._voice_test_button is not None:
            self._voice_test_button.pack_forget()

    def _check_updates_silently(self) -> None:
        self._check_for_updates(silent=True)

    def _on_update_button(self) -> None:
        if self._busy:
            return
        if self._update_info is None:
            self._check_for_updates(silent=False)
            return
        release = self._update_info
        if not messagebox.askyesno(
            "Update tersedia",
            f"Recon Builder v{release.version} tersedia. Unduh dan pasang sekarang?\n\n"
            "Aplikasi akan ditutup dan dibuka kembali setelah proses update.",
        ):
            self.audit.record(
                self.operator, "update_declined", version=release.version
            )
            return
        self._busy = True
        self._update_button.configure(state="disabled", text="Mengunduh update…")
        self._terminal_write(f"Mengunduh Recon Builder v{release.version}…")
        self.audit.record(self.operator, "update_started", version=release.version)
        threading.Thread(
            target=self._download_and_start_update,
            args=(release,),
            daemon=True,
        ).start()

    def _check_for_updates(self, *, silent: bool) -> None:
        if self._update_checking or self._busy:
            return
        self._update_checking = True
        self.audit.record(
            self.operator, "update_check_started", automatic=silent
        )
        if self._update_button is not None:
            self._update_button.configure(state="disabled", text="Memeriksa…")
        self._release_status_var.set(
            f"APP v{__version__}  ·  RELEASE sedang diperiksa"
        )
        threading.Thread(
            target=self._update_check_worker, args=(silent,), daemon=True
        ).start()

    def _update_check_worker(self, silent: bool) -> None:
        try:
            release = latest_update(__version__)
            error = None
        except Exception as exc:
            release = None
            error = str(exc)
        self.after(0, self._update_check_finished, release, error, silent)

    def _update_check_finished(
        self, release: UpdateRelease | None, error: str | None, silent: bool
    ) -> None:
        self._update_checking = False
        if self._update_button is not None:
            self._update_button.configure(state="normal")
        if error:
            self._update_info = None
            if self._update_button is not None:
                self._update_button.configure(text="Periksa Update")
            self._release_status_var.set(
                f"APP v{__version__}  ·  RELEASE belum dapat diperiksa"
            )
            self.audit.record(self.operator, "update_check_failed", error=error)
            if not silent:
                messagebox.showwarning(
                    "Update",
                    "Tidak dapat memeriksa versi terbaru. "
                    "Pastikan komputer terhubung ke internet, "
                    "lalu klik Periksa Update lagi.",
                )
            return
        self._update_info = release
        if release is not None:
            if self._update_button is not None:
                self._update_button.configure(
                    text=f"Update v{release.version}",
                    style="Header.TButton",
                )
            self._release_status_var.set(
                f"APP v{__version__}  →  RELEASE v{release.version} tersedia"
            )
            self.audit.record(
                self.operator, "update_available", version=release.version,
                release_url=release.html_url,
            )
            if not silent:
                messagebox.showinfo(
                    "Update tersedia",
                    f"Versi terbaru v{release.version} siap dipasang.",
                )
        else:
            self.audit.record(
                self.operator, "update_check_completed",
                current_version=__version__, update_available=False,
            )
            if self._update_button is not None:
                self._update_button.configure(
                    text="Periksa Update",
                    style="Header.TButton",
                )
            self._release_status_var.set(
                f"APP v{__version__}  ·  RELEASE v{__version__} terbaru"
            )
            if not silent:
                messagebox.showinfo(
                    "Aplikasi terbaru",
                    f"Recon Builder v{__version__} sudah versi terbaru.",
                )

    def _download_and_start_update(self, release: UpdateRelease) -> None:
        try:
            installer = download_update(
                release,
                progress_callback=lambda completed, total: self.after(
                    0, self._update_download_progress, completed, total
                ),
            )
            start_installer(installer, os.getpid())
            error = None
        except Exception as exc:
            error = str(exc)
        self.after(0, self._update_install_finished, release, error)

    def _update_download_progress(self, completed: int, total: int) -> None:
        if total <= 0:
            return
        percent = min(100, int(completed * 100 / total))
        if self._update_button is not None:
            self._update_button.configure(text=f"Mengunduh update… {percent}%")

    def _update_install_finished(
        self, release: UpdateRelease, error: str | None
    ) -> None:
        self._busy = False
        if error:
            if self._update_button is not None:
                self._update_button.configure(
                    state="normal", text=f"Update v{release.version}"
                )
            self.audit.record(
                self.operator, "update_failed", version=release.version,
                error=error,
            )
            self._refresh_history()
            messagebox.showerror("Update gagal", error)
            return
        self.audit.record(
            self.operator, "update_installer_started", version=release.version
        )
        self.audit.record(self.operator, "app_closed_for_update")
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
                "Tunggu sampai pembacaan semua file selesai.",
            )
            return
        lb_path = Path(self.lb_path.get().strip()).expanduser()
        rb_path = Path(self.rb_path.get().strip()).expanduser()
        recon_path = Path(self.recon_path.get().strip()).expanduser()
        backdate_value = self.backdate_path.get().strip()
        output_dir = Path(self.out_path.get().strip()).expanduser()
        backdate_frames = None

        backdate_path: Path | None = None
        if backdate_value:
            backdate_path = Path(backdate_value).expanduser()
            if backdate_path.suffix.lower() != ".xlsx":
                messagebox.showerror(
                    "Format tidak didukung",
                    "File Backdate harus berformat Excel .xlsx.",
                )
                return
            if not backdate_path.is_file():
                messagebox.showerror(
                    "File Backdate",
                    "File Backdate yang dipilih tidak ditemukan.",
                )
                return
            backdate_frames = {
                sheet: self._frames[f"Backdate {sheet}"]
                for sheet in ("LB", "RB", "RECON")
            }
            if any(frame is None for frame in backdate_frames.values()):
                messagebox.showerror(
                    "File Backdate",
                    "Pastikan semua sheet Backdate selesai dimuat sebelum Generate.",
                )
                return

        if self._frames["LB"] is None or not lb_path.is_file():
            messagebox.showerror(
                "Upload LB", "Upload dan muat file LB terlebih dahulu."
            )
            return
        if self._frames["RB"] is None or not rb_path.is_file():
            messagebox.showerror(
                "Upload RB", "Upload dan muat file RB terlebih dahulu."
            )
            return
        if (
            self._frames["Recon Summary"] is None
            or self._frames["Recon Records"] is None
            or not recon_path.is_file()
        ):
            messagebox.showerror(
                "File recon",
                "Upload file recon yang memiliki sheet Summary dan Records.",
            )
            return
        if not self.out_path.get().strip():
            messagebox.showerror("File output", "Tentukan lokasi file output.")
            return
        if output_dir.exists() and not output_dir.is_dir():
            messagebox.showerror("Folder output", "Lokasi output harus berupa folder.")
            return

        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("Folder output", str(exc))
            return
        output = output_dir / dt.datetime.now().astimezone().strftime(
            "ReconBuilder_%Y%m%d_%H%M%S_%f.xlsx"
        )
        self._pending_output_path = output
        self._terminal.configure(state="normal")
        self._terminal.delete("1.0", "end")
        self._terminal.configure(state="disabled")
        self._terminal_write("Generate dimulai")
        self.audit.record(
            self.operator, "generate_started", output=output.name,
            output_path=str(output),
            lb_rows=self._frames["LB"].height,
            rb_rows=self._frames["RB"].height,
            backdate_file=backdate_path.name if backdate_path else None,
            backdate_rows=(
                {sheet: frame.height for sheet, frame in backdate_frames.items()}
                if backdate_frames else None
            ),
        )

        self._busy = True
        self.generate_button.configure(state="disabled")
        self.progress.start(10)
        announce("Mohon tunggu sebentar ya, file rekonsiliasi sedang dibuat.")
        self.status.set(
            "Membuat file recon gabungan Backdate dan End Date…"
            if backdate_path else
            "Membuat file recon dari data yang sudah ditinjau…"
        )
        started_at = time.perf_counter()
        threading.Thread(
            target=self._export_worker,
            args=(
                self._frames["LB"], self._frames["RB"], output,
                backdate_path, started_at, backdate_frames,
                self._frames.get("Recon Records"),
            ),
            daemon=True,
        ).start()

    def _export_worker(
        self, lb_frame: pl.DataFrame | None, rb_frame: pl.DataFrame | None,
        output: Path, backdate_path: Path | None, started_at: float,
        backdate_frames: dict[str, pl.DataFrame | None] | None,
        recon_records_frame: pl.DataFrame | None,
    ) -> None:
        try:
            if lb_frame is None or rb_frame is None:
                raise ValueError("File LB dan RB harus dimuat sebelum export.")
            if backdate_path is not None and backdate_frames is not None:
                self._terminal_from_worker(
                    f"Membaca Backdate: {backdate_path.name}"
                )
                self._terminal_from_worker(
                    "Sheet Backdate: " + ", ".join(
                        f"{name} {frame.height:,} baris"
                        for name, frame in backdate_frames.items()
                        if frame is not None
                    )
                )
            result = build_recon_workbook(
                lb_frame,
                rb_frame,
                output,
                backdate_lb_df=backdate_frames["LB"] if backdate_frames else None,
                backdate_rb_df=backdate_frames["RB"] if backdate_frames else None,
                backdate_recon_df=backdate_frames["RECON"] if backdate_frames else None,
                recon_records_df=recon_records_frame,
                progress_callback=self._terminal_from_worker,
            )
            self.after(0, self._export_finished, result, None, started_at)
        except Exception as exc:
            self.after(0, self._export_finished, None, str(exc), started_at)

    def _export_finished(
        self, result: Path | None, error: str | None, started_at: float
    ) -> None:
        elapsed = time.perf_counter() - started_at
        self.progress.stop()
        self._busy = False
        self.generate_button.configure(state="normal")
        if error is not None:
            self.status.set("Export gagal.")
            self._terminal_write(f"GAGAL: {error}\n")
            self.audit.record(
                self.operator, "generate_failed",
                output=(
                    self._pending_output_path.name
                    if self._pending_output_path else None
                ),
                output_path=(
                    str(self._pending_output_path)
                    if self._pending_output_path else None
                ),
                error=error, elapsed_seconds=round(elapsed, 3),
            )
            self._pending_output_path = None
            self._refresh_history()
            messagebox.showerror("Gagal membuat recon", error)
            return
        self.last_output_path = result
        self._pending_output_path = None
        self.status.set(f"Export selesai: {result.name}")
        self._terminal_write(
            f"Selesai dalam {elapsed:.1f} detik\nFile: {result}\n"
        )
        self.audit.record(
            self.operator, "generate_succeeded", output=result.name,
            output_path=str(result),
            elapsed_seconds=round(elapsed, 3), size_bytes=result.stat().st_size,
        )
        announce("Selesai. File rekonsiliasi dari BO Biller sudah berhasil dibuat.")
        self._refresh_history()
        self._refresh_generated_files()
        messagebox.showinfo("Recon selesai", f"File berhasil dibuat:\n{result}")


def main() -> None:
    ReconApp().mainloop()


if __name__ == "__main__":
    main()
