"""Tkinter desktop interface for creating reconciliation workbooks."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from src.recon_exporter import (
    LB_COL_MAP,
    RB_COL_MAP,
    build_recon_workbook,
    read_raw_excel,
)


class ReconApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Recon Builder - LB & RB Excel Splitter")
        self.geometry("760x500")
        self.resizable(False, False)

        self.lb_path = tk.StringVar()
        self.rb_path = tk.StringVar()
        self.out_path = tk.StringVar(
            value=str(Path.home() / "Desktop" / "recon-output.xlsx")
        )
        self._build_ui()

    def _build_ui(self) -> None:
        padding = {"padx": 14, "pady": 7}
        ttk.Label(
            self, text="Recon Builder", font=("Helvetica", 19, "bold")
        ).pack(pady=(18, 4))
        ttk.Label(
            self,
            text=(
                "Buat workbook recon dari file LB dan RB, lengkap dengan "
                "formula dan format sesuai file acuan."
            ),
            foreground="#555",
        ).pack(pady=(0, 10))

        fields = ttk.Frame(self)
        fields.pack(fill="x", **padding)
        self._add_file_row(fields, 0, "File LB (Loket Bayar):", self.lb_path,
                           self._browse_lb)
        self._add_file_row(fields, 1, "File RB (Rajabiller):", self.rb_path,
                           self._browse_rb)
        self._add_file_row(fields, 2, "File output recon:", self.out_path,
                           self._browse_out)

        actions = ttk.Frame(self)
        actions.pack(pady=12)
        self.generate_button = ttk.Button(
            actions, text="Generate Recon", command=self._on_generate
        )
        self.generate_button.pack(side="left", padx=6)
        ttk.Button(
            actions, text="Buka Folder Output", command=self._open_output_folder
        ).pack(side="left", padx=6)

        self.progress = ttk.Progressbar(
            self, mode="indeterminate", length=700
        )
        self.progress.pack(padx=14, pady=(5, 3))
        self.log = tk.Text(
            self, height=11, state="disabled", font=("Menlo", 10),
            bg="#1e1e1e", fg="#d4d4d4",
        )
        self.log.pack(fill="both", expand=True, padx=14, pady=(3, 14))
        self._log("Siap. Pilih file LB dan RB, lalu klik Generate Recon.")

    @staticmethod
    def _add_file_row(
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: object,
    ) -> None:
        ttk.Label(parent, text=label, width=22, anchor="w").grid(
            row=row, column=0, sticky="w", pady=5
        )
        ttk.Entry(parent, textvariable=variable, width=67).grid(
            row=row, column=1, padx=(0, 7), pady=5
        )
        ttk.Button(parent, text="Browse...", command=command).grid(
            row=row, column=2, pady=5
        )

    @staticmethod
    def _pick_xlsx(title: str) -> str:
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )

    def _browse_lb(self) -> None:
        selected = self._pick_xlsx("Pilih file LB (Loket Bayar)")
        if selected:
            self.lb_path.set(selected)

    def _browse_rb(self) -> None:
        selected = self._pick_xlsx("Pilih file RB (Rajabiller)")
        if selected:
            self.rb_path.set(selected)

    def _browse_out(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="Simpan file recon sebagai",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if selected:
            self.out_path.set(selected)

    def _open_output_folder(self) -> None:
        folder = Path(self.out_path.get()).expanduser().parent
        if not folder.is_dir():
            messagebox.showwarning(
                "Folder tidak ada", f"Folder output belum ada:\n{folder}"
            )
            return
        import subprocess
        import sys

        if sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        elif sys.platform.startswith("win"):
            subprocess.Popen(["explorer", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])

    def _log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _worker(self, lb: Path, rb: Path, output: Path) -> None:
        try:
            self.after(0, self._log, f"Membaca LB: {lb}")
            lb_df = read_raw_excel(lb, col_map=LB_COL_MAP)
            self.after(0, self._log, f"  -> {lb_df.height:,} baris")

            self.after(0, self._log, f"Membaca RB: {rb}")
            rb_df = read_raw_excel(
                rb, col_map=RB_COL_MAP, drop_col_non_numeric="no"
            )
            self.after(0, self._log, f"  -> {rb_df.height:,} baris")

            self.after(0, self._log, "Membuat workbook dengan Polars + XlsxWriter...")
            result = build_recon_workbook(lb_df, rb_df, output)
            self.after(0, self._log, f"Selesai -> {result}")
            self.after(
                0, lambda: messagebox.showinfo(
                    "Sukses", f"File recon berhasil dibuat:\n{result}"
                )
            )
        except Exception as error:
            message = str(error)
            self.after(0, self._log, f"ERROR: {message}")
            self.after(
                0, lambda: messagebox.showerror("Gagal", message)
            )
        finally:
            self.after(0, self._finish)

    def _on_generate(self) -> None:
        lb = Path(self.lb_path.get().strip()).expanduser()
        rb = Path(self.rb_path.get().strip()).expanduser()
        output = Path(self.out_path.get().strip()).expanduser()

        if not lb.is_file():
            messagebox.showerror("Error", "File LB belum dipilih / tidak ada.")
            return
        if not rb.is_file():
            messagebox.showerror("Error", "File RB belum dipilih / tidak ada.")
            return
        if lb.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format file tidak didukung",
                "File LB harus berformat Excel .xlsx.",
            )
            return
        if rb.suffix.lower() != ".xlsx":
            messagebox.showerror(
                "Format file tidak didukung",
                "File RB harus berformat Excel .xlsx.",
            )
            return
        if not self.out_path.get().strip():
            messagebox.showerror("Error", "Path output belum ditentukan.")
            return
        if output.resolve() in {lb.resolve(), rb.resolve()}:
            messagebox.showerror(
                "Error", "Path output tidak boleh sama dengan file input."
            )
            return
        if output.suffix.lower() != ".xlsx":
            messagebox.showerror("Error", "File output harus berakhiran .xlsx.")
            return

        self.generate_button.configure(state="disabled")
        self.progress.start(10)
        threading.Thread(
            target=self._worker, args=(lb, rb, output), daemon=True
        ).start()

    def _finish(self) -> None:
        self.progress.stop()
        self.generate_button.configure(state="normal")


def main() -> None:
    ReconApp().mainloop()


if __name__ == "__main__":
    main()
