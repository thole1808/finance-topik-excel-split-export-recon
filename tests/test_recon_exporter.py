from __future__ import annotations

import datetime as dt
import math
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import polars as pl

from src.recon_exporter import build_recon_workbook


MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _lb_row(
    timestamp: dt.datetime,
    status: str,
    reference: str,
    vendor_reference: str,
    amount: int,
) -> dict[str, object]:
    return {
        "tanggal": timestamp,
        "operator": "test",
        "nama": "test",
        "idpel": "123",
        "produk": "TEST",
        "periode": None,
        "status": status,
        "ref": reference,
        "lembar": 1,
        "tagihan": amount,
        "denda": 0,
        "admin": 0,
        "harga": amount + 10,
        "fee_loket": 2,
        "fee_bulan": 3,
        "total": amount + 10,
        "ref_id": vendor_reference,
    }


def _rb_row(
    timestamp: dt.datetime,
    status: str,
    reference: str,
    amount: int,
) -> dict[str, object]:
    return {
        "no": 1,
        "id_trx": f"trx-{reference}",
        "tanggal": timestamp,
        "kode_produk": "TEST",
        "pelanggan_1": "123",
        "pelanggan_2": None,
        "pelanggan_3": None,
        "pelanggan_4": None,
        "via": "test",
        "sn": "test",
        "ref": reference,
        "status": status,
        "nominal": amount,
    }


class ReconExporterBackdateTests(unittest.TestCase):
    def _sheet_cells(self, path: Path, sheet_number: int) -> dict[str, ET.Element]:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(
                archive.read(f"xl/worksheets/sheet{sheet_number}.xml")
            )
        return {
            cell.attrib["r"]: cell
            for cell in root.findall(f".//{MAIN_NS}c")
        }

    def _subtotal_amounts(
        self, path: Path, sheet_number: int, tomorrow_amount_column: str
    ) -> list[float]:
        values: list[float] = []
        for address, cell in self._sheet_cells(path, sheet_number).items():
            if not address.startswith("G") or not address[1:].isdigit():
                continue
            formula = cell.findtext(f"{MAIN_NS}f") or ""
            if not formula.startswith("SUM(G") or (
                f"SUM({tomorrow_amount_column}" not in formula
                and f"{tomorrow_amount_column}$4:" not in formula
            ):
                continue
            cached_value = cell.findtext(f"{MAIN_NS}v")
            if cached_value is not None:
                values.append(float(cached_value))
        return sorted(values)

    def _has_date(
        self, path: Path, sheet_number: int, column: str, timestamp: dt.datetime
    ) -> bool:
        expected_serial = (
            timestamp - dt.datetime(1899, 12, 30)
        ).total_seconds() / 86400
        for address, cell in self._sheet_cells(path, sheet_number).items():
            if not address.rstrip("0123456789") == column:
                continue
            value = cell.findtext(f"{MAIN_NS}v")
            if value is not None and math.isclose(
                float(value), expected_serial, abs_tol=1e-9
            ):
                return True
        return False

    def _cells_in_row(self, path: Path, sheet_number: int, row: int) -> dict[str, str]:
        cells = self._sheet_cells(path, sheet_number)
        result: dict[str, str] = {}
        for address, cell in cells.items():
            column = address.rstrip("0123456789")
            row_text = address[len(column):]
            if column not in {
                "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU",
                "AV", "AW", "AX", "AY", "AZ", "BA",
            } or not row_text.isdigit() or int(row_text) != row:
                continue
            value = cell.findtext(f"{MAIN_NS}v")
            if value is None:
                value = "".join(
                    cell.itertext()
                )
            result[column] = value
        return result

    def _find_lb_tomorrow_row(
        self, path: Path, timestamp: dt.datetime
    ) -> dict[str, str] | None:
        expected = (
            timestamp - dt.datetime(1899, 12, 30)
        ).total_seconds() / 86400
        for address, cell in self._sheet_cells(path, 3).items():
            if not address.startswith("AN") or not address[2:].isdigit():
                continue
            value = cell.findtext(f"{MAIN_NS}v")
            if value is not None and math.isclose(
                float(value), expected, abs_tol=1e-9
            ):
                return self._cells_in_row(path, 3, int(address[2:]))
        return None

    def test_combines_subtotals_keeps_source_references_and_end_date_23(self) -> None:
        backdate_lb = pl.DataFrame([
            _lb_row(
                dt.datetime(2026, 10, 14, 23, 30), "SUCCESS",
                "LB-BD-SOURCE", "LB-BD-VENDOR", 100,
            ),
            _lb_row(
                dt.datetime(2026, 10, 14, 22, 30), "FAILED",
                "LB-BD-FAILED", "LB-BD-VENDOR-FAILED", 200,
            ),
        ])
        enddate_lb = pl.DataFrame([
            _lb_row(
                dt.datetime(2026, 10, 15, 23, 30), "SUCCESS",
                "LB-END-SOURCE", "LB-END-VENDOR", 300,
            ),
            _lb_row(
                dt.datetime(2026, 10, 15, 12, 30), "FAILED",
                "LB-END-FAILED", "LB-END-VENDOR-FAILED", 400,
            ),
        ])
        backdate_rb = pl.DataFrame([
            _rb_row(
                dt.datetime(2026, 10, 14, 22, 30), "SUCCESS",
                "RB-BD-SOURCE", 50,
            ),
            _rb_row(
                dt.datetime(2026, 10, 14, 21, 30), "FAILED",
                "RB-BD-FAILED", 60,
            ),
        ])
        enddate_rb = pl.DataFrame([
            _rb_row(
                dt.datetime(2026, 10, 15, 23, 30), "SUCCESS",
                "RB-END-SOURCE", 70,
            ),
            _rb_row(
                dt.datetime(2026, 10, 15, 12, 30), "FAILED",
                "RB-END-FAILED", 80,
            ),
        ])
        backdate_recon = pl.DataFrame({
            "Datetime": [
                dt.datetime(2026, 10, 14, 23, 30),
                dt.datetime(2026, 10, 14, 22, 30),
            ],
            "Product": ["LB", "LB"],
            "Billing Number": ["LB-BD-BILLING-1", "LB-BD-BILLING-2"],
            "Billing Name": ["LB-BD-NAME-1", "LB-BD-NAME-2"],
            "Ref": ["RECON-BD-REF-1", "RECON-BD-REF-2"],
            "Vendor Ref": ["RECON-BD-VENDOR-1", "RECON-BD-VENDOR-2"],
            "Amount": [100, 200],
            "Total Fee": [5, 10],
            "Customer Fee": [2, 4],
            "LST Fee": [1, 2],
            "Margin Biller": [2, 4],
            "Transaction Status": ["SUCCESS", "FAILED"],
            "Vendor": ["Loket Bayar", "Loket Bayar"],
            "Recon Result": ["MATCHED", "INTERNAL"],
        })
        recon_records = pl.DataFrame({
            "Datetime": [
                dt.datetime(2026, 10, 15, 23, 30),
                dt.datetime(2026, 10, 15, 12, 30),
            ],
            "Product": ["LB", "LB"],
            "Billing Number": ["LB-END-BILLING-1", "LB-END-BILLING-2"],
            "Billing Name": ["LB-END-NAME-1", "LB-END-NAME-2"],
            "Ref": ["RECON-END-REF-1", "RECON-END-REF-2"],
            "Vendor Ref": ["RECON-END-VENDOR-1", "RECON-END-VENDOR-2"],
            "Amount": [300, 400],
            "Total Fee": [10, 15],
            "Customer Fee": [4, 6],
            "LST Fee": [2, 3],
            "Margin Biller": [4, 6],
            "Transaction Status": ["SUCCESS", "FAILED"],
            "Vendor": ["Loket Bayar", "Loket Bayar"],
            "Recon Result": ["MATCHED", "INTERNAL"],
        })

        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "recon.xlsx"
            build_recon_workbook(
                enddate_lb,
                enddate_rb,
                output,
                backdate_lb_df=backdate_lb,
                backdate_rb_df=backdate_rb,
                backdate_recon_df=backdate_recon,
                recon_records_df=recon_records,
            )

            sheets = pl.read_excel(output, sheet_id=0, engine="calamine")
            records = sheets["Records"]

            # Backdate records: Ref dan Vendor Ref harus dari sheet RECON (bukan LB/RB)
            lb_backdate = records.filter(pl.col("Ref") == "RECON-BD-REF-1")
            self.assertEqual(lb_backdate["Vendor Ref"].to_list(), ["RECON-BD-VENDOR-1"])

            # End date records: Ref dan Vendor Ref harus dari sheet Records upload (bukan LB/RB)
            lb_enddate = records.filter(pl.col("Ref") == "RECON-END-REF-1")
            self.assertEqual(lb_enddate["Vendor Ref"].to_list(), ["RECON-END-VENDOR-1"])

            # Pastikan Ref dan Vendor Ref di output TIDAK sama dengan nilai asli LB/RB
            lb_enddate_wrong = records.filter(pl.col("Ref") == "LB-END-SOURCE")
            self.assertEqual(lb_enddate_wrong.height, 0)

            self.assertEqual(
                records.filter(
                    pl.col("Datetime") == dt.datetime(2026, 10, 15, 23, 30)
                ).height,
                1,
            )
            self.assertEqual(
                self._subtotal_amounts(output, 3, "AT"),
                [400.0, 600.0],
            )
            self.assertEqual(
                self._subtotal_amounts(output, 4, "BD"),
                [120.0, 140.0],
            )

            self.assertFalse(
                self._has_date(
                    output, 3, "A", dt.datetime(2026, 10, 15, 23, 30)
                )
            )
            self.assertFalse(
                self._has_date(
                    output, 3, "U", dt.datetime(2026, 10, 15, 23, 30)
                )
            )

    def test_regular_export_keeps_end_date_23_without_backdate(self) -> None:
        lb = pl.DataFrame([
            _lb_row(
                dt.datetime(2026, 10, 15, 23, 30), "SUCCESS",
                "LB-END-SOURCE", "LB-END-VENDOR", 300,
            ),
        ])

        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "regular.xlsx"
            build_recon_workbook(lb, pl.DataFrame(), output)

            records = pl.read_excel(
                output, sheet_id=0, engine="calamine"
            )["Records"]
            self.assertEqual(records.height, 1)
            self.assertEqual(records["Ref"].to_list(), ["LB-END-SOURCE"])
            self.assertTrue(
                self._has_date(
                    output, 3, "AN", dt.datetime(2026, 10, 15, 23, 30)
                )
            )
            self.assertFalse(
                self._has_date(
                    output, 3, "A", dt.datetime(2026, 10, 15, 23, 30)
                )
            )
            self.assertFalse(
                self._has_date(
                    output, 3, "U", dt.datetime(2026, 10, 15, 23, 30)
                )
            )


if __name__ == "__main__":
    unittest.main()
