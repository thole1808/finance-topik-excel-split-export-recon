"""Create the LB/RB reconciliation workbook with Polars and XlsxWriter."""

from __future__ import annotations

import datetime as dt
import json
import math
import sys
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from typing import Any

import polars as pl
import xlsxwriter

from src.version import __version__


def _resource_path(name: str) -> Path:
    bundle_root = getattr(sys, "_MEIPASS", None)
    candidates = (
        (Path(bundle_root) / "src" / name) if bundle_root else Path("/nonexistent"),
        Path(__file__).resolve().parent / name,
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(f"File referensi aplikasi tidak ditemukan: {name}")


def _load_json(name: str) -> dict[str, Any]:
    with _resource_path(name).open(encoding="utf-8") as source:
        return json.load(source)


REFERENCE_STYLES = _load_json("ref_styles.json")
FEE_TABLES = _load_json("fee_tables.json")
REFERENCE_DATA = _load_json("reference_data.json")

DATETIME_NUMBER_FORMAT = "yyyy-mm-dd hh:mm:ss"
RECORDS_HEADERS = (
    "Datetime", "Product", "Billing Number", "Billing Name", "Ref",
    "Vendor Ref", "Amount", "Total Fee", "Customer Fee", "LST Fee",
    "Margin Biller", "Transaction Status", "Vendor", "Recon Result",
)

LB_SHEET_LAYOUT = {
    "A": {
        1: "Datetime", 2: "Product", 3: "Billing Number",
        4: "Billing Name", 5: "Ref", 6: "Vendor Ref", 7: "Amount",
        8: "Total Fee", 9: "Customer Fee", 10: "LST Fee",
        11: "Margin Biller", 12: "Transaction Status", 13: "Vendor",
        14: "Recon Result", 16: "Selisih Fee", 17: "Selisih Fee",
        18: "STATUS", 19: "ADA REFUND ? (N/A)",
    },
    "B": {
        21: "Tanggal", 22: "Operator", 23: "Nama", 24: "Idpel",
        25: "Produk", 26: "Periode", 27: "Status", 28: "Ref ID",
        29: "Lembar", 30: "Tagihan", 31: "Denda", 32: "Admin",
        33: "Harga", 34: "Fee Loket", 35: "Fee Bulanan", 36: "Total",
        37: "No. Reff",
    },
    "C": {
        40: "Datetime", 41: "Product", 42: "Billing Number",
        43: "Billing Name", 44: "Ref", 45: "Vendor Ref",
        46: "Amount", 47: "Total Fee", 48: "Customer Fee",
        49: "LST Fee", 50: "Margin Biller", 51: "Transaction Status",
        52: "Vendor", 53: "Recon Result", 56: "Tanggal",
        57: "Operator", 58: "Nama", 59: "Idpel", 60: "Produk",
        61: "Periode", 62: "Status", 63: "No. Reff (SN)",
        64: "Lembar", 65: "Tagihan", 66: "Denda", 67: "Admin",
        68: "Harga", 69: "Fee Loket", 70: "Fee Bulan", 71: "Total",
        72: "Ref ID",
    },
}

RB_SHEET_LAYOUT = {
    "A": {
        1: "Datetime", 2: "Product", 3: "Billing Number",
        4: "Billing Name", 5: "Ref", 6: "Vendor Ref", 7: "Amount",
        8: "Total Fee", 9: "Customer Fee", 10: "LST Fee",
        11: "Margin Biller", 12: "Transaction Status", 13: "Vendor",
        14: "Recon Result", 16: "STATUS", 18: "No", 19: "ID Trx",
        20: "Tanggal", 21: "Kode Produk", 22: "No Pelanggan-1",
        23: "No Pelanggan-2", 24: "No Pelanggan-3",
        25: "No Pelanggan-4", 26: "Via", 27: "SN", 28: "Ref",
        29: "Status", 30: "Total / Nominal",
    },
    "B": {
        34: "No", 35: "ID Trx", 36: "Tanggal", 37: "Kode Produk",
        38: "No Pelanggan-1", 39: "No Pelanggan-2",
        40: "No Pelanggan-3", 41: "No Pelanggan-4", 42: "Via",
        43: "SN", 44: "Ref", 45: "Status", 46: "Total / Nominal",
    },
    "C": {
        50: "Datetime", 51: "Product", 52: "Billing Number",
        53: "Billing Name", 54: "Ref", 55: "Vendor Ref",
        56: "Amount", 57: "Total Fee", 58: "Customer Fee",
        59: "LST Fee", 60: "Margin Biller",
        61: "Transaction Status", 62: "Vendor", 63: "Recon Result",
    },
}

LB_COL_MAP = {
    "Tanggal": "tanggal", "Operator": "operator", "Nama": "nama",
    "Idpel": "idpel", "Produk": "produk", "Periode": "periode",
    "Status": "status", "No. Reff (SN)": "ref", "Lembar": "lembar",
    "Tagihan": "tagihan", "Denda": "denda", "Admin": "admin",
    "Harga": "harga", "Fee Loket": "fee_loket",
    "Fee Bulan": "fee_bulan", "Total": "total", "Ref ID": "ref_id",
}
RB_COL_MAP = {
    "No": "no", "ID Trx": "id_trx", "Tanggal": "tanggal",
    "Kode Produk": "kode_produk", "No Pelanggan-1": "pelanggan_1",
    "No Pelanggan-2": "pelanggan_2", "No Pelanggan-3": "pelanggan_3",
    "No Pelanggan-4": "pelanggan_4", "Via": "via", "SN": "sn",
    "Ref": "ref", "Status": "status", "Nominal": "nominal",
}


def _clean(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def _as_datetime(value: Any) -> dt.datetime | None:
    value = _clean(value)
    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time())
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            return dt.datetime.fromisoformat(text.replace("Z", "+00:00")).replace(
                tzinfo=None
            )
        except ValueError:
            return None
    return None


def _datetime_value(value: Any) -> Any:
    parsed = _as_datetime(value)
    return parsed if parsed is not None else _clean(value)


def _is_tomorrow_transaction(row: Mapping[str, Any]) -> bool:
    if row.get("_period_order") == 0:
        return False
    vendor_row = row.get("_pair_vendor")
    bo_record = row.get("_pair_bo")
    source_time = (
        vendor_row.get("tanggal") if isinstance(vendor_row, Mapping)
        else bo_record.get("datetime") if isinstance(bo_record, Mapping)
        else row.get("tanggal")
    )
    timestamp = _as_datetime(source_time)
    return timestamp is not None and timestamp.time() >= dt.time(23, 0)


def _transaction_status_sort_key(row: Mapping[str, Any]) -> int:
    status = _transaction_status(row.get("status"))
    return {"SUCCESS": 0, "FAILED": 1}.get(status, 2)


def _status_highlight(status: Any) -> dict[str, Any] | None:
    normalized = _transaction_status(status)
    colors = {
        "SUCCESS": ("#C6EFCE", "#006100"),
        "FAILED": ("#FFC7CE", "#9C0006"),
    }
    color = colors.get(normalized)
    if color is None:
        return None
    return {
        "bg_color": color[0],
        "font_color": color[1],
        "pattern": 1,
    }


def _status_text(value: Any) -> str:
    cleaned = _clean(value)
    return str(cleaned).strip().upper() if cleaned is not None else ""


def _transaction_status(value: Any) -> str:
    status = _status_text(value)
    if status in {"SUKSES", "BERHASIL", "SUCCESS"}:
        return "SUCCESS"
    if status in {"GAGAL", "FAILED", "REFUND"}:
        return "FAILED"
    return status


def _recon_result(value: Any) -> str:
    if _transaction_status(value) == "SUCCESS":
        return "MATCHED"
    return "INTERNAL"


def read_raw_excel(
    path: Path,
    col_map: Mapping[str, str] | None = None,
    drop_col_non_numeric: str | None = None,
) -> pl.DataFrame:
    """Read the first worksheet, rename source headers, and discard footer rows."""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"File tidak ditemukan: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("File input harus berformat .xlsx.")

    frame = pl.read_excel(path, sheet_id=1, engine="calamine")
    if col_map:
        if (
            "Nominal" in col_map and "Nominal" not in frame.columns
            and "Total" in frame.columns
        ):
            # Some RB vendor exports use Total where the standard export uses
            # Nominal. Keep Total intact and add the required display alias.
            frame = frame.with_columns(pl.col("Total").alias("Nominal"))
        rename_map = {source: target for source, target in col_map.items()
                      if source in frame.columns}
        frame = frame.rename(rename_map)
        required = set(col_map.values())
        missing = sorted(required.difference(frame.columns))
        if missing:
            raise ValueError(
                f"Kolom wajib tidak ditemukan pada {path.name}: "
                f"{', '.join(missing)}"
            )

    if drop_col_non_numeric and drop_col_non_numeric in frame.columns:
        frame = frame.filter(
            pl.col(drop_col_non_numeric)
            .cast(pl.Float64, strict=False)
            .is_not_null()
        )

    return frame.with_row_index("_row_idx")


def read_recon_reference(path: Path) -> dict[str, pl.DataFrame]:
    """Read the Summary and Records sheets from an existing recon workbook."""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"File tidak ditemukan: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("File input harus berformat .xlsx.")

    frames: dict[str, pl.DataFrame] = {}
    for sheet_name in ("Summary", "Records"):
        try:
            frame = pl.read_excel(
                path, sheet_name=sheet_name, engine="calamine"
            )
        except Exception as exc:
            raise ValueError(
                f"Sheet wajib '{sheet_name}' tidak dapat dibaca dari "
                f"{path.name}: {exc}"
            ) from exc
        if not frame.columns:
            raise ValueError(
                f"Sheet '{sheet_name}' pada {path.name} tidak memiliki header."
            )
        frames[f"Recon {sheet_name}"] = frame.with_row_index("_row_idx")
    return frames


def read_backdate_workbook(path: Path) -> dict[str, pl.DataFrame]:
    """Read and validate the LB, RB, and RECON sheets from a Backdate workbook."""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"File Backdate tidak ditemukan: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError("File Backdate harus berformat .xlsx.")

    try:
        sheets = pl.read_excel(path, sheet_id=0, engine="calamine")
    except Exception as exc:
        raise ValueError(
            f"Workbook Backdate tidak dapat dibaca: {path.name}: {exc}"
        ) from exc

    missing_sheets = {"LB", "RB", "RECON"}.difference(sheets)
    if missing_sheets:
        raise ValueError(
            f"Sheet wajib tidak ditemukan pada {path.name}: "
            f"{', '.join(sorted(missing_sheets))}"
        )

    def normalize_source(
        sheet_name: str, column_map: Mapping[str, str]
    ) -> pl.DataFrame:
        frame = sheets[sheet_name]
        if (
            sheet_name == "RB" and "Nominal" not in frame.columns
            and "Total" in frame.columns
        ):
            frame = frame.with_columns(pl.col("Total").alias("Nominal"))
        rename_map = {
            source: target for source, target in column_map.items()
            if source in frame.columns
        }
        frame = frame.rename(rename_map)
        missing_columns = sorted(
            set(column_map.values()).difference(frame.columns)
        )
        if missing_columns:
            raise ValueError(
                f"Kolom wajib tidak ditemukan pada sheet {sheet_name} "
                f"{path.name}: {', '.join(missing_columns)}"
            )
        if sheet_name == "RB":
            frame = frame.filter(
                pl.col("no").cast(pl.Float64, strict=False).is_not_null()
            )
        return frame.with_row_index("_row_idx")

    recon = sheets["RECON"]
    missing_recon_columns = sorted(
        set(RECORDS_HEADERS).difference(recon.columns)
    )
    if missing_recon_columns:
        raise ValueError(
            f"Kolom wajib tidak ditemukan pada sheet RECON {path.name}: "
            f"{', '.join(missing_recon_columns)}"
        )

    return {
        "LB": normalize_source("LB", LB_COL_MAP),
        "RB": normalize_source("RB", RB_COL_MAP),
        "RECON": recon,
    }


def _number_key(value: Any) -> str:
    value = _clean(value)
    if isinstance(value, (int, float)) and float(value).is_integer():
        return str(int(value))
    return str(value)


def _fees(side: str, row: Mapping[str, Any], code: Any, amount: Any) -> tuple[Any, ...]:
    code_key = str(_clean(code) or "").strip()
    amount_key = _number_key(amount)
    reference = REFERENCE_DATA.get(f"{side}_fees", {}).get(
        f"{code_key}|{amount_key}"
    )
    if reference is not None:
        return tuple(reference)

    table = FEE_TABLES.get(side, {}).get(code_key)
    if table is not None:
        return tuple(table)

    customer_fee = _clean(row.get("fee_loket"))
    lst_fee = _clean(row.get("fee_bulan"))
    price = _clean(row.get("harga"))
    margin = None
    if isinstance(price, (int, float)) and isinstance(amount, (int, float)):
        margin = price - amount - (customer_fee or 0) - (lst_fee or 0)
    total_fee = None
    if all(isinstance(item, (int, float)) for item in (customer_fee, lst_fee, margin)):
        total_fee = customer_fee + lst_fee + margin
    return total_fee, customer_fee, lst_fee, margin


def _product_name(side: str, code: Any) -> Any:
    if code is None:
        return None
    normalized = str(code).strip()
    return REFERENCE_DATA.get(f"{side}_names", {}).get(normalized, normalized)


def _records(
    lb_rows: Sequence[Mapping[str, Any]],
    rb_rows: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in lb_rows:
        fees = _fees("lb", row, row.get("produk"), row.get("tagihan"))
        status = _transaction_status(row.get("status"))
        result.append({
            "datetime": _datetime_value(row.get("tanggal")),
            "product": _product_name("lb", row.get("produk")),
            "billing_number": _clean(row.get("idpel")),
            "billing_name": _clean(row.get("nama")),
            "ref": _clean(row.get("ref")),
            "vendor_ref": _clean(row.get("ref_id")),
            "amount": _clean(row.get("tagihan")),
            "total_fee": fees[0],
            "customer_fee": fees[1],
            "lst_fee": fees[2],
            "margin_biller": fees[3],
            "transaction_status": status,
            "vendor": "Loket Bayar",
            "recon_result": _recon_result(row.get("status")),
        })
    for row in rb_rows:
        fees = _fees("rb", row, row.get("kode_produk"), row.get("nominal"))
        status = _transaction_status(row.get("status"))
        result.append({
            "datetime": _datetime_value(row.get("tanggal")),
            "product": _product_name("rb", row.get("kode_produk")),
            "billing_number": _clean(row.get("pelanggan_1")),
            "billing_name": None,
            "ref": _clean(row.get("ref")),
            "vendor_ref": _clean(row.get("ref")),
            "amount": _clean(row.get("nominal")),
            "total_fee": fees[0],
            "customer_fee": fees[1],
            "lst_fee": fees[2],
            "margin_biller": fees[3],
            "transaction_status": status,
            "vendor": "Rajabiller",
            "recon_result": _recon_result(row.get("status")),
        })
    result.sort(
        key=lambda item: (
            _transaction_status_sort_key(
                {"status": item["transaction_status"]}
            ),
            _as_datetime(item["datetime"]) or dt.datetime.min,
        )
    )
    return result


def _format_options(style: Mapping[str, Any]) -> dict[str, Any]:
    options: dict[str, Any] = {}
    if style.get("fill"):
        options["bg_color"] = f"#{style['fill']}"
        options["pattern"] = 1
    if style.get("bold"):
        options["bold"] = True
    if style.get("size"):
        options["font_size"] = style["size"]
    if style.get("font"):
        options["font_name"] = style["font"]
    if style.get("color"):
        options["font_color"] = f"#{style['color']}"
    if style.get("numfmt"):
        options["num_format"] = style["numfmt"]
    if style.get("halign"):
        options["align"] = style["halign"]
    if style.get("valign"):
        options["valign"] = style["valign"]
    if style.get("wrap"):
        options["text_wrap"] = True
    border_style = {
        "hair": 7,
        "dotted": 4,
        "dashDot": 9,
        "dashed": 3,
        "thin": 1,
        "medium": 2,
        "thick": 5,
        "double": 6,
    }
    if style.get("border_sides"):
        for side, value in style["border_sides"].items():
            options[side] = border_style.get(value, 1)
    elif style.get("border"):
        options["border"] = 1
    return options


def _format_for(
    workbook: xlsxwriter.Workbook,
    sheet: str,
    template_row: int,
    column: int,
    cache: dict[tuple[str, int, int], xlsxwriter.format.Format],
) -> xlsxwriter.format.Format | None:
    key = (sheet, template_row, column)
    if key not in cache:
        row_styles = REFERENCE_STYLES[sheet]["styles"].get(str(template_row), {})
        style = row_styles.get(str(column), {})
        options = _format_options(style)
        cache[key] = workbook.add_format(options)
    return cache[key]


def _write_cell(
    worksheet: xlsxwriter.worksheet.Worksheet,
    workbook: xlsxwriter.Workbook,
    sheet: str,
    row: int,
    column: int,
    template_row: int,
    value: Any = None,
    formula: str | None = None,
    cached_value: Any = None,
    format_cache: dict[tuple[str, int, int], xlsxwriter.format.Format] | None = None,
    format_overrides: Mapping[str, Any] | None = None,
) -> None:
    cache = format_cache if format_cache is not None else {}
    cell_format = _format_for(workbook, sheet, template_row, column, cache)
    if format_overrides:
        style = (
            REFERENCE_STYLES[sheet]["styles"]
            .get(str(template_row), {})
            .get(str(column), {})
        )
        options = _format_options(style)
        options.update(format_overrides)
        cell_format = workbook.add_format(options)
    if formula is not None:
        worksheet.write_formula(row, column - 1, formula, cell_format, cached_value)
        return

    value = _clean(value)
    if value is None:
        worksheet.write_blank(row, column - 1, None, cell_format)
    elif isinstance(value, (dt.datetime, dt.date)):
        worksheet.write_datetime(row, column - 1, value, cell_format)
    elif isinstance(value, bool):
        worksheet.write_boolean(row, column - 1, value, cell_format)
    elif isinstance(value, (int, float)):
        worksheet.write_number(row, column - 1, value, cell_format)
    else:
        worksheet.write_string(row, column - 1, str(value), cell_format)


def _write_template_rows(
    worksheet: xlsxwriter.worksheet.Worksheet,
    workbook: xlsxwriter.Workbook,
    sheet: str,
    max_column: int,
    values_by_row: Mapping[int, Mapping[int, Any]],
    format_cache: dict[tuple[str, int, int], xlsxwriter.format.Format],
    formulas_by_row: Mapping[
        int, Mapping[int, tuple[str, Any]]
    ] | None = None,
    format_overrides_by_row: Mapping[
        int, Mapping[int, Mapping[str, Any]]
    ] | None = None,
) -> None:
    for row_number in (1, 2, 3):
        row_values = values_by_row.get(row_number, {})
        row_formulas = (
            formulas_by_row.get(row_number, {})
            if formulas_by_row is not None else {}
        )
        template_row = row_number
        for column in range(1, max_column + 1):
            formula_data = row_formulas.get(column)
            _write_cell(
                worksheet, workbook, sheet, row_number - 1, column,
                template_row, row_values.get(column),
                formula=formula_data[0] if formula_data else None,
                cached_value=formula_data[1] if formula_data else None,
                format_cache=format_cache,
                format_overrides=(
                    format_overrides_by_row.get(row_number, {}).get(column)
                    if format_overrides_by_row is not None else None
                ),
            )
def _set_widths(
    worksheet: xlsxwriter.worksheet.Worksheet, sheet: str
) -> None:
    for letter, width in REFERENCE_STYLES[sheet]["widths"].items():
        # XlsxWriter adds Excel's 5px column padding when serializing widths.
        worksheet.set_column(f"{letter}:{letter}", max(0, width - 5 / 7))


def _to_rows(frame: pl.DataFrame) -> list[dict[str, Any]]:
    return [
        {key: _clean(value) for key, value in row.items()}
        for row in frame.iter_rows(named=True)
    ]


def _vendor_amount(row: Mapping[str, Any], side: str) -> Any:
    """Return the Vendor Total compared to the BO Biller Amount.

    Standard Rajabiller workbooks have Nominal instead of Total, so Nominal
    remains the fallback for that source format.
    """
    for key in ("total", "Total"):
        value = _clean(row.get(key))
        if value is not None:
            return value
    return _clean(row.get("nominal")) if side == "RB" else None


def _lb_values(row: Mapping[str, Any]) -> dict[int, Any]:
    if "_pair_vendor" in row or "_pair_bo" in row:
        vendor = row.get("_pair_vendor")
        bo = row.get("_pair_bo")
        vendor = vendor if isinstance(vendor, Mapping) else {}
        bo = bo if isinstance(bo, Mapping) else {}
        fees = tuple(bo.get(key) for key in (
            "total_fee", "customer_fee", "lst_fee", "margin_biller",
        ))
        date_value = _datetime_value(bo.get("datetime"))
        status = _transaction_status(bo.get("transaction_status"))
        recon = {
            1: date_value, 2: bo.get("product"),
            3: bo.get("billing_number"), 4: bo.get("billing_name"),
            5: bo.get("ref"), 6: bo.get("vendor_ref"),
            7: bo.get("amount"), 8: fees[0], 9: fees[1],
            10: fees[2], 11: fees[3], 12: status,
            13: bo.get("vendor"),
            14: row.get("_pair_issue") or bo.get("recon_result"),
            40: date_value, 41: bo.get("product"),
            42: bo.get("billing_number"), 43: bo.get("billing_name"),
            44: bo.get("ref"), 45: bo.get("vendor_ref"),
            46: bo.get("amount"), 47: fees[0], 48: fees[1],
            49: fees[2], 50: fees[3], 51: status,
            52: bo.get("vendor"),
            53: row.get("_pair_issue") or bo.get("recon_result"),
        }
        raw = {
            21: _datetime_value(vendor.get("tanggal")),
            22: vendor.get("operator"), 23: vendor.get("nama"),
            24: vendor.get("idpel"), 25: vendor.get("produk"),
            26: vendor.get("periode"), 27: vendor.get("status"),
            28: vendor.get("ref_id"), 29: vendor.get("lembar"),
            30: vendor.get("tagihan"), 31: vendor.get("denda"),
            32: vendor.get("admin"), 33: vendor.get("harga"),
            34: vendor.get("fee_loket"), 35: vendor.get("fee_bulan"),
            36: vendor.get("total"), 37: vendor.get("ref"),
        }
        raw_next = {
            56: _datetime_value(vendor.get("tanggal")),
            57: vendor.get("operator"), 58: vendor.get("nama"),
            59: vendor.get("idpel"), 60: vendor.get("produk"),
            61: vendor.get("periode"), 62: vendor.get("status"),
            63: vendor.get("ref"), 64: vendor.get("lembar"),
            65: vendor.get("tagihan"), 66: vendor.get("denda"),
            67: vendor.get("admin"), 68: vendor.get("harga"),
            69: vendor.get("fee_loket"), 70: vendor.get("fee_bulan"),
            71: vendor.get("total"), 72: vendor.get("ref_id"),
        }
        return {**raw, **recon, **raw_next}

    fees = (
        row.get("_sync_total_fee"), row.get("_sync_customer_fee"),
        row.get("_sync_lst_fee"), row.get("_sync_margin"),
    )
    if not any(value is not None for value in fees):
        fees = _fees("lb", row, row.get("produk"), row.get("tagihan"))
    date_value = _datetime_value(row.get("tanggal"))
    status = _transaction_status(row.get("status"))
    result = _recon_result(row.get("status"))
    product = _product_name("lb", row.get("produk"))
    raw = {
        21: date_value, 22: row.get("operator"), 23: row.get("nama"),
        24: row.get("idpel"), 25: row.get("produk"),
        26: row.get("periode"), 27: row.get("status"),
        28: row.get("ref_id"), 29: row.get("lembar"),
        30: row.get("tagihan"), 31: row.get("denda"),
        32: row.get("admin"), 33: row.get("harga"),
        34: row.get("fee_loket"), 35: row.get("fee_bulan"),
        36: row.get("total"), 37: row.get("ref"),
    }
    recon = {
        1: date_value, 2: product, 3: row.get("idpel"),
        4: row.get("nama"), 5: row.get("ref"), 6: row.get("ref_id"),
        7: row.get("tagihan"), 8: fees[0], 9: fees[1], 10: fees[2],
        11: fees[3], 12: status, 13: "Loket Bayar", 14: result,
        40: date_value, 41: product, 42: row.get("idpel"),
        43: row.get("nama"), 44: row.get("ref"), 45: row.get("ref_id"),
        46: row.get("tagihan"), 47: fees[0], 48: fees[1],
        49: fees[2], 50: fees[3], 51: status, 52: "Loket Bayar",
        53: result,
    }
    raw_next = {
        56: date_value, 57: row.get("operator"),
        58: row.get("nama"), 59: row.get("idpel"),
        60: row.get("produk"), 61: row.get("periode"),
        62: row.get("status"), 63: row.get("ref"), 64: row.get("lembar"),
        65: row.get("tagihan"), 66: row.get("denda"),
        67: row.get("admin"), 68: row.get("harga"),
        69: row.get("fee_loket"), 70: row.get("fee_bulan"),
        71: row.get("total"), 72: row.get("ref_id"),
    }
    return {**raw, **recon, **raw_next}


def _rb_values(row: Mapping[str, Any]) -> dict[int, Any]:
    if "_pair_vendor" in row or "_pair_bo" in row:
        vendor = row.get("_pair_vendor")
        bo = row.get("_pair_bo")
        vendor = vendor if isinstance(vendor, Mapping) else {}
        bo = bo if isinstance(bo, Mapping) else {}
        fees = tuple(bo.get(key) for key in (
            "total_fee", "customer_fee", "lst_fee", "margin_biller",
        ))
        date_value = _datetime_value(bo.get("datetime"))
        status = _transaction_status(bo.get("transaction_status"))
        recon = {
            1: date_value, 2: bo.get("product"),
            3: bo.get("billing_number"), 4: bo.get("billing_name"),
            5: bo.get("ref"), 6: bo.get("vendor_ref"),
            7: bo.get("amount"), 8: fees[0], 9: fees[1],
            10: fees[2], 11: fees[3], 12: status,
            13: bo.get("vendor"),
            14: row.get("_pair_issue") or bo.get("recon_result"),
            50: date_value, 51: bo.get("product"),
            52: bo.get("billing_number"), 53: bo.get("billing_name"),
            54: bo.get("ref"), 55: bo.get("vendor_ref"),
            56: bo.get("amount"), 57: fees[0], 58: fees[1],
            59: fees[2], 60: fees[3], 61: status,
            62: bo.get("vendor"),
            63: row.get("_pair_issue") or bo.get("recon_result"),
        }
        raw = {
            18: vendor.get("no"), 19: vendor.get("id_trx"),
            20: _datetime_value(vendor.get("tanggal")),
            21: vendor.get("kode_produk"), 22: vendor.get("pelanggan_1"),
            23: vendor.get("pelanggan_2"), 24: vendor.get("pelanggan_3"),
            25: vendor.get("pelanggan_4"), 26: vendor.get("via"),
            27: vendor.get("sn"), 28: vendor.get("ref"),
            29: vendor.get("status"), 30: _vendor_amount(vendor, "RB"),
            34: vendor.get("no"), 35: vendor.get("id_trx"),
            36: _datetime_value(vendor.get("tanggal")),
            37: vendor.get("kode_produk"),
            38: vendor.get("pelanggan_1"), 39: vendor.get("pelanggan_2"),
            40: vendor.get("pelanggan_3"), 41: vendor.get("pelanggan_4"),
            42: vendor.get("via"), 43: vendor.get("sn"),
            44: vendor.get("ref"), 45: vendor.get("status"),
            46: _vendor_amount(vendor, "RB"),
        }
        return {**raw, **recon}

    fees = (
        row.get("_sync_total_fee"), row.get("_sync_customer_fee"),
        row.get("_sync_lst_fee"), row.get("_sync_margin"),
    )
    if not any(value is not None for value in fees):
        fees = _fees("rb", row, row.get("kode_produk"), row.get("nominal"))
    date_value = _datetime_value(row.get("tanggal"))
    status = _transaction_status(row.get("status"))
    result = _recon_result(row.get("status"))
    product = _product_name("rb", row.get("kode_produk"))
    raw = {
        18: row.get("no"), 19: row.get("id_trx"),
        20: date_value,
        21: row.get("kode_produk"), 22: row.get("pelanggan_1"),
        23: row.get("pelanggan_2"), 24: row.get("pelanggan_3"),
        25: row.get("pelanggan_4"), 26: row.get("via"), 27: row.get("sn"),
        28: row.get("ref"), 29: row.get("status"),
        30: _vendor_amount(row, "RB"),
        34: row.get("no"), 35: row.get("id_trx"),
        36: date_value,
        37: row.get("kode_produk"), 38: row.get("pelanggan_1"),
        39: row.get("pelanggan_2"), 40: row.get("pelanggan_3"),
        41: row.get("pelanggan_4"), 42: row.get("via"), 43: row.get("sn"),
        44: row.get("ref"), 45: row.get("status"),
        46: _vendor_amount(row, "RB"),
    }
    recon = {
        1: date_value, 2: product, 3: row.get("pelanggan_1"),
        4: None, 5: row.get("ref"), 6: row.get("ref"),
        7: row.get("nominal"), 8: fees[0], 9: fees[1], 10: fees[2],
        11: fees[3], 12: status, 13: "Rajabiller", 14: result,
        50: date_value, 51: product, 52: row.get("pelanggan_1"),
        53: None, 54: row.get("ref"), 55: row.get("ref"),
        56: row.get("nominal"), 57: fees[0], 58: fees[1],
        59: fees[2], 60: fees[3], 61: status, 62: "Rajabiller",
        63: result,
    }
    return {**raw, **recon}


def _write_status_total_row(
    worksheet: xlsxwriter.worksheet.Worksheet,
    workbook: xlsxwriter.Workbook,
    sheet: str,
    total_row: int,
    first_data_row: int,
    last_data_row: int,
    group_rows: Sequence[Mapping[str, Any]],
    values_function: Any,
    format_cache: dict[tuple[str, int, int], xlsxwriter.format.Format],
) -> None:
    current_columns = (7, 8, 9, 10, 11)
    totals = [0.0] * len(current_columns)
    for record in group_rows:
        values = values_function(record)
        for index, column in enumerate(current_columns):
            value = _clean(values.get(column))
            if isinstance(value, (int, float)):
                totals[index] += float(value)

    amount_style = (
        REFERENCE_STYLES[sheet]["styles"].get("4", {}).get("7", {})
    )
    common_overrides: dict[str, Any] = {"bold": True, "top": 1}
    if amount_style.get("numfmt"):
        common_overrides["num_format"] = amount_style["numfmt"]
    workbook_amount = totals[0] + totals[4]
    _write_cell(
        worksheet, workbook, sheet, total_row - 1, 6, 4,
        formula=f"=G{total_row}+K{total_row}",
        cached_value=workbook_amount,
        format_cache=format_cache,
        format_overrides=common_overrides,
    )
    for index, current_column in enumerate(current_columns):
        current_letter = xlsxwriter.utility.xl_col_to_name(current_column - 1)
        formula = (
            f"=SUM({current_letter}{first_data_row}:"
            f"{current_letter}{last_data_row})"
        )
        _write_cell(
            worksheet, workbook, sheet, total_row - 1, current_column, 4,
            formula=formula,
            cached_value=totals[index],
            format_cache=format_cache,
            format_overrides=common_overrides,
        )


def _write_wide_sheet(
    worksheet: xlsxwriter.worksheet.Worksheet,
    workbook: xlsxwriter.Workbook,
    sheet: str,
    rows: Sequence[Mapping[str, Any]],
    values_function: Any,
    layout: Mapping[str, Mapping[int, str]],
    max_column: int,
    format_cache: dict[tuple[str, int, int], xlsxwriter.format.Format],
) -> None:
    rows = sorted(
        rows,
        key=lambda row: (
            row.get("_period_order", 0),
            _transaction_status_sort_key(row),
            _as_datetime(row.get("tanggal")) or dt.datetime.min,
        ),
    )
    current_day_rows = [
        record for record in rows if not _is_tomorrow_transaction(record)
    ]
    title_columns = (
        (1, "Current Day from bo biller recon"),
        (21 if sheet == "LB" else 19, f"Current Day from Vendor {sheet}"),
        (
            40 if sheet == "LB" else 34,
            f"Vendor {sheet} for tomorrow because different time",
        ),
    )
    values_by_row: dict[int, dict[int, Any]] = {
        1: {column: text for column, text in title_columns},
        2: {1: "SUCCESS", 3: "FAILED"},
        3: {},
    }
    normalized_statuses = [
        _transaction_status(row.get("status")) for row in current_day_rows
    ]
    spacer_rows = 3
    output_slots: list[
        tuple[int, Mapping[str, Any] | None, Mapping[str, Any] | None]
    ] = []
    subtotal_after_slot: dict[
        int, tuple[int, int, Sequence[Mapping[str, Any]]]
    ] = {}
    next_row = 4
    for status in ("SUCCESS", "FAILED", "OTHER"):
        def matches_status(row: Mapping[str, Any]) -> bool:
            row_status = _transaction_status(row.get("status"))
            return (
                row_status == status
                if status != "OTHER"
                else row_status not in {"SUCCESS", "FAILED"}
            )

        first_data_row = next_row
        status_rows: list[Mapping[str, Any]] = []
        current_status_rows = [
            row for row in rows
            if matches_status(row) and not _is_tomorrow_transaction(row)
        ]
        tomorrow_status_rows = [
            row for row in rows
            if matches_status(row) and _is_tomorrow_transaction(row)
        ]
        status_rows.extend(current_status_rows)
        status_rows.extend(tomorrow_status_rows)
        slot_count = max(len(current_status_rows), len(tomorrow_status_rows))
        for index in range(slot_count):
            output_slots.append((
                next_row,
                current_status_rows[index]
                if index < len(current_status_rows) else None,
                tomorrow_status_rows[index]
                if index < len(tomorrow_status_rows) else None,
            ))
            next_row += 1
        if not status_rows:
            continue
        last_data_row = next_row - 1
        if status in {"SUCCESS", "FAILED"}:
            subtotal_after_slot[len(output_slots) - 1] = (
                first_data_row,
                last_data_row,
                status_rows,
            )
            next_row += 1 + spacer_rows
    source_data_end_row = output_slots[-1][0] if output_slots else 3
    current_day_end_row = max(
        (excel_row for excel_row, current, _ in output_slots
         if current is not None),
        default=3,
    )
    formulas_by_row = {
        2: {
            2: (
                f'=COUNTIF($L$4:$L${current_day_end_row},"SUCCESS")',
                normalized_statuses.count("SUCCESS"),
            ),
            4: (
                f'=COUNTIF($L$4:$L${current_day_end_row},"FAILED")',
                normalized_statuses.count("FAILED"),
            ),
        }
    }
    for section in ("A", "B", "C"):
        values_by_row[3].update(layout[section])
    output_max_column = max(max_column, 79) if sheet == "LB" else max_column
    _write_template_rows(
        worksheet,
        workbook,
        sheet,
        output_max_column,
        values_by_row,
        format_cache,
        formulas_by_row=formulas_by_row,
    )

    tomorrow_recon_start = 40 if sheet == "LB" else 50
    tomorrow_recon_end = 53 if sheet == "LB" else 63
    tomorrow_rows = (
        sorted(
            (row for row in rows if _is_tomorrow_transaction(row)),
            key=lambda row: (
                _transaction_status_sort_key(row),
                _as_datetime(row.get("tanggal")) or dt.datetime.min,
            ),
        )
        if sheet == "LB" else []
    )
    current_records_by_row = {
        excel_row: current_record
        for excel_row, current_record, _ in output_slots
    }
    paired_tomorrow_by_row = {
        excel_row: tomorrow_record
        for excel_row, _, tomorrow_record in output_slots
        if tomorrow_record is not None
    }
    subtotal_rows = {
        output_slots[index][0] + 1: subtotal_data
        for index, subtotal_data in subtotal_after_slot.items()
    }
    last_output_row = max(
        [3 + len(tomorrow_rows), *current_records_by_row, *subtotal_rows],
        default=3,
    )

    for excel_row in range(4, last_output_row + 1):
        current_day_record = current_records_by_row.get(excel_row)
        tomorrow_index = excel_row - 4
        if sheet == "LB":
            tomorrow_record = (
                tomorrow_rows[tomorrow_index]
                if tomorrow_index < len(tomorrow_rows) else None
            )
        else:
            tomorrow_record = paired_tomorrow_by_row.get(excel_row)
        current_day_values = (
            values_function(current_day_record)
            if current_day_record is not None else {}
        )
        tomorrow_values = (
            values_function(tomorrow_record)
            if tomorrow_record is not None else {}
        )
        formulas: dict[int, tuple[str, Any]] = {}
        if sheet == "LB" and current_day_record is not None:
            formula_row = excel_row
            source_row = current_day_record.get("_pair_vendor")
            if not isinstance(source_row, Mapping):
                source_row = current_day_record
            amount = _clean(source_row.get("tagihan"))
            price = _clean(source_row.get("harga"))
            margin = _clean(current_day_values.get(11))
            p_value = (
                (price - amount) - margin
                if isinstance(price, (int, float))
                and isinstance(amount, (int, float))
                and isinstance(margin, (int, float))
                else 0
            )
            q_value = (
                (amount + margin) - price
                if isinstance(price, (int, float))
                and isinstance(amount, (int, float))
                and isinstance(margin, (int, float))
                else 0
            )
            formulas = {
                16: (
                    f"=(AG{formula_row}-AD{formula_row})"
                    f"-K{formula_row}",
                    p_value,
                ),
                17: (
                    f"=(G{formula_row}+K{formula_row})"
                    f"-AG{formula_row}",
                    q_value,
                ),
                18: (
                    f'=OR(L{formula_row}="SUCCESS",'
                    f'AA{formula_row}="SUKSES")',
                    _transaction_status(current_day_record.get("status"))
                    == "SUCCESS",
                ),
            }
            if (
                current_day_rows
                and current_day_record is current_day_rows[0]
            ):
                formulas[19] = (
                    f'=_xlfn._xlws.FILTER(Z4:Z{source_data_end_row},'
                    f'Z4:Z{source_data_end_row}="N/A","Tidak ditemukan")',
                    "Tidak ditemukan",
                )
        elif sheet == "RB" and current_day_record is not None:
            row_number = excel_row
            formulas = {
                16: (
                    f'=OR(L{row_number}="SUCCESS",'
                    f'AC{row_number}="BERHASIL")',
                    _transaction_status(current_day_record.get("status"))
                    == "SUCCESS",
                ),
            }

        for column in range(1, output_max_column + 1):
            status_columns = (
                {12, 51} if sheet == "LB"
                else {12, 61}
            )
            if sheet == "LB":
                is_main_column = 1 <= column <= 19
                is_current_source_column = 21 <= column <= 37
                is_tomorrow_source_column = 56 <= column <= 72
                is_tomorrow_recon_column = 40 <= column <= 53
            else:
                is_main_column = 1 <= column <= 16
                is_current_source_column = 18 <= column <= 30
                is_tomorrow_source_column = 34 <= column <= 46
                is_tomorrow_recon_column = 50 <= column <= 63
            is_tomorrow_column = (
                tomorrow_recon_start <= column <= tomorrow_recon_end
            )
            if is_main_column or is_current_source_column:
                value = current_day_values.get(column)
                status_row = current_day_record
                has_status_row = status_row is not None
            elif is_tomorrow_column:
                value = tomorrow_values.get(
                    column - tomorrow_recon_start + 1
                )
                status_row = tomorrow_record
                has_status_row = status_row is not None
            elif is_tomorrow_source_column:
                value = tomorrow_values.get(column)
                status_row = tomorrow_record
                has_status_row = status_row is not None
            else:
                value = None
                status_row = None
                has_status_row = False
            row_highlight = (
                _status_highlight(status_row.get("status"))
                if (
                    column in status_columns
                    or (sheet == "LB" and column == 51)
                ) and has_status_row
                else None
            )
            formula_data = formulas.get(column)
            date_columns = (
                {1, 21, 40, 56} if sheet == "LB"
                else {1, 20, 36, 50}
            )
            if sheet == "LB" and is_tomorrow_recon_column:
                date_columns = {40}
            cell_overrides = (
                {"num_format": DATETIME_NUMBER_FORMAT}
                if column in date_columns else None
            )
            template_row = (
                5 if sheet == "RB" and 34 <= column <= 46
                and excel_row % 2 == 1 else 4
            )
            _write_cell(
                worksheet, workbook, sheet, excel_row - 1, column,
                4 if is_tomorrow_recon_column else template_row,
                value,
                formula=formula_data[0] if formula_data else None,
                cached_value=formula_data[1] if formula_data else None,
                format_cache=format_cache,
                format_overrides={
                    **(cell_overrides or {}),
                    **(row_highlight or {}),
                } or None,
            )
        if excel_row in subtotal_rows:
            first_data_row, last_data_row, status_rows = subtotal_rows[excel_row]
            _write_status_total_row(
                worksheet, workbook, sheet, excel_row,
                first_data_row, last_data_row, status_rows,
                values_function, format_cache,
            )

    worksheet.autofilter(
        2, 0, max(2, current_day_end_row - 1), 13
    )


def _records_from_recon_sheet(
    df: pl.DataFrame,
) -> list[dict[str, Any]]:
    """Convert a RECON/Records sheet DataFrame directly into records dicts.

    Column names match RECORDS_HEADERS exactly:
    Datetime, Product, Billing Number, Billing Name, Ref, Vendor Ref,
    Amount, Total Fee, Customer Fee, LST Fee, Margin Biller,
    Transaction Status, Vendor, Recon Result.
    """
    result: list[dict[str, Any]] = []
    for row in df.iter_rows(named=True):
        result.append({
            "datetime": _datetime_value(row.get("Datetime")),
            "product": _clean(row.get("Product")),
            "billing_number": _clean(row.get("Billing Number")),
            "billing_name": _clean(row.get("Billing Name")),
            "ref": _clean(row.get("Ref")),
            "vendor_ref": _clean(row.get("Vendor Ref")),
            "amount": _clean(row.get("Amount")),
            "total_fee": _clean(row.get("Total Fee")),
            "customer_fee": _clean(row.get("Customer Fee")),
            "lst_fee": _clean(row.get("LST Fee")),
            "margin_biller": _clean(row.get("Margin Biller")),
            "transaction_status": _transaction_status(row.get("Transaction Status")),
            "vendor": _clean(row.get("Vendor")),
            "recon_result": _clean(row.get("Recon Result")),
        })
    return result


def _vendor_rows_from_source(
    df: pl.DataFrame, side: str
) -> list[dict[str, Any]]:
    """Normalize raw or Backdate vendor sheets to the internal source schema."""
    mapping = LB_COL_MAP if side == "LB" else RB_COL_MAP
    rename = {header: field for header, field in mapping.items()
              if header in df.columns}
    normalized = df.rename(rename)
    return _to_rows(normalized)


def _record_vendor_side(record: Mapping[str, Any]) -> str | None:
    vendor = " ".join(str(record.get("vendor") or "").casefold().split())
    product = " ".join(str(record.get("product") or "").casefold().split())
    if "loket bayar" in vendor or vendor == "lb":
        return "LB"
    if "rajabiller" in vendor or vendor == "rb":
        return "RB"
    if "loket bayar" in product or product == "lb" or product.startswith("lb "):
        return "LB"
    if "rajabiller" in product or product == "rb" or product.startswith("rb "):
        return "RB"
    return None


def _pair_vendor_and_bo(
    vendor_rows: Sequence[Mapping[str, Any]],
    bo_records: Sequence[Mapping[str, Any]],
    vendor_side: str,
    period_order: int,
) -> tuple[list[dict[str, Any]], int, int, int]:
    """Pair source rows one-to-one; ambiguous candidates stay unmatched."""
    def text_key(value: Any) -> str:
        value = _clean(value)
        return "" if value is None else "".join(str(value).split()).casefold()

    def billing_key(value: Any) -> str:
        value = _clean(value)
        if value is None:
            return ""
        normalized = "".join(str(value).split()).casefold()
        digits = "".join(ch for ch in normalized if ch.isdigit())
        return (digits.lstrip("0") or digits) if digits else normalized

    def amount_key(value: Any) -> int | None:
        value = _clean(value)
        if value is None:
            return None
        try:
            return round(float(value) * 100)
        except (TypeError, ValueError):
            return None

    def timestamp_key(value: Any) -> str:
        parsed = _as_datetime(value)
        return parsed.strftime("%Y-%m-%d %H:%M:%S") if parsed else ""

    def keys(row: Mapping[str, Any], bo: bool) -> tuple[Any, ...]:
        if bo:
            ref, vendor_ref = row.get("ref"), row.get("vendor_ref")
            bn, amount, timestamp = (
                row.get("billing_number"), row.get("amount"),
                row.get("datetime"),
            )
        elif vendor_side == "LB":
            ref, vendor_ref = row.get("ref"), row.get("ref_id")
            bn, amount, timestamp = (
                row.get("idpel"), _vendor_amount(row, vendor_side),
                row.get("tanggal"),
            )
        else:
            ref, vendor_ref = row.get("ref"), row.get("sn")
            bn, amount, timestamp = (
                row.get("pelanggan_1"), _vendor_amount(row, vendor_side),
                row.get("tanggal"),
            )
        ref_key, vendor_ref_key = text_key(ref), text_key(vendor_ref)
        bn_key, amt_key = billing_key(bn), amount_key(amount)
        time_key = timestamp_key(timestamp)
        date_key = time_key[:10] if time_key else ""
        if bo:
            # BO Vendor Ref may correspond to either LB's Ref ID or its
            # transaction serial (No. Reff/SN). BO Ref is a biller reference.
            vendor_reference = vendor_ref_key
            alternate_reference = vendor_ref_key
        elif vendor_side == "LB":
            vendor_reference = text_key(row.get("ref_id"))
            alternate_reference = text_key(row.get("ref"))
        else:
            vendor_reference = text_key(row.get("sn"))
            alternate_reference = text_key(row.get("id_trx"))
        return (
            ("vendor_ref", vendor_reference) if vendor_reference else None,
            ("vendor_ref", alternate_reference)
            if alternate_reference else None,
            ("bn_amount_time", bn_key, amt_key, time_key)
            if bn_key and amt_key is not None and time_key else None,
            ("bn_amount_date", bn_key, amt_key, date_key)
            if bn_key and amt_key is not None and date_key else None,
            ("amount_time", amt_key, time_key)
            if amt_key is not None and time_key else None,
            ("amount_date", amt_key, date_key)
            if amt_key is not None and date_key else None,
            ("bn_time", bn_key, time_key)
            if bn_key and time_key else None,
            ("bn_date", bn_key, date_key)
            if bn_key and date_key else None,
        )

    vendor_keys = [keys(row, False) for row in vendor_rows]
    bo_keys = [keys(row, True) for row in bo_records]
    matches: dict[int, int] = {}
    used_bo: set[int] = set()
    for tier in range(8):
        vendor_groups: dict[Any, list[int]] = {}
        bo_groups: dict[Any, list[int]] = {}
        for index, row_keys in enumerate(vendor_keys):
            if index in matches or row_keys[tier] is None:
                continue
            vendor_groups.setdefault(row_keys[tier], []).append(index)
        for index, row_keys in enumerate(bo_keys):
            if index in used_bo or row_keys[tier] is None:
                continue
            bo_groups.setdefault(row_keys[tier], []).append(index)
        for key, vendor_indices in vendor_groups.items():
            bo_indices = bo_groups.get(key, [])
            if len(vendor_indices) == len(bo_indices) == 1:
                vendor_index, bo_index = vendor_indices[0], bo_indices[0]
                vendor, bo = vendor_rows[vendor_index], bo_records[bo_index]
                vendor_date = _as_datetime(vendor.get("tanggal"))
                bo_date = _as_datetime(bo.get("datetime"))
                vendor_amount_value = amount_key(
                    _vendor_amount(vendor, vendor_side)
                )
                bo_amount = amount_key(bo.get("amount"))
                vendor_number = billing_key(
                    vendor.get("idpel") if vendor_side == "LB"
                    else vendor.get("pelanggan_1")
                )
                bo_number = billing_key(bo.get("billing_number"))
                vendor_status = _transaction_status(vendor.get("status"))
                bo_status = _transaction_status(bo.get("transaction_status"))
                # Exact transaction references are authoritative. Other
                # heuristic keys still require matching date, bill number,
                # and status to avoid pairing unrelated repeated payments.
                if tier <= 1:
                    if vendor_status and bo_status and vendor_status != bo_status:
                        continue
                elif tier >= 6:
                    if (
                        not vendor_number or vendor_number != bo_number
                        or vendor_date is None or bo_date is None
                        or vendor_date.date() != bo_date.date()
                        or (vendor_status and bo_status and vendor_status != bo_status)
                    ):
                        continue
                elif (
                    not vendor_number or vendor_number != bo_number
                    or vendor_date is None or bo_date is None
                    or vendor_date.date() != bo_date.date()
                    or (vendor_status and bo_status and vendor_status != bo_status)
                    or vendor_amount_value is None or vendor_amount_value != bo_amount
                ):
                    continue
                matches[vendor_index] = bo_index
                used_bo.add(bo_index)

    result: list[dict[str, Any]] = []
    for index, vendor in enumerate(vendor_rows):
        bo_index = matches.get(index)
        bo = bo_records[bo_index] if bo_index is not None else None
        date_value = (
            bo.get("datetime") if bo else vendor.get("tanggal")
        )
        status = (
            bo.get("transaction_status") if bo
            else vendor.get("status")
        )
        result.append({
            "_pair_vendor": vendor,
            "_pair_bo": bo,
            "_period_order": period_order,
            "tanggal": date_value,
            "status": status,
            "tagihan": vendor.get("tagihan") if vendor_side == "LB" else None,
            "harga": vendor.get("harga") if vendor_side == "LB" else None,
            "matched": bo is not None,
            "_pair_issue": None,
        })
    for index, bo in enumerate(bo_records):
        if index not in used_bo:
            result.append({
                "_pair_vendor": None,
                "_pair_bo": bo,
                "_period_order": period_order,
                "tanggal": bo.get("datetime"),
                "status": bo.get("transaction_status"),
                "tagihan": None,
                "harga": None,
                "matched": False,
                "_pair_issue": "NO VENDOR MATCH",
            })
    return result, len(matches), len(vendor_rows) - len(matches), len(bo_records) - len(used_bo)


def build_recon_workbook(
    lb_df: pl.DataFrame,
    rb_df: pl.DataFrame,
    output_path: Path,
    backdate_lb_df: pl.DataFrame | None = None,
    backdate_rb_df: pl.DataFrame | None = None,
    backdate_recon_df: pl.DataFrame | None = None,
    recon_records_df: pl.DataFrame | None = None,
    audit_metadata: Mapping[str, str] | None = None,
    progress_callback: Callable[[str], None] | None = None,
) -> Path:
    """Write a recon workbook, optionally merging Backdate rows before End Date.

    Perlindungan nilai Ref dan Vendor Ref dari BO BILLER:
    - Backdate Records diambil dari sheet RECON workbook Backdate (kode asli BO BILLER),
      JANGAN dari sheet LB/RB. Nilai Ref dan Vendor Ref dipertahankan asli dari data RECON.
    - End Date Records diambil dari sheet Records file recon referensi, nilai
      Ref dan Vendor Ref dipertahankan asli (murni dari BO BILLER), tidak
      dimodifikasi oleh aplikasi.
    """
    output_path = Path(output_path).expanduser().resolve()
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Path output harus berakhiran .xlsx.")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    backdate_frames = (backdate_lb_df, backdate_rb_df, backdate_recon_df)
    has_backdate = all(frame is not None for frame in backdate_frames)
    if any(frame is not None for frame in backdate_frames) and not has_backdate:
        raise ValueError(
            "Data Backdate harus menyertakan sheet LB, RB, dan RECON."
        )
    lb_rows = (
        [{**row, "_period_order": 0}
         for row in _vendor_rows_from_source(backdate_lb_df, "LB")]
        if backdate_lb_df is not None else []
    )
    lb_rows.extend(
        [{**row, "_period_order": 1} for row in _vendor_rows_from_source(lb_df, "LB")]
        if has_backdate else _vendor_rows_from_source(lb_df, "LB")
    )
    rb_rows = (
        [{**row, "_period_order": 0}
         for row in _vendor_rows_from_source(backdate_rb_df, "RB")]
        if backdate_rb_df is not None else []
    )
    rb_rows.extend(
        [{**row, "_period_order": 1} for row in _vendor_rows_from_source(rb_df, "RB")]
        if has_backdate else _vendor_rows_from_source(rb_df, "RB")
    )
    bo_records_by_period: dict[int, list[dict[str, Any]]] = {0: [], 1: []}
    if progress_callback:
        progress_callback("Menyusun sumber BO Biller dan data vendor.")

    # Backdate RECON adalah sumber BO Biller untuk vendor Backdate.
    if backdate_recon_df is not None and backdate_recon_df.height > 0:
        bo_records_by_period[0] = _records_from_recon_sheet(backdate_recon_df)

    # Records dari file recon adalah sumber BO Biller untuk End Date.
    if recon_records_df is not None and recon_records_df.height > 0:
        bo_records_by_period[1] = _records_from_recon_sheet(recon_records_df)

    has_bo_references = any(bo_records_by_period.values())
    if has_bo_references:
        unclassified = [
            row for period_rows in bo_records_by_period.values()
            for row in period_rows if _record_vendor_side(row) is None
        ]
        if unclassified:
            raise ValueError(
                f"{len(unclassified):,} baris Records BO Biller tidak dapat "
                "ditentukan sebagai LB atau RB. Pastikan kolom Vendor berisi "
                "Loket Bayar/Rajabiller, atau Product berisi penanda LB/RB."
            )
        for side, source_rows in (("LB", lb_rows), ("RB", rb_rows)):
            paired_rows: list[dict[str, Any]] = []
            matched_count = vendor_only_count = bo_only_count = 0
            for period_order in (0, 1):
                vendor_period = [
                    row for row in source_rows
                    if row.get("_period_order", 1) == period_order
                ]
                bo_period = [
                    row for row in bo_records_by_period[period_order]
                    if _record_vendor_side(row) == side
                ]
                result, matched, vendor_only, bo_only = _pair_vendor_and_bo(
                    vendor_period, bo_period, side, period_order
                )
                paired_rows.extend(result)
                matched_count += matched
                vendor_only_count += vendor_only
                bo_only_count += bo_only
            if side == "LB":
                lb_rows = paired_rows
            else:
                rb_rows = paired_rows
            if progress_callback:
                progress_callback(
                    f"{side}: {matched_count:,} pasangan cocok; "
                    f"{vendor_only_count:,} baris vendor tanpa pasangan BO; "
                    f"{bo_only_count:,} baris BO tanpa pasangan vendor."
                )

    workbook = xlsxwriter.Workbook(
        str(output_path),
        {"constant_memory": True, "strings_to_formulas": False},
    )
    workbook.set_properties({
        "title": "Recon Builder - LB & RB",
        "subject": "Reconciliation workbook",
        "author": (audit_metadata or {}).get("Generated By", "Recon Builder"),
    })
    workbook.set_custom_property("Recon Builder Version", __version__)
    for key, value in (audit_metadata or {}).items():
        workbook.set_custom_property(key, value[:255])
    format_cache: dict[
        tuple[str, int, int], xlsxwriter.format.Format
    ] = {}

    ws_lb = workbook.add_worksheet("LB")
    ws_lb.hide_gridlines(2)
    ws_lb.freeze_panes(3, 0)
    _set_widths(ws_lb, "LB")
    if progress_callback:
        progress_callback(f"Membangun sheet LB ({len(lb_rows):,} baris sumber).")
    _write_wide_sheet(
        ws_lb, workbook, "LB", lb_rows, _lb_values, LB_SHEET_LAYOUT, 72,
        format_cache,
    )

    ws_rb = workbook.add_worksheet("RB")
    ws_rb.hide_gridlines(2)
    ws_rb.freeze_panes(3, 0)
    _set_widths(ws_rb, "RB")
    if progress_callback:
        progress_callback(f"Membangun sheet RB ({len(rb_rows):,} baris sumber).")
    _write_wide_sheet(
        ws_rb, workbook, "RB", rb_rows, _rb_values, RB_SHEET_LAYOUT, 63,
        format_cache,
    )
    workbook.close()
    if progress_callback:
        progress_callback("Workbook ditutup dan file Excel siap.")
    return output_path


__all__ = [
    "LB_COL_MAP", "RB_COL_MAP", "build_recon_workbook", "read_raw_excel",
    "read_backdate_workbook", "read_recon_reference",
]
