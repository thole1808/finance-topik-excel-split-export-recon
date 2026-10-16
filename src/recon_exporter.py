"""Create the LB/RB reconciliation workbook with Polars and XlsxWriter."""

from __future__ import annotations

import datetime as dt
import json
import math
import sys
from collections.abc import Mapping, Sequence
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

SUMMARY_HEADERS = (
    "Category", "Count", "Amount", "Amount Success", "Amount Failed",
)
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
        29: "Status", 30: "Nominal",
    },
    "B": {
        34: "No", 35: "ID Trx", 36: "Tanggal", 37: "Kode Produk",
        38: "No Pelanggan-1", 39: "No Pelanggan-2",
        40: "No Pelanggan-3", 41: "No Pelanggan-4", 42: "Via",
        43: "SN", 44: "Ref", 45: "Status", 46: "Nominal",
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
    timestamp = _as_datetime(row.get("tanggal"))
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


def _write_summary(
    worksheet: xlsxwriter.worksheet.Worksheet,
    workbook: xlsxwriter.Workbook,
    records: Sequence[Mapping[str, Any]],
    format_cache: dict[tuple[str, int, int], xlsxwriter.format.Format],
) -> None:
    for col, header in enumerate(SUMMARY_HEADERS, start=1):
        _write_cell(
            worksheet, workbook, "Summary", 0, col, 1, header,
            format_cache=format_cache,
        )

    last_record_row = max(2, len(records) + 1)
    end = f"{last_record_row}"
    categories = (
        "Internal Transactions",
        "Rajabiller Transactions",
        "Loket Bayar Transactions",
        "Matched Transactions",
        "Unmatched (Internal Only)",
        "Unmatched (Rajabiller Only)",
        "Unmatched (Loket Bayar Only)",
    )
    for summary_row, category in enumerate(categories, start=2):
        write_row = summary_row - 1
        _write_cell(
            worksheet, workbook, "Summary", write_row, 1, 2, category,
            format_cache=format_cache,
        )
        recon_range = f"Records!$N$2:$N${end}"
        vendor_range = f"Records!$M$2:$M${end}"
        status_range = f"Records!$L$2:$L${end}"
        amount_range = f"Records!$G$2:$G${end}"
        fee_range = f"Records!$H$2:$H${end}"
        vendor = None
        recon = None
        if category == "Internal Transactions":
            count_formula = (
                f'=COUNTIF({recon_range},"MATCHED")+'
                f'COUNTIF({recon_range},"INTERNAL")'
            )
            amount_formula = (
                f'=SUMIF({recon_range},"MATCHED",{amount_range})+'
                f'SUMIF({recon_range},"INTERNAL",{amount_range})+'
                f'SUMIF({recon_range},"MATCHED",{fee_range})+'
                f'SUMIF({recon_range},"INTERNAL",{fee_range})'
            )
            recon = {"MATCHED", "INTERNAL"}
        elif category == "Rajabiller Transactions":
            vendor, recon = "Rajabiller", {"MATCHED", "EXTERNAL"}
        elif category == "Loket Bayar Transactions":
            vendor, recon = "Loket Bayar", {"MATCHED", "EXTERNAL"}
        elif category == "Matched Transactions":
            recon = {"MATCHED"}
        elif category == "Unmatched (Internal Only)":
            recon = {"INTERNAL"}
        elif category == "Unmatched (Rajabiller Only)":
            vendor, recon = "Rajabiller", {"EXTERNAL"}
        else:
            vendor, recon = "Loket Bayar", {"EXTERNAL"}

        cached_records = [
            item for item in records
            if (recon is None or item["recon_result"] in recon)
            and (vendor is None or item["vendor"] == vendor)
        ]
        if category != "Internal Transactions":
            criteria = ",".join(
                f'{recon_range},"{result}"' for result in sorted(recon or ())
            )
            if vendor:
                vendor_criteria = f',{vendor_range},"{vendor}"'
                count_formula = (
                    f"=COUNTIFS({criteria}{vendor_criteria})"
                    if len(recon or ()) == 1
                    else f"=COUNTIFS({criteria},{vendor_range},\"{vendor}\")"
                )
                # COUNTIFS criteria pairs must be complete for each result.
                if len(recon or ()) > 1:
                    count_formula = "+".join(
                        f'COUNTIFS({recon_range},"{result}",'
                        f'{vendor_range},"{vendor}")'
                        for result in sorted(recon or ())
                    )
                    count_formula = f"={count_formula}"
            elif len(recon or ()) == 1:
                count_formula = (
                    f'=COUNTIF({recon_range},"{next(iter(recon or ())) }")'
                )
            else:
                count_formula = "+".join(
                    f'COUNTIF({recon_range},"{result}")'
                    for result in sorted(recon or ())
                )
                count_formula = f"={count_formula}"

            if vendor:
                amount_formula = "+".join(
                    f'SUMIFS({amount_range},{recon_range},"{result}",'
                    f'{vendor_range},"{vendor}")+'
                    f'SUMIFS({fee_range},{recon_range},"{result}",'
                    f'{vendor_range},"{vendor}")'
                    for result in sorted(recon or ())
                )
                amount_formula = f"={amount_formula}"
            elif len(recon or ()) == 1:
                result = next(iter(recon or ()))
                amount_formula = (
                    f'=SUMIF({recon_range},"{result}",{amount_range})+'
                    f'SUMIF({recon_range},"{result}",{fee_range})'
                )
            else:
                amount_formula = "+".join(
                    f'SUMIF({recon_range},"{result}",{amount_range})+'
                    f'SUMIF({recon_range},"{result}",{fee_range})'
                    for result in sorted(recon or ())
                )
                amount_formula = f"={amount_formula}"

        cached_count = len(cached_records)
        cached_amount = sum(
            float(item["amount"] or 0) + float(item["total_fee"] or 0)
            for item in cached_records
        )
        success_amount = sum(
            float(item["amount"] or 0) + float(item["total_fee"] or 0)
            for item in cached_records
            if item["transaction_status"] == "SUCCESS"
        )
        failed_amount = sum(
            float(item["amount"] or 0) + float(item["total_fee"] or 0)
            for item in cached_records
            if item["transaction_status"] == "FAILED"
        )
        if category == "Internal Transactions":
            success_formula = (
                f'=SUMIFS({amount_range},{recon_range},"MATCHED",'
                f'{status_range},"SUCCESS")+SUMIFS({amount_range},'
                f'{recon_range},"INTERNAL",{status_range},"SUCCESS")+'
                f'SUMIFS({fee_range},{recon_range},"MATCHED",'
                f'{status_range},"SUCCESS")+SUMIFS({fee_range},'
                f'{recon_range},"INTERNAL",{status_range},"SUCCESS")'
            )
            failed_formula = (
                f'=SUMIFS({amount_range},{recon_range},"MATCHED",'
                f'{status_range},"FAILED")+SUMIFS({amount_range},'
                f'{recon_range},"INTERNAL",{status_range},"FAILED")+'
                f'SUMIFS({fee_range},{recon_range},"MATCHED",'
                f'{status_range},"FAILED")+SUMIFS({fee_range},'
                f'{recon_range},"INTERNAL",{status_range},"FAILED")'
            )
        else:
            recon_criteria = sorted(recon or ())
            vendor_criteria = f',{vendor_range},"{vendor}"' if vendor else ""
            success_terms: list[str] = []
            failed_terms: list[str] = []
            for result in recon_criteria:
                success_terms.extend((
                    f'SUMIFS({amount_range},{recon_range},"{result}",'
                    f'{status_range},"SUCCESS"{vendor_criteria})',
                    f'SUMIFS({fee_range},{recon_range},"{result}",'
                    f'{status_range},"SUCCESS"{vendor_criteria})',
                ))
                failed_terms.extend((
                    f'SUMIFS({amount_range},{recon_range},"{result}",'
                    f'{status_range},"FAILED"{vendor_criteria})',
                    f'SUMIFS({fee_range},{recon_range},"{result}",'
                    f'{status_range},"FAILED"{vendor_criteria})',
                ))
            success_formula = "=" + "+".join(success_terms)
            failed_formula = "=" + "+".join(failed_terms)

        formulas = {
            2: (count_formula, cached_count),
            3: (amount_formula, cached_amount),
            4: (success_formula, success_amount),
            5: (failed_formula, failed_amount),
        }
        for column, (formula, cached) in formulas.items():
            _write_cell(
                worksheet, workbook, "Summary", write_row, column, 2,
                formula=formula, cached_value=cached,
                format_cache=format_cache,
            )


def _to_rows(frame: pl.DataFrame) -> list[dict[str, Any]]:
    return [
        {key: _clean(value) for key, value in row.items()}
        for row in frame.iter_rows(named=True)
    ]


def _lb_values(row: Mapping[str, Any]) -> dict[int, Any]:
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
        28: row.get("ref"), 29: row.get("status"), 30: row.get("nominal"),
        34: row.get("no"), 35: row.get("id_trx"),
        36: date_value,
        37: row.get("kode_produk"), 38: row.get("pelanggan_1"),
        39: row.get("pelanggan_2"), 40: row.get("pelanggan_3"),
        41: row.get("pelanggan_4"), 42: row.get("via"), 43: row.get("sn"),
        44: row.get("ref"), 45: row.get("status"), 46: row.get("nominal"),
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
            _transaction_status_sort_key(row),
            _as_datetime(row.get("tanggal")) or dt.datetime.min,
        ),
    )
    tomorrow_rows = [
        record for record in rows if _is_tomorrow_transaction(record)
    ]
    current_day_rows = [
        record for record in rows if not _is_tomorrow_transaction(record)
    ]
    title_columns = (
        (1, "Current Day from bo biller recon"),
        (21 if sheet == "LB" else 19, "Current Day from bo biller apps"),
        (40 if sheet == "LB" else 34, "For tomorrow because different time"),
    )
    values_by_row: dict[int, dict[int, Any]] = {
        1: {column: text for column, text in title_columns},
        2: {1: "SUCCESS", 3: "FAILED"},
        3: {},
    }
    normalized_statuses = [
        _transaction_status(row.get("status")) for row in current_day_rows
    ]
    current_day_end_row = max(4, len(current_day_rows) + 3)
    source_data_end_row = max(4, len(rows) + 3)
    summary_start = max_column + 2
    summary_end = summary_start + 2
    amount_column = 7
    status_column = 12
    tomorrow_amount_column = 46 if sheet == "LB" else 56
    tomorrow_status_column = 51 if sheet == "LB" else 61
    last_data_row = max(4, len(rows) + 3)
    amount_letter = xlsxwriter.utility.xl_col_to_name(amount_column - 1)
    status_letter = xlsxwriter.utility.xl_col_to_name(status_column - 1)
    tomorrow_amount_letter = xlsxwriter.utility.xl_col_to_name(
        tomorrow_amount_column - 1
    )
    tomorrow_status_letter = xlsxwriter.utility.xl_col_to_name(
        tomorrow_status_column - 1
    )
    summary_labels = ("SUCCESS", "FAILED", "TOTAL AUTO SUM")
    status_sums = {"SUCCESS": 0.0, "FAILED": 0.0}
    total_amount = 0.0
    for record in rows:
        amount = _clean(record.get(
            "tagihan" if sheet == "LB" else "nominal"
        ))
        if isinstance(amount, (int, float)):
            total_amount += float(amount)
            status = _transaction_status(record.get("status"))
            if status in status_sums:
                status_sums[status] += float(amount)

    values_by_row[1][summary_start] = "TOTAL AUTO SUM (AMOUNT)"
    values_by_row[2].update({
        summary_start + offset: label
        for offset, label in enumerate(summary_labels)
    })
    summary_formulas = {}
    for offset, label in enumerate(summary_labels):
        column = summary_start + offset
        if label == "TOTAL AUTO SUM":
            formula = (
                f"=SUM({amount_letter}4:{amount_letter}{last_data_row})"
                f"+SUM({tomorrow_amount_letter}4:"
                f"{tomorrow_amount_letter}{last_data_row})"
            )
            cached_value = total_amount
        else:
            formula = (
                f'=SUMIF(${status_letter}$4:${status_letter}${last_data_row},'
                f'"{label}",${amount_letter}$4:${amount_letter}${last_data_row})'
                f'+SUMIF(${tomorrow_status_letter}$4:'
                f'${tomorrow_status_letter}${last_data_row},"{label}",'
                f'${tomorrow_amount_letter}$4:'
                f'${tomorrow_amount_letter}${last_data_row})'
            )
            cached_value = status_sums[label]
        summary_formulas[column] = (formula, cached_value)
    summary_title_options = {
        "bold": True, "bg_color": "#D9EAF7", "font_color": "#17365D",
        "border": 1,
    }
    summary_header_options = {
        "bold": True, "bg_color": "#17365D", "font_color": "#FFFFFF",
        "border": 1, "align": "center", "valign": "vcenter",
    }
    summary_value_options = {
        "num_format": "#,##0.00", "bg_color": "#EAF1F8",
        "font_color": "#17365D", "border": 1,
    }
    worksheet.set_column(summary_start - 1, summary_end - 1, 19)
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
    formulas_by_row[3] = summary_formulas
    for section in ("A", "B", "C"):
        values_by_row[3].update(layout[section])
    _write_template_rows(
        worksheet,
        workbook,
        sheet,
        summary_end,
        values_by_row,
        format_cache,
        formulas_by_row=formulas_by_row,
        format_overrides_by_row={
            1: {summary_start: summary_title_options},
            2: {
                summary_start + offset: summary_header_options
                for offset in range(len(summary_labels))
            },
            3: {
                summary_start + offset: summary_value_options
                for offset in range(len(summary_labels))
            },
        },
    )

    tomorrow_start = 40 if sheet == "LB" else 50
    tomorrow_end = 72 if sheet == "LB" else 63

    for index in range(len(rows)):
        excel_row = index + 4
        current_day_record = (
            current_day_rows[index] if index < len(current_day_rows) else None
        )
        current_day_values = (
            values_function(current_day_record)
            if current_day_record is not None else {}
        )
        tomorrow_values = (
            values_function(tomorrow_rows[index])
            if index < len(tomorrow_rows) else {}
        )
        formulas: dict[int, tuple[str, Any]] = {}
        if sheet == "LB" and current_day_record is not None:
            formula_row = excel_row
            amount = _clean(current_day_record.get("tagihan"))
            fees = _fees(
                "lb", current_day_record,
                current_day_record.get("produk"), amount,
            )
            price = _clean(current_day_record.get("harga"))
            margin = fees[3]
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

        for column in range(1, max_column + 1):
            status_columns = (
                {12, 51} if sheet == "LB"
                else {12, 61}
            )
            if sheet == "LB":
                is_main_column = 1 <= column <= 19
                is_current_source_column = 21 <= column <= 37
                is_tomorrow_source_column = 56 <= column <= 72
            else:
                is_main_column = 1 <= column <= 16
                is_current_source_column = 18 <= column <= 30
                is_tomorrow_source_column = 34 <= column <= 46
            is_tomorrow_column = tomorrow_start <= column <= tomorrow_end
            if is_main_column or is_current_source_column:
                value = current_day_values.get(column)
                status_row = current_day_record
                has_status_row = status_row is not None
            elif is_tomorrow_column or is_tomorrow_source_column:
                value = tomorrow_values.get(column)
                tomorrow_record = (
                    tomorrow_rows[index]
                    if index < len(tomorrow_rows) else None
                )
                status_row = tomorrow_record
                has_status_row = status_row is not None
            else:
                value = None
                status_row = None
                has_status_row = False
            row_highlight = (
                _status_highlight(status_row.get("status"))
                if column in status_columns and has_status_row
                else None
            )
            formula_data = formulas.get(column)
            date_columns = (
                {1, 21, 40, 56} if sheet == "LB"
                else {1, 20, 36, 50}
            )
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
                template_row,
                value,
                formula=formula_data[0] if formula_data else None,
                cached_value=formula_data[1] if formula_data else None,
                format_cache=format_cache,
                format_overrides={
                    **(cell_overrides or {}),
                    **(row_highlight or {}),
                } or None,
            )

    worksheet.autofilter(
        2, 0, max(2, len(current_day_rows) + 2), 13
    )


def build_recon_workbook(
    lb_df: pl.DataFrame, rb_df: pl.DataFrame, output_path: Path
) -> Path:
    """Write Summary, Records, LB, and RB sheets based on the reference workbook."""
    output_path = Path(output_path).expanduser().resolve()
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Path output harus berakhiran .xlsx.")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    lb_rows = _to_rows(lb_df)
    rb_rows = _to_rows(rb_df)
    records = _records(lb_rows, rb_rows)
    workbook = xlsxwriter.Workbook(
        str(output_path),
        {"constant_memory": True, "strings_to_formulas": False},
    )
    workbook.set_properties({
        "title": "Recon Builder - LB & RB",
        "subject": "Reconciliation workbook",
        "author": "Recon Builder",
    })
    workbook.set_custom_property("Recon Builder Version", __version__)
    format_cache: dict[
        tuple[str, int, int], xlsxwriter.format.Format
    ] = {}

    ws_summary = workbook.add_worksheet("Summary")
    ws_summary.hide_gridlines(2)
    ws_summary.freeze_panes(1, 0)
    _set_widths(ws_summary, "Summary")
    _write_summary(ws_summary, workbook, records, format_cache)

    ws_records = workbook.add_worksheet("Records")
    ws_records.hide_gridlines(2)
    ws_records.freeze_panes(1, 0)
    _set_widths(ws_records, "Records")
    for column, header in enumerate(RECORDS_HEADERS, start=1):
        _write_cell(
            ws_records, workbook, "Records", 0, column, 1, header,
            format_cache=format_cache,
        )
    for index, record in enumerate(records, start=1):
        for column, key in enumerate((
            "datetime", "product", "billing_number", "billing_name", "ref",
            "vendor_ref", "amount", "total_fee", "customer_fee", "lst_fee",
            "margin_biller", "transaction_status", "vendor", "recon_result",
        ), start=1):
            format_overrides = (
                _status_highlight(record.get("transaction_status")) or {}
                if key == "transaction_status" else {}
            )
            if column == 1:
                format_overrides["num_format"] = DATETIME_NUMBER_FORMAT
            _write_cell(
                ws_records, workbook, "Records", index, column, 2,
                record.get(key), format_cache=format_cache,
                format_overrides=format_overrides or None,
            )
    ws_records.autofilter(0, 0, len(records), len(RECORDS_HEADERS) - 1)

    ws_lb = workbook.add_worksheet("LB")
    ws_lb.hide_gridlines(2)
    ws_lb.freeze_panes(3, 0)
    _set_widths(ws_lb, "LB")
    _write_wide_sheet(
        ws_lb, workbook, "LB", lb_rows, _lb_values, LB_SHEET_LAYOUT, 72,
        format_cache,
    )

    ws_rb = workbook.add_worksheet("RB")
    ws_rb.hide_gridlines(2)
    ws_rb.freeze_panes(3, 0)
    _set_widths(ws_rb, "RB")
    _write_wide_sheet(
        ws_rb, workbook, "RB", rb_rows, _rb_values, RB_SHEET_LAYOUT, 63,
        format_cache,
    )
    workbook.close()
    return output_path


__all__ = [
    "LB_COL_MAP", "RB_COL_MAP", "build_recon_workbook", "read_raw_excel",
    "read_recon_reference",
]
