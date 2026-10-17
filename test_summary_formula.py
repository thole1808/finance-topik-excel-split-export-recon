import sys, tempfile, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
import polars as pl
from src.recon_exporter import build_recon_workbook

lb = pl.DataFrame([
    {
        'tanggal': '2026-10-15 10:00:00', 'operator': 'test', 'nama': 'test',
        'idpel': '123', 'produk': 'TEST', 'periode': None, 'status': 'SUCCESS',
        'ref': 'REF-001', 'lembar': 1, 'tagihan': 100000, 'denda': 0,
        'admin': 0, 'harga': 102000, 'fee_loket': 2000, 'fee_bulan': 3000,
        'total': 102000, 'ref_id': 'VEND-001',
    },
    {
        'tanggal': '2026-10-15 11:00:00', 'operator': 'test', 'nama': 'test',
        'idpel': '124', 'produk': 'TEST', 'periode': None, 'status': 'FAILED',
        'ref': 'REF-002', 'lembar': 1, 'tagihan': 50000, 'denda': 0,
        'admin': 0, 'harga': 51000, 'fee_loket': 1000, 'fee_bulan': 2000,
        'total': 51000, 'ref_id': 'VEND-002',
    },
])
rb = pl.DataFrame([])

with tempfile.TemporaryDirectory() as td:
    output = Path(td) / 'test.xlsx'
    build_recon_workbook(lb, rb, output)

    with zipfile.ZipFile(output) as z:
        root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        print("=== Sheet Summary (sheet1) ===")
        for cell in root.findall(f'.//{ns}c'):
            addr = cell.get('r')
            if addr and addr[0] in 'ABCDE' and addr[1:].isdigit() and int(addr[1:]) >= 2:
                f_el = cell.find(f'{ns}f')
                v_el = cell.find(f'{ns}v')
                formula = f_el.text if f_el is not None else ''
                value = v_el.text if v_el is not None else ''
                print(f'{addr}: formula={formula} cached={value}')