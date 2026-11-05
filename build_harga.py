"""Bangun data/harga_dasar.csv dari file acuan (HSP Jakarta & Tender Andalas).

Format keluaran (long): wilayah, kategori (Upah|Bahan|Alat), uraian, satuan, harga, sumber
"""
import argparse
import csv
import re
import openpyxl


def rp(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.-]", "", str(v))
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        a, b = s.rsplit(",", 1)
        s = a.replace(".", "") + "." + b if len(b) <= 2 else s.replace(",", "")
    elif "." in s:
        a, b = s.rsplit(".", 1)
        s = s.replace(".", "") if len(b) == 3 else s
    try:
        return float(s)
    except ValueError:
        return None


def jakarta(path, out):
    wb = openpyxl.load_workbook(path, data_only=True)
    for sheet, kat in (("Tukang", "Upah"), ("Material", "Bahan"), ("Alat", "Alat")):
        for r in wb[sheet].iter_rows(min_row=2, values_only=True):
            if r[1] and r[3] is not None and rp(r[3]):
                out.append(["DKI Jakarta", kat, re.sub(r"\s+", " ", str(r[1])).strip(), str(r[2]).strip(), rp(r[3]), "HSP DKI Jakarta"])


def padang(path, out, wil="Kota Padang (Sumatera Barat)"):
    wb = openpyxl.load_workbook(path, data_only=True)
    for r in wb["UPAH"].iter_rows(min_row=13, max_col=6, values_only=True):
        if r[1] and rp(r[5]):
            out.append([wil, "Upah", str(r[1]).strip(), "OH", rp(r[5]), "Tender Andalas - UPAH"])
            if rp(r[4]):
                out.append([wil, "Upah", str(r[1]).strip(), "OJ", rp(r[4]), "Tender Andalas - UPAH"])
    for r in wb["BAHAN"].iter_rows(min_row=14, max_col=7, values_only=True):
        if r[1] and r[3] and isinstance(r[2], (int, float)) and r[2] > 0:
            out.append([wil, "Bahan", re.sub(r"\s+", " ", str(r[1])).strip(), str(r[3]).strip(), float(r[2]), r[6] or "Tender Andalas - BAHAN"])
    # kolom ALAT: B=jenis, D=kode, E=satuan, F=harga  (bergeser 1 kolom karena kolom A kosong)
    for r in wb["ALAT"].iter_rows(min_row=13, values_only=True):
        nm, sat, hrg = r[2], r[5], r[6]
        if nm and sat and rp(hrg):
            out.append([wil, "Alat", re.sub(r"\s+", " ", str(nm)).strip(), str(sat).strip(), rp(hrg), "Tender Andalas - ALAT"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--jakarta")
    ap.add_argument("--padang")
    ap.add_argument("-o", required=True)
    a = ap.parse_args()
    rows = []
    if a.jakarta:
        jakarta(a.jakarta, rows)
    if a.padang:
        padang(a.padang, rows)
    with open(a.o, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["wilayah", "kategori", "uraian", "satuan", "harga", "sumber"])
        w.writerows(rows)
    print(len(rows), "baris harga")
