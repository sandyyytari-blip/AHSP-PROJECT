"""Parser lampiran SE DJBK 47/2026 (hasil ekstraksi teks PDF) -> data/ahsp_db.json

Pemakaian:
    python tools/parse_ahsp.py file1.txt file2.txt -o data/ahsp_db.json
"""
import argparse
import json
import re
import sys

RE_ITEM = re.compile(r"^(\d+(?:\.\d+){2,})\s+(\S.*)$")
RE_PAGE = re.compile(r"^\s*-\s*\d+\s*-\s*$")
RE_SECTION = {"A": "tenaga", "B": "bahan", "C": "alat"}
RE_NUM = re.compile(r"^-?\d+(?:[.,]\d+)?$")
RE_HEAD = re.compile(r"^(\d+(?:\.\d+){0,2})\s+([A-Z].*)$")

SKIP_PREFIX = (
    "No Uraian", "Harga", "Satuan", "(Rp)", "Jumlah", "Harga Satuan Pekerjaan",
    "D Jumlah", "E Biaya", "F Harga",
)


def num(s):
    return float(s.replace(".", "").replace(",", ".")) if re.search(r",", s) and re.search(r"\.\d{3},", s) \
        else float(s.replace(",", "."))


UNITS = {
    "OH", "OJ", "Unit", "unit", "m'", "m’", "buah", "Buah", "m3", "m2", "kg", "Kg", "m", "M", "hari", "Hari",
    "jam", "Jam", "liter", "Lot", "lot", "lembar", "batang", "set", "Set", "daun", "Gulung", "bh", "titik",
    "ton", "Ls", "ls", "Btg", "btg", "m²", "m³", "Lembar", "Batang", "Titik", "Liter", "Bh", "Ton", "paket", "Paket",
}


def parse_line(buf):
    """buf: teks gabungan (mungkin multi-baris). Kembalikan (uraian, kode, satuan, koef) atau None."""
    toks = buf.split()
    if len(toks) < 3 or not RE_NUM.match(toks[-1]):
        return None
    satuan = toks[-2]
    if satuan not in UNITS:
        return None
    koef = num(toks[-1])
    satuan = satuan.replace("’", "'")
    rest = toks[:-2]
    kode = ""
    if rest and re.fullmatch(r"L\.\d{2}", rest[-1]):
        kode = rest.pop()
    uraian = " ".join(rest).strip()
    if not uraian:
        return None
    return uraian, kode, satuan, koef


def parse(files):
    lines = []
    for f in files:
        with open(f, encoding="utf-8", errors="ignore") as fh:
            lines.extend(l.rstrip("\n") for l in fh)

    items, heads = [], {}
    cur = None
    state = "title"  # title | table
    sect = None
    buf = ""
    divisi_stack = {}

    def flush_buf():
        nonlocal buf
        if not buf:
            return
        r = parse_line(buf)
        if r and cur is not None and sect:
            u, k, s, c = r
            cur[RE_SECTION_KEY[sect]].append({"uraian": u, "kode": k, "satuan": s, "koef": c})
        buf = ""

    RE_SECTION_KEY = {"A": "tenaga", "B": "bahan", "C": "alat"}

    for raw in lines:
        l = raw.strip()
        if not l or RE_PAGE.match(l):
            continue
        m = RE_ITEM.match(l)
        if m and (state != "table" or sect is None or l[0].isdigit()):
            # judul item baru hanya jika diikuti tabel; dibedakan lewat state
            if not (state == "table" and sect in ("A", "B", "C") and RE_NUM.match(l.split()[-1]) and len(m.group(1).split(".")) < 4):
                flush_buf()
                cur = {"kode": m.group(1), "uraian": m.group(2), "tenaga": [], "bahan": [], "alat": []}
                items.append(cur)
                state, sect = "title", None
                continue
        if state == "title":
            if l.startswith("No Uraian"):
                state, sect = "table", None
                if cur is not None:
                    cur["_t"] = True
            else:
                if cur is not None:
                    cur["uraian"] += " " + l
                else:
                    mh = RE_HEAD.match(l)
                    if mh:
                        heads[mh.group(1)] = mh.group(2)
            continue
        # state == table
        if l.startswith("F Harga Satuan Pekerjaan"):
            flush_buf()
            state, sect = "after", None
            continue
        if state == "after":
            mh = RE_HEAD.match(l)
            if mh and not RE_ITEM.match(l):
                heads[mh.group(1)] = mh.group(2)
            continue
        if re.fullmatch(r"[ABC]\s+(Tenaga Kerja|Bahan|Peralatan)\s*", l, re.I):
            flush_buf()
            sect = l[0]
            continue
        if l.lower().startswith(("jumlah harga", "d jumlah", "e biaya")):
            flush_buf()
            continue
        if l in ("Harga", "Satuan", "(Rp)", "Jumlah", "Harga Satuan", "Harga", "Satuan (Rp)", "Harga (Rp)") or l.startswith("Harga "):
            continue
        if sect is None:
            continue
        if not buf:
            mn = re.match(r"^(\d+)\s+(\S.*)$", l)
            if mn and int(mn.group(1)) == len(cur[RE_SECTION_KEY[sect]]) + 1:
                l = mn.group(2)
        buf = (buf + " " + l).strip() if buf else l
        r = parse_line(buf)
        if r and (RE_NUM.match(l.split()[-1])):
            flush_buf()

    flush_buf()
    # kamus judul bagian: hanya tingkat 1-3 yang valid (abaikan sisa baris tabel/teks yang terbungkus)
    heads2 = {}
    for n, raw in enumerate(lines):
        l = raw.strip()
        mh = RE_HEAD.match(l)
        if not mh or RE_ITEM.match(l):
            continue
        kode, txt = mh.group(1), mh.group(2)
        segs = kode.split(".")
        if not (1 <= int(segs[0]) <= 10) or len(segs) > 2:
            continue
        up = sum(c.isupper() for c in txt) / max(1, sum(c.isalpha() for c in txt))
        if len(segs) == 1 and (up < 0.8 or len(txt) < 8):
            continue
        if segs[0] != "10" and up < 0.8:
            continue
        if len(txt) < 4 or any(ch.isdigit() for ch in txt[:3]):
            continue
        if len(segs) == 2 and (segs[1].startswith("0") or not (1 <= int(segs[1]) <= 25)):
            continue
        # sambung baris lanjutan huruf kapital
        k = n + 1
        while k < len(lines) and lines[k].strip() and lines[k].strip().isupper() and not RE_HEAD.match(lines[k].strip()):
            txt += " " + lines[k].strip()
            k += 1
        heads2.setdefault(kode, re.sub(r"\s+", " ", txt).strip())
    for it in items:
        if not it.get("_t") and it["kode"].count(".") == 2:
            heads2.setdefault(it["kode"], re.sub(r"\s+", " ", it["uraian"]).strip())
    return items, heads2


_NOT_ENUM = {"inci", "inch", "mm", "cm", "m", "kg", "mm2", "buah", "unit", "btg", "batang", "liter", "ltr", "kw", "hp",
             "pk", "watt", "volt", "amp", "ton", "gr", "gram", "mtr", "meter", "lembar", "set", "bh", "pcs", "phase",
             "fasa", "way", "pole", "core", "kali", "x", "pair", "kva", "w", "a", "v", "mpa", "psi", "bar", "lumen"}


def strip_enum(u):
    m = re.match(r"^(\d{1,2})\s+([A-Za-z][A-Za-z/.-]{2,})(.*)$", u)
    if m and int(m.group(1)) <= 40 and m.group(2).lower().rstrip(".") not in _NOT_ENUM:
        return (m.group(2) + m.group(3)).strip()
    return u


def clean_uraian(u):
    u = re.sub(r"\(Lihat\s+Permen.*$", "", u).strip()
    u = re.sub(r"\s+", " ", u)
    return u


SAT_MAP = {"m3": "m3", "m³": "m3", "m2": "m2", "m²": "m2", "m'": "m'", "m’": "m'", "m": "m'", "meter": "m'",
           "kg": "kg", "ton": "ton", "buah": "buah", "bh": "buah", "unit": "unit", "set": "set", "titik": "titik",
           "lembar": "lembar", "liter": "liter", "batang": "batang", "ls": "Ls", "lot": "Ls", "paket": "paket",
           "bidang": "bidang", "stel": "set", "kali": "kali", "orang": "orang", "hari": "hari", "bulan": "bulan",
           "pasang": "pasang", "roll": "roll", "panel": "panel", "lubang": "lubang", "stasiun": "stasiun"}
RE_SAT = re.compile(r"(?<![\d.,])1(?:\s|\s*)(m3|m³|m2|m²|m['’]|m|meter|kg|ton|buah|bh|unit|set|titik|lembar|liter|batang|ls|lot|paket|bidang|stel|"
                    r"kali|orang|hari|bulan|pasang|roll|panel|lubang|stasiun)\b", re.I)


def deteksi_satuan(u):
    m = RE_SAT.search(u)
    if m:
        return SAT_MAP.get(m.group(1).lower().replace("’", "'"), m.group(1))
    m = re.search(r"(?<![\d.,])1\s+([A-Za-z][A-Za-z²³'’]*)", u)
    return m.group(1).lower() if m else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    items, heads = parse(a.files)
    out = []
    seen = set()
    for it in items:
        if not it.pop("_t", False):
            heads.setdefault(it["kode"], clean_uraian(it["uraian"]))
            continue
        if it["kode"] in seen:
            continue
        seen.add(it["kode"])
        for kk in ("tenaga", "bahan", "alat"):
            for c in it[kk]:
                c["uraian"] = strip_enum(c["uraian"])
        it["uraian_penuh"] = re.sub(r"\s+", " ", it["uraian"]).strip()
        it["uraian"] = clean_uraian(it["uraian"])
        it["satuan"] = deteksi_satuan(it["uraian"])
        out.append(it)
    json.dump({"heads": heads, "items": out}, open(a.out, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"{len(out)} item AHSP, {len(heads)} judul bagian", file=sys.stderr)


if __name__ == "__main__":
    main()
