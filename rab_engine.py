# -*- coding: utf-8 -*-
"""
rab_engine.py - Aplikasi RAB Konstruksi (AHSP SE DJBK No. 47 Tahun 2026, Bidang Cipta Karya)

Alur:
  1. Pengguna memilih Wilayah Proyek  -> harga satuan dasar upah / bahan / alat.
  2. Pengguna mengisi VOLUME pekerjaan (unggah Excel/CSV atau pilih dari katalog AHSP) -
     tanpa mengetik kode AHSP; aplikasi mencocokkan uraian pekerjaan secara otomatis.
  3. Aplikasi menghasilkan file Excel (rumus hidup) berisi sheet:
        Rekap RAB | BoQ | Kurva S | DHSP | AHSP | UPAH | BAHAN | ALAT

Menjalankan:   streamlit run rab_engine.py
Struktur data:
  data/ahsp_db.json     koefisien AHSP hasil ekstraksi lampiran SE DJBK 47/2026
  data/harga_dasar.csv  harga dasar upah/bahan/alat per wilayah (long format)
"""
from __future__ import annotations

import datetime as dt
import difflib
import io
import json
import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import pandas as pd

try:  # Streamlit hanya dibutuhkan untuk antarmuka; engine tetap bisa dipakai tanpa itu
    import streamlit as st
except Exception:  # pragma: no cover
    st = None

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data" if (BASE_DIR / "data").exists() else BASE_DIR

APP_TITLE = "Generator RAB Konstruksi - AHSP SE DJBK 47/2026"
JAM_PER_HARI = 7.0  # jam kerja efektif per hari (sesuai acuan UPAH: 7 jam/hari)
BERAT_ISI_DEFAULT = {"pasir": 1400.0, "kerikil": 1400.0, "batu": 1500.0}  # kg/m3 (asumsi, dapat diubah)

STATUS_OK = "Cocok"
STATUS_PERIKSA = "Perkiraan (periksa)"
STATUS_ACUAN = "Acuan wilayah lain (periksa)"
STATUS_MANUAL = "Input manual"
STATUS_NOL = "Belum ada harga"
STATUS_NEEDS_REVIEW = {STATUS_PERIKSA, STATUS_ACUAN, STATUS_NOL}

DIV_DEFAULT = {
    "1": "PERSIAPAN LAPANGAN/SITE WORK", "2": "PEKERJAAN STRUKTUR", "3": "PEKERJAAN ARSITEKTUR",
    "4": "PEKERJAAN LANSEKAP", "5": "PEKERJAAN MEKANIKAL DAN ELEKTRIKAL", "6": "PEKERJAAN PLAMBING",
    "7": "JALAN PADA PERMUKIMAN", "8": "PEKERJAAN SALURAN", "9": "PEKERJAAN JARINGAN PIPA DI LUAR GEDUNG",
    "10": "SISTEM STRUKTUR RISHA",
}
ROMAN = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI",
         "XVII", "XVIII", "XIX", "XX"]

# --------------------------------------------------------------------------------------
# 1. PEMUATAN DATA
# --------------------------------------------------------------------------------------


def _kode_key(kode: str):
    return tuple(int(x) if x.isdigit() else 0 for x in str(kode).split("."))


@lru_cache(maxsize=1)
def load_ahsp() -> dict:
    with open(DATA_DIR / "ahsp_db.json", encoding="utf-8") as fh:
        raw = json.load(fh)
    heads = dict(raw["heads"])
    items = {}
    for it in raw["items"]:
        items[it["kode"]] = it
    return {"heads": heads, "items": items}


def nama_divisi(kode_div: str) -> str:
    h = load_ahsp()["heads"]
    return (h.get(kode_div) or DIV_DEFAULT.get(kode_div) or f"DIVISI {kode_div}").strip()


def nama_kelompok(kode_kel: str) -> str:
    h = load_ahsp()["heads"]
    if kode_kel in h:
        return h[kode_kel].strip()
    div = kode_kel.split(".")[0]
    return f"{nama_divisi(div)} ({kode_kel})"


def info_kode(kode: str) -> dict:
    segs = kode.split(".")
    div = segs[0]
    kel = ".".join(segs[:2])
    return {"div": div, "div_nama": nama_divisi(div), "kel": kel, "kel_nama": nama_kelompok(kel)}


@lru_cache(maxsize=1)
def katalog_df() -> pd.DataFrame:
    rows = []
    for kode, it in load_ahsp()["items"].items():
        inf = info_kode(kode)
        rows.append({
            "kode": kode, "uraian": it["uraian"], "satuan": it.get("satuan", ""),
            "div": inf["div"], "div_nama": inf["div_nama"], "kel": inf["kel"], "kel_nama": inf["kel_nama"],
            "_k": _kode_key(kode),
        })
    df = pd.DataFrame(rows).sort_values("_k").reset_index(drop=True)
    return df.drop(columns="_k")


@lru_cache(maxsize=1)
def harga_dasar_df() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "harga_dasar.csv")
    df["harga"] = pd.to_numeric(df["harga"], errors="coerce")
    return df.dropna(subset=["harga"])


def daftar_wilayah() -> list[str]:
    return sorted(harga_dasar_df()["wilayah"].unique().tolist())


# --------------------------------------------------------------------------------------
# 2. TEKS, SATUAN, DAN PENCARI KECOCOKAN (tanpa dependensi eksternal)
# --------------------------------------------------------------------------------------

_SYN = [
    (r"baja tulangan polos\s*\(?bjtp\)?|\bbjtp\b", "besi beton polos"),
    (r"baja tulangan ulir\s*\(?bjts\)?|\bbjts\b|\bbjtd\b", "besi beton ulir"),
    (r"batu pecah|batu split|\bsplit\b|\bkoral\b", "kerikil"),
    (r"multiplek|multipleks|multiplex", "multipleks"),
    (r"triplek|tripleks|triplex", "triplek"),
    (r"\bdiameter\b|\bdia\b|[ø∅]", "dia"),
    (r"\binchi\b|\binch\b|\binci\b|\"|”|″", " inch "),
    (r"≥", " >= "), (r"≤", " <= "),
    (r"\bltr\b|\blt\b", "liter"),
    (r"\bzak\b|\bsak\b", "sak"),
    (r"\bpek\.", "pekerjaan"), (r"\bpas\.", "pasangan"),
    (r"\bbata\s+ringan\b|\bhebel\b", "bata ringan"),
    (r"kelas\s*(?:iii|3)\b", "kelas iii"), (r"kelas\s*(?:ii|2)\b", "kelas ii"), (r"kelas\s*(?:i|1)\b", "kelas i"),
]
_STOP = {"dan", "atau", "yang", "untuk", "dengan", "dari", "di", "ke", "sewa", "aksesoris", "aksesori", "material",
         "bantu", "pc", "termasuk", "tipe", "type", "jenis", "per", "satu", "1", "pekerjaan", "pemasangan",
         "pembuatan", "cara", "secara"}
_TOK = re.compile(r"[<>]=?|[a-z]+\d*|\d+(?:[.,]\d+)?")


def _stem(t: str) -> str:
    if len(t) <= 4 or t[0].isdigit() or t[0] in "<>":
        return t
    for p in ("penge", "menge", "meny", "peny", "peng", "meng", "pem", "mem", "pen", "men", "ber", "ter", "per",
              ):
        if t.startswith(p) and len(t) - len(p) >= 3:
            rest = t[len(p):]
            if p in ("pem", "mem") and rest[0] in "aiueo":
                rest = "p" + rest
            elif p in ("pen", "men") and rest[0] in "aiueo":
                rest = "t" + rest
            elif p in ("peng", "meng") and rest[0] in "aiueo":
                rest = "k" + rest
            elif p in ("peny", "meny") and rest[0] in "aiueo":
                rest = "s" + rest
            t = rest
            break
    for s in ("kan", "an"):
        if t.endswith(s) and len(t) - len(s) >= 3:
            t = t[: -len(s)]
            break
    return t


K_KE_MPA = {100: 8, 125: 10, 150: 12, 175: 15, 200: 17, 225: 19, 250: 21, 275: 23, 300: 25, 350: 30, 400: 35, 450: 38, 500: 42}


def _k_ke_mpa(m):
    k = int(m.group(1))
    mpa = K_KE_MPA.get(k)
    return f" k{k} fc {mpa} mpa " if mpa else m.group(0)


def tokenize(s: str, stem: bool = False) -> list[str]:
    s = str(s).lower()
    s = re.sub(r"\bk[- ]?(\d{3})\b", _k_ke_mpa, s)
    s = re.sub(r"\bf['’`´]?c\b", "fc", s).replace("’", "'").replace("²", "2").replace("³", "3")
    for pat, rep in _SYN:
        s = re.sub(pat, rep, s)
    s = re.sub(r"(\d),(\d)", r"\1.\2", s)
    out = []
    for t in _TOK.findall(s):
        if t in _STOP:
            continue
        if stem:
            t = _stem(t)
            if t in ("pasang", "buat", "kerja"):
                continue
        out.append(t)
    return out


_UNIT_ALIAS = {
    "m3": "m3", "m³": "m3", "m2": "m2", "m²": "m2", "m'": "m", "m": "m", "mtr": "m", "meter": "m", "lm": "m",
    "kg": "kg", "kilogram": "kg", "ton": "ton", "liter": "liter", "ltr": "liter", "l": "liter", "lt": "liter",
    "buah": "pcs", "bh": "pcs", "unit": "pcs", "pcs": "pcs", "biji": "pcs", "titik": "titik", "set": "set",
    "stel": "set", "lembar": "lembar", "lbr": "lembar", "batang": "batang", "btg": "batang", "hari": "hari",
    "jam": "jam", "oh": "oh", "oj": "oj", "lot": "ls", "ls": "ls", "lumpsum": "ls", "gulung": "gulung",
    "roll": "gulung", "daun": "daun", "tube": "tube", "dus": "dus", "zak": "sak", "sak": "sak", "truck": "truk",
    "truk": "truk", "pasang": "pasang", "bulan": "bulan", "ha": "ha", "kali": "kali", "orang": "orang",
}


def unit_canon(u) -> tuple[str, dict]:
    s = str(u or "").lower().strip().replace("’", "'")
    m = re.match(r"^(?:sak|zak|karung)\s*\(?\s*(\d+(?:[.,]\d+)?)\s*kg\)?$", s)
    if m:
        return "sak", {"kg": float(m.group(1).replace(",", "."))}
    m = re.match(r"^(?:batang|btg)\s*\(?\s*(\d+(?:[.,]\d+)?)\s*(?:m|meter)\)?$", s)
    if m:
        return "batang", {"m": float(m.group(1).replace(",", "."))}
    s2 = re.sub(r"[\s().]", "", s)
    return _UNIT_ALIAS.get(s2, s2), {}


def _agregat(nama: str) -> str | None:
    n = nama.lower()
    if "pasir" in n:
        return "pasir"
    if any(k in n for k in ("kerikil", "split", "batu pecah", "abu batu", "sirtu", "agregat")):
        return "kerikil"
    if "batu" in n and "bata" not in n:
        return "batu"
    return None


def faktor_konversi(nama: str, au: tuple, su: tuple, berat_isi: dict) -> float | None:
    """Faktor f sehingga harga per satuan AHSP = harga satuan sumber x f. None bila tidak dapat dikonversi."""
    a, ae = au
    s, se = su
    if a == s and (a not in ("sak", "batang") or ae == se):
        return 1.0
    if a == "kg" and s == "sak" and se.get("kg"):
        return 1.0 / se["kg"]
    if a == "sak" and s == "kg" and ae.get("kg"):
        return ae["kg"]
    if a == "kg" and s == "ton":
        return 1.0 / 1000
    if a == "ton" and s == "kg":
        return 1000.0
    if a == "liter" and s == "m3":
        return 1.0 / 1000
    if a == "m3" and s == "liter":
        return 1000.0
    if a == "jam" and s == "hari":
        return 1.0 / JAM_PER_HARI
    if a == "hari" and s == "jam":
        return JAM_PER_HARI
    if a == "oj" and s == "oh":
        return 1.0 / JAM_PER_HARI
    if a == "oh" and s == "oj":
        return JAM_PER_HARI
    if a == "m" and s == "batang" and se.get("m"):
        return 1.0 / se["m"]
    if a == "batang" and s == "m" and ae.get("m"):
        return ae["m"]
    ag = _agregat(nama)
    if ag and ag in berat_isi:
        if a == "kg" and s == "m3":
            return 1.0 / berat_isi[ag]
        if a == "m3" and s == "kg":
            return berat_isi[ag]
    return None


class Corpus:
    """Indeks token berbobot IDF untuk pencarian uraian terdekat."""

    def __init__(self, names: list[str], ctx: list[str] | None = None, stem: bool = False):
        self.stem = stem
        self.names = names
        self.toks = [set(tokenize(n, stem)) for n in names]
        self.ctx = [set(tokenize(c, stem)) for c in ctx] if ctx else None
        self.N = max(1, len(names))
        df = Counter()
        self.inv = defaultdict(list)
        for i, ts in enumerate(self.toks):
            for t in ts:
                df[t] += 1
                self.inv[t].append(i)
        self.df = df

    def w(self, t: str) -> float:
        base = math.log(1 + self.N / (1 + self.df.get(t, 0)))
        return base * (1.4 if t[0].isdigit() else 1.0)

    def search(self, q: str, k: int = 10) -> list[tuple[int, float]]:
        qt = set(tokenize(q, self.stem))
        if not qt:
            return []
        wq = sum(self.w(t) for t in qt)
        cand = set()
        rare = [t for t in qt if self.df.get(t, 0) and self.df[t] <= 0.25 * self.N]
        for t in (rare or qt):
            cand.update(self.inv.get(t, ()))
        qn = {t for t in qt if t[0].isdigit()}
        res = []
        for i in cand:
            nt = self.toks[i]
            common = qt & nt
            if not common:
                continue
            wc = sum(self.w(t) for t in common)
            wn = sum(self.w(t) for t in nt)
            dice = 2 * wc / (wq + wn)
            recall = wc / wq
            if self.ctx is not None:
                cover = sum(self.w(t) for t in qt if t in nt or t in self.ctx[i]) / wq
                sc = 0.35 * dice + 0.65 * cover
            else:
                sc = 0.5 * dice + 0.5 * recall
            cn = {t for t in nt if t[0].isdigit()}
            if qn and cn and qn != cn:
                sc *= 0.8 if (qn < cn or cn < qn) else 0.5
            elif qn and not cn:
                sc *= 0.85
            # kata (non-angka) pada uraian yang dicari harus tercakup: mencegah "Pipa Beton" ~ "Pipa Galvanis"
            qw = [t for t in qt if not t[0].isdigit() and t[0] not in "<>"]
            if len(qw) >= 2:
                wqw = sum(self.w(t) for t in qw)
                wcw = sum(self.w(t) for t in qw if t in nt or (self.ctx is not None and t in self.ctx[i]))
                sc *= (wcw / wqw) ** 1.5
            res.append((i, sc * 100))
        res.sort(key=lambda x: -x[1])
        return res[:k]


@dataclass
class HargaRes:
    harga: float = 0.0
    status: str = STATUS_NOL
    skor: float = 0.0
    src_nama: str = ""
    src_satuan: str = ""
    faktor: float = 1.0
    harga_src: float = 0.0
    wilayah: str = ""
    faktor_wil: float = 1.0


class BukuHarga:
    """Harga dasar satu wilayah (upah, bahan, alat)."""

    def __init__(self, df: pd.DataFrame, wilayah: str):
        self.wilayah = wilayah
        self.kat = {}
        for kat in ("Upah", "Bahan", "Alat"):
            sub = df[df["kategori"] == kat].reset_index(drop=True)
            ent = sub.to_dict("records")
            self.kat[kat] = (ent, Corpus([e["uraian"] for e in ent]))

    def cari(self, kat: str, nama: str, satuan: str, berat_isi: dict, kode: str = "") -> HargaRes | None:
        ent, corp = self.kat[kat]
        if not ent:
            return None
        au = unit_canon(satuan)
        best = None
        for i, sc in corp.search(nama, k=25):
            e = ent[i]
            su = unit_canon(e["satuan"])
            f = faktor_konversi(nama, au, su, berat_isi)
            if f is None:
                continue
            adj = sc + (8 if su == au else 0) - (0 if su == au else 4)
            if best is None or adj > best[0]:
                best = (adj, sc, e, f)
        if kat == "Upah":
            best = self._upah_generik(nama, au, kode, best, ent)
        if best is None:
            return None
        adj, sc, e, f = best
        return HargaRes(harga=float(e["harga"]) * f, skor=min(sc, 100.0), src_nama=e["uraian"],
                        src_satuan=e["satuan"], faktor=f, harga_src=float(e["harga"]), wilayah=self.wilayah)

    def _upah_generik(self, nama, au, kode, best, ent):
        """Tenaga: bila nama khusus (mis. Tukang Besi) tidak ada, pakai kelompok generik menurut kode L.0x."""
        if best is not None and best[1] >= 85:
            return best
        n = nama.lower().strip()
        gen = None
        if n.startswith("tukang"):
            gen = "tukang"
        elif n.startswith("kepala"):
            gen = "kepala tukang"
        elif n.startswith("mandor"):
            gen = "mandor"
        elif n.startswith("pekerja"):
            gen = "pekerja"
        else:
            gen = {"L.01": "pekerja", "L.02": "tukang", "L.03": "kepala tukang", "L.04": "mandor"}.get(kode)
        if gen:
            for e in ent:
                if e["uraian"].lower().strip() == gen:
                    su = unit_canon(e["satuan"])
                    f = faktor_konversi(nama, au, su, {})
                    if f is not None and (su == au):
                        return (60.0, 60.0, e, f)
        return best


def _buku_dari_wilayah(wilayah: str) -> BukuHarga:
    df = harga_dasar_df()
    return BukuHarga(df[df["wilayah"] == wilayah], wilayah)


@lru_cache(maxsize=8)
def buku_harga(wilayah: str) -> BukuHarga:
    return _buku_dari_wilayah(wilayah)


THR_OK = 78.0
THR_LOW = 52.0


def resolve_harga(kat: str, nama: str, satuan: str, kode: str, buku: BukuHarga, buku_ref: BukuHarga | None,
                  berat_isi: dict, faktor_ref: float = 1.0) -> HargaRes:
    r = buku.cari(kat, nama, satuan, berat_isi, kode)
    if r and r.skor >= THR_OK:
        r.status = STATUS_OK
        return r
    ref = buku_ref.cari(kat, nama, satuan, berat_isi, kode) if buku_ref else None
    if ref and ref.skor >= THR_LOW and (r is None or r.skor < THR_LOW or ref.skor > r.skor + 10):
        ref.status = STATUS_ACUAN
        ref.harga *= faktor_ref
        ref.faktor_wil = faktor_ref
        ref.wilayah = f"{ref.wilayah} (acuan)"
        return ref
    if r and r.skor >= THR_LOW:
        r.status = STATUS_PERIKSA
        return r
    return HargaRes(status=STATUS_NOL)


# --------------------------------------------------------------------------------------
# 3. MODEL & PERHITUNGAN AHSP
# --------------------------------------------------------------------------------------

KAT_AHSP = (("tenaga", "Upah", "A", "Tenaga Kerja"), ("bahan", "Bahan", "B", "Bahan"), ("alat", "Alat", "C", "Peralatan"))


def komp_key(uraian: str, satuan: str) -> str:
    return f"{uraian}|{satuan}"


@dataclass
class Params:
    wilayah: str
    wilayah_acuan: str = "DKI Jakarta"
    pakai_acuan: bool = True
    faktor_acuan: float = 1.0
    overhead_profit: float = 0.15
    ppn: float = 0.12
    pembulatan_digit: int = -3
    berat_isi: dict = field(default_factory=lambda: dict(BERAT_ISI_DEFAULT))
    harga_manual: dict = field(default_factory=dict)  # {(kategori, key): harga}
    harga_wilayah_df: pd.DataFrame | None = None  # harga unggahan pengguna (opsional)
    info: dict = field(default_factory=dict)  # program, kegiatan, pekerjaan, lokasi, tahun
    durasi_minggu: int = 24
    tgl_mulai: dt.date = field(default_factory=dt.date.today)
    jadwal: dict = field(default_factory=dict)  # {kel: (mulai, selesai)}
    penandatangan: dict = field(default_factory=dict)


@dataclass
class Baris:  # satu baris BoQ
    kode: str
    uraian: str
    satuan: str
    volume: float
    harga_manual: float | None = None
    skor: float = 100.0


@dataclass
class HasilAHSP:
    kode: str
    uraian: str
    satuan: str
    rows: dict  # {"tenaga": [...], "bahan": [...], "alat": [...]}; tiap row: dict(uraian,kode,satuan,koef,harga,jumlah,key)
    A: float
    B: float
    C: float
    D: float
    E: float
    F: float


def _siapkan_buku(p: Params):
    if p.harga_wilayah_df is not None and p.wilayah == "Unggahan Anda":
        buku = BukuHarga(p.harga_wilayah_df, "Unggahan Anda")
    else:
        buku = buku_harga(p.wilayah)
    ref = None
    if p.pakai_acuan and p.wilayah_acuan and p.wilayah_acuan != p.wilayah:
        ref = buku_harga(p.wilayah_acuan)
    return buku, ref


def hitung_rab(baris: list[Baris], p: Params) -> dict:
    """Hitung seluruh RAB (Python) - dipakai untuk pratinjau, nilai-cache Excel, dan terbilang."""
    ahsp = load_ahsp()["items"]
    buku, ref = _siapkan_buku(p)
    harga_cache: dict[tuple, HargaRes] = {}
    komponen: dict[tuple, dict] = {}  # (kat, key) -> info
    hasil_ahsp: dict[str, HasilAHSP] = {}
    items = []
    dilewati = []
    n_manual = 0

    for b in baris:
        if b.harga_manual is not None and (b.kode not in ahsp):
            n_manual += 1
            items.append({"b": b, "ahsp": None, "hs": float(b.harga_manual), "kode_out": f"M-{n_manual:03d}"})
            continue
        if b.kode not in ahsp:
            dilewati.append(b)
            continue
        if b.kode not in hasil_ahsp:
            src = ahsp[b.kode]
            rows = {}
            tot = {}
            for kk, kat, _, _ in KAT_AHSP:
                rr = []
                for c in src.get(kk, []):
                    key = komp_key(c["uraian"], c["satuan"])
                    ck = (kat, key)
                    if ck not in harga_cache:
                        if ck in p.harga_manual:
                            hv = float(p.harga_manual[ck])
                            harga_cache[ck] = HargaRes(harga=hv, status=STATUS_MANUAL, skor=100, wilayah="Manual")
                        else:
                            harga_cache[ck] = resolve_harga(kat, c["uraian"], c["satuan"], c.get("kode", ""), buku,
                                                            ref, p.berat_isi, p.faktor_acuan)
                    hr = harga_cache[ck]
                    komponen.setdefault(ck, {"kat": kat, "uraian": c["uraian"], "satuan": c["satuan"],
                                             "kode": c.get("kode", ""), "res": hr})
                    rr.append({"uraian": c["uraian"], "kode": c.get("kode", ""), "satuan": c["satuan"],
                               "koef": c["koef"], "harga": hr.harga, "jumlah": c["koef"] * hr.harga, "key": key})
                rows[kk] = rr
                tot[kk] = sum(x["jumlah"] for x in rr)
            D = tot["tenaga"] + tot["bahan"] + tot["alat"]
            E = p.overhead_profit * D
            hasil_ahsp[b.kode] = HasilAHSP(b.kode, src["uraian"], src.get("satuan", ""), rows, tot["tenaga"],
                                           tot["bahan"], tot["alat"], D, E, D + E)
        items.append({"b": b, "ahsp": hasil_ahsp[b.kode], "hs": hasil_ahsp[b.kode].F, "kode_out": b.kode})

    # kelompokkan: divisi -> kelompok -> item (urutan kode, urutan input dipertahankan di dalam kelompok)
    grup: dict[str, dict[str, list]] = defaultdict(lambda: defaultdict(list))
    for it in items:
        k = it["b"].kode
        inf = info_kode(k) if (k in ahsp and it["ahsp"] is not None) else {"div": "99", "kel": "99.1", "div_nama": "PEKERJAAN LAIN-LAIN",
                                              "kel_nama": "Pekerjaan Lain-lain (Harga Manual)"}
        it["inf"] = inf
        it["jumlah"] = it["b"].volume * it["hs"]
        grup[inf["div"]][inf["kel"]].append(it)
    divisi = []
    total = 0.0
    for dk in sorted(grup, key=_kode_key):
        kels = []
        sd = 0.0
        for kk in sorted(grup[dk], key=_kode_key):
            its = grup[dk][kk]
            sk = sum(i["jumlah"] for i in its)
            kels.append({"kode": kk, "nama": its[0]["inf"]["kel_nama"], "items": its, "subtotal": sk})
            sd += sk
        divisi.append({"kode": dk, "nama": grup[dk][sorted(grup[dk], key=_kode_key)[0]][0]["inf"]["div_nama"],
                       "kelompok": kels, "subtotal": sd})
        total += sd
    ppn = p.ppn * total
    grand = total + ppn
    dibulatkan = _roundup(grand, p.pembulatan_digit)
    return {"divisi": divisi, "ahsp": hasil_ahsp, "komponen": komponen, "jumlah": total, "ppn": ppn,
            "total": grand, "dibulatkan": dibulatkan, "n_item": len(items), "dilewati": dilewati}


def _roundup(x: float, digits: int) -> float:
    f = 10 ** (-digits)
    return math.ceil(round(x / f, 9)) * f if digits < 0 else math.ceil(round(x * 10 ** digits, 9)) / 10 ** digits


# --------------------------------------------------------------------------------------
# 4. UTILITAS: TERBILANG, PENCOCOKAN URAIAN -> KODE AHSP, JADWAL OTOMATIS
# --------------------------------------------------------------------------------------

_SATUAN = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"]


def _terbilang(n: int) -> str:
    if n < 12:
        return _SATUAN[n]
    if n < 20:
        return _terbilang(n - 10) + " belas"
    if n < 100:
        return _terbilang(n // 10) + " puluh" + (" " + _terbilang(n % 10) if n % 10 else "")
    if n < 200:
        return "seratus" + (" " + _terbilang(n - 100) if n > 100 else "")
    if n < 1000:
        return _terbilang(n // 100) + " ratus" + (" " + _terbilang(n % 100) if n % 100 else "")
    if n < 2000:
        return "seribu" + (" " + _terbilang(n - 1000) if n > 1000 else "")
    if n < 1_000_000:
        return _terbilang(n // 1000) + " ribu" + (" " + _terbilang(n % 1000) if n % 1000 else "")
    for div, nm in ((10 ** 12, "triliun"), (10 ** 9, "miliar"), (10 ** 6, "juta")):
        if n >= div:
            return _terbilang(n // div) + " " + nm + (" " + _terbilang(n % div) if n % div else "")
    return ""


def terbilang_rupiah(x: float) -> str:
    n = int(round(x))
    return "Terbilang : " + (_terbilang(n).title() if n else "Nol") + " Rupiah"


@lru_cache(maxsize=1)
def _corpus_katalog() -> tuple[Corpus, pd.DataFrame]:
    df = katalog_df()
    return Corpus(df["uraian"].tolist(), ctx=(df["kel_nama"] + " " + df["div_nama"]).tolist(), stem=True), df


def cocokkan_uraian(uraian: str, satuan: str = "", k: int = 6) -> list[dict]:
    """Cari kandidat AHSP untuk uraian bebas pengguna. Mengembalikan daftar {kode, uraian, satuan, skor}."""
    corp, df = _corpus_katalog()
    us = unit_canon(satuan)[0] if satuan else None
    hasil = []
    for i, sc in corp.search(uraian, k=40):
        row = df.iloc[i]
        if us and row["satuan"]:
            sc += 10 if unit_canon(row["satuan"])[0] == us else -12
        hasil.append({"kode": row["kode"], "uraian": row["uraian"], "satuan": row["satuan"], "skor": max(0.0, min(sc, 100.0))})
    hasil.sort(key=lambda x: -x["skor"])
    return hasil[:k]


# jendela waktu awal per divisi (fraksi durasi proyek) - hanya usulan, dapat diubah pengguna
_JENDELA_DIV = {"1": (0.00, 0.20), "2": (0.10, 0.60), "3": (0.45, 0.92), "4": (0.82, 1.00), "5": (0.40, 0.92),
                "6": (0.40, 0.92), "7": (0.60, 1.00), "8": (0.55, 0.95), "9": (0.55, 0.95), "10": (0.10, 0.80),
                "99": (0.80, 1.00)}


def cari_katalog(q: str, maks: int = 60) -> pd.DataFrame:
    """Pencarian bebas pada seluruh katalog (semua kata harus ada; K-xxx diterjemahkan ke f'c)."""
    df = katalog_df()
    q = re.sub(r"\bk[- ]?(\d{3})\b", lambda m: f"{m.group(0)} " + (f"{K_KE_MPA[int(m.group(1))]} mpa" if int(m.group(1)) in K_KE_MPA else ""), q.lower())
    kata = [t for t in re.split(r"\s+", q.strip()) if t]
    m = pd.Series(True, index=df.index)
    txt = df["uraian"].str.lower().str.replace(r"f['’`´]c\b", "fc", regex=True).str.replace("’", "'", regex=False)
    for t in kata:
        if re.fullmatch(r"k[- ]?\d{3}", t):
            continue
        m &= txt.str.contains(re.escape(t.replace("’", "'")), regex=True)
    return df[m].head(maks)


def jadwal_otomatis(kelompok: list[tuple[str, float]], minggu: int) -> dict:
    """Usulan jadwal awal {kode_kelompok: (mulai, selesai)} berdasarkan jendela waktu tiap divisi.
    Kelompok dalam satu divisi disusun berurutan dengan tumpang-tindih 30%."""
    minggu = max(1, int(minggu))
    per_div: dict[str, list[str]] = defaultdict(list)
    for kd, _ in kelompok:
        per_div[kd.split(".")[0]].append(kd)
    hasil = {}
    ov = 0.30
    for dv, kels in per_div.items():
        f0, f1 = _JENDELA_DIV.get(dv, (0.0, 1.0))
        a = f0 * minggu
        bnd = max(f1 * minggu, a + 1)
        n = len(kels)
        span = bnd - a
        L = span / (n * (1 - ov) + ov)
        for i, kd in enumerate(sorted(kels, key=_kode_key)):
            st = a + i * L * (1 - ov)
            s_ = int(round(st)) + 1
            e_ = int(round(st + L))
            s_ = min(max(1, s_), minggu)
            e_ = min(max(s_, e_), minggu)
            hasil[kd] = (s_, e_)
    return hasil


# --------------------------------------------------------------------------------------
# 5. PEMBUAT FILE EXCEL (XlsxWriter - rumus hidup + nilai cache)
# --------------------------------------------------------------------------------------
import xlsxwriter  # noqa: E402
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell  # noqa: E402

S_REKAP, S_BOQ, S_KS, S_DHSP, S_AHSP = "Rekap RAB", "BoQ", "Kurva S", "DHSP", "AHSP"
S_UPAH, S_BAHAN, S_ALAT, S_CAT = "UPAH", "BAHAN", "ALAT", "Catatan"
SHEET_OF_KAT = {"Upah": S_UPAH, "Bahan": S_BAHAN, "Alat": S_ALAT}
JUDUL_KAT = {"Upah": "DAFTAR HARGA SATUAN DASAR UPAH", "Bahan": "DAFTAR HARGA SATUAN DASAR BAHAN",
             "Alat": "DAFTAR HARGA SATUAN DASAR ALAT"}


def _q(sheet: str) -> str:
    return f"'{sheet}'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet


def _san(t: str, maxlen: int = 200) -> str:
    """Bersihkan teks agar aman sebagai kunci MATCH (tanpa wildcard * ? ~) dan tidak terlalu panjang."""
    t = re.sub(r"[*?~]", "", str(t))
    t = re.sub(r"\s+", " ", t).strip()
    return t[:maxlen]


def _kunci(uraian: str, satuan: str) -> str:
    return f"{_san(uraian)}|{_san(satuan, 30)}"


def _fmt(wb, **kw):
    base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
    base.update(kw)
    return wb.add_format(base)


class _F:
    """Kumpulan format."""

    def __init__(self, wb):
        b = {"border": 1}
        self.title = _fmt(wb, bold=True, font_size=14, align="center")
        self.sub = _fmt(wb, bold=True, font_size=11)
        self.txt = _fmt(wb)
        self.txt_b = _fmt(wb, bold=True)
        self.txt_i = _fmt(wb, italic=True, font_color="#595959")
        self.wrap = _fmt(wb, text_wrap=True, valign="top")
        self.head = _fmt(wb, bold=True, align="center", text_wrap=True, bg_color="#D9E1F2", **b)
        self.num_hdr = _fmt(wb, italic=True, align="center", font_size=8, **b)
        self.cell = _fmt(wb, **b)
        self.cell_wrap = _fmt(wb, text_wrap=True, valign="top", **b)
        self.cell_c = _fmt(wb, align="center", **b)
        self.cell_b = _fmt(wb, bold=True, **b)
        self.cell_bc = _fmt(wb, bold=True, align="center", **b)
        self.money = _fmt(wb, num_format="#,##0.00", **b)
        self.money_b = _fmt(wb, num_format="#,##0.00", bold=True, **b)
        self.money_tot = _fmt(wb, num_format="#,##0.00", bold=True, bg_color="#F2F2F2", **b)
        self.money_in = _fmt(wb, num_format="#,##0.00", font_color="#0000FF", **b)
        self.int0 = _fmt(wb, num_format="#,##0", **b)
        self.int0_b = _fmt(wb, num_format="#,##0", bold=True, **b)
        self.coef = _fmt(wb, num_format="0.00####", **b)
        self.vol_in = _fmt(wb, num_format="#,##0.00##", font_color="#0000FF", **b)
        self.pct = _fmt(wb, num_format="0.00%", **b)
        self.pct_in = _fmt(wb, num_format="0.00%", font_color="#0000FF", bold=True, **b)
        self.pct2 = _fmt(wb, num_format="0.00", **b)
        self.pct2_b = _fmt(wb, num_format="0.00", bold=True, bg_color="#F2F2F2", **b)
        self.pct2_tot = _fmt(wb, num_format="0.00", bold=True, bg_color="#F2F2F2", **b)
        self.int_in = _fmt(wb, num_format="0", font_color="#0000FF", align="center", **b)
        self.date_in = _fmt(wb, num_format="dd-mmm-yyyy", font_color="#0000FF", bold=True, align="left", **b)
        self.div = _fmt(wb, bold=True, bg_color="#BDD7EE", **b)
        self.kel = _fmt(wb, bold=True, bg_color="#DDEBF7", **b)
        self.lbl_r = _fmt(wb, bold=True, align="right", **b)
        self.lbl_tot = _fmt(wb, bold=True, align="right", bg_color="#F2F2F2", **b)
        self.week = _fmt(wb, num_format="0.00;-0.00;", font_size=8, **b)
        self.week_h = _fmt(wb, bold=True, align="center", bg_color="#D9E1F2", font_size=8, **b)
        self.week_d = _fmt(wb, align="center", font_size=7, text_wrap=True, **b)
        self.stat = {
            STATUS_OK: _fmt(wb, **b), STATUS_PERIKSA: _fmt(wb, bg_color="#FFF2CC", **b),
            STATUS_ACUAN: _fmt(wb, bg_color="#FCE4D6", **b), STATUS_NOL: _fmt(wb, bg_color="#F8CBAD", bold=True, **b),
            STATUS_MANUAL: _fmt(wb, bg_color="#DDEBF7", **b),
        }
        self.key = _fmt(wb, font_color="#A6A6A6", font_size=8, **b)
        self.warn = _fmt(wb, bold=True, font_color="#C00000", text_wrap=True)


def _info_block(ws, F, info: dict, r0: int, c_lbl: int, c_val: int = None):
    rows = [("Program", info.get("program", "")), ("Kegiatan", info.get("kegiatan", "")),
            ("Pekerjaan", info.get("pekerjaan", "")), ("Lokasi", info.get("lokasi", "")),
            ("Tahun Anggaran", info.get("tahun", ""))]
    for i, (a, b) in enumerate(rows):
        ws.write_string(r0 + i, c_lbl, f"{a} : {b}", F.txt)


def buat_excel(baris: list[Baris], p: Params, hasil: dict | None = None) -> bytes:
    """Bangun workbook RAB lengkap dan kembalikan sebagai bytes (.xlsx)."""
    hasil = hasil or hitung_rab(baris, p)
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    wb.set_calc_mode("auto")
    F = _F(wb)
    ws = {n: wb.add_worksheet(n) for n in (S_REKAP, S_BOQ, S_KS, S_DHSP, S_AHSP, S_UPAH, S_BAHAN, S_ALAT, S_CAT)}
    for w in ws.values():
        w.hide_gridlines(2)
        w.set_landscape()
        w.set_paper(9)
        w.fit_to_pages(1, 0)
    info = p.info

    # ---------------- 5a. sheet harga dasar (UPAH / BAHAN / ALAT) ----------------
    kat_rows: dict[str, dict[str, int]] = {}  # kat -> key -> row(0-based)
    kat_rng: dict[str, tuple[int, int]] = {}
    HR0 = 12  # baris data pertama (0-based) = baris 13 di Excel
    for kat in ("Upah", "Bahan", "Alat"):
        w = ws[SHEET_OF_KAT[kat]]
        w.write_string(1, 0, JUDUL_KAT[kat], F.sub)
        _info_block(w, F, info, 3, 0, 1)
        w.write_string(9, 0, "Wilayah harga: " + p.wilayah, F.txt_b)
        hdr = ["No", "Uraian (sesuai AHSP)", "Satuan", "Harga Sumber (Rp)", "Satuan Sumber", "Faktor Konversi Satuan",
               "Faktor Penyesuaian Wilayah", "Harga Satuan (Rp)", "Uraian pada Daftar Harga Sumber", "Sumber Wilayah",
               "Status", "Kunci Pencarian"]
        for c, h in enumerate(hdr):
            w.write_string(10, c, h, F.head)
        w.set_row(10, 42)
        comps = sorted([v for k, v in hasil["komponen"].items() if k[0] == kat],
                       key=lambda v: (v["uraian"].lower(), v["satuan"]))
        seen = {}
        for v in comps:
            seen.setdefault(_kunci(v["uraian"], v["satuan"]), v)
        rows = {}
        for i, (key, v) in enumerate(seen.items()):
            r = HR0 + i
            rows[key] = r
            hr: HargaRes = v["res"]
            xl = r + 1
            w.write_number(r, 0, i + 1, F.cell_c)
            w.write_string(r, 1, _san(v["uraian"]), F.cell_wrap)
            w.write_string(r, 2, _san(v["satuan"], 30), F.cell_c)
            manual = hr.status == STATUS_MANUAL
            hs = hr.harga if manual else hr.harga_src
            if hr.status == STATUS_NOL:
                hs = 0.0
            w.write_number(r, 3, hs, F.money_in)
            w.write_string(r, 4, _san(hr.src_satuan or v["satuan"], 30), F.cell_c)
            w.write_number(r, 5, 1.0 if (manual or hr.status == STATUS_NOL) else hr.faktor, _fmt(wb, num_format="0.######", border=1, font_color="#0000FF"))
            w.write_number(r, 6, hr.faktor_wil, _fmt(wb, num_format="0.####", border=1, font_color="#0000FF"))
            w.write_formula(r, 7, f"=D{xl}*F{xl}*G{xl}", F.money_b, hr.harga)
            w.write_string(r, 8, hr.src_nama or ("(input manual)" if manual else "-"), F.cell_wrap)
            w.write_string(r, 9, hr.wilayah or "-", F.cell)
            w.write_string(r, 10, hr.status, F.stat.get(hr.status, F.cell))
            w.write_formula(r, 11, f'=B{xl}&"|"&C{xl}', F.key, key)
        last = HR0 + max(len(seen), 1) - 1
        kat_rows[kat], kat_rng[kat] = rows, (HR0, last)
        for c, wd in enumerate([5, 48, 10, 16, 14, 14, 14, 17, 48, 24, 22, 30]):
            w.set_column(c, c, wd)
        w.freeze_panes(11, 2)
        w.autofilter(10, 0, max(last, 11), 11)
        w.write_string(last + 2, 1, "Sel biru = input yang dapat diubah; Harga Satuan menghitung ulang seluruh AHSP, DHSP, BoQ, Rekap dan Kurva S.", F.txt_i)
        w.write_string(last + 3, 1, "Kolom Status berwarna kuning/oranye/merah = harga perlu diperiksa atau belum tersedia - isi kolom Harga Sumber.", F.txt_i)

    def look(kat: str, r_xl: int) -> str:
        a, b = kat_rng[kat]
        sh = SHEET_OF_KAT[kat]
        return (f"=IFERROR(INDEX({sh}!$H${a + 1}:$H${b + 1},MATCH($C{r_xl}&\"|\"&$E{r_xl},"
                f"{sh}!$L${a + 1}:$L${b + 1},0)),0)")

    # ---------------- 5b. AHSP ----------------
    w = ws[S_AHSP]
    for c, wd in enumerate([13, 7, 52, 9, 10, 12, 17, 18]):
        w.set_column(c, c, wd)
    w.write_string(1, 1, "ANALISA HARGA SATUAN PEKERJAAN (AHSP)", F.title)
    w.merge_range(1, 1, 1, 7, "ANALISA HARGA SATUAN PEKERJAAN (AHSP)", F.title)
    _info_block(w, F, info, 3, 1, 2)
    w.merge_range(8, 1, 8, 2, "Biaya Umum + Keuntungan (input, SE: 10% - 15% x D)  >>", _fmt(wb, bold=True, align="right"))
    w.write_number(8, 3, p.overhead_profit, F.pct_in)
    w.write_string(8, 4, "berlaku untuk seluruh AHSP", F.txt_i)
    OHP = "$D$9"
    ahsp_end_row = {}  # kode -> baris Excel (1-based) 'F' harga satuan
    r = 10
    last_div = last_kel = None
    for dv in hasil["divisi"]:
        for kl in dv["kelompok"]:
            for it in kl["items"]:
                a: HasilAHSP | None = it["ahsp"]
                if a is None or it["kode_out"] in ahsp_end_row:
                    continue
                if dv["kode"] != last_div:
                    w.write_string(r, 1, f"DIVISI {dv['kode']}", F.div)
                    w.merge_range(r, 2, r, 7, dv["nama"], F.div)
                    r += 1
                    last_div = dv["kode"]
                if kl["kode"] != last_kel:
                    w.write_string(r, 1, kl["kode"], F.kel)
                    w.merge_range(r, 2, r, 7, kl["nama"], F.kel)
                    r += 1
                    last_kel = kl["kode"]
                w.write_string(r, 1, a.kode, F.cell_b)
                w.merge_range(r, 2, r, 7, a.uraian, _fmt(wb, bold=True, text_wrap=True, border=1, valign="top"))
                w.set_row(r, 15 * max(1, math.ceil(len(a.uraian) / 85)))
                r += 1
                for c, h in enumerate(["No", "Uraian", "Kode", "Satuan", "Koefisien", "Harga Satuan\n(Rp)", "Jumlah\n(Rp)"]):
                    w.write_string(r, 1 + c, h, F.head)
                w.set_row(r, 28)
                r += 1
                for c in range(7):
                    w.write_number(r, 1 + c, c + 1, F.num_hdr)
                r += 1
                sub_rows = {}
                for kk, kat, huruf, label in KAT_AHSP:
                    w.write_string(r, 1, huruf, F.cell_bc)
                    w.write_string(r, 2, label, F.cell_b)
                    for c in range(3, 8):
                        w.write_blank(r, c, None, F.cell)
                    r += 1
                    first = r
                    for n, c_ in enumerate(a.rows[kk], 1):
                        xl = r + 1
                        w.write_number(r, 1, n, F.cell_c)
                        w.write_string(r, 2, _san(c_["uraian"]), F.cell_wrap)
                        w.write_string(r, 3, c_["kode"], F.cell_c)
                        w.write_string(r, 4, _san(c_["satuan"], 30), F.cell_c)
                        w.write_number(r, 5, c_["koef"], F.coef)
                        w.write_formula(r, 6, look(kat, xl), F.money, c_["harga"])
                        w.write_formula(r, 7, f"=F{xl}*G{xl}", F.money, c_["jumlah"])
                        r += 1
                    last = r - 1
                    for cc in (1, 2, 3):
                        w.write_blank(r, cc, None, F.cell)
                    w.merge_range(r, 4, r, 6, f"Jumlah Harga {label}", F.lbl_r)
                    tot = {"tenaga": a.A, "bahan": a.B, "alat": a.C}[kk]
                    if a.rows[kk]:
                        w.write_formula(r, 7, f"=SUM(H{first + 1}:H{last + 1})", F.money_b, tot)
                    else:
                        w.write_number(r, 7, 0, F.money_b)
                    sub_rows[kk] = r + 1
                    r += 1
                xl = r + 1
                w.write_string(r, 1, "D", F.cell_bc)
                w.merge_range(r, 2, r, 6, "Jumlah Harga Tenaga, Bahan dan Peralatan (A+B+C)", F.cell_b)
                w.write_formula(r, 7, f"=H{sub_rows['tenaga']}+H{sub_rows['bahan']}+H{sub_rows['alat']}", F.money_b, a.D)
                rD = xl
                r += 1
                xl = r + 1
                w.write_string(r, 1, "E", F.cell_bc)
                w.merge_range(r, 2, r, 4, "Biaya Umum dan Keuntungan", F.cell_b)
                w.write_formula(r, 5, f'=ROUND({OHP}*100,2)&"% x D"', F.cell_c, f"{round(p.overhead_profit * 100, 2):g}% x D")
                w.write_blank(r, 6, None, F.cell)
                w.write_formula(r, 7, f"={OHP}*H{rD}", F.money_b, a.E)
                r += 1
                xl = r + 1
                w.write_string(r, 0, a.kode, F.key)
                w.write_string(r, 1, "F", F.cell_bc)
                w.merge_range(r, 2, r, 6, "Harga Satuan Pekerjaan (D+E)", F.cell_b)
                w.write_formula(r, 7, f"=H{rD}+H{rD + 1}", F.money_tot, a.F)
                ahsp_end_row[a.kode] = xl
                r += 2
    ahsp_last = max(r, 12)
    w.freeze_panes(9, 0)
    w.repeat_rows(1, 8)

    # ---------------- 5c. DHSP ----------------
    w = ws[S_DHSP]
    for c, wd in enumerate([2, 6, 14, 70, 10, 20]):
        w.set_column(c, c, wd)
    w.merge_range(1, 1, 1, 5, "DAFTAR HARGA SATUAN PEKERJAAN (DHSP)", F.title)
    _info_block(w, F, info, 3, 1, 2)
    for c, h in enumerate(["No", "Kode", "URAIAN PEKERJAAN", "Satuan", "Harga Satuan (Rp)"]):
        w.write_string(9, 1 + c, h, F.head)
    for c in range(5):
        w.write_number(10, 1 + c - 0, c + 1, F.num_hdr) if c < 5 else None
    r = 11
    dhsp_first = r + 1
    n = 0
    dhsp_done = set()
    last_div = last_kel = None
    for dv in hasil["divisi"]:
        for kl in dv["kelompok"]:
            for it in kl["items"]:
                kode = it["kode_out"]
                if kode in dhsp_done:
                    continue
                if dv["kode"] != last_div:
                    w.write_string(r, 1, "", F.div)
                    w.write_string(r, 2, dv["kode"], F.div)
                    w.write_string(r, 3, dv["nama"], F.div)
                    w.write_blank(r, 4, None, F.div)
                    w.write_blank(r, 5, None, F.div)
                    r += 1
                    last_div = dv["kode"]
                if kl["kode"] != last_kel:
                    w.write_string(r, 1, "", F.kel)
                    w.write_string(r, 2, kl["kode"], F.kel)
                    w.write_string(r, 3, kl["nama"], F.kel)
                    w.write_blank(r, 4, None, F.kel)
                    w.write_blank(r, 5, None, F.kel)
                    r += 1
                    last_kel = kl["kode"]
                dhsp_done.add(kode)
                n += 1
                xl = r + 1
                w.write_number(r, 1, n, F.cell_c)
                w.write_string(r, 2, kode, F.cell_c)
                a = it["ahsp"]
                uraian = a.uraian if a else it["b"].uraian
                satuan = (a.satuan if a else "") or it["b"].satuan
                w.write_string(r, 3, uraian, F.cell_wrap)
                w.write_string(r, 4, satuan, F.cell_c)
                if a is not None:
                    w.write_formula(r, 5, f"=VLOOKUP(C{xl},{S_AHSP}!$A$1:$H${ahsp_last},8,FALSE)", F.money, it["hs"])
                else:
                    w.write_number(r, 5, it["hs"], F.money_in)
                r += 1
    dhsp_last = max(r, dhsp_first + 1)
    w.freeze_panes(11, 0)
    w.repeat_rows(9, 10)

    # ---------------- 5d. BoQ ----------------
    w = ws[S_BOQ]
    for c, wd in enumerate([2, 6, 13, 66, 9, 12, 19, 21]):
        w.set_column(c, c, wd)
    w.merge_range(1, 1, 1, 7, "BILL OF QUANTITY (BOQ)", F.title)
    _info_block(w, F, info, 3, 2, 3)
    for c, h in enumerate(["NO", "KODE", "ITEM PEKERJAAN", "Satuan", "Volume", "Harga Satuan Pekerjaan", "Jumlah Harga (Rp)"]):
        w.write_string(8, 1 + c, h, F.head)
    w.set_row(8, 30)
    r = 9
    div_tot_row, kel_sub_row = {}, {}
    for di, dv in enumerate(hasil["divisi"], 1):
        w.write_string(r, 1, ROMAN[di] if di < len(ROMAN) else str(di), F.div)
        w.write_blank(r, 2, None, F.div)
        w.write_string(r, 3, dv["nama"], F.div)
        for c in range(4, 8):
            w.write_blank(r, c, None, F.div)
        r += 1
        subs = []
        for ki, kl in enumerate(dv["kelompok"], 1):
            w.write_number(r, 1, ki, F.kel)
            w.write_blank(r, 2, None, F.kel)
            w.write_string(r, 3, kl["nama"], F.kel)
            for c in range(4, 8):
                w.write_blank(r, c, None, F.kel)
            r += 1
            first = r + 1
            for it in kl["items"]:
                b: Baris = it["b"]
                a = it["ahsp"]
                xl = r + 1
                w.write_string(r, 1, "•", F.cell_c)
                w.write_string(r, 2, it["kode_out"], F.cell_c)
                uraian = b.uraian.strip() if b.uraian and b.uraian.strip() else (a.uraian if a else "")
                satuan = b.satuan or (a.satuan if a else "")
                w.write_string(r, 3, uraian, F.cell_wrap)
                w.set_row(r, 15 * max(1, math.ceil(len(uraian) / 72)))
                w.write_string(r, 4, satuan, F.cell_c)
                w.write_number(r, 5, b.volume, F.vol_in)
                w.write_formula(r, 6, f"=VLOOKUP(C{xl},{S_DHSP}!$C$12:$F${dhsp_last},4,FALSE)", F.money, it["hs"])
                w.write_formula(r, 7, f"=F{xl}*G{xl}", F.money, it["jumlah"])
                r += 1
            last = r
            w.write_blank(r, 1, None, F.cell)
            w.write_blank(r, 2, None, F.cell)
            w.merge_range(r, 3, r, 6, f"Subtotal {kl['nama'].title()}", F.lbl_r)
            w.write_formula(r, 7, f"=SUM(H{first}:H{last})", F.money_b, kl["subtotal"])
            kel_sub_row[(dv["kode"], kl["kode"])] = r + 1
            subs.append(r + 1)
            r += 1
        for c in (1, 2):
            w.write_blank(r, c, None, F.cell_b)
        w.merge_range(r, 3, r, 6, f"JUMLAH DIVISI {ROMAN[di] if di < len(ROMAN) else di}", F.lbl_tot)
        w.write_formula(r, 7, "=" + "+".join(f"H{x}" for x in subs), F.money_tot, dv["subtotal"])
        div_tot_row[dv["kode"]] = r + 1
        r += 2
    w.freeze_panes(9, 4)
    w.repeat_rows(8, 8)

    # ---------------- 5e. Rekap RAB ----------------
    w = ws[S_REKAP]
    for c, wd in enumerate([2, 7, 62, 22, 12, 3, 32, 14]):
        w.set_column(c, c, wd)
    w.merge_range(1, 1, 1, 4, "REKAPITULASI RENCANA ANGGARAN BIAYA (RAB)", F.title)
    _info_block(w, F, info, 3, 1, 2)
    w.write_string(3, 6, "Parameter", F.txt_b)
    w.write_string(4, 6, "PPN", F.txt)
    w.write_number(4, 7, p.ppn, F.pct_in)
    w.write_string(5, 6, "Pembulatan ROUNDUP (digit)", F.txt)
    w.write_number(5, 7, p.pembulatan_digit, F.int_in)
    w.write_string(6, 6, "Biaya Umum + Keuntungan", F.txt)
    w.write_formula(6, 7, f"={S_AHSP}!{OHP}", F.pct, p.overhead_profit)
    w.write_string(7, 6, "Wilayah harga", F.txt)
    w.write_string(7, 7, p.wilayah, F.txt_b)
    for c, h in enumerate(["NO", "URAIAN PEKERJAAN", "JUMLAH HARGA (Rp)", "BOBOT (%)"]):
        w.write_string(9, 1 + c, h, F.head)
    r = 10
    first = r + 1
    rek_rows = []
    for di, dv in enumerate(hasil["divisi"], 1):
        w.write_string(r, 1, ROMAN[di] if di < len(ROMAN) else str(di), F.cell_bc)
        w.write_string(r, 2, dv["nama"], F.cell_b)
        w.write_formula(r, 3, f"={S_BOQ}!H{div_tot_row[dv['kode']]}", F.money, dv["subtotal"])
        rek_rows.append((r + 1, dv["subtotal"]))
        r += 1
    last = r
    rJ = r + 1
    for c in (1,):
        w.write_blank(r, c, None, F.cell)
    w.write_string(r, 2, "JUMLAH", F.lbl_tot)
    w.write_formula(r, 3, f"=SUM(D{first}:D{last})", F.money_tot, hasil["jumlah"])
    w.write_formula(r, 4, f"=SUM(E{first}:E{last})", F.pct2_tot, 100.0 if hasil["jumlah"] else 0.0)
    for (rx, val) in rek_rows:
        bobot = val / hasil["jumlah"] * 100 if hasil["jumlah"] else 0
        w.write_formula(rx - 1, 4, f"=IF($D${rJ}=0,0,D{rx}/$D${rJ}*100)", F.pct2, bobot)
    r += 1
    w.write_blank(r, 1, None, F.cell)
    w.write_formula(r, 2, '="PPN "&ROUND($H$5*100,2)&"%"', F.lbl_r, f"PPN {round(p.ppn * 100, 2):g}%")
    w.write_formula(r, 3, f"=$H$5*D{rJ}", F.money_b, hasil["ppn"])
    rP = r + 1
    r += 1
    w.write_blank(r, 1, None, F.cell)
    w.write_string(r, 2, "TOTAL", F.lbl_tot)
    w.write_formula(r, 3, f"=D{rJ}+D{rP}", F.money_tot, hasil["total"])
    rT = r + 1
    r += 1
    w.write_blank(r, 1, None, F.cell)
    w.write_string(r, 2, "DIBULATKAN", F.lbl_tot)
    w.write_formula(r, 3, f"=ROUNDUP(D{rT},$H$6)", F.money_tot, hasil["dibulatkan"])
    rB = r + 1
    r += 2
    w.merge_range(r, 1, r, 4, terbilang_rupiah(hasil["dibulatkan"]), _fmt(wb, italic=True, bold=True, text_wrap=True, valign="top"))
    w.set_row(r, 32)
    r += 3
    ttd = p.penandatangan or {}
    if any(ttd.values()):
        w.write_string(r, 3, f"{ttd.get('kota', '')}, {ttd.get('tanggal', '')}".strip(", "), F.txt)
        w.write_string(r + 1, 3, ttd.get("jabatan", ""), F.txt)
        w.write_string(r + 5, 3, ttd.get("nama", ""), F.txt_b)
        r += 7
    w.freeze_panes(10, 0)

    # ---------------- 5f. Kurva S ----------------
    w = ws[S_KS]
    nweek = max(1, int(p.durasi_minggu))
    W0 = 6  # kolom minggu pertama (G)
    w.set_column(0, 0, 5)
    w.set_column(1, 1, 52)
    w.set_column(2, 2, 20)
    w.set_column(3, 3, 9)
    w.set_column(4, 5, 8)
    w.set_column(W0, W0 + nweek - 1, 8.5)
    w.merge_range(1, 0, 1, 7, "TIME SCHEDULE / KURVA S", F.title)
    _info_block(w, F, info, 3, 0, 1)
    w.write_string(8, 1, "Tanggal mulai pekerjaan (Senin minggu ke-1)", F.txt_b)
    w.write_datetime(8, 2, dt.datetime.combine(p.tgl_mulai, dt.time()), F.date_in)
    w.write_string(9, 1, "Durasi (minggu)", F.txt_b)
    w.write_number(9, 2, nweek, F.int0)
    HDR = 11
    for c, h in enumerate(["NO", "JENIS PEKERJAAN", "JUMLAH HARGA (Rp)", "BOBOT (%)", "MULAI MINGGU KE", "SELESAI MINGGU KE"]):
        w.merge_range(HDR, c, HDR + 2, c, h, F.head)
    w.merge_range(HDR, W0, HDR, W0 + nweek - 1, "JADWAL PELAKSANAAN (MINGGU)", F.head)
    for k in range(nweek):
        c = W0 + k
        w.write_number(HDR + 1, c, k + 1, F.week_h)
        d0 = p.tgl_mulai + dt.timedelta(days=7 * k)
        d1 = d0 + dt.timedelta(days=6)
        col = xl_col_to_name(c)
        w.write_formula(HDR + 2, c, f'=TEXT($C$9+7*({col}${HDR + 2}-1),"dd/mm")&"-"&TEXT($C$9+7*{col}${HDR + 2}-1,"dd/mm")',
                        F.week_d, f"{d0:%d/%m}-{d1:%d/%m}")
    w.set_row(HDR + 2, 28)
    wk_row_xl = HDR + 2  # baris Excel (1-based) berisi nomor minggu
    r = HDR + 3
    ks_first = r + 1
    ks_rows = []  # (excel_row, nilai, mulai, selesai)
    # jadwal
    allkel = [(kl["kode"], kl["subtotal"]) for dv in hasil["divisi"] for kl in dv["kelompok"]]
    auto = jadwal_otomatis(allkel, nweek)
    for di, dv in enumerate(hasil["divisi"], 1):
        w.write_string(r, 0, chr(64 + di) if di <= 26 else str(di), F.div)
        w.write_string(r, 1, dv["nama"], F.div)
        for c in range(2, W0 + nweek):
            w.write_blank(r, c, None, F.div)
        r += 1
        for ki, kl in enumerate(dv["kelompok"], 1):
            xl = r + 1
            s_, e_ = p.jadwal.get(kl["kode"]) or auto.get(kl["kode"]) or (1, nweek)
            s_ = int(min(max(1, s_), nweek))
            e_ = int(min(max(s_, e_), nweek))
            w.write_number(r, 0, ki, F.cell_c)
            w.write_string(r, 1, kl["nama"], F.cell)
            w.write_formula(r, 2, f"={S_BOQ}!H{kel_sub_row[(dv['kode'], kl['kode'])]}", F.money, kl["subtotal"])
            w.write_number(r, 4, s_, F.int_in)
            w.write_number(r, 5, e_, F.int_in)
            w.data_validation(r, 4, r, 5, {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": nweek,
                                           "error_message": f"Isi bilangan bulat 1 - {nweek}."})
            ks_rows.append((xl, kl["subtotal"], s_, e_))
            r += 1
    ks_last = r
    rt = r + 1  # baris total (Excel)
    # bobot (kolom D) - memerlukan baris total -> tulis setelah diketahui
    total_nilai = sum(v for _, v, _, _ in ks_rows)
    for (xl, val, s_, e_) in ks_rows:
        w.write_formula(xl - 1, 3, f"=IF($C${rt}=0,0,C{xl}/$C${rt}*100)", F.pct2, val / total_nilai * 100 if total_nilai else 0)
        bobot = val / total_nilai * 100 if total_nilai else 0
        for k in range(nweek):
            col = xl_col_to_name(W0 + k)
            v = bobot / (e_ - s_ + 1) if s_ <= k + 1 <= e_ else 0.0
            w.write_formula(xl - 1, W0 + k,
                            f"=IF(AND({col}${wk_row_xl}>=$E{xl},{col}${wk_row_xl}<=$F{xl}),$D{xl}/($F{xl}-$E{xl}+1),0)",
                            F.week, v)
    w.write_string(r, 0, "", F.lbl_tot)
    w.write_string(r, 1, "JUMLAH / Rencana Bobot (%)", F.lbl_tot)
    w.write_formula(r, 2, f"=SUM(C{ks_first}:C{ks_last})", F.money_tot, total_nilai)
    w.write_formula(r, 3, f"=SUM(D{ks_first}:D{ks_last})", F.pct2_tot, 100.0 if total_nilai else 0)
    w.write_blank(r, 4, None, F.lbl_tot)
    w.write_blank(r, 5, None, F.lbl_tot)
    weekly = []
    for k in range(nweek):
        tot_k = 0.0
        for (xl, val, s_, e_) in ks_rows:
            if s_ <= k + 1 <= e_ and total_nilai:
                tot_k += val / total_nilai * 100 / (e_ - s_ + 1)
        weekly.append(tot_k)
        col = xl_col_to_name(W0 + k)
        w.write_formula(r, W0 + k, f"=SUM({col}{ks_first}:{col}{ks_last})", F.pct2_tot, tot_k)
    r += 1
    w.write_string(r, 1, "Akumulatif Bobot Rencana (%)", F.lbl_tot)
    for c in (0, 2, 3, 4, 5):
        w.write_blank(r, c, None, F.lbl_tot)
    cum, acc = [], 0.0
    for k in range(nweek):
        acc += weekly[k]
        cum.append(acc)
        col = xl_col_to_name(W0 + k)
        prev = xl_col_to_name(W0 + k - 1)
        w.write_formula(r, W0 + k, f"={col}{rt}" if k == 0 else f"={prev}{r + 1}+{col}{rt}", F.pct2_tot, acc)
    r_cum = r + 1
    r += 1
    w.write_string(r, 1, "Rencana Biaya Mingguan (Rp juta)", F.lbl_tot)
    for c in (0, 2, 3, 4, 5):
        w.write_blank(r, c, None, F.lbl_tot)
    for k in range(nweek):
        col = xl_col_to_name(W0 + k)
        w.write_formula(r, W0 + k, f"={col}{rt}/100*$C${rt}/1000000", _fmt(wb, num_format="#,##0.0", bold=True, font_size=8, border=1, bg_color="#F2F2F2"),
                        weekly[k] / 100 * total_nilai / 1e6)
    r += 1
    w.write_string(r, 1, "Akumulatif Biaya (Rp juta)", F.lbl_tot)
    for c in (0, 2, 3, 4, 5):
        w.write_blank(r, c, None, F.lbl_tot)
    for k in range(nweek):
        col = xl_col_to_name(W0 + k)
        w.write_formula(r, W0 + k, f"={col}{r_cum}/100*$C${rt}/1000000", _fmt(wb, num_format="#,##0.0", bold=True, font_size=8, border=1, bg_color="#F2F2F2"),
                        cum[k] / 100 * total_nilai / 1e6)
    w.conditional_format(ks_first - 1, W0, ks_last - 1, W0 + nweek - 1,
                         {"type": "cell", "criteria": ">", "value": 0, "format": wb.add_format({"bg_color": "#9DC3E6"})})
    # grafik
    cats = [S_KS, wk_row_xl - 1, W0, wk_row_xl - 1, W0 + nweek - 1]
    col_chart = wb.add_chart({"type": "column"})
    col_chart.add_series({"name": "Bobot Mingguan (%)", "categories": cats,
                          "values": [S_KS, rt - 1, W0, rt - 1, W0 + nweek - 1], "fill": {"color": "#9DC3E6"}, "gap": 60})
    line = wb.add_chart({"type": "line"})
    line.add_series({"name": "Kumulatif Rencana (%)", "categories": cats,
                     "values": [S_KS, r_cum - 1, W0, r_cum - 1, W0 + nweek - 1], "y2_axis": True,
                     "line": {"color": "#C00000", "width": 2.5}, "marker": {"type": "circle", "size": 4}})
    col_chart.combine(line)
    col_chart.set_title({"name": "KURVA S - Rencana Bobot Pekerjaan", "name_font": {"size": 12}})
    col_chart.set_x_axis({"name": "Minggu ke-"})
    col_chart.set_y_axis({"name": "Bobot mingguan (%)", "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
    line.set_y2_axis({"name": "Kumulatif (%)", "min": 0, "max": 100})
    col_chart.set_legend({"position": "bottom"})
    col_chart.set_size({"width": 980, "height": 380})
    w.insert_chart(r + 3, 1, col_chart)
    w.freeze_panes(HDR + 3, 2)
    w.repeat_rows(HDR, HDR + 2)

    # ---------------- 5g. Catatan ----------------
    w = ws[S_CAT]
    w.set_column(0, 0, 4)
    w.set_column(1, 1, 36)
    w.set_column(2, 2, 90)
    w.write_string(1, 1, "CATATAN & ASUMSI PERHITUNGAN", F.sub)
    stat = Counter(v["res"].status for v in hasil["komponen"].values())
    lines = [
        ("Dasar analisis", "AHSP Bidang Cipta Karya & Perumahan - Lampiran VI SE DJBK No. 47 Tahun 2026 (koefisien tenaga, bahan, alat)."),
        ("Wilayah harga dasar", p.wilayah),
        ("Wilayah acuan (cadangan)", f"{p.wilayah_acuan} (faktor {p.faktor_acuan:g})" if p.pakai_acuan else "Tidak dipakai"),
        ("Dibuat pada", dt.datetime.now().strftime("%d-%m-%Y %H:%M")),
        ("Biaya umum + keuntungan", f"{p.overhead_profit * 100:g}% (ubah di sheet AHSP sel D9)"),
        ("PPN", f"{p.ppn * 100:g}% (ubah di sheet Rekap RAB sel H5)"),
        ("Jam kerja efektif / hari", f"{JAM_PER_HARI:g} jam (konversi OH<->OJ dan hari<->jam)"),
        ("Berat isi (konversi m3<->kg)", ", ".join(f"{k} {v:g} kg/m3" for k, v in p.berat_isi.items())),
        ("Komponen harga - Cocok", str(stat.get(STATUS_OK, 0))),
        ("Komponen harga - Perkiraan (periksa)", str(stat.get(STATUS_PERIKSA, 0))),
        ("Komponen harga - Acuan wilayah lain (periksa)", str(stat.get(STATUS_ACUAN, 0))),
        ("Komponen harga - Input manual", str(stat.get(STATUS_MANUAL, 0))),
        ("Komponen harga - BELUM ADA HARGA", str(stat.get(STATUS_NOL, 0))),
    ]
    for i, (a_, b_) in enumerate(lines):
        w.write_string(3 + i, 1, a_, F.cell_b)
        w.write_string(3 + i, 2, b_, F.cell_wrap)
    rr = 4 + len(lines)
    msgs = ["Pencocokan harga dilakukan otomatis berdasarkan kemiripan uraian & satuan. Periksa sheet UPAH / BAHAN / ALAT kolom Status.",
            "Harga bersatuan berbeda dikonversi (mis. sak->kg, m3->kg memakai berat isi di atas) - lihat kolom Faktor Konversi.",
            "Seluruh angka di BoQ, DHSP, Rekap RAB dan Kurva S berupa rumus: ubah sel biru (volume, harga, jadwal) dan hasil menghitung ulang.",
            "Terbilang pada Rekap RAB bersifat teks tetap (dihitung saat file dibuat)."]
    if stat.get(STATUS_NOL, 0):
        msgs.insert(0, f"PERHATIAN: {stat[STATUS_NOL]} komponen belum memiliki harga (bernilai 0). Lengkapi di sheet UPAH/BAHAN/ALAT.")
    for i, m in enumerate(msgs):
        w.merge_range(rr + i, 1, rr + i, 2, m, F.warn if m.startswith("PERHATIAN") else F.wrap)
        w.set_row(rr + i, 30)

    ws[S_REKAP].activate()
    ws[S_REKAP].set_first_sheet()
    wb.close()
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# 6. IMPOR VOLUME (Excel/CSV), TEMPLATE, DAN HARGA UNGGAHAN
# --------------------------------------------------------------------------------------
_KOL_ALIAS = {
    "uraian": ("uraian", "pekerjaan", "item", "nama", "deskripsi", "jenis pekerjaan"),
    "satuan": ("satuan", "sat", "unit"),
    "volume": ("volume", "vol", "qty", "kuantitas", "jumlah"),
    "kode": ("kode", "kode ahsp", "no ahsp"),
    "harga_manual": ("harga satuan manual", "harga manual", "harga satuan", "harga"),
}


def template_volume_xlsx() -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = wb.add_worksheet("Volume")
    hd = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "text_wrap": True, "valign": "vcenter"})
    inp = wb.add_format({"font_color": "#0000FF", "border": 1})
    ws.write_row(0, 0, ["Uraian Pekerjaan", "Satuan", "Volume", "Kode AHSP (opsional)", "Harga Satuan Manual (opsional)"], hd)
    contoh = [("Pembersihan lapangan", "m2", 1783.24, "", ""), ("Pasangan bata merah dinding 1 batu", "m2", 420.5, "", ""),
              ("Plesteran 1:4 tebal 15 mm", "m2", 840.0, "", ""), ("Pengecatan dinding", "m2", 840.0, "", ""),
              ("Pembesian slab BjTP diameter < 12 mm", "kg", 5200, "", ""),
              ("Papan nama proyek (lumpsum, harga manual)", "Ls", 1, "", 2500000)]
    for i, row in enumerate(contoh, 1):
        for c, v in enumerate(row):
            if v == "":
                ws.write_blank(i, c, None, inp)
            else:
                ws.write(i, c, v, inp)
    ws.set_column(0, 0, 58)
    ws.set_column(1, 2, 12)
    ws.set_column(3, 4, 26)
    ws.set_row(0, 30)
    h = wb.add_worksheet("Petunjuk")
    h.set_column(0, 0, 110)
    for i, t in enumerate([
        "PETUNJUK PENGISIAN",
        "1. Isi 'Uraian Pekerjaan', 'Satuan' dan 'Volume'. Tidak perlu mengetik kode AHSP - aplikasi mencocokkan otomatis.",
        "2. Kolom 'Kode AHSP' bersifat opsional (isi bila Anda sudah tahu kodenya, mis. 3.7.4); kode itu akan dipakai langsung.",
        "3. Untuk pekerjaan di luar AHSP (lumpsum), isi 'Harga Satuan Manual'. Baris tersebut tidak dianalisis AHSP.",
        "4. Setelah diunggah, periksa baris dengan tingkat kecocokan rendah dan pilih AHSP yang tepat pada aplikasi.",
        "5. Hapus baris contoh sebelum mengunggah file Anda."]):
        h.write(i, 0, t)
    wb.close()
    return buf.getvalue()


def baca_tabel_volume(nama: str, data: bytes) -> pd.DataFrame:
    """Baca file volume (xlsx/csv). Header dicari otomatis; kolom dinormalkan ke uraian/satuan/volume/kode/harga_manual."""
    if nama.lower().endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(data), header=None, dtype=object, sep=None, engine="python")
    else:
        raw = pd.read_excel(io.BytesIO(data), header=None, dtype=object, sheet_name=0)
    hdr = None
    for i in range(min(25, len(raw))):
        vals = [str(v).strip().lower() for v in raw.iloc[i].tolist() if pd.notna(v)]
        if any(v in _KOL_ALIAS["volume"] for v in vals) and any(any(a in v for a in _KOL_ALIAS["uraian"]) for v in vals):
            hdr = i
            break
    if hdr is None:
        raise ValueError("Header tidak ditemukan. Pastikan ada kolom 'Uraian Pekerjaan' dan 'Volume' (gunakan template).")
    cols = [str(v).strip().lower() if pd.notna(v) else "" for v in raw.iloc[hdr].tolist()]
    df = raw.iloc[hdr + 1:].copy()
    df.columns = cols
    peta = {}
    for tujuan in ("kode", "harga_manual", "volume", "satuan", "uraian"):
        for c in cols:
            if c and c not in peta.values() and any(c == a or (len(a) > 3 and a in c) for a in _KOL_ALIAS[tujuan]):
                if tujuan == "volume" and "harga" in c:
                    continue
                if tujuan == "uraian" and ("kode" in c or "harga" in c or c == "satuan"):
                    continue
                if tujuan == "satuan" and "harga" in c:
                    continue
                peta[tujuan] = c
                break
    out = pd.DataFrame()
    for t in ("uraian", "satuan", "volume", "kode", "harga_manual"):
        out[t] = df[peta[t]].values if t in peta else None
    out["uraian"] = out["uraian"].astype(object).where(out["uraian"].notna(), "").astype(str).str.strip()
    out["satuan"] = out["satuan"].astype(object).where(out["satuan"].notna(), "").astype(str).str.strip()
    out["kode"] = out["kode"].astype(object).where(out["kode"].notna(), "").astype(str).str.strip()
    out["volume"] = pd.to_numeric(out["volume"], errors="coerce")
    out["harga_manual"] = pd.to_numeric(out["harga_manual"], errors="coerce")
    out = out[(out["uraian"] != "") & (out["volume"].fillna(0) > 0)].reset_index(drop=True)
    return out


def cocokkan_tabel(df: pd.DataFrame) -> list[dict]:
    """Cocokkan setiap baris volume dengan AHSP. Hasil: list dict siap dipakai UI."""
    ahsp = load_ahsp()["items"]
    rows = []
    for _, r in df.iterrows():
        kode = str(r["kode"]).strip()
        kand = cocokkan_uraian(r["uraian"], r["satuan"], k=8)
        row = {"uraian": r["uraian"], "satuan": r["satuan"], "volume": float(r["volume"]),
               "harga_manual": None if pd.isna(r["harga_manual"]) else float(r["harga_manual"]),
               "kandidat": kand, "kode": "", "skor": 0.0, "sumber": "otomatis"}
        if row["harga_manual"]:
            row["sumber"] = "manual"
        elif kode in ahsp:
            row.update(kode=kode, skor=100.0, sumber="kode")
        elif kand:
            row.update(kode=kand[0]["kode"], skor=kand[0]["skor"])
        rows.append(row)
    return rows


def baca_harga_unggahan(nama: str, data: bytes) -> pd.DataFrame:
    """Format: kolom kategori (Upah/Bahan/Alat), uraian, satuan, harga."""
    if nama.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(data), sep=None, engine="python")
    else:
        xl = pd.ExcelFile(io.BytesIO(data))
        df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
    df.columns = [str(c).strip().lower() for c in df.columns]
    kolom = {"kategori": ("kategori", "jenis"), "uraian": ("uraian", "nama", "nama material"), "satuan": ("satuan",),
             "harga": ("harga", "harga satuan")}
    peta = {}
    for t, alias in kolom.items():
        for c in df.columns:
            if c in alias:
                peta[t] = c
                break
    if set(peta) != set(kolom):
        raise ValueError("Kolom wajib: kategori (Upah/Bahan/Alat), uraian, satuan, harga.")
    out = pd.DataFrame({t: df[c] for t, c in peta.items()})
    out["kategori"] = out["kategori"].astype(str).str.strip().str.capitalize()
    out["harga"] = pd.to_numeric(out["harga"], errors="coerce")
    out = out[out["kategori"].isin(["Upah", "Bahan", "Alat"]) & out["harga"].gt(0) & out["uraian"].notna()].copy()
    out["uraian"] = out["uraian"].astype(str).str.strip()
    out["satuan"] = out["satuan"].astype(str).str.strip()
    out["wilayah"] = "Unggahan Anda"
    out["sumber"] = "Unggahan pengguna"
    return out.reset_index(drop=True)


def hitung_kurva_s(hasil: dict, jadwal: dict, nweek: int) -> pd.DataFrame:
    kel = [(k["kode"], k["subtotal"]) for d in hasil["divisi"] for k in d["kelompok"]]
    total = sum(v for _, v in kel)
    auto = jadwal_otomatis(kel, nweek)
    weekly = [0.0] * nweek
    for kd, v in kel:
        s_, e_ = jadwal.get(kd) or auto.get(kd) or (1, nweek)
        s_ = min(max(1, int(s_)), nweek)
        e_ = min(max(s_, int(e_)), nweek)
        for k in range(s_ - 1, e_):
            weekly[k] += (v / total * 100 / (e_ - s_ + 1)) if total else 0.0
    cum, acc = [], 0.0
    for x in weekly:
        acc += x
        cum.append(acc)
    return pd.DataFrame({"Minggu": list(range(1, nweek + 1)), "Bobot mingguan (%)": weekly, "Kumulatif (%)": cum}).set_index("Minggu")


def rp(x: float) -> str:
    return "Rp " + f"{x:,.0f}".replace(",", ".")


# --------------------------------------------------------------------------------------
# 7. ANTARMUKA STREAMLIT
# --------------------------------------------------------------------------------------
THR_REVIEW = 72.0  # di bawah ini baris unggahan ditandai "perlu ditinjau"
WIL_CUSTOM = "Unggahan Anda"


def _ss_init():
    ss = st.session_state
    ss.setdefault("upload_rows", [])
    ss.setdefault("upload_sig", None)
    ss.setdefault("vol_katalog", {})
    ss.setdefault("harga_manual", {})
    ss.setdefault("jadwal", {})
    ss.setdefault("xlsx", None)
    ss.setdefault("xlsx_sig", None)
    ss.setdefault("cat_seed", {})
    ss.setdefault("cat_sig", None)


def _sidebar() -> Params:
    sb = st.sidebar
    sb.header("⚙️ Parameter Proyek")
    with sb.expander("Identitas proyek", expanded=True):
        info = {
            "program": st.text_input("Program", ""),
            "kegiatan": st.text_input("Kegiatan", ""),
            "pekerjaan": st.text_input("Pekerjaan", "Pembangunan Gedung"),
            "lokasi": st.text_input("Lokasi", ""),
            "tahun": st.text_input("Tahun Anggaran", str(dt.date.today().year)),
        }
    sb.subheader("📍 Wilayah Proyek")
    wil_list = daftar_wilayah()
    pilihan = wil_list + [WIL_CUSTOM + " (unggah daftar harga sendiri)"]
    idx = sb.selectbox("Wilayah (menentukan harga upah, bahan & alat)", range(len(pilihan)), format_func=lambda i: pilihan[i])
    harga_df, wilayah = None, wil_list[idx] if idx < len(wil_list) else WIL_CUSTOM
    if wilayah == WIL_CUSTOM:
        f = sb.file_uploader("Daftar harga wilayah (xlsx/csv)", type=["xlsx", "csv"], key="up_harga",
                             help="Kolom wajib: kategori (Upah/Bahan/Alat), uraian, satuan, harga.")
        if f is not None:
            try:
                harga_df = baca_harga_unggahan(f.name, f.getvalue())
                sb.success(f"{len(harga_df)} baris harga terbaca.")
            except Exception as e:  # noqa: BLE001
                sb.error(f"Gagal membaca harga: {e}")
        if harga_df is None:
            sb.warning("Unggah daftar harga untuk memakai wilayah ini; sementara harga acuan dipakai.")
            wilayah = wil_list[0]
    pakai_acuan = sb.checkbox("Pakai harga wilayah acuan bila harga tidak ditemukan", value=True,
                              help="Komponen yang tidak ada pada daftar harga wilayah diisi dari wilayah acuan dan ditandai 'periksa'.")
    acuan, faktor = "DKI Jakarta", 1.0
    if pakai_acuan:
        acuan = sb.selectbox("Wilayah acuan", wil_list, index=wil_list.index("DKI Jakarta") if "DKI Jakarta" in wil_list else 0)
        faktor = sb.number_input("Faktor penyesuaian harga acuan", 0.3, 3.0, 1.0, 0.01,
                                 help="Mis. 0,95 bila harga di wilayah proyek ±5% lebih murah dari wilayah acuan.")
    sb.subheader("🧮 Parameter Biaya")
    ohp = sb.number_input("Biaya umum + keuntungan (%)", 0.0, 40.0, 15.0, 0.5, help="SE DJBK: 10% - 15% dari jumlah tenaga+bahan+alat.")
    ppn = sb.number_input("PPN (%)", 0.0, 25.0, 12.0, 0.5)
    digit = sb.selectbox("Pembulatan ke atas (ROUNDUP)", [0, -3, -5], index=1,
                         format_func=lambda d: {0: "Rupiah penuh", -3: "Ribuan", -5: "Ratusan ribu"}[d])
    with sb.expander("Berat isi bahan (konversi m³ ↔ kg)"):
        bi = {"pasir": st.number_input("Pasir (kg/m³)", 800.0, 2200.0, BERAT_ISI_DEFAULT["pasir"], 10.0),
              "kerikil": st.number_input("Kerikil/split (kg/m³)", 800.0, 2200.0, BERAT_ISI_DEFAULT["kerikil"], 10.0),
              "batu": st.number_input("Batu kali/belah (kg/m³)", 800.0, 2800.0, BERAT_ISI_DEFAULT["batu"], 10.0)}
        st.caption("Dipakai bila AHSP memakai kg sedangkan harga wilayah per m³ (atau sebaliknya). Asumsi - sesuaikan dengan hasil uji bahan.")
    sb.subheader("🗓️ Jadwal")
    durasi = int(sb.number_input("Durasi pelaksanaan (minggu)", 2, 120, 24, 1))
    tgl = sb.date_input("Tanggal mulai", dt.date.today())
    with sb.expander("Penandatangan (opsional)"):
        ttd = {"kota": st.text_input("Kota", ""), "tanggal": st.text_input("Tanggal", dt.date.today().strftime("%d %B %Y")),
               "jabatan": st.text_input("Jabatan", ""), "nama": st.text_input("Nama", "")}
        if not ttd["kota"] and not ttd["jabatan"] and not ttd["nama"]:
            ttd = {}
    return Params(wilayah=wilayah if harga_df is None else WIL_CUSTOM, wilayah_acuan=acuan, pakai_acuan=pakai_acuan,
                  faktor_acuan=float(faktor), overhead_profit=ohp / 100, ppn=ppn / 100, pembulatan_digit=int(digit),
                  berat_isi=bi, harga_manual=dict(st.session_state.harga_manual), harga_wilayah_df=harga_df, info=info,
                  durasi_minggu=durasi, tgl_mulai=tgl, jadwal=dict(st.session_state.jadwal), penandatangan=ttd)


def _tab_upload():
    ss = st.session_state
    st.markdown("Unggah daftar pekerjaan beserta **volume**. Tidak perlu kode AHSP - aplikasi mencocokkan uraian otomatis.")
    c1, c2 = st.columns([1, 3])
    c1.download_button("⬇️ Unduh template", template_volume_xlsx(), "template_volume_rab.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    f = c2.file_uploader("File volume (xlsx / csv)", type=["xlsx", "csv"], key="up_vol")
    if f is not None:
        sig = (f.name, f.size)
        if ss.upload_sig != sig:
            try:
                with st.spinner("Mencocokkan uraian dengan AHSP..."):
                    ss.upload_rows = cocokkan_tabel(baca_tabel_volume(f.name, f.getvalue()))
                ss.upload_sig = sig
            except Exception as e:  # noqa: BLE001
                st.error(f"Gagal membaca file: {e}")
                return
    if not ss.upload_rows:
        return
    rows = ss.upload_rows
    n_rev = sum(1 for r in rows if r["sumber"] == "otomatis" and r["skor"] < THR_REVIEW and r["kode"] != "")
    st.success(f"{len(rows)} baris terbaca. " + (f"**{n_rev} baris perlu ditinjau** (kecocokan rendah) - lihat di bawah." if n_rev else "Semua baris tercocokkan dengan yakin."))
    ahsp = load_ahsp()["items"]
    tbl = pd.DataFrame([{
        "No": i + 1, "Uraian (file Anda)": r["uraian"], "Satuan": r["satuan"], "Volume": r["volume"],
        "Dicocokkan ke AHSP": ("(harga manual)" if r["sumber"] == "manual" else
                               (f"{r['kode']} - {ahsp[r['kode']]['uraian'][:90]}" if r["kode"] in ahsp else "(diabaikan)")),
        "Kecocokan": "manual" if r["sumber"] == "manual" else ("kode" if r["sumber"] == "kode" else f"{r['skor']:.0f}%"),
    } for i, r in enumerate(rows)])
    st.dataframe(tbl, hide_index=True, use_container_width=True, height=min(420, 38 + 35 * len(tbl)))
    flagged = [i for i, r in enumerate(rows) if r["sumber"] in ("otomatis", "dipilih") and (r["skor"] < THR_REVIEW or r["sumber"] == "dipilih") or (r["sumber"] == "otomatis" and not r["kode"])]
    with st.expander(f"🔍 Tinjau & ubah pencocokan ({len(flagged)} baris kecocokan rendah)", expanded=bool(flagged)):
        pilih_semua = st.checkbox("Tampilkan semua baris (bukan hanya yang kecocokannya rendah)")
        daftar = range(len(rows)) if pilih_semua else flagged
        for i in daftar:
            r = rows[i]
            opsi = [k["kode"] for k in r["kandidat"]]
            lab = {k["kode"]: f"{k['kode']} · {k['uraian'][:100]} ({k['satuan']}) - {k['skor']:.0f}%" for k in r["kandidat"]}
            cari = st.text_input(f"{i + 1}. {r['uraian']} - {r['volume']:g} {r['satuan']}  |  cari AHSP lain di katalog",
                                 key=f"cari_{ss.upload_sig}_{i}", placeholder="ketik kata kunci, mis. ready mixed 25 MPa  atau  K-300")
            if cari.strip():
                for hit in cari_katalog(cari).itertuples():
                    if hit.kode not in lab:
                        opsi.insert(0, hit.kode)
                        lab[hit.kode] = f"🔎 {hit.kode} · {hit.uraian[:100]} ({hit.satuan})"
            if r["kode"] and r["kode"] not in opsi:
                opsi.insert(0, r["kode"])
                lab[r["kode"]] = f"{r['kode']} · {ahsp[r['kode']]['uraian'][:100]} (dipilih)"
            opsi.append("__abaikan__")
            lab["__abaikan__"] = "(abaikan baris ini)"
            cur = r["kode"] if r["kode"] in opsi else opsi[0]
            sel = st.selectbox("Pilih AHSP", opsi, index=opsi.index(cur), format_func=lambda k, lab=lab: lab[k],
                               key=f"pil_{ss.upload_sig}_{i}_{hash(cari)}", label_visibility="collapsed")
            if sel != r["kode"]:
                r["sumber"] = "dipilih"
                r["skor"] = 100.0
            r["kode"] = "" if sel == "__abaikan__" else sel
        st.caption("Kolom pencarian memakai seluruh katalog AHSP. Pilihan yang sudah Anda tentukan tidak lagi ditandai 'perlu ditinjau'.")
    if st.button("🗑️ Hapus data unggahan"):
        ss.upload_rows, ss.upload_sig = [], None
        st.rerun()


def _tab_katalog():
    ss = st.session_state
    df = katalog_df()
    c1, c2 = st.columns([2, 3])
    divs = ["(Semua divisi)"] + [f"{k} - {n}" for k, n in df[["div", "div_nama"]].drop_duplicates().itertuples(index=False)]
    dsel = c1.selectbox("Divisi", divs)
    sub = df if dsel == divs[0] else df[df["div"] == dsel.split(" - ")[0]]
    kels = ["(Semua kelompok)"] + [f"{k} - {n}" for k, n in sub[["kel", "kel_nama"]].drop_duplicates().itertuples(index=False)]
    ksel = c2.selectbox("Kelompok pekerjaan", kels)
    if ksel != kels[0]:
        sub = sub[sub["kel"] == ksel.split(" - ")[0]]
    q = st.text_input("🔎 Cari uraian pekerjaan (semua kata harus ada)", placeholder="mis. plesteran 1:4  |  bata merah  |  keramik 40")
    if q.strip():
        kata = [t for t in re.split(r"\s+", q.lower().strip()) if t]
        m = pd.Series(True, index=sub.index)
        for t in kata:
            m &= sub["uraian"].str.lower().str.contains(re.escape(t), regex=True)
        sub = sub[m]
    MAXR = 300
    st.caption(f"{len(sub)} analisis ditemukan" + (f" (ditampilkan {MAXR} pertama - persempit pencarian)" if len(sub) > MAXR else "") +
               ". Isi kolom **Volume** pada pekerjaan yang dibutuhkan.")
    sub = sub.head(MAXR)
    sig = (dsel, ksel, q.strip().lower())
    if ss.cat_sig != sig:
        ss.cat_sig = sig
        ss.cat_seed = dict(ss.vol_katalog)
    view = pd.DataFrame({"Kode": sub["kode"].values, "Uraian": sub["uraian"].values, "Satuan": sub["satuan"].values,
                         "Volume": [float(ss.cat_seed.get(k, 0.0)) for k in sub["kode"]]})
    ed = st.data_editor(view, hide_index=True, use_container_width=True, height=430, key=f"ed_{hash(sig)}",
                        disabled=["Kode", "Uraian", "Satuan"],
                        column_config={"Volume": st.column_config.NumberColumn("Volume", min_value=0.0, format="%.2f"),
                                       "Uraian": st.column_config.TextColumn("Uraian", width="large")})
    for k, v in zip(ed["Kode"], ed["Volume"]):
        if v and v > 0:
            ss.vol_katalog[k] = float(v)
        else:
            ss.vol_katalog.pop(k, None)


def _kumpulkan_baris() -> list[Baris]:
    ss = st.session_state
    ahsp = load_ahsp()["items"]
    baris = []
    for r in ss.upload_rows:
        if r.get("harga_manual"):
            baris.append(Baris("", r["uraian"], r["satuan"], r["volume"], harga_manual=r["harga_manual"]))
        elif r["kode"] in ahsp:
            baris.append(Baris(r["kode"], r["uraian"], r["satuan"] or ahsp[r["kode"]]["satuan"], r["volume"], skor=r["skor"]))
    for kode, vol in sorted(ss.vol_katalog.items(), key=lambda kv: _kode_key(kv[0])):
        if kode in ahsp:
            baris.append(Baris(kode, ahsp[kode]["uraian"], ahsp[kode]["satuan"], vol))
    return baris


def _tab_ringkasan(baris: list[Baris]):
    st.subheader("📋 Daftar pekerjaan yang akan dihitung")
    if not baris:
        st.info("Belum ada pekerjaan. Unggah file volume atau isi volume dari katalog AHSP.")
        return
    ahsp = load_ahsp()["items"]
    df = pd.DataFrame([{"Kode": b.kode or "(manual)", "Uraian": b.uraian, "Satuan": b.satuan, "Volume": b.volume,
                        "AHSP": (ahsp[b.kode]["uraian"][:80] if b.kode in ahsp else "-")} for b in baris])
    st.dataframe(df, hide_index=True, use_container_width=True, height=min(360, 38 + 35 * len(df)))
    if st.button("Kosongkan seluruh daftar volume"):
        st.session_state.upload_rows, st.session_state.upload_sig = [], None
        st.session_state.vol_katalog, st.session_state.cat_sig = {}, None
        st.rerun()


def _tab_jadwal(hasil: dict, p: Params):
    ss = st.session_state
    kel = [(k["kode"], k["nama"], k["subtotal"]) for d in hasil["divisi"] for k in d["kelompok"]]
    if not kel:
        st.info("Isi volume pekerjaan terlebih dulu (tab 1).")
        return
    auto = jadwal_otomatis([(k, v) for k, _, v in kel], p.durasi_minggu)
    total = sum(v for _, _, v in kel) or 1
    df = pd.DataFrame([{"Kode": k, "Kelompok pekerjaan": n, "Nilai (Rp)": v, "Bobot (%)": v / total * 100,
                        "Mulai (minggu ke-)": int((ss.jadwal.get(k) or auto[k])[0]),
                        "Selesai (minggu ke-)": int((ss.jadwal.get(k) or auto[k])[1])} for k, n, v in kel])
    st.caption(f"Usulan jadwal otomatis untuk durasi {p.durasi_minggu} minggu. Ubah kolom Mulai/Selesai sesuai metode kerja Anda.")
    ed = st.data_editor(df, hide_index=True, use_container_width=True, key="ed_jadwal",
                        disabled=["Kode", "Kelompok pekerjaan", "Nilai (Rp)", "Bobot (%)"],
                        column_config={"Nilai (Rp)": st.column_config.NumberColumn(format="%.0f"),
                                       "Bobot (%)": st.column_config.NumberColumn(format="%.2f"),
                                       "Mulai (minggu ke-)": st.column_config.NumberColumn(min_value=1, max_value=p.durasi_minggu, step=1),
                                       "Selesai (minggu ke-)": st.column_config.NumberColumn(min_value=1, max_value=p.durasi_minggu, step=1)})
    baru = {r["Kode"]: (int(r["Mulai (minggu ke-)"]), int(r["Selesai (minggu ke-)"])) for _, r in ed.iterrows()}
    if baru != {k: tuple(v) for k, v in ss.jadwal.items() if k in baru} and baru != auto:
        ss.jadwal = baru
        st.rerun()
    st.caption("Grafik Kurva S (pratinjau):")
    ks = hitung_kurva_s(hasil, {k: tuple(v) for k, v in ss.jadwal.items()}, p.durasi_minggu)
    st.line_chart(ks[["Kumulatif (%)"]])
    st.bar_chart(ks[["Bobot mingguan (%)"]])


def _tab_hasil(baris: list[Baris], hasil: dict, p: Params):
    ss = st.session_state
    if not baris or hasil["n_item"] == 0:
        st.info("Belum ada pekerjaan untuk dihitung.")
        return
    c = st.columns(4)
    c[0].metric("Jumlah (sebelum PPN)", rp(hasil["jumlah"]))
    c[1].metric(f"PPN {p.ppn * 100:g}%", rp(hasil["ppn"]))
    c[2].metric("Total", rp(hasil["total"]))
    c[3].metric("Dibulatkan", rp(hasil["dibulatkan"]))
    st.caption(terbilang_rupiah(hasil["dibulatkan"]))
    rekap = pd.DataFrame([{"Divisi": f"{d['kode']} - {d['nama']}", "Jumlah Harga (Rp)": d["subtotal"],
                           "Bobot (%)": d["subtotal"] / hasil["jumlah"] * 100 if hasil["jumlah"] else 0} for d in hasil["divisi"]])
    st.dataframe(rekap, hide_index=True, use_container_width=True,
                 column_config={"Jumlah Harga (Rp)": st.column_config.NumberColumn(format="%.0f"),
                                "Bobot (%)": st.column_config.NumberColumn(format="%.2f")})
    if hasil["dilewati"]:
        st.warning(f"{len(hasil['dilewati'])} baris tidak dihitung karena tidak ada AHSP terpilih / harga manual.")

    # ---- status harga ----
    stat = Counter(v["res"].status for v in hasil["komponen"].values())
    st.subheader("💲 Status harga satuan dasar")
    cols = st.columns(5)
    for col, key in zip(cols, (STATUS_OK, STATUS_PERIKSA, STATUS_ACUAN, STATUS_MANUAL, STATUS_NOL)):
        col.metric(key, stat.get(key, 0))
    if stat.get(STATUS_NOL, 0):
        st.error(f"{stat[STATUS_NOL]} komponen belum memiliki harga (dihitung Rp 0) sehingga RAB bisa lebih rendah dari seharusnya. "
                 "Isi harganya di bawah ini atau unggah daftar harga wilayah yang lebih lengkap.")
    perlu = [(k, v) for k, v in hasil["komponen"].items()
             if v["res"].status in STATUS_NEEDS_REVIEW or v["res"].status == STATUS_MANUAL]
    if perlu:
        with st.expander(f"✏️ Periksa / isi harga manual ({len(perlu)} komponen)", expanded=bool(stat.get(STATUS_NOL, 0))):
            perlu.sort(key=lambda kv: (kv[1]["res"].status != STATUS_NOL, kv[0][0], kv[1]["uraian"]))
            st.caption("Ubah kolom **Harga baru** (per satuan AHSP, mis. per kg / m³ / OH / jam), lalu tekan **Terapkan harga**. "
                       "Kosongkan kolom untuk membatalkan harga manual.")
            d = pd.DataFrame([{"Kategori": k[0], "Uraian (AHSP)": v["uraian"], "Satuan": v["satuan"], "Status": v["res"].status,
                               "Harga saat ini": float(round(v["res"].harga, 2)),
                               "Harga baru": float(ss.harga_manual[k]) if k in ss.harga_manual else None,
                               "Dicocokkan dengan": v["res"].src_nama or "-"} for k, v in perlu])
            sig_ed = hash(tuple(k for k, _ in perlu))
            with st.form(f"form_harga_{sig_ed}"):
                ed = st.data_editor(d, hide_index=True, use_container_width=True, height=min(450, 38 + 35 * len(d)),
                                    key=f"ed_harga_{sig_ed}",
                                    disabled=["Kategori", "Uraian (AHSP)", "Satuan", "Status", "Harga saat ini", "Dicocokkan dengan"],
                                    column_config={"Harga baru": st.column_config.NumberColumn("Harga baru (Rp)", min_value=0.0, format="%.2f"),
                                                   "Harga saat ini": st.column_config.NumberColumn(format="%.2f")})
                ok = st.form_submit_button("Terapkan harga", type="primary")
            if ok:
                for (k, _), (_, r) in zip(perlu, ed.iterrows()):
                    nv = r["Harga baru"]
                    if pd.notna(nv) and float(nv) >= 0:
                        ss.harga_manual[k] = float(nv)
                    else:
                        ss.harga_manual.pop(k, None)
                st.rerun()
    if ss.harga_manual and st.button("Hapus semua harga manual"):
        ss.harga_manual = {}
        st.rerun()

    # ---- ekspor ----
    st.subheader("📥 Unduh Excel")
    sig = repr((tuple((b.kode, b.uraian, b.satuan, b.volume, b.harga_manual) for b in baris), p.wilayah, p.wilayah_acuan,
                p.pakai_acuan, p.faktor_acuan, p.overhead_profit, p.ppn, p.pembulatan_digit, sorted(p.berat_isi.items()),
                sorted(p.harga_manual.items()), sorted(p.info.items()), p.durasi_minggu, str(p.tgl_mulai),
                sorted(p.jadwal.items()), sorted(p.penandatangan.items()), None if p.harga_wilayah_df is None else len(p.harga_wilayah_df)))
    if ss.xlsx_sig != sig:
        ss.xlsx, ss.xlsx_sig = None, None
    if st.button("⚙️ Buat file Excel RAB", type="primary"):
        with st.spinner("Menyusun AHSP, DHSP, BoQ, Rekap RAB dan Kurva S..."):
            ss.xlsx = buat_excel(baris, p, hasil)
            ss.xlsx_sig = sig
    if ss.xlsx:
        nama = re.sub(r"[^A-Za-z0-9]+", "_", p.wilayah).strip("_")
        st.download_button("⬇️ Unduh RAB.xlsx", ss.xlsx, f"RAB_{nama}_{dt.date.today():%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.caption("Sheet: Rekap RAB · BoQ · Kurva S · DHSP · AHSP · UPAH · BAHAN · ALAT · Catatan. Semua angka berupa rumus hidup.")


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🏗️", layout="wide")
    _ss_init()
    st.title("🏗️ Generator RAB Konstruksi")
    st.caption("Berdasarkan AHSP Bidang Cipta Karya - Lampiran VI SE DJBK No. 47 Tahun 2026 · Pilih wilayah → isi volume → unduh Excel "
               "(AHSP, DHSP, BoQ, Rekap RAB, Kurva S)")
    p = _sidebar()
    t1, t2, t3, t4 = st.tabs(["1️⃣ Volume Pekerjaan", "2️⃣ Jadwal & Kurva S", "3️⃣ Hitung & Unduh", "ℹ️ Panduan"])
    with t1:
        s1, s2 = st.tabs(["📤 Unggah Volume (Excel/CSV)", "🔎 Pilih dari Katalog AHSP"])
        with s1:
            _tab_upload()
        with s2:
            _tab_katalog()
        baris = _kumpulkan_baris()
        _tab_ringkasan(baris)
    baris = _kumpulkan_baris()
    hasil = hitung_rab(baris, p) if baris else {"divisi": [], "n_item": 0, "komponen": {}, "dilewati": []}
    with t2:
        _tab_jadwal(hasil, p)
    with t3:
        _tab_hasil(baris, hasil, p)
    with t4:
        st.markdown(PANDUAN)


PANDUAN = """
### Cara pakai
1. **Pilih wilayah proyek** di panel kiri - menentukan harga satuan dasar upah, bahan, dan alat.
2. **Isi volume** (tab 1): unggah Excel/CSV (uraian, satuan, volume) *atau* cari pekerjaan di katalog AHSP lalu isi volumenya.
   Aplikasi mencocokkan uraian ke AHSP SE DJBK 47/2026 secara otomatis; baris yang ragu-ragu ditandai untuk ditinjau.
3. **Atur jadwal** (tab 2) - usulan otomatis per kelompok pekerjaan, dapat diubah.
4. **Hitung & unduh** (tab 3): periksa status harga, lalu buat file Excel.

### Isi file Excel
| Sheet | Isi |
|---|---|
| Rekap RAB | Rekap per divisi, PPN, total, pembulatan, terbilang |
| BoQ | Bill of Quantities: kode, uraian, satuan, volume, harga satuan, jumlah |
| Kurva S | Jadwal mingguan (Gantt), bobot, kumulatif, grafik |
| DHSP | Daftar harga satuan pekerjaan |
| AHSP | Analisis tenaga/bahan/alat + biaya umum & keuntungan |
| UPAH / BAHAN / ALAT | Harga dasar yang dipakai beserta sumber & status pencocokan |

### Catatan penting
- Harga **bukan** bagian dari SE (SE hanya memuat koefisien). Harga berasal dari daftar harga wilayah yang tersedia atau yang Anda unggah.
- Komponen berstatus *Perkiraan*, *Acuan wilayah lain*, atau *Belum ada harga* **wajib diperiksa** sebelum RAB dipakai sebagai dokumen resmi.
- Hasil pencocokan uraian bersifat usulan; pastikan AHSP yang terpilih sesuai spesifikasi pekerjaan.
"""

if st is not None and __name__ == "__main__":
    main()
