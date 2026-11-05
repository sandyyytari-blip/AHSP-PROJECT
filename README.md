# Generator RAB Konstruksi (AHSP SE DJBK 47/2026)

Jalankan lokal:  `pip install -r requirements.txt && streamlit run rab_engine.py`
Deploy Streamlit Cloud: push folder ini ke GitHub, Main file path = `rab_engine.py`.

Struktur:
- rab_engine.py      aplikasi + engine (hitung AHSP, cocokkan harga, buat Excel)
- data/ahsp_db.json  2.797 analisis AHSP hasil ekstraksi lampiran SE (koefisien)
- data/harga_dasar.csv  harga dasar upah/bahan/alat: DKI Jakarta & Kota Padang
- tools/parse_ahsp.py, tools/build_harga.py  skrip pembuat kedua data di atas

Tambah wilayah: tambahkan baris (wilayah,kategori,uraian,satuan,harga,sumber) ke data/harga_dasar.csv,
atau unggah daftar harga via sidebar (kolom: kategori, uraian, satuan, harga).
