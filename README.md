# Dashboard Rekomendasi Saham IDX

Dashboard Streamlit interaktif untuk screening dan rekomendasi awal saham Indonesia
berdasarkan file data `Ringkasan.xlsx`.

## Fitur

- Deduplikasi saham berdasarkan `Kode`, karena saham yang sama bisa muncul di beberapa indeks.
- Scoring multi-factor: valuasi, kualitas profit, risiko, likuiditas, momentum, dan kekuatan indeks.
- Filter threshold dari sheet `NonBank` dan `Banking`.
- Histori return 4, 13, 26, dan 52 minggu dari sheet `Metrik`.
- Grafik histori online fleksibel untuk saham IDX dengan ticker `KODE.JK`.
- Sumber histori online utama: `yfinance`, fallback: `pandas-datareader`, lalu cache lokal.
- Profil scoring: Balanced, Defensive, Growth, dan Value.
- Bobot scoring bisa diatur langsung dari sidebar.
- Dashboard dinamis dengan tab Rekomendasi, Explorer, Histori Harga, Sektor, dan Metodologi.
- Tabel hasil bisa di-download sebagai CSV.

## File Utama

```text
streamlit_app.py      # Aplikasi utama Streamlit
Ringkasan.xlsx        # Data saham yang dibaca aplikasi
requirements.txt      # Dependency Python
README.md             # Dokumentasi singkat
```

## Menjalankan Lokal

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Aplikasi akan terbuka di:

```text
http://localhost:8501
```

Jika command `streamlit` tidak tersedia di PATH:

```bash
python -m streamlit run streamlit_app.py
```

Fitur histori online membutuhkan koneksi internet. Jika sumber online gagal,
dashboard mencoba cache lokal di `history_cache/`; jika cache belum ada,
dashboard tetap memakai data Excel untuk rekomendasi utama.

## Histori Harga

Tab `Histori Harga` menyediakan dua mode:

- `Excel Metrik 4W-52W`: memakai return historis dari sheet `Metrik` untuk grafik ringkas 4, 13, 26, dan 52 minggu.
- `Online yfinance KODE.JK`: mengambil harga historis IDX secara online dengan format ticker seperti `BBCA.JK`, `BMRI.JK`, atau `TLKM.JK`.

Mode online mendukung rentang 6 bulan, 1 tahun, 2 tahun, 5 tahun, 10 tahun,
dan all/sepanjang masa. Jika koneksi live gagal, dashboard mencoba fallback
`pandas-datareader` dan cache lokal di folder `history_cache/`.

## Metodologi Singkat

Score akhir memakai normalisasi percentile yang dipotong di persentil 3 dan 97
agar outlier ekstrem tidak mendominasi ranking.

Faktor yang dihitung:

- Valuasi: PER rendah dan PBV rendah.
- Kualitas profit: ROE, ROA, dan NPM positif.
- Risiko: DER rendah dan intraday range rendah.
- Likuiditas: volume dan turnover harga x volume.
- Momentum: histori 4, 13, 26, 52 minggu dan perubahan harga harian yang tidak ekstrem.
- Kekuatan indeks: jumlah indeks tempat saham masuk.
- Threshold: batas rasio dari sheet `NonBank` atau `Banking`.

Penalti diterapkan untuk data yang kurang sehat, seperti PER/PBV negatif,
profitabilitas negatif, volume rendah, harga nol, perubahan harian ekstrem,
dan kelulusan threshold yang terlalu rendah.

## Catatan

Dashboard ini adalah alat screening kuantitatif awal, bukan saran investasi.
Tetap cek laporan keuangan terbaru, aksi korporasi, berita material, manajemen,
dan risiko portofolio sebelum mengambil keputusan.
