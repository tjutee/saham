# Dashboard Rekomendasi Saham IDX

Dashboard Streamlit interaktif untuk screening dan rekomendasi awal saham Indonesia
dengan alur data online-first. Universe kode saham diprioritaskan dari daftar
resmi BEI/IDX, harga dan histori diprioritaskan dari `yfinance`, dan
`Ringkasan.xlsx` dipakai sebagai fallback untuk rasio fundamental, metrik bank,
sektor, serta data yang belum tersedia online.

## Fitur

- Universe kode saham online dari endpoint resmi BEI/IDX, dengan fallback TradingView, StockAnalysis, dan Excel.
- Audit sumber kode menandai kode yang match BEI/IDX resmi dan kode yang hanya tersedia dari fallback.
- Deduplikasi saham berdasarkan `Kode`, karena saham yang sama bisa muncul di beberapa indeks.
- Scoring multi-factor: valuasi, kualitas profit, risiko, likuiditas, momentum, dan kekuatan indeks.
- Filter threshold dari sheet `NonBank` dan `Banking`.
- Histori return 4, 13, 26, 52 minggu, dan YTD dihitung dari yfinance bila tersedia.
- Grafik histori online fleksibel untuk saham IDX dengan ticker `KODE.JK`.
- Sumber harga/histori online utama: `yfinance`, fallback: `pandas-datareader`, cache lokal, lalu Excel bila tersedia.
- Profil scoring: Balanced, Defensive, Growth, dan Value.
- Bobot scoring bisa diatur langsung dari sidebar.
- Dashboard dinamis dengan tab Ringkasan, Rekomendasi, Explorer, Histori Harga, Sektor, Data Quality, dan Metodologi.
- Ringkasan eksekutif berisi distribusi rekomendasi, komposisi risiko, sumber data, top kandidat, dan heatmap faktor.
- Data Quality menampilkan audit sumber kode, audit filter, coverage kolom, serta campuran sumber harga/volume.
- Tabel hasil bisa di-download sebagai CSV.

## File Utama

```text
streamlit_app.py      # Aplikasi utama Streamlit
Ringkasan.xlsx        # Fallback rasio/fundamental dan cadangan data
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

Fitur online membutuhkan koneksi internet. Jika sumber online gagal, dashboard
mencoba fallback online lain, cache lokal di `history_cache/`, lalu mengisi
kolom yang kosong dari `Ringkasan.xlsx` bila file tersedia.

## Prioritas Data

1. Universe kode saham: endpoint resmi BEI/IDX.
2. Cadangan universe: TradingView, StockAnalysis, lalu `Ringkasan.xlsx`.
3. Harga, volume, OHLC, dan histori: `yfinance`.
4. Cadangan histori: `pandas-datareader`, cache lokal, lalu sheet `Metrik`.
5. Fundamental dan metrik bank: `Ringkasan.xlsx` sampai tersedia API fundamental yang lebih lengkap.

Tab `Data Quality` menampilkan audit sumber kode agar perbedaan antara daftar
resmi BEI/IDX dan fallback tetap terlihat, bukan diabaikan.

## Histori Harga

Tab `Histori Harga` menyediakan dua mode:

- `Excel Metrik 4W-52W`: memakai return historis dari sheet `Metrik` sebagai mode pembanding/cadangan.
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
- Momentum: histori online 4, 13, 26, 52 minggu dan perubahan harga harian yang tidak ekstrem, dengan sheet `Metrik` sebagai fallback.
- Kekuatan indeks: nilai kolom Sigma i >= 7 dari Excel bila tersedia, lalu jumlah indeks/tempat kemunculan dari fallback, atau minimal satu untuk kode yang hanya tersedia dari universe resmi.
- Threshold: batas rasio dari sheet `NonBank` atau `Banking`.

Penalti diterapkan untuk data yang kurang sehat, seperti PER/PBV negatif,
profitabilitas negatif, volume rendah, harga nol, perubahan harian ekstrem,
dan kelulusan threshold yang terlalu rendah.

## Catatan

Dashboard ini adalah alat screening kuantitatif awal, bukan saran investasi.
Tetap cek laporan keuangan terbaru, aksi korporasi, berita material, manajemen,
dan risiko portofolio sebelum mengambil keputusan.
