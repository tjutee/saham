# Dashboard Rekomendasi Saham IDX

Dashboard Streamlit interaktif untuk screening dan rekomendasi awal saham Indonesia
dengan alur data online-first. Universe kode saham diprioritaskan dari daftar
resmi BEI/IDX, harga dan histori diprioritaskan dari `yfinance`, dan
fundamental online dilengkapi dari TradingView scanner. `Ringkasan.xlsx`
dipakai sebagai fallback, pembanding, dan sumber ide algoritme untuk data yang
belum stabil tersedia online.

## Fitur

- Universe kode saham online dari endpoint resmi BEI/IDX, dengan fallback TradingView, StockAnalysis, dan Excel.
- Audit sumber kode menandai kode yang match BEI/IDX resmi dan kode yang hanya tersedia dari fallback.
- Fundamental online massal dari TradingView scanner untuk PER, PBV, ROE, ROA, DER, NPM, market cap, revenue, sektor, dan industri bila tersedia.
- Deduplikasi saham berdasarkan `Kode`, karena saham yang sama bisa muncul di beberapa indeks.
- Scoring multi-factor: valuasi, kualitas profit, risiko, likuiditas, momentum, dan kekuatan indeks.
- Sector-relative scoring untuk membandingkan valuasi dan kualitas saham terhadap sektor yang sama.
- Explainability layer: `Decision_Summary`, `Top_Strengths`, `Top_Risks`, dan `Action_Checklist`.
- Market regime dari IHSG `^JKSE`: Risk-On, Neutral, atau Risk-Off berdasarkan MA50/MA200 dan momentum 20D/60D.
- Market breadth: persentase saham sample yang berada di atas MA50 dan MA200.
- Data freshness guard untuk melihat tanggal online terakhir, lag data, coverage online, dan fallback Excel.
- Filter threshold dari sheet `NonBank` dan `Banking`.
- Histori return 4, 13, 26, 52 minggu, dan YTD dihitung dari yfinance bila tersedia.
- Grafik utama memprioritaskan harga/volume yfinance atau cache online; Excel hanya dipakai bila online/cache kosong.
- Sheet `Metrik` dimanfaatkan untuk market cap, revenue, subsektor, industri, subindustri, dan daftar indeks gabungan.
- Grafik histori online fleksibel untuk saham IDX dengan ticker `KODE.JK`.
- Analisa teknikal berbasis OHLCV yfinance/cache: candlestick, MA20/50/200, RSI, MACD, ATR, Technical Score, Entry Action, Position Action, dan ATR stop zone.
- Sumber harga/histori online utama: `yfinance`, fallback: `pandas-datareader`, cache lokal, lalu Excel bila tersedia.
- Profil scoring: Balanced, Defensive, Growth, dan Value.
- Bobot scoring bisa diatur langsung dari sidebar.
- Dashboard dinamis dengan tab Ringkasan, Rekomendasi, Harga & Teknikal, Explorer, Sektor, Kualitas Data, dan Metodologi.
- Ringkasan eksekutif berisi distribusi rekomendasi, komposisi risiko, sumber data, top kandidat, dan heatmap faktor.
- Kualitas Data menampilkan audit sumber kode, audit filter, coverage kolom, serta campuran sumber harga/volume.
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

1. Universe kode saham dan metadata listing: endpoint resmi BEI/IDX.
2. Cadangan universe: TradingView, StockAnalysis, lalu `Ringkasan.xlsx`.
3. Harga, volume, OHLC, dan histori: `yfinance`.
4. Fundamental massal: TradingView scanner bila tersedia.
5. Cadangan histori: `pandas-datareader`, cache lokal, lalu sheet `Metrik`.
6. Excel fallback: rasio, metrik bank, market cap, revenue, hierarki industri, kode khusus, dan ide algoritme yang belum punya sumber online stabil.

Tab `Kualitas Data` menampilkan audit sumber kode agar perbedaan antara daftar
resmi BEI/IDX, sumber online pelengkap, dan fallback Excel tetap terlihat,
bukan diabaikan.

## Market Regime & Freshness

Dashboard menambahkan konteks pasar besar agar saham bagus tidak dibaca lepas
dari kondisi IHSG.

- `Market_Regime`: `Risk-On`, `Neutral`, atau `Risk-Off` dari posisi IHSG terhadap MA50/MA200 dan return 20D/60D.
- `Market_Breadth`: kondisi breadth dari persentase saham sample yang berada di atas MA50/MA200.
- `Freshness`: status data online seperti `Fresh`, `Stale`, atau `Needs Refresh`.
- Jika market melemah, `Action_Checklist` menambahkan konteks seperti `market risk-off`, tanpa mengubah `Score` fundamental.

Tab `Kualitas Data` menampilkan audit market/freshness agar sumber, tanggal,
lag data, dan fallback tetap transparan.

## Harga & Teknikal

Tab `Harga & Teknikal` menggabungkan histori harga dan analisa teknikal karena
keduanya memakai OHLCV online/cache dari yfinance.

Bagian histori menyediakan dua mode:

- `Online yfinance KODE.JK`: mengambil harga historis IDX secara online dengan format ticker seperti `BBCA.JK`, `BMRI.JK`, atau `TLKM.JK`.
- `Excel Metrik 4W-52W`: memakai return historis dari sheet `Metrik` hanya sebagai mode pembanding/cadangan.

Mode online mendukung rentang 1 minggu, 2 minggu, 1/3/6 bulan, 1/2/5/10 tahun,
dan all/sepanjang masa. Jika koneksi live gagal, dashboard mencoba fallback
`pandas-datareader` dan cache lokal di folder `history_cache/`.

Bagian teknikal membaca timing dari OHLCV online/cache, bukan untuk mengganti
score fundamental utama.

- `Technical_Score`: trend MA20/50/200, RSI, MACD, volume ratio, dan ATR.
- `Entry_Action`: rekomendasi untuk calon pembelian, misalnya `Buy Candidate`, `Wait Confirmation`, `Wait Pullback`, atau `Avoid Entry`.
- `Position_Action`: rekomendasi untuk saham yang sudah dimiliki, misalnya `Hold`, `Take Profit`, `Reduce`, atau `Exit / Sell`.
- `Exit_Risk`: risiko keluar/pengetatan posisi dari kombinasi fundamental dan teknikal.
- `ATR_Stop_2x`: zona risiko teknikal berbasis dua kali ATR dari harga terakhir, bukan instruksi order otomatis.

Dashboard tidak memakai harga beli pribadi, sehingga `Position_Action` adalah
arah umum berbasis kondisi saham terbaru.

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
- Relatif sektor: PER/PBV dan ROE/ROA/NPM dibanding saham lain dalam sektor yang sama.
- Konteks ukuran: market cap, revenue, dan MCap/Revenue ditampilkan untuk analisis, tetapi belum menjadi faktor utama score.
- Threshold: batas rasio dari sheet `NonBank` atau `Banking`.

Layer teknikal terpisah dari score fundamental:

- Technical Score memakai trend MA20/50/200, RSI, MACD, volume ratio, dan ATR.
- Entry Action memakai fundamental sebagai gerbang awal, lalu teknikal untuk timing pembelian.
- Position Action memberi arahan umum hold/reduce/take profit/exit untuk saham yang sudah dimiliki tanpa memakai harga beli pribadi.
- ATR Stop 2x membantu membaca zona risiko teknikal.

Layer explainability:

- `Decision_Summary` merangkum rekomendasi, posisi relatif sektor, dan risiko.
- `Top_Strengths` menunjukkan faktor skor paling kuat.
- `Top_Risks` menunjukkan faktor skor paling lemah.
- `Action_Checklist` menunjukkan hal yang perlu dikonfirmasi sebelum entry atau mempertahankan posisi.

Penalti diterapkan untuk data yang kurang sehat, seperti PER/PBV negatif,
profitabilitas negatif, volume rendah, harga nol, perubahan harian ekstrem,
dan kelulusan threshold yang terlalu rendah.

## Catatan

Dashboard ini adalah alat screening kuantitatif awal, bukan saran investasi.
Tetap cek laporan keuangan terbaru, aksi korporasi, berita material, manajemen,
dan risiko portofolio sebelum mengambil keputusan.
