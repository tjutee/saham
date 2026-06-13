# Dashboard Rekomendasi Saham IDX

Dashboard Streamlit interaktif untuk screening dan rekomendasi awal saham Indonesia
dengan alur data online-first. Universe kode saham diprioritaskan dari daftar
resmi BEI/IDX, harga dan histori diprioritaskan dari `yfinance`, dan
fundamental online dilengkapi dari TradingView scanner. `Ringkasan.xlsx`
dipakai sebagai fallback, pembanding, dan sumber ide algoritme untuk data yang
belum stabil tersedia online.

Dashboard memakai pola cache-first untuk mengurangi lag saat jam bursa:
startup membaca snapshot pasar dari `data_cache/market_snapshot_1y.csv` dan
snapshot fundamental dari `data_cache/fundamental_snapshot.csv` bila tersedia,
lalu refresh online dilakukan manual/terkontrol.

## Fitur

- Universe kode saham online dari endpoint resmi BEI/IDX, dengan fallback TradingView, StockAnalysis, dan Excel.
- Audit sumber kode menandai kode yang match BEI/IDX resmi dan kode yang hanya tersedia dari fallback.
- Fundamental online massal dari TradingView scanner untuk PER, PBV, ROE, ROA, DER, NPM, market cap, revenue, sektor, dan industri bila tersedia.
- Deduplikasi saham berdasarkan `Kode`, karena saham yang sama bisa muncul di beberapa indeks.
- Scoring multi-factor: valuasi, kualitas profit, risiko, likuiditas, momentum, dan kekuatan indeks.
- Sector-relative scoring untuk membandingkan valuasi dan kualitas saham terhadap sektor yang sama.
- Explainability layer: `Decision_Summary`, `Top_Strengths`, `Top_Risks`, dan `Action_Checklist`.
- Final decision layer: `Final_Action`, `Decision_Confidence`, `Decision_Blockers`, dan `Next_Step` sebagai playbook praktis setelah score, risiko, sektor, data, momentum, dan market regime digabung.
- Market regime dari IHSG `^JKSE`: Risk-On, Neutral, atau Risk-Off berdasarkan MA50/MA200 dan momentum 20D/60D.
- Market breadth: persentase saham sample yang berada di atas MA50 dan MA200.
- Data freshness guard untuk melihat tanggal online terakhir, lag data, coverage online, dan fallback Excel.
- Snapshot pasar repo untuk harga, volume, dan return 1 tahun agar Streamlit Cloud tidak selalu mengambil semua ticker saat cold start.
- Snapshot fundamental repo untuk rasio massal TradingView agar startup tidak selalu menunggu request online besar.
- Perencana portofolio untuk membaca alokasi kandidat, konsentrasi sektor, risk mix, final action mix, dan estimasi lot.
- Filter threshold dari sheet `NonBank` dan `Banking`.
- Backtest event-based untuk menguji sinyal teknikal historis terhadap return 5/20/60 hari.
- Histori return 4, 13, 26, 52 minggu, dan YTD dihitung dari yfinance bila tersedia.
- Grafik utama memprioritaskan harga/volume yfinance atau cache online; Excel hanya dipakai bila online/cache kosong.
- Sheet `Metrik` dimanfaatkan untuk market cap, revenue, subsektor, industri, subindustri, dan daftar indeks gabungan.
- Grafik histori online fleksibel untuk saham IDX dengan ticker `KODE.JK`.
- Analisa teknikal berbasis OHLCV yfinance/cache: candlestick, MA20/50/200, RSI, MACD, ATR, Fibonacci Confluence, Astro-Fibo Timing, Technical Score, Entry Action, Position Action, dan ATR stop zone.
- Sumber harga/histori online utama: `yfinance`, fallback: `pandas-datareader`, cache lokal, lalu Excel bila tersedia.
- Profil scoring: Balanced, Defensive, Growth, dan Value.
- Bobot scoring bisa diatur langsung dari sidebar.
- Dashboard dinamis dengan tab Ringkasan, Rekomendasi, Portofolio, Harga & Teknikal, Validasi & Prediksi, Explorer & Sektor, dan Data & Metodologi.
- Ringkasan eksekutif berisi distribusi rekomendasi, komposisi risiko, sumber data, top kandidat, dan heatmap faktor.
- Data & Metodologi menampilkan audit sumber kode, audit filter, coverage kolom, serta campuran sumber harga/volume.
- Tabel hasil bisa di-download sebagai CSV.

## File Utama

```text
streamlit_app.py      # Aplikasi utama Streamlit
Ringkasan.xlsx        # Fallback rasio/fundamental dan cadangan data
data_cache/           # Snapshot pasar ringkas yang boleh di-commit
history_cache/*.csv   # Cache histori OHLCV yang boleh di-commit bila ingin histori tersedia di cloud
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
3. Snapshot pasar repo: `data_cache/market_snapshot_1y.csv` untuk startup cepat.
4. Snapshot fundamental repo: `data_cache/fundamental_snapshot.csv` untuk rasio massal.
5. Cache histori repo: CSV OHLCV di `history_cache/` untuk grafik/histori tanpa fetch ulang.
6. Refresh harga, volume, OHLC, dan histori: `yfinance`.
7. Fundamental massal live: TradingView scanner bila perlu refresh.
8. Cadangan histori: `pandas-datareader`, cache lokal, lalu sheet `Metrik`.
9. Excel fallback: rasio, metrik bank, market cap, revenue, hierarki industri, kode khusus, dan ide algoritme yang belum punya sumber online stabil.

Catatan realtime praktis: `yfinance`/TradingView adalah sumber online praktis
dan bisa delayed, bukan feed bursa resmi realtime tick-by-tick. Dashboard
menampilkan status jam bursa WIB, freshness, lag tanggal online, dan sumber
data. Saat bursa berjalan, refresh histori/harga dilakukan terkontrol untuk
saham prioritas lalu snapshot pasar dibangun dari cache. Saat jeda, akhir pekan,
atau pasca-penutupan, cache/snapshot menjaga dashboard tetap cepat dan stabil.
Untuk realtime resmi perlu provider data pasar/IDX feed berlisensi.

Catatan jumlah data Excel: sheet `Ringkasan`/`Draft` bisa berisi 959 kode unik,
sedangkan sheet `Metrik` dapat berisi 957 baris/kode karena cakupannya berbeda.
Dashboard menampilkan audit jumlah kode per sheet di tab `Data & Metodologi` agar
perbedaan ini terlihat, bukan disamarkan.

Tab `Data & Metodologi` menampilkan audit sumber kode agar perbedaan antara daftar
resmi BEI/IDX, sumber online pelengkap, dan fallback Excel tetap terlihat,
bukan diabaikan.

## Market Regime & Freshness

Dashboard menambahkan konteks pasar besar agar saham bagus tidak dibaca lepas
dari kondisi IHSG.

- `Market_Regime`: `Risk-On`, `Neutral`, atau `Risk-Off` dari posisi IHSG terhadap MA50/MA200 dan return 20D/60D.
- `Market_Breadth`: kondisi breadth dari persentase saham sample yang berada di atas MA50/MA200.
- `Freshness`: status data online seperti `Fresh`, `Stale`, atau `Needs Refresh`.
- Jika market melemah, `Action_Checklist` menambahkan konteks seperti `market risk-off`, tanpa mengubah `Score` fundamental.

Tab `Data & Metodologi` menampilkan audit market/freshness agar sumber, tanggal,
lag data, dan fallback tetap transparan.

## Backtest

Tab `Validasi & Prediksi` menguji sinyal historis dari OHLCV online/cache. Backtest awal
ini event-based, bukan simulasi broker penuh.

- Sinyal yang bisa diuji: `Bullish`, `Constructive`, `Weak`, `Overbought`, `MA50 Recovery`, dan `MA50 Breakdown`.
- Metrik hasil: event count, hit rate 5/20/60 hari, average return, median return, dan max drawdown setelah event.
- Evidence label: `Strong evidence`, `Mixed positive`, `Low sample`, atau `Weak evidence`.
- Walk-forward validation membagi event secara berurutan untuk membandingkan performa train dan out-of-sample.
- Hasil fundamental saat ini ditampilkan sebagai konteks, tetapi backtest belum memakai fundamental historis point-in-time.

## Prediksi

Area prediksi pada tab `Validasi & Prediksi` memakai setup historis yang mirip dengan kondisi teknikal saat
ini untuk menghitung probabilitas statistik.

- Output: `Prediction_Bias`, `Probability_Up_20D/60D`, `Expected_Return`, `Downside_Risk`, `Model_Confidence`, dan sample historis.
- Similarity memakai `Technical_Signal`, `Fibo_Zone`, `Technical_Score`, `RSI`, jarak harga ke level Fibonacci, dan Astro-Fibo timing bila sampel historis cukup.
- Periode prediksi memakai opsi yang sama dengan Histori/Harga & Teknikal: 1 minggu, 2 minggu, 1/3/6 bulan, 1/2/5/10 tahun, dan all/sepanjang masa.
- Default 2 tahun dipakai sebagai titik seimbang; periode pendek lebih cepat tetapi sering minim sample, sedangkan 5-10 tahun/All lebih kaya sample tetapi bisa mencampur rezim pasar lama.
- Ini baseline probabilistik berbasis histori, bukan prediksi harga pasti dan belum memakai model ML berat seperti LightGBM.
- Gunakan bersama fundamental, market regime, backtest, dan trade plan.

## Portofolio

Tab `Portofolio` mengubah kandidat hasil filter menjadi simulasi alokasi.

- Pilih saham dari kandidat `Final_Action` yang paling layak, lalu tentukan modal, metode alokasi, batas per saham, dan batas konsentrasi sektor.
- Metode alokasi: equal weight, score-weighted, atau berbasis aksi akhir.
- Output: nilai teralokasi, sisa kas, posisi terbesar, jumlah sektor, jumlah saham high risk, alokasi sektor, alokasi menurut aksi akhir, estimasi lot, dan next step.
- Simulasi ini tidak mengirim order dan belum memperhitungkan fee, slippage, gap harga, pajak, atau tujuan pribadi.

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
- `Trade Plan & Position Sizing`: estimasi lot dari modal, risiko per transaksi, batas posisi maksimum, harga terakhir, dan stop plan. Hasil belum memperhitungkan fee, slippage, atau gap harga.
- `Fibonacci Confluence`: level retracement/extension dari swing high-low periode teknikal untuk membaca support/resistance dan nearest level. Ini bukan prediksi pasti.
- `Astro-Fibo Timing`: layer terinspirasi Astronacci yang menggabungkan jendela waktu Fibonacci dari swing terakhir, fase bulan sederhana, siklus Sun/zodiak musiman, Fibonacci price zone, dan konfirmasi teknikal. Ini bukan formula proprietary Astronacci dan bukan ramalan pasti.

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
- Position sizing menghitung lot estimasi dengan risk budget / risk per share dan dibatasi maksimum nilai posisi.
- Fibonacci Confluence memakai level 23.6%, 38.2%, 50%, 61.8%, 78.6%, 127.2%, dan 161.8% sebagai area support/resistance yang perlu dikonfirmasi oleh trend, RSI/MACD, volume, dan backtest.
- Astro-Fibo Timing memakai hitungan hari Fibonacci 5, 8, 13, 21, 34, 55, 89, dan 144 dari swing terakhir, fase bulan sederhana, siklus Sun/zodiak musiman, Fibo score, dan technical score untuk membaca jendela timing yang perlu dikonfirmasi.

Layer explainability dan final decision:

- `Decision_Summary` merangkum rekomendasi, posisi relatif sektor, dan risiko.
- `Top_Strengths` menunjukkan faktor skor paling kuat.
- `Top_Risks` menunjukkan faktor skor paling lemah.
- `Action_Checklist` menunjukkan hal yang perlu dikonfirmasi sebelum entry atau mempertahankan posisi.
- `Final_Action` menerjemahkan hasil screening menjadi playbook: `Accumulate Candidate`, `Wait Market Confirmation`, `Watchlist`, `Speculative Monitor`, atau `Avoid / Review`.
- `Decision_Confidence` membaca seberapa bersih sinyal akhir setelah blocker seperti data belum bersih, risiko tinggi, threshold rendah, lemah relatif sektor, atau market risk-off.
- `Next_Step` menunjukkan langkah berikutnya, misalnya cek Harga & Teknikal, tunggu konfirmasi market, atau review ulang data.

Layer portofolio:

- Alokasi memakai harga terakhir dari yfinance/cache bila tersedia, dengan Excel sebagai fallback harga.
- Batas per saham membuat sebagian modal bisa tetap menjadi kas bila kandidat terlalu sedikit atau harga per lot terlalu besar.
- Konsentrasi sektor ditandai bila melewati batas yang dipilih pengguna.

Penalti diterapkan untuk data yang kurang sehat, seperti PER/PBV negatif,
profitabilitas negatif, volume rendah, harga nol, perubahan harian ekstrem,
dan kelulusan threshold yang terlalu rendah.

## Catatan

Dashboard ini adalah alat screening kuantitatif awal, bukan saran investasi.
Tetap cek laporan keuangan terbaru, aksi korporasi, berita material, manajemen,
dan risiko portofolio sebelum mengambil keputusan.
