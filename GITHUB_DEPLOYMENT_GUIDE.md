# Panduan Deploy Streamlit

Panduan singkat untuk upload proyek dashboard saham ini ke GitHub dan deploy ke
Streamlit Cloud.

## File yang Perlu Di-upload

Pastikan file berikut ada di repository:

```text
streamlit_app.py
requirements.txt
Ringkasan.xlsx
data_cache/market_snapshot_1y.csv
data_cache/fundamental_snapshot.csv
history_cache/*.csv
README.md
.gitignore
```

`Ringkasan.xlsx` berfungsi sebagai fallback. Aplikasi mengambil universe kode
saham dari endpoint resmi BEI/IDX lebih dulu, mengambil harga/histori dari
snapshot repo `data_cache/market_snapshot_1y.csv` agar startup cepat. Cache
histori `history_cache/*.csv` boleh ikut repository agar grafik histori tidak
selalu fetch ulang di Streamlit Cloud. Refresh online dari yfinance dilakukan
manual/terkontrol. Fundamental massal dibaca dari snapshot repo
`data_cache/fundamental_snapshot.csv` bila tersedia, lalu bisa di-refresh dari
TradingView scanner. Excel mengisi rasio/metrik yang masih kosong, metrik bank,
sektor, dan cadangan data yang belum tersedia online. Sheet `Metrik` juga
dipakai untuk market cap, revenue, subsektor, industri, subindustri, dan daftar
indeks.

File yang tidak perlu di-upload:

```text
__pycache__/
*.pyc
streamlit_stdout.log
streamlit_stderr.log
.venv/
```

## Push ke GitHub

```powershell
git init
git remote add origin https://github.com/tjutee/saham.git
git add .
git commit -m "Dashboard rekomendasi saham Streamlit"
git branch -M main
git push -u origin main
```

## Deploy ke Streamlit Cloud

1. Buka `https://streamlit.io/cloud`.
2. Login dengan GitHub.
3. Klik `Create app` atau `New app`.
4. Pilih repository `tjutee/saham`.
5. Set `Main file path` ke `streamlit_app.py`.
6. Klik `Deploy`.

## Jika Data Online Tidak Terbaca

Pastikan koneksi internet Streamlit Cloud aktif dan dependency di
`requirements.txt` berhasil ter-install. Dashboard akan mencoba:

1. daftar kode saham online dari endpoint resmi BEI/IDX, lalu fallback TradingView/StockAnalysis,
2. snapshot repo `data_cache/market_snapshot_1y.csv` untuk startup cepat,
3. snapshot fundamental repo `data_cache/fundamental_snapshot.csv`,
4. cache histori repo `history_cache/*.csv`,
5. refresh harga/histori dari `yfinance` bila cache belum tersedia atau dipaksa live,
6. fundamental massal dari TradingView scanner bila snapshot belum tersedia atau dipaksa live,
7. fallback `pandas-datareader`,
8. cache lokal,
9. `Ringkasan.xlsx` untuk kolom yang masih kosong.

Grafik utama memakai harga/volume yfinance atau cache online lebih dulu. Excel
hanya menjadi fallback bila online/cache tidak tersedia, atau mode pembanding
manual pada tab `Harga & Teknikal`.

Setelah deploy, lakukan smoke test cepat:

1. Cek tab `Ringkasan` untuk memastikan sumber harga, status kode, market
   regime, distribusi rekomendasi, dan distribusi aksi akhir terbaca wajar.
2. Cek tab `Rekomendasi` untuk memastikan kolom `Aksi Akhir`, `Keyakinan`,
   `Penghambat`, dan `Langkah Berikutnya` muncul.
3. Cek tab `Portofolio` dengan beberapa saham default untuk memastikan estimasi
   alokasi, lot, risk mix, dan konsentrasi sektor muncul tanpa error.
4. Cek tab `Harga & Teknikal` untuk memastikan histori yfinance/cache dan
   indikator teknikal tampil.
5. Cek tab `Kualitas Data` -> `Audit sumber kode saham` dan `Kelengkapan kolom
   & sumber data`.
6. Cek `Kualitas Data` -> `Jumlah kode per sheet Excel`: `Ringkasan`/`Draft`
   dapat berisi 959 kode unik, sedangkan `Metrik` dapat berisi 957 karena
   cakupannya berbeda.

Di tab `Kualitas Data` akan terlihat jumlah kode yang match daftar resmi
BEI/IDX, kode yang hanya berasal dari fallback, serta kolom fundamental,
histori, market cap, revenue, dan hierarki industri yang masih kosong.

Jika fallback Excel ingin dipakai, pastikan `Ringkasan.xlsx` ada di root
repository, satu folder dengan `streamlit_app.py`. Nama file harus sama persis
karena app membaca:

```python
DATA_FILE = "Ringkasan.xlsx"
```

## Update Data Fallback

Harga dan histori utama akan di-refresh dari online/cache aplikasi. Jika ingin
memperbaiki rasio fundamental, metrik bank, sektor, atau cadangan Excel-only,
timpa file `Ringkasan.xlsx`, lalu push ulang:

```powershell
git add Ringkasan.xlsx
git commit -m "Update data Ringkasan"
git push origin main
```

Streamlit Cloud akan otomatis redeploy setelah push.

## Update Snapshot Pasar

Untuk mengurangi lag saat jam bursa, simpan snapshot pasar ringkas ke repository:

1. Jalankan aplikasi lokal.
2. Buka sidebar `Workflow Update`.
3. Jalankan `Refresh cache histori top saham` bila perlu.
4. Klik `Bangun snapshot pasar dari cache`.
5. Commit dan push file `data_cache/market_snapshot_1y.csv` dan `data_cache/fundamental_snapshot.csv`.

Jika ingin memaksa aplikasi mengambil data live penuh saat startup, set
environment variable:

```text
SAHAM_LIVE_ON_START=1
```

Untuk memaksa refresh fundamental live saat startup:

```text
SAHAM_FUNDAMENTAL_LIVE_ON_START=1
```

Default keduanya `0`, yaitu cache-first agar Streamlit Cloud lebih cepat dan
stabil.

Catatan realtime: `yfinance`/TradingView adalah sumber online praktis dan dapat
delayed. Jika membutuhkan feed bursa resmi realtime tick-by-tick, gunakan
provider data pasar/IDX feed berlisensi dan sambungkan sebagai adapter data baru.
