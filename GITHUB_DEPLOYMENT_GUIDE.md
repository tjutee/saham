# Panduan Deploy Streamlit

Panduan singkat untuk upload proyek dashboard saham ini ke GitHub dan deploy ke
Streamlit Cloud.

## File yang Perlu Di-upload

Pastikan file berikut ada di repository:

```text
streamlit_app.py
requirements.txt
Ringkasan.xlsx
README.md
.gitignore
```

`Ringkasan.xlsx` berfungsi sebagai fallback. Aplikasi mengambil universe kode
saham dari endpoint resmi BEI/IDX lebih dulu, mengambil harga/histori dari
online, melengkapi fundamental massal dari TradingView scanner, lalu memakai
Excel untuk rasio/metrik yang masih kosong, metrik bank, sektor, dan cadangan
data yang belum tersedia online. Sheet `Metrik` juga dipakai untuk market cap,
revenue, subsektor, industri, subindustri, dan daftar indeks.

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
2. harga/histori dari `yfinance`,
3. fundamental massal dari TradingView scanner,
4. fallback `pandas-datareader`,
5. cache lokal,
6. `Ringkasan.xlsx` untuk kolom yang masih kosong.

Grafik utama memakai harga/volume yfinance atau cache online lebih dulu. Excel
hanya menjadi fallback bila online/cache tidak tersedia, atau mode pembanding
manual pada tab `Harga & Teknikal`.

Setelah deploy, cek tab `Ringkasan` untuk memastikan sumber harga, status kode,
dan distribusi rekomendasi terbaca wajar. Lalu cek tab `Kualitas Data` ->
`Audit sumber kode saham` dan `Kelengkapan kolom & sumber data`. Di sana akan
terlihat jumlah kode yang match daftar resmi BEI/IDX, kode yang hanya berasal
dari fallback, serta kolom fundamental, histori, market cap, revenue, dan
hierarki industri yang masih kosong.

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
