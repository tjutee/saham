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

## Jika Data Tidak Terbaca

Pastikan `Ringkasan.xlsx` ada di root repository, satu folder dengan
`streamlit_app.py`. Nama file harus sama persis karena app membaca:

```python
DATA_FILE = "Ringkasan.xlsx"
```

## Update Data

Jika ingin mengganti data, timpa file `Ringkasan.xlsx`, lalu push ulang:

```powershell
git add Ringkasan.xlsx
git commit -m "Update data Ringkasan"
git push origin main
```

Streamlit Cloud akan otomatis redeploy setelah push.
