# Bank Metrics Connector

This connector keeps the dashboard fast by importing bank-specific metrics into a committed snapshot:

`data_cache/bank_metrics_snapshot.csv`

## Realistic Source Strategy

The dashboard should not scrape OJK/IDX pages on every Streamlit rerun. NIM, CAR, LDR, NPL, BOPO, CIR, and LAR are reporting-period metrics, not intraday market data.

Recommended flow:

1. Download or export OJK/IDX/bank publication files as CSV/XLSX.
2. Place them in `data_sources/bank_metrics/`.
3. Run:

```powershell
python tools/update_bank_metrics_snapshot.py
```

4. Review `data_cache/bank_metrics_snapshot.csv`.
5. Commit the snapshot.

If no official/imported file is available, the script falls back to `Ringkasan.xlsx` and labels the rows as `Fallback`.

## Scheduled Update

GitHub Actions workflow:

`.github/workflows/update-bank-metrics.yml`

Schedule:

- Weekly, Sunday 23:15 UTC.
- Manual run via `workflow_dispatch`.

The workflow rebuilds `data_cache/bank_metrics_snapshot.csv`, validates the schema, and commits only the snapshot when it changes.

This schedule is intentionally weekly because bank-specific ratios are reporting-period metrics. They should not be refreshed like intraday price data.

## Accepted Columns

Code columns can be named:

- `Kode`
- `Kode Saham`
- `Ticker`
- `Symbol`
- `Emiten`
- `Kode Emiten`
- `Stock Code`

Metric columns are detected when their header contains:

- `NIM`
- `CAR`
- `LDR`
- `NPL`
- `BOPO`
- `CIR`
- `LAR`

Optional metadata columns:

- `Period` / `Periode` / `Tanggal` / `Date` / `Report Period` / `Periode Laporan`
- `Source_URL` / `URL` / `Source URL`

## Confidence Labels

- Files with `ojk` in the filename are labeled `Regulatory`.
- Files with `idx` or `bei` in the filename are labeled `Exchange`.
- Other imported files are labeled `Imported`.
- Workbook fallback is labeled `Fallback`.

## Why Snapshot-Based

OJK exposes official banking statistics and publication pages, and IDX exposes listed-company reporting pages, but neither should be treated as a lightweight real-time field API for these bank-specific ratios. Snapshot import is more reliable, auditable, and much lighter for Streamlit.
