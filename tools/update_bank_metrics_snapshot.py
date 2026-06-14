from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "Ringkasan.xlsx"
OUTPUT_FILE = ROOT / "data_cache" / "bank_metrics_snapshot.csv"
BANK_METRICS = ["NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]


def normalize_metric_name(column):
    text = str(column).upper()
    for metric in BANK_METRICS:
        if metric in text:
            return metric
    return None


def read_metric_sheet(sheet_name):
    frame = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
    code_column = "Kode" if "Kode" in frame.columns else "Kode Saham" if "Kode Saham" in frame.columns else None
    if code_column is None:
        return pd.DataFrame(columns=["Kode"])

    output = pd.DataFrame({"Kode": frame[code_column].astype(str).str.strip().str.upper()})
    for column in frame.columns:
        metric = normalize_metric_name(column)
        if metric:
            output[metric] = pd.to_numeric(frame[column], errors="coerce")

    output["Bank_Metric_Source"] = f"Excel {sheet_name} fallback"
    output["Bank_Metric_Last_Update"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    return output


def main():
    rows = []
    for sheet_name in ["Banking", "NonBank"]:
        rows.append(read_metric_sheet(sheet_name))

    combined = pd.concat(rows, ignore_index=True)
    combined = combined[combined["Kode"].notna() & combined["Kode"].str.match(r"^[A-Z0-9]{4,}$", na=False)]
    keep_columns = ["Kode"] + BANK_METRICS + ["Bank_Metric_Source", "Bank_Metric_Last_Update"]
    combined = combined[[column for column in keep_columns if column in combined.columns]]
    metric_columns = [column for column in BANK_METRICS if column in combined.columns]
    combined = combined.dropna(subset=metric_columns, how="all")
    combined = combined.sort_values(["Kode", "Bank_Metric_Source"]).drop_duplicates("Kode", keep="first")

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    combined.to_csv(OUTPUT_FILE, index=False)
    print(f"Wrote {len(combined):,} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
