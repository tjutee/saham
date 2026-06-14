import argparse
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "Ringkasan.xlsx"
OUTPUT_FILE = ROOT / "data_cache" / "bank_metrics_snapshot.csv"
DEFAULT_SOURCE_DIR = ROOT / "data_sources" / "bank_metrics"
BANK_METRICS = ["NIM", "CAR", "LDR", "NPL", "BOPO", "CIR", "LAR"]
SNAPSHOT_COLUMNS = [
    "Kode",
    *BANK_METRICS,
    "Period",
    "Bank_Metric_Source",
    "Bank_Metric_Source_URL",
    "Bank_Metric_Last_Update",
    "Bank_Metric_Confidence",
    "Bank_Metric_Notes",
]
METRIC_RANGES = {
    "NIM": (0, 100),
    "CAR": (0, 250),
    "LDR": (0, 250),
    "NPL": (0, 100),
    "BOPO": (0, 250),
    "CIR": (0, 250),
    "LAR": (0, 100),
}
CODE_ALIASES = ["Kode", "Kode Saham", "Ticker", "Symbol", "Emiten", "Kode Emiten", "Stock Code"]
PERIOD_ALIASES = ["Period", "Periode", "Tanggal", "Date", "Report Period", "Periode Laporan"]
SOURCE_URL_ALIASES = ["Bank_Metric_Source_URL", "Source_URL", "URL", "Source URL", "Link"]


def normalize_metric_name(column):
    text = str(column).upper()
    for metric in BANK_METRICS:
        if metric in text:
            return metric
    return None


def first_existing_column(columns, aliases):
    normalized = {str(column).strip().lower(): column for column in columns}
    for alias in aliases:
        match = normalized.get(alias.lower())
        if match is not None:
            return match
    return None


def source_confidence_for(path):
    text = path.name.lower()
    if "ojk" in text:
        return "Regulatory"
    if "idx" in text or "bei" in text:
        return "Exchange"
    return "Imported"


def source_label_for(path):
    confidence = source_confidence_for(path)
    if confidence == "Regulatory":
        return "OJK file import"
    if confidence == "Exchange":
        return "IDX/BEI file import"
    return "Bank metric file import"


def read_table_file(path):
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return [pd.read_csv(path)]
    if suffix in {".xlsx", ".xls"}:
        sheets = pd.read_excel(path, sheet_name=None)
        return list(sheets.values())
    return []


def read_external_metric_file(path):
    frames = []
    for frame in read_table_file(path):
        if frame.empty:
            continue
        code_column = first_existing_column(frame.columns, CODE_ALIASES)
        if code_column is None:
            continue
        output = pd.DataFrame({"Kode": frame[code_column].astype(str).str.strip().str.upper()})
        for column in frame.columns:
            metric = normalize_metric_name(column)
            if metric:
                output[metric] = pd.to_numeric(frame[column], errors="coerce")

        period_column = first_existing_column(frame.columns, PERIOD_ALIASES)
        source_url_column = first_existing_column(frame.columns, SOURCE_URL_ALIASES)
        output["Period"] = frame[period_column].astype(str) if period_column is not None else path.stem
        output["Bank_Metric_Source"] = source_label_for(path)
        output["Bank_Metric_Source_URL"] = frame[source_url_column].astype(str) if source_url_column is not None else str(path.relative_to(ROOT))
        output["Bank_Metric_Last_Update"] = pd.Timestamp.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        output["Bank_Metric_Confidence"] = source_confidence_for(path)
        output["Bank_Metric_Notes"] = f"Imported from {path.name}."
        frames.append(output)
    return frames


def read_external_metric_sources(source_dir):
    if not source_dir.exists():
        return []
    frames = []
    for path in sorted(source_dir.glob("*")):
        if not path.is_file() or path.suffix.lower() not in {".csv", ".xlsx", ".xls"}:
            continue
        frames.extend(read_external_metric_file(path))
    return frames


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
    output["Bank_Metric_Source_URL"] = str(DATA_FILE.name)
    output["Period"] = "Latest workbook snapshot"
    output["Bank_Metric_Last_Update"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    output["Bank_Metric_Confidence"] = "Fallback"
    output["Bank_Metric_Notes"] = "Extracted from Ringkasan.xlsx; replace with OJK/IDX source when available."
    return output


def priority_for_source(source):
    text = str(source).lower()
    if "ojk" in text:
        return 0
    if "idx" in text or "bei" in text:
        return 1
    if "import" in text:
        return 2
    return 9


def coalesce_group(group):
    ordered = group.sort_values("Source_Priority")
    row = {"Kode": ordered["Kode"].iloc[0]}
    for metric in BANK_METRICS:
        row[metric] = pd.to_numeric(ordered.get(metric), errors="coerce").dropna().iloc[0] if metric in ordered and pd.to_numeric(ordered.get(metric), errors="coerce").notna().any() else pd.NA
    metadata_source = ordered.iloc[0]
    for column in [
        "Period",
        "Bank_Metric_Source",
        "Bank_Metric_Source_URL",
        "Bank_Metric_Last_Update",
        "Bank_Metric_Confidence",
        "Bank_Metric_Notes",
    ]:
        row[column] = metadata_source.get(column, pd.NA)
    return pd.Series(row)


def validate_ranges(frame):
    output = frame.copy()
    notes = output.get("Bank_Metric_Notes", pd.Series("", index=output.index)).fillna("").astype(str)
    for metric, (lower, upper) in METRIC_RANGES.items():
        if metric not in output.columns:
            continue
        values = pd.to_numeric(output[metric], errors="coerce")
        invalid = values.notna() & ~values.between(lower, upper, inclusive="both")
        output.loc[invalid, metric] = pd.NA
        notes = notes.where(~invalid, notes + f" {metric} outside expected range removed.")
    output["Bank_Metric_Notes"] = notes.str.strip()
    return output


def build_snapshot(source_dir=DEFAULT_SOURCE_DIR, include_excel_fallback=True):
    rows = []
    rows.extend(read_external_metric_sources(source_dir))
    if include_excel_fallback:
        for sheet_name in ["Banking", "NonBank"]:
            rows.append(read_metric_sheet(sheet_name))
    if not rows:
        return pd.DataFrame(columns=SNAPSHOT_COLUMNS)

    combined = pd.concat(rows, ignore_index=True)
    combined = combined[combined["Kode"].notna() & combined["Kode"].str.match(r"^[A-Z0-9]{4,}$", na=False)]
    combined = validate_ranges(combined)
    combined["Source_Priority"] = combined["Bank_Metric_Source"].map(priority_for_source)
    keep_columns = SNAPSHOT_COLUMNS
    combined = combined[[column for column in keep_columns + ["Source_Priority"] if column in combined.columns]]
    metric_columns = [column for column in BANK_METRICS if column in combined.columns]
    combined = combined.dropna(subset=metric_columns, how="all")
    combined = pd.DataFrame([coalesce_group(group) for _, group in combined.groupby("Kode", sort=False)])
    for column in SNAPSHOT_COLUMNS:
        if column not in combined.columns:
            combined[column] = pd.NA
    return combined[SNAPSHOT_COLUMNS].sort_values("Kode")


def main():
    parser = argparse.ArgumentParser(description="Build bank metrics snapshot from OJK/IDX/imported files with Excel fallback.")
    parser.add_argument("--source-dir", default=str(DEFAULT_SOURCE_DIR), help="Folder containing OJK/IDX CSV/XLSX files to import.")
    parser.add_argument("--no-excel-fallback", action="store_true", help="Do not fill from Ringkasan.xlsx fallback sheets.")
    parser.add_argument("--allow-empty", action="store_true", help="Allow writing an empty snapshot.")
    args = parser.parse_args()

    combined = build_snapshot(Path(args.source_dir), include_excel_fallback=not args.no_excel_fallback)
    if combined.empty and OUTPUT_FILE.exists() and not args.allow_empty:
        print(f"No rows found; keeping existing snapshot at {OUTPUT_FILE}. Use --allow-empty to overwrite.")
        return

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    combined.to_csv(OUTPUT_FILE, index=False)
    print(f"Wrote {len(combined):,} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
