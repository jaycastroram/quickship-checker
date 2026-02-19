import pandas as pd
from pathlib import Path

# -----------------------------
# Config – update column names if needed
# -----------------------------
INPUT_CSV = r"/mnt/data/All Takeoff Submissions.csv"
OUTPUT_CSV = "takeoff_expedited_5wk_check.csv"

REQUEST_COL = "Request Date"
DELIVERY_COL = "Delivery Date"
STORE_COL = "Store Number"

PROJECT_COL = "Project Code"
TAKEOFF_TYPE_COL = "Take Off Type"
NOTES_COL = "Notes"

VALID_PROJECTS = {"Remodel_26", "SpringNewStore_26"}

HARD_RULE_DAYS = 35  # 5 weeks


def parse_date_series(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)

    # Fallback for Zoho-style dates like 13-Jan-26
    if parsed.isna().mean() > 0.5:
        parsed = pd.to_datetime(series, errors="coerce", format="%d-%b-%y")

    return parsed


def main():
    df = pd.read_csv(INPUT_CSV)

    # Validate required columns
    required_cols = {
        REQUEST_COL,
        DELIVERY_COL,
        PROJECT_COL
    }

    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Parse dates
    df[REQUEST_COL] = parse_date_series(df[REQUEST_COL])
    df[DELIVERY_COL] = parse_date_series(df[DELIVERY_COL])

    # Calculate difference
    df["Days_Between_Request_and_Delivery"] = (
        df[DELIVERY_COL] - df[REQUEST_COL]
    ).dt.days

    # Apply hard 5-week rule
    df["Expedited_5wk_HardRule"] = df["Days_Between_Request_and_Delivery"] < HARD_RULE_DAYS

    # Cutoff date for auditing
    df["FiveWeek_Cutoff_Date"] = (
        df[DELIVERY_COL] - pd.to_timedelta(HARD_RULE_DAYS, unit="D")
    )

    # Reason column
    def reason(row):
        if pd.isna(row[REQUEST_COL]) or pd.isna(row[DELIVERY_COL]):
            return "Missing date(s)"
        if row["Days_Between_Request_and_Delivery"] < 0:
            return "Delivery before request"
        if row["Days_Between_Request_and_Delivery"] < HARD_RULE_DAYS:
            return "Takeoff late (<5 weeks)"
        return ""

    df["Expedited_Reason"] = df.apply(reason, axis=1)

    # Focus only on valid projects
    df_filtered = df[df[PROJECT_COL].isin(VALID_PROJECTS)].copy()

    # Order columns for review
    output_columns = [
        STORE_COL,
        PROJECT_COL,
        TAKEOFF_TYPE_COL,
        REQUEST_COL,
        DELIVERY_COL,
        "Days_Between_Request_and_Delivery",
        "FiveWeek_Cutoff_Date",
        "Expedited_5wk_HardRule",
        "Expedited_Reason",
        NOTES_COL
    ]

    output_columns = [c for c in output_columns if c in df_filtered.columns]

    df_filtered = df_filtered[output_columns]

    # Save output
    df_filtered.to_csv(OUTPUT_CSV, index=False)

    # Console summary
    print(f"Rows evaluated (Remodel_26 / SpringNewStore_26): {len(df_filtered)}")
    print(f"Flagged expedited (<5 weeks): {df_filtered['Expedited_5wk_HardRule'].sum()}")

    # Optional: show flagged rows
    flagged = df_filtered[df_filtered["Expedited_5wk_HardRule"]]
    if not flagged.empty:
        print("\nFlagged orders:")
        print(flagged.to_string(index=False))

    print(f"\nSaved output to: {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
