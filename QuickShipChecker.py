
from __future__ import annotations

import argparse
import json
import math
import os
from dataclasses import dataclass
from typing import Optional, List, Dict, Any, Tuple, Union

import pandas as pd


# -----------------------------------------------------------------------------
# Configuration (tweak these to match your business rules)
# -----------------------------------------------------------------------------

# Approximate number of cases that fit on one pallet (rough estimate)
DEFAULT_CASES_PER_PALLET = 40

# Approximate number of SKUs that typically fit on one pallet (rough estimate)
DEFAULT_SKUS_PER_PALLET = 12

# If estimated pallets >= this number, recommend Work Order
DEFAULT_PALLET_THRESHOLD = 10

# Column fallbacks to find SKU and quantity
SKU_COLUMNS = [
    "Product #",
    "product #",
    "sku",
    "item",
    "item_id",
    "item_code",
    "ns_sku",
    "normalized_sku",
    "product #",
    "product",
    "product number",
]
QTY_COLUMNS = [
    "FSI Total Order Qty (in Packs)",
    "fsi total order qty (in packs)",
    "order_cases",
    "cases",
    "case_qty",
    "qty_cases",
    "quantity",
    "qty",
    "fsi total order qty (in packs)",
    "fsi total order qty (in cases)",
]

EACHES_COLUMNS = [
    "FSI Total Order Qty - Use for RELEASE (in Eaches)",
    "fsi total order qty - use for release (in eaches)",
    "Enter Order Qty",
    "enter order qty",
    "fsi total order qty (in eaches)",
    "eaches",
]

PACK_QTY_COLUMNS = [
    "Pack Qty",
    "pack qty",
    "pack_qty",
    "unit of measure (uom)",
    "uom",
]


@dataclass
class QuickShipDecision:
    work_order: bool
    reason: str
    estimated_pallets: float
    pt_pallets: int
    ap_pallets: int
    ap_fraction_sum: float
    total_cases: float
    total_skus: int
    total_lines: int
    missing_skus: int
    missing_uom_count: int
    missing_uom_samples: List[str]
    used_sku_db: bool = False


def _first_existing_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    col_set = set(columns)
    col_lower_to_orig = {c.lower(): c for c in columns}
    for candidate in candidates:
        if candidate in col_set:
            return candidate
        if candidate.lower() in col_lower_to_orig:
            return col_lower_to_orig[candidate.lower()]
    return None


def _normalize_name(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def _find_by_keywords(columns: List[str], keywords: List[str]) -> Optional[str]:
    """
    Find first column containing all keywords (case/whitespace-insensitive).
    """
    normalized = {col: _normalize_name(col) for col in columns}
    for col, norm in normalized.items():
        if all(keyword in norm for keyword in keywords):
            return col
    return None


def load_takeoff(path: str) -> pd.DataFrame:
    """Load takeoff file from CSV or Excel."""
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)
    return pd.read_csv(path)


def load_sku_database(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)


def estimate_pt_ap_pallets(
    df: pd.DataFrame,
    sku_column: str,
    cases_column: str,
    sku_db: Dict[str, Any],
    cases_per_pallet: int,
) -> Tuple[int, int, float, int]:
    """
    Estimate PT and AP pallets using SKU PT_QTY values.
    Returns (pt_pallets, ap_pallets, ap_fraction_sum, missing_skus).
    """
    grouped = df.groupby(sku_column, as_index=False)[cases_column].sum()
    pt_pallets = 0
    ap_fraction_sum = 0.0
    missing_skus = 0

    for _, row in grouped.iterrows():
        sku = str(row[sku_column]).strip()
        cases = float(row[cases_column])
        pt_qty_raw = sku_db.get(sku)

        if pt_qty_raw is None:
            missing_skus += 1
            ap_fraction_sum += cases / max(cases_per_pallet, 1)
            continue

        try:
            pt_qty = float(pt_qty_raw)
        except (TypeError, ValueError):
            pt_qty = 0

        if pt_qty <= 0:
            ap_fraction_sum += cases / max(cases_per_pallet, 1)
            continue

        pt_pallets += int(cases // pt_qty)
        remainder = cases - (pt_qty * int(cases // pt_qty))
        if remainder > 0:
            ap_fraction_sum += remainder / pt_qty

    ap_pallets = int(math.ceil(ap_fraction_sum)) if ap_fraction_sum > 0 else 0
    return pt_pallets, ap_pallets, ap_fraction_sum, missing_skus


def _normalize_sku_value(value: Any) -> str:
    return str(value).strip()


def _clean_sku_series(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.strip()
    cleaned = cleaned.replace(["", "nan", "none", "None"], pd.NA)
    return cleaned


def _parse_uom(uom_value: Any, display_name: Optional[str] = None) -> str:
    if uom_value is not None and pd.notna(uom_value):
        uom = str(uom_value).strip().lower()
        if uom:
            return uom
    if display_name:
        display = str(display_name).strip().lower()
        if display.endswith("-case") or " case" in display:
            return "case"
        if display.endswith("-ea") or " ea" in display or " each" in display:
            return "ea"
    return ""


def _sales_order_cases(quantity: float, uom: str) -> float:
    """
    Convert sales order quantity to cases.
    - case/cs: quantity already in cases
    - ea/each: assume 1 each per case (best-effort)
    - otherwise: treat as cases
    """
    uom = (uom or "").lower()
    if uom in {"case", "cases", "cs"}:
        return quantity
    if uom in {"ea", "each", "eaches"}:
        return quantity
    return quantity


def evaluate_sales_orders(
    orders_df: pd.DataFrame,
    cases_per_pallet: int = DEFAULT_CASES_PER_PALLET,
    pallet_threshold: int = DEFAULT_PALLET_THRESHOLD,
    sku_column: Optional[str] = None,
    qty_column: Optional[str] = None,
    uom_column: Optional[str] = None,
    display_name_column: Optional[str] = None,
    sku_prefix_to_strip: Optional[str] = "3172_",
    sku_db: Optional[Union[Dict[str, Any], str]] = None,
) -> QuickShipDecision:
    """
    Estimate PT/AP pallets from Sales Order exports.
    """
    df = orders_df.copy()

    # Detect columns
    if sku_column is None:
        for col in ["Item", "item", "sku", "SKU"]:
            if col in df.columns:
                sku_column = col
                break
    if qty_column is None:
        for col in ["Quantity", "quantity", "qty", "Qty"]:
            if col in df.columns:
                qty_column = col
                break
    if display_name_column is None:
        for col in ["Display Name", "display name", "Display", "display"]:
            if col in df.columns:
                display_name_column = col
                break

    if sku_column is None or qty_column is None:
        raise ValueError("Could not find required Sales Order columns (Item/Quantity).")

    # Normalize quantity
    df[qty_column] = pd.to_numeric(df[qty_column], errors="coerce").fillna(0)

    # Normalize SKU and strip prefix
    df[sku_column] = _clean_sku_series(df[sku_column])
    if sku_prefix_to_strip:
        df[sku_column] = df[sku_column].str.replace(
            f"^{sku_prefix_to_strip}", "", regex=True
        )

    # Build cases from UOM or display name
    missing_uom_set = set()

    def to_cases(row) -> float:
        qty = float(row[qty_column]) if pd.notna(row[qty_column]) else 0
        uom_val = row[uom_column] if uom_column in df.columns else None
        display_val = (
            row[display_name_column] if display_name_column in df.columns else None
        )
        uom = _parse_uom(uom_val, display_val)
        if not uom:
            sku_val = row[sku_column]
            if pd.notna(sku_val):
                missing_uom_set.add(str(sku_val).strip())
        return _sales_order_cases(qty, uom)

    df["_cases"] = df.apply(to_cases, axis=1)
    df = df[(df["_cases"] > 0) & (df[sku_column].notna())].copy()

    total_lines = len(df)
    total_skus = df[sku_column].nunique(dropna=True)
    total_cases = float(df["_cases"].sum())

    sku_db_map: Dict[str, Any] = {}
    if isinstance(sku_db, str):
        sku_db_map = load_sku_database(sku_db)
    elif isinstance(sku_db, dict):
        sku_db_map = sku_db

    if sku_db_map:
        pt_pallets, ap_pallets, ap_fraction_sum, missing_skus = estimate_pt_ap_pallets(
            df=df,
            sku_column=sku_column,
            cases_column="_cases",
            sku_db=sku_db_map,
            cases_per_pallet=cases_per_pallet,
        )
        estimated_pallets = float(pt_pallets + ap_pallets)
        work_order = estimated_pallets >= pallet_threshold
        reason = (
            f"PT={pt_pallets}, AP≈{ap_pallets} (AP fill={ap_fraction_sum:.2f}) "
            f"vs threshold {pallet_threshold}"
        )
    else:
        estimated_pallets = total_cases / max(cases_per_pallet, 1)
        pt_pallets = 0
        ap_pallets = int(math.ceil(estimated_pallets))
        ap_fraction_sum = estimated_pallets
        missing_skus = 0
        work_order = estimated_pallets >= pallet_threshold
        reason = (
            f"Estimated pallets {estimated_pallets:.2f} (volume only) "
            f"vs threshold {pallet_threshold}"
        )

    missing_uom_list = sorted(missing_uom_set)
    return QuickShipDecision(
        work_order=work_order,
        reason=reason,
        estimated_pallets=estimated_pallets,
        pt_pallets=pt_pallets,
        ap_pallets=ap_pallets,
        ap_fraction_sum=ap_fraction_sum,
        total_cases=total_cases,
        total_skus=total_skus,
        total_lines=total_lines,
        missing_skus=missing_skus,
        missing_uom_count=len(missing_uom_list),
        missing_uom_samples=missing_uom_list[:10],
        used_sku_db=bool(sku_db_map),
    )


def evaluate_takeoff(
    takeoff_df: pd.DataFrame,
    cases_per_pallet: int = DEFAULT_CASES_PER_PALLET,
    skus_per_pallet: int = DEFAULT_SKUS_PER_PALLET,
    pallet_threshold: int = DEFAULT_PALLET_THRESHOLD,
    sku_column: Optional[str] = None,
    qty_column: Optional[str] = None,
    eaches_column: Optional[str] = None,
    pack_qty_column: Optional[str] = None,
    sku_db: Optional[Union[Dict[str, Any], str]] = None,
) -> QuickShipDecision:
    """
    Estimate pallet count and return Work Order vs Quick Pick recommendation.
    """
    df = takeoff_df.copy()

    # Identify SKU column
    if sku_column is None:
        sku_column = _first_existing_column(df.columns, SKU_COLUMNS)
        if sku_column is None:
            sku_column = _find_by_keywords(df.columns, ["product", "#"])
    if sku_column is None:
        raise ValueError(
            f"Could not find SKU column. Checked: {', '.join(SKU_COLUMNS)}"
        )

    # Identify quantity column (cases/packs preferred)
    if qty_column is None:
        qty_column = _first_existing_column(df.columns, QTY_COLUMNS)
        if qty_column is None:
            qty_column = _find_by_keywords(df.columns, ["fsi", "total", "order", "qty", "pack"])
        if qty_column is None:
            qty_column = _find_by_keywords(df.columns, ["fsi", "total", "order", "qty", "case"])

    # Identify eaches column (optional fallback)
    if eaches_column is None:
        eaches_column = _first_existing_column(df.columns, EACHES_COLUMNS)
        if eaches_column is None:
            eaches_column = _find_by_keywords(df.columns, ["fsi", "total", "order", "qty", "each"])

    # Identify pack qty / UOM column (optional, used to convert eaches -> cases)
    if pack_qty_column is None:
        pack_qty_column = _first_existing_column(df.columns, PACK_QTY_COLUMNS)
        if pack_qty_column is None:
            pack_qty_column = _find_by_keywords(df.columns, ["unit", "measure"])

    if qty_column is None and eaches_column is None:
        raise ValueError(
            "Could not find quantity columns. "
            f"Checked cases: {', '.join(QTY_COLUMNS)}; eaches: {', '.join(EACHES_COLUMNS)}"
        )

    # Normalize quantities
    if qty_column is not None:
        df[qty_column] = pd.to_numeric(df[qty_column], errors="coerce").fillna(0)
    if eaches_column is not None:
        df[eaches_column] = pd.to_numeric(df[eaches_column], errors="coerce").fillna(0)
    if pack_qty_column is not None and pack_qty_column in df.columns:
        df[pack_qty_column] = pd.to_numeric(df[pack_qty_column], errors="coerce").fillna(0)

    # Build cases column
    if qty_column is not None:
        df["_cases"] = df[qty_column].copy()
    else:
        df["_cases"] = 0

    # If cases missing/zero and eaches exists, convert eaches -> cases using pack qty/UOM
    if eaches_column is not None:
        pack_qty = df[pack_qty_column] if pack_qty_column in df.columns else 1
        pack_qty = pack_qty.replace(0, 1)
        eaches_to_cases = df[eaches_column] / pack_qty
        df["_cases"] = df["_cases"].where(df["_cases"] > 0, eaches_to_cases)

    # Normalize SKU values (trim, drop blanks)
    df[sku_column] = df[sku_column].astype(str).str.strip()
    df.loc[df[sku_column].isin(["", "nan", "none", "None"]), sku_column] = pd.NA

    df = df[(df["_cases"] > 0) & (df[sku_column].notna())].copy()

    total_lines = len(df)
    total_skus = df[sku_column].nunique(dropna=True)
    total_cases = float(df["_cases"].sum())

    # PT/AP estimation using SKU database when available
    sku_db_map: Dict[str, Any] = {}
    if isinstance(sku_db, str):
        sku_db_map = load_sku_database(sku_db)
    elif isinstance(sku_db, dict):
        sku_db_map = sku_db

    if sku_db_map:
        pt_pallets, ap_pallets, ap_fraction_sum, missing_skus = estimate_pt_ap_pallets(
            df=df,
            sku_column=sku_column,
            cases_column="_cases",
            sku_db=sku_db_map,
            cases_per_pallet=cases_per_pallet,
        )
        estimated_pallets = float(pt_pallets + ap_pallets)
        work_order = estimated_pallets >= pallet_threshold
        reason = (
            f"PT={pt_pallets}, AP≈{ap_pallets} (AP fill={ap_fraction_sum:.2f}) "
            f"vs threshold {pallet_threshold}"
        )
    else:
        # Fallback: estimate pallets based on volume and SKU complexity.
        # The larger of the two becomes the estimate.
        pallets_by_volume = total_cases / max(cases_per_pallet, 1)
        pallets_by_sku = total_skus / max(skus_per_pallet, 1)
        estimated_pallets = max(pallets_by_volume, pallets_by_sku)
        pt_pallets = 0
        ap_pallets = int(math.ceil(estimated_pallets))
        ap_fraction_sum = estimated_pallets
        missing_skus = 0
        work_order = estimated_pallets >= pallet_threshold
        reason = (
            f"Estimated pallets {estimated_pallets:.2f} "
            f"(volume={pallets_by_volume:.2f}, sku={pallets_by_sku:.2f}) "
            f"vs threshold {pallet_threshold}"
        )

    return QuickShipDecision(
        work_order=work_order,
        reason=reason,
        estimated_pallets=estimated_pallets,
        pt_pallets=pt_pallets,
        ap_pallets=ap_pallets,
        ap_fraction_sum=ap_fraction_sum,
        total_cases=total_cases,
        total_skus=total_skus,
        total_lines=total_lines,
        missing_skus=missing_skus,
        missing_uom_count=0,
        missing_uom_samples=[],
        used_sku_db=bool(sku_db_map),
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Lite takeoff checker: Work Order vs Quick Pick"
    )
    parser.add_argument("path", help="Path to takeoff file (CSV or Excel)")
    parser.add_argument(
        "--cases-per-pallet",
        type=int,
        default=DEFAULT_CASES_PER_PALLET,
        help="Estimated cases per pallet",
    )
    parser.add_argument(
        "--skus-per-pallet",
        type=int,
        default=DEFAULT_SKUS_PER_PALLET,
        help="Estimated SKUs per pallet",
    )
    parser.add_argument(
        "--pallet-threshold",
        type=int,
        default=DEFAULT_PALLET_THRESHOLD,
        help="Work order threshold (estimated pallets)",
    )
    parser.add_argument(
        "--sku-column",
        type=str,
        default=None,
        help="Override SKU column name",
    )
    parser.add_argument(
        "--qty-column",
        type=str,
        default=None,
        help="Override quantity column name (cases)",
    )
    parser.add_argument(
        "--eaches-column",
        type=str,
        default=None,
        help="Optional: eaches column (used if cases missing)",
    )
    parser.add_argument(
        "--pack-qty-column",
        type=str,
        default=None,
        help="Optional: pack/UOM column for eaches->cases conversion",
    )
    parser.add_argument(
        "--sku-db",
        type=str,
        default="SKU_DATABASE.JSON",
        help="Path to SKU database JSON (pt_qty per SKU)",
    )

    args = parser.parse_args()

    df = load_takeoff(args.path)
    decision = evaluate_takeoff(
        df,
        cases_per_pallet=args.cases_per_pallet,
        skus_per_pallet=args.skus_per_pallet,
        pallet_threshold=args.pallet_threshold,
        sku_column=args.sku_column,
        qty_column=args.qty_column,
        eaches_column=args.eaches_column,
        pack_qty_column=args.pack_qty_column,
        sku_db=args.sku_db,
    )

    decision_label = "WORK ORDER" if decision.work_order else "QUICK PICK"
    print(f"Decision: {decision_label}")
    print(decision.reason)
    print(
        "Stats:",
        f"cases={decision.total_cases:.0f}",
        f"skus={decision.total_skus}",
        f"lines={decision.total_lines}",
        f"pt={decision.pt_pallets}",
        f"ap={decision.ap_pallets}",
        f"missing_skus={decision.missing_skus}",
    )


if __name__ == "__main__":
    main()
