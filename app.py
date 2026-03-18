from __future__ import annotations

import io
import json
import os
from typing import Optional

import pandas as pd
import streamlit as st

from QuickShipChecker import (
    evaluate_takeoff,
    evaluate_sales_orders,
    load_sku_database,
    DEFAULT_CASES_PER_PALLET,
    DEFAULT_SKUS_PER_PALLET,
    DEFAULT_PALLET_THRESHOLD,
    SKU_COLUMNS,
    QTY_COLUMNS,
    EACHES_COLUMNS,
    PACK_QTY_COLUMNS,
    _find_by_keywords,
)


def _read_uploaded_file(uploaded_file, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Read uploaded CSV or Excel file. For Excel, an optional sheet name can be provided."""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    buffer = io.BytesIO(data)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(buffer, sheet_name=sheet_name)
    return pd.read_csv(buffer)


def _pick_default_column(columns, candidates) -> Optional[str]:
    """Pick first candidate that exists in columns (exact or case-insensitive)."""
    col_set = {c: c for c in columns}
    col_lower = {c.lower(): c for c in columns}
    for candidate in candidates:
        if candidate in col_set:
            return candidate
        if candidate.lower() in col_lower:
            return col_lower[candidate.lower()]
    return None


def _filter_ordered_rows(
    df: pd.DataFrame,
    sku_column: str,
    qty_column: Optional[str],
    eaches_column: Optional[str],
    pack_qty_column: Optional[str],
) -> pd.DataFrame:
    filtered = df.copy()
    if qty_column is not None:
        filtered[qty_column] = pd.to_numeric(filtered[qty_column], errors="coerce").fillna(0)
    if eaches_column is not None:
        filtered[eaches_column] = pd.to_numeric(filtered[eaches_column], errors="coerce").fillna(0)
    if pack_qty_column is not None:
        filtered[pack_qty_column] = pd.to_numeric(filtered[pack_qty_column], errors="coerce").fillna(0)

    if qty_column is not None:
        filtered["_cases"] = filtered[qty_column].copy()
    else:
        filtered["_cases"] = 0

    if eaches_column is not None:
        pack_qty = filtered[pack_qty_column] if pack_qty_column in filtered.columns else 1
        pack_qty = pack_qty.replace(0, 1)
        eaches_to_cases = filtered[eaches_column] / pack_qty
        filtered["_cases"] = filtered["_cases"].where(filtered["_cases"] > 0, eaches_to_cases)

    filtered[sku_column] = filtered[sku_column].astype(str).str.strip()
    filtered.loc[filtered[sku_column].isin(["", "nan", "none", "None"]), sku_column] = pd.NA

    filtered = filtered[(filtered["_cases"] > 0) & (filtered[sku_column].notna())].copy()
    return filtered.drop(columns=["_cases"])


st.set_page_config(page_title="QuickShip Checker", layout="wide")

STYLE_PATH = "Style.css"
if os.path.exists(STYLE_PATH):
    with open(STYLE_PATH, "r", encoding="utf-8") as css_file:
        st.markdown(
            f"<style>{css_file.read()}</style>",
            unsafe_allow_html=True,
        )

st.markdown(
    """
    <div class="fsi-splash-top"></div>
    <div class="fsi-splash-side"></div>
    """,
    unsafe_allow_html=True,
)

LOGO_PATH = "logo.png"

st.markdown('<div class="app-content">', unsafe_allow_html=True)

logo_exists = os.path.exists(LOGO_PATH)
header_left, header_right = st.columns([1, 6], vertical_alignment="center")
with header_left:
    if logo_exists:
        st.image(LOGO_PATH, use_container_width=True)
    else:
        st.markdown('<div class="logo-slot">Logo</div>', unsafe_allow_html=True)
with header_right:
    st.markdown(
        """
        <div class="app-header">
          <div>
            <p class="title">QuickShip Checker (Lite)</p>
            <p class="subtitle">Upload a takeoff to decide Work Order vs Quick Pick.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

uploaded_file = st.file_uploader(
    "Upload CSV or Excel",
    type=["csv", "xlsx", "xls"],
    help="Drag and drop a file here or click to browse.",
    key=f"upload_{st.session_state.uploader_key}",
)

st.link_button(
    "GET SALES ORDER HERE",
    "https://4119972.app.netsuite.com/app/common/search/searchresults.nl?searchid=6337&whence=",
)

# Default SKU database path (app directory); used unless user uploads an override
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SKU_DB_PATH = os.path.join(_APP_DIR, "SKU_DATABASE.JSON")

with st.expander("Advanced settings", expanded=False):
    cases_per_pallet = st.number_input(
        "Estimated cases per pallet",
        min_value=1,
        value=DEFAULT_CASES_PER_PALLET,
        step=1,
    )
    skus_per_pallet = st.number_input(
        "Estimated SKUs per pallet",
        min_value=1,
        value=DEFAULT_SKUS_PER_PALLET,
        step=1,
    )
    pallet_threshold = st.number_input(
        "Work order threshold (estimated pallets)",
        min_value=1,
        value=DEFAULT_PALLET_THRESHOLD,
        step=1,
    )
    sku_db_file = st.file_uploader(
        "SKU database JSON (optional override)",
        type=["json"],
        help="Uses SKU_DATABASE.JSON from app folder by default. Upload here to override.",
    )
    sku_prefix = st.text_input(
        "SKU prefix to strip (Sales Order)",
        value="3172_",
        help="Prefix removed from Item/SKU values when using Sales Order input.",
    )

input_mode = st.radio(
    "Input type",
    options=["Takeoff", "Sales Order"],
    horizontal=True,
)

if uploaded_file is not None:
    try:
        # Allow selecting a worksheet when an Excel file has multiple sheets.
        sheet_name: Optional[str] = None
        file_name = uploaded_file.name.lower()
        if file_name.endswith((".xlsx", ".xls")):
            data = uploaded_file.getvalue()
            excel = pd.ExcelFile(io.BytesIO(data))
            sheet_names = excel.sheet_names
            sheet_name = st.selectbox(
                "Worksheet",
                options=sheet_names,
                index=0,
            )
        df = _read_uploaded_file(uploaded_file, sheet_name=sheet_name)
    except Exception as exc:
        st.error(f"Failed to read file: {exc}")
        st.stop()

    columns = list(df.columns)

    if input_mode == "Takeoff":
        # Auto-map takeoff: Product #, FSI Total Order Qty (in Packs), eaches, Pack Qty
        sku_column = _pick_default_column(columns, SKU_COLUMNS) or columns[0]
        qty_column = _pick_default_column(columns, QTY_COLUMNS)
        if qty_column is None:
            qty_column = _find_by_keywords(columns, ["fsi", "total", "order", "qty", "pack"])
        if qty_column is None:
            qty_column = _find_by_keywords(columns, ["fsi", "total", "order", "qty", "case"])
        eaches_column = _pick_default_column(columns, EACHES_COLUMNS)
        if eaches_column is None:
            eaches_column = _find_by_keywords(columns, ["fsi", "total", "order", "qty", "each"])
        if eaches_column is None:
            eaches_column = _find_by_keywords(columns, ["fsi", "release", "eaches"])
        pack_qty_column = _pick_default_column(columns, PACK_QTY_COLUMNS)
        if pack_qty_column is None:
            pack_qty_column = _find_by_keywords(columns, ["pack", "qty"])
        if pack_qty_column is None:
            pack_qty_column = _find_by_keywords(columns, ["unit", "measure"])

        preview_df = _filter_ordered_rows(
            df,
            sku_column=sku_column,
            qty_column=qty_column,
            eaches_column=eaches_column,
            pack_qty_column=pack_qty_column,
        )
        st.subheader("Preview (ordered items only)")
        st.dataframe(preview_df.head(50), use_container_width=True)
    else:
        # Auto-map sales order columns: SKU, Quantity, Item (display), UOM (no UI)
        item_column = _pick_default_column(columns, ["SKU", "sku", "Item", "item"]) or columns[0]
        qty_column = _pick_default_column(columns, ["Quantity", "quantity", "Qty", "qty"]) or columns[0]
        display_column = _pick_default_column(
            columns, ["Item", "item", "Display Name", "display name", "Display", "display"]
        )
        uom_column = _pick_default_column(
            columns, ["UOM", "uom", "Unit of Measure", "unit of measure"]
        )

        preview_df = df.copy()
        preview_df[item_column] = preview_df[item_column].astype(str).str.strip()
        preview_df[qty_column] = pd.to_numeric(
            preview_df[qty_column], errors="coerce"
        ).fillna(0)
        preview_df = preview_df[preview_df[qty_column] > 0]
        st.subheader("Preview (ordered items only)")
        st.dataframe(preview_df.head(50), use_container_width=True)

    if st.button("Evaluate"):
        try:
            # Use uploaded JSON if provided; otherwise default to SKU_DATABASE.JSON in app dir
            sku_db = None
            if sku_db_file is not None:
                sku_db = json.loads(sku_db_file.getvalue().decode("utf-8"))
            else:
                sku_db = load_sku_database(DEFAULT_SKU_DB_PATH)
            if input_mode == "Takeoff":
                decision = evaluate_takeoff(
                    df,
                    cases_per_pallet=int(cases_per_pallet),
                    skus_per_pallet=int(skus_per_pallet),
                    pallet_threshold=int(pallet_threshold),
                    sku_column=sku_column,
                    qty_column=qty_column,
                    eaches_column=eaches_column,
                    pack_qty_column=pack_qty_column,
                    sku_db=sku_db,
                )
            else:
                decision = evaluate_sales_orders(
                    df,
                    cases_per_pallet=int(cases_per_pallet),
                    pallet_threshold=int(pallet_threshold),
                    sku_column=item_column,
                    qty_column=qty_column,
                    uom_column=uom_column,
                    display_name_column=display_column,
                    sku_prefix_to_strip=sku_prefix,
                    sku_db=sku_db,
                )
        except Exception as exc:
            st.error(f"Failed to evaluate: {exc}")
            st.stop()

        decision_label = "WORK ORDER" if decision.work_order else "QUICK SHIP"
        if decision.work_order:
            st.success("Decision: WORK ORDER")
        else:
            st.info("Decision: QUICK SHIP")

        st.subheader("Results")
        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
        r1c1.metric("Decision", decision_label)
        r1c2.metric("Estimated pallets", f"{decision.estimated_pallets:.2f}")
        r1c3.metric("PT pallets", f"{decision.pt_pallets}")
        r1c4.metric("AP pallets", f"{decision.ap_pallets}")

        r2c1, r2c2, r2c3, r2c4 = st.columns(4)
        r2c1.metric("Total cases", f"{decision.total_cases:.0f}")
        r2c2.metric("Total SKUs", f"{decision.total_skus}")
        r2c3.metric("Total lines", f"{decision.total_lines}")
        r2c4.metric("Missing SKUs", f"{decision.missing_skus}")

        st.caption(decision.reason)
        st.caption(f"AP fill sum: {decision.ap_fraction_sum:.2f}")

        # Explain why PT pallets might be 0 so users can fix data or load SKU DB
        if decision.pt_pallets == 0:
            if not decision.used_sku_db:
                st.info(
                    "**PT pallets = 0:** No SKU database was loaded. Upload a SKU database JSON (SKU → PT_QTY) in Advanced settings to calculate PT (full) vs AP (partial) pallets."
                )
            elif decision.missing_skus >= decision.total_skus:
                st.warning(
                    f"**PT pallets = 0:** None of the {decision.total_skus} order SKUs were found in the SKU database. Add these SKUs to the database with a PT_QTY (cases per full pallet) to see PT pallets."
                )
            else:
                st.info(
                    f"**PT pallets = 0:** SKU database was used ({decision.total_skus - decision.missing_skus} SKUs matched, {decision.missing_skus} missing). No full pallets: for every matched SKU, order cases were less than PT_QTY (cases per full pallet). Remaining volume is counted as AP pallets."
                )

        if st.button("Clear / Start over"):
            st.session_state.uploader_key += 1
            st.rerun()
        if decision.missing_uom_count > 0:
            st.warning(
                "Missing UOM for "
                f"{decision.missing_uom_count} item(s). "
                "They were treated as cases. "
                f"Examples: {', '.join(decision.missing_uom_samples)}"
            )
else:
    st.info("Upload a file to get started.")

st.markdown("</div>", unsafe_allow_html=True)
