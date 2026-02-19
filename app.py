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
    DEFAULT_CASES_PER_PALLET,
    DEFAULT_SKUS_PER_PALLET,
    DEFAULT_PALLET_THRESHOLD,
    SKU_COLUMNS,
    QTY_COLUMNS,
    EACHES_COLUMNS,
    PACK_QTY_COLUMNS,
)


def _read_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    buffer = io.BytesIO(data)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(buffer)
    return pd.read_csv(buffer)


def _pick_default_column(columns, candidates) -> Optional[str]:
    for candidate in candidates:
        if candidate in columns:
            return candidate
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
        "SKU database JSON (optional)",
        type=["json"],
        help="JSON map of SKU -> PT_QTY",
    )

input_mode = st.radio(
    "Input type",
    options=["Takeoff", "Sales Order"],
    horizontal=True,
)

if uploaded_file is not None:
    try:
        df = _read_uploaded_file(uploaded_file)
    except Exception as exc:
        st.error(f"Failed to read file: {exc}")
        st.stop()

    columns = list(df.columns)

    if input_mode == "Takeoff":
        default_sku = _pick_default_column(columns, SKU_COLUMNS)
        default_qty = _pick_default_column(columns, QTY_COLUMNS)
        default_eaches = _pick_default_column(columns, EACHES_COLUMNS)
        default_pack_qty = _pick_default_column(columns, PACK_QTY_COLUMNS)

        sku_column = st.selectbox(
            "SKU column",
            options=columns,
            index=columns.index(default_sku) if default_sku in columns else 0,
        )
        qty_options = ["(auto)"] + columns
        qty_default_index = (
            qty_options.index(default_qty) if default_qty in qty_options else 0
        )
        qty_column = st.selectbox(
            "Quantity (cases/packs) column",
            options=qty_options,
            index=qty_default_index,
        )
        eaches_options = ["(auto/none)"] + columns
        eaches_default_index = (
            eaches_options.index(default_eaches) if default_eaches in eaches_options else 0
        )
        eaches_column = st.selectbox(
            "Eaches column (optional)",
            options=eaches_options,
            index=eaches_default_index,
        )
        pack_qty_options = ["(auto/none)"] + columns
        pack_qty_default_index = (
            pack_qty_options.index(default_pack_qty)
            if default_pack_qty in pack_qty_options
            else 0
        )
        pack_qty_column = st.selectbox(
            "Pack Qty / UOM column (optional)",
            options=pack_qty_options,
            index=pack_qty_default_index,
        )

        preview_df = _filter_ordered_rows(
            df,
            sku_column=sku_column,
            qty_column=None if qty_column == "(auto)" else qty_column,
            eaches_column=None if eaches_column == "(auto/none)" else eaches_column,
            pack_qty_column=None if pack_qty_column == "(auto/none)" else pack_qty_column,
        )
        st.subheader("Preview (ordered items only)")
        st.dataframe(preview_df.head(50), use_container_width=True)
    else:
        item_default = "Item" if "Item" in columns else columns[0]
        qty_default = "Quantity" if "Quantity" in columns else columns[0]
        display_default = "Display Name" if "Display Name" in columns else columns[0]
        uom_default = "UOM" if "UOM" in columns else None

        item_column = st.selectbox(
            "Item/SKU column",
            options=columns,
            index=columns.index(item_default) if item_default in columns else 0,
        )
        qty_column = st.selectbox(
            "Quantity column",
            options=columns,
            index=columns.index(qty_default) if qty_default in columns else 0,
        )
        display_column = st.selectbox(
            "Display Name column (optional)",
            options=["(auto/none)"] + columns,
            index=0 if display_default not in columns else (columns.index(display_default) + 1),
        )
        uom_column = st.selectbox(
            "UOM column (optional)",
            options=["(auto/none)"] + columns,
            index=0 if uom_default not in columns else (columns.index(uom_default) + 1),
        )
        sku_prefix = st.text_input("SKU prefix to strip", value="3172_")

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
            sku_db = None
            if sku_db_file is not None:
                sku_db = json.loads(sku_db_file.getvalue().decode("utf-8"))
            if input_mode == "Takeoff":
                decision = evaluate_takeoff(
                    df,
                    cases_per_pallet=int(cases_per_pallet),
                    skus_per_pallet=int(skus_per_pallet),
                    pallet_threshold=int(pallet_threshold),
                    sku_column=sku_column,
                    qty_column=None if qty_column == "(auto)" else qty_column,
                    eaches_column=None
                    if eaches_column == "(auto/none)"
                    else eaches_column,
                    pack_qty_column=None
                    if pack_qty_column == "(auto/none)"
                    else pack_qty_column,
                    sku_db=sku_db,
                )
            else:
                decision = evaluate_sales_orders(
                    df,
                    cases_per_pallet=int(cases_per_pallet),
                    pallet_threshold=int(pallet_threshold),
                    sku_column=item_column,
                    qty_column=qty_column,
                    uom_column=None if uom_column == "(auto/none)" else uom_column,
                    display_name_column=None
                    if display_column == "(auto/none)"
                    else display_column,
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
