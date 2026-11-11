import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
import hashlib
import os
import zipfile
from pathlib import Path
from copy import deepcopy
from io import BytesIO
from lxml import etree

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="RDC orders - BOL Generator",
    page_icon="📦",
    layout="wide"
)

# =========================================================
# CONSTANTS / FILE PATHS
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR

LTL_QTY_PATH = ASSETS_DIR / "LTL_qty_updated_7.28.26.xlsx"
RDC_LIST_PATH = ASSETS_DIR / "RDC_list.xlsx"
CARRIER_MAP_PATH = ASSETS_DIR / "Carrier List 2026 Rev 10.27.26.xlsx"
CUBE_PATH = ASSETS_DIR / "Cube_mapping.xlsx"
BOL_TEMPLATE_PATH = ASSETS_DIR / "BOL_RDC_template.docx"
MANIFEST_TEMPLATE_PATH = ASSETS_DIR / "Master Manifest_template.docx"

MAX_TEMPLATE_ROWS = 5
MAX_TRUCK_WEIGHT_LBS = 44000
INT_TRUCK_WEIGHT_LBS = 42500
TRUCK_ANALYSIS_VERSION = 4
TRUCK_SPLIT_COLUMNS = [
    "SID",
    "DN#",
    "Truck",
    "Purchase Order No",
    "Customer",
    "Name 1",
    "Description",
    "Sales document",
    "Material",
    "Material Description",
    "Batch",
    "Sales unit",
    "Order Quantity",
    "Gross weight (kg)",
    "Gross weight (lbs)"
]
TRUCK_SPLIT_INPUT_COLUMNS = [
    "Purchase order no.",
    "Customer",
    "Name 1",
    "Description",
    "Sales document",
    "Material",
    "Material Description",
    "Batch",
    "Sales unit",
    "Order Quantity",
    "Gross weight"
]

# =========================================================
# WORD DOCUMENT PROCESSING FUNCTIONS
# =========================================================
def _replace_in_paragraph(paragraph, replacements):
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    replaced = full_text
    for k, v in replacements.items():
        replaced = replaced.replace(k, str(v))

    if replaced != full_text and paragraph.runs:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run.text = ""


def _replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                _replace_in_paragraph(p, replacements)
            for t in cell.tables:
                _replace_in_table(t, replacements)


def fill_template_doc(template_path, replacements):
    doc = Document(template_path)

    for p in doc.paragraphs:
        _replace_in_paragraph(p, replacements)

    for t in doc.tables:
        _replace_in_table(t, replacements)

    for section in doc.sections:
        for p in section.header.paragraphs:
            _replace_in_paragraph(p, replacements)
        for t in section.header.tables:
            _replace_in_table(t, replacements)

        for p in section.footer.paragraphs:
            _replace_in_paragraph(p, replacements)
        for t in section.footer.tables:
            _replace_in_table(t, replacements)

    return doc


def append_doc_content(master_doc, sub_doc):
    for element in sub_doc.element.body:
        if element.tag.endswith("sectPr"):
            continue
        master_doc.element.body.append(deepcopy(element))


def prepend_doc_content(main_doc, prefix_doc):
    body = main_doc.element.body

    to_insert = [
        deepcopy(el) for el in prefix_doc.element.body
        if not el.tag.endswith("sectPr")
    ]

    pgbrk_xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:r><w:br w:type="page"/></w:r></w:p>'
    )
    to_insert.append(etree.fromstring(pgbrk_xml))

    for el in reversed(to_insert):
        body.insert(0, el)


# =========================================================
# MANIFEST FUNCTIONS
# =========================================================
def build_po_summary(sid, upload_merged, truck=None):
    sid_df = upload_merged[upload_merged["SID"] == sid]
    if truck is not None and "Truck" in sid_df.columns:
        sid_df = sid_df[sid_df["Truck"] == truck]
    po_summary = (
        sid_df
        .groupby(["SID", "Purchase order no."], as_index=False)
        .agg({
            "Gross weight": "sum",
            "RDC names": "first",
            "Codice RDC": "first",
            "Carrier_mapped": "first",
            "SCAC Code": "first",
            "Cube_po": "sum",
            "Pallet_qty": "sum"
        })
    )
    # Convert from kg to lbs + add pallet weight
    po_summary["Gross weight"] = po_summary["Gross weight"]*2.205 + po_summary["Pallet_qty"]*46
    po_summary["Gross weight"] = po_summary["Gross weight"].astype(int)
    po_summary["Pallet_qty"] = po_summary["Pallet_qty"].fillna(0).astype(int)
    return po_summary


def build_manifest_replacements(sid, po_summary, carrier_name, scac, shipment_rows):
    pickup_date = pd.to_datetime(shipment_rows["Pickup Date"].iloc[0], errors="coerce")
    delivery_date = pickup_date + pd.Timedelta(days=1) if pd.notna(pickup_date) else pd.NaT

    # instead of using tot weight from VPX, sum the POs weight from SAP
    total_weight = po_summary["Gross weight"].sum()

    replacements = {
        "{{SID}}": sid,
        "{{CARRIER NAME}}": carrier_name,
        "{{SCAC}}": scac,
        "{{TOTWEIGHT}}": f"{int(total_weight):,}",
        "{{TOTCUBE}}": f"{int(shipment_rows['Cube'].iloc[0]):,}",
        "{{PICKUP}}": pickup_date.strftime("%m/%d/%Y") if pd.notna(pickup_date) else "",
        "{{DELIVERY}}": delivery_date.strftime("%m/%d/%Y") if pd.notna(delivery_date) else "",
        "{{RECEIVER}}": "MCDONOUGH IFC",
        **{f"{{{{PO_{i}}}}}": "" for i in range(1, MAX_TEMPLATE_ROWS + 1)},
        **{f"{{{{RDC_{i}}}}}": "" for i in range(1, MAX_TEMPLATE_ROWS + 1)},
        **{f"{{{{Weight_{i}}}}}": "" for i in range(1, MAX_TEMPLATE_ROWS + 1)},
        **{f"{{{{Cube_{i}}}}}": "" for i in range(1, MAX_TEMPLATE_ROWS + 1)},
    }
    return replacements


def _get_row_data_for_po(po_row):
    return {
        "stop": "FINAL",
        "sid": str(po_row["SID"]),
        "po": str(po_row["Purchase order no."]),
        "supplier": "ATLAS CONCORDE - MT PLEASANT",
        "dest": str(po_row["Codice RDC"]),
        "qty": "",
        "haz": "",
        "weight": f"{int(po_row['Gross weight']):,}",
        "cube": f"{int(po_row['Cube_po']):,}",
        "pct_truck": "",
        "case": "",
        "exp_pallets": str(int(po_row["Pallet_qty"])),
        "rec_pallets": ""
    }


def _fill_detail_row(table_row, rd):
    cell_values = [
        rd["stop"], rd["sid"], rd["po"], rd["supplier"], rd["dest"],
        rd["qty"], rd["haz"], rd["weight"], rd["cube"],
        rd["pct_truck"], rd["case"], rd["exp_pallets"], rd["rec_pallets"]
    ]

    cells = table_row.cells

    for j, val in enumerate(cell_values):
        if j < len(cells):
            cell = cells[j]

            if not cell.paragraphs:
                p = cell.add_paragraph()
            else:
                p = cell.paragraphs[0]

            p.text = str(val)

            for extra_p in cell.paragraphs[1:]:
                extra_p._element.getparent().remove(extra_p._element)


def _clear_detail_row(table_row):
    empty = {
        "stop": "", "sid": "", "po": "", "supplier": "", "dest": "",
        "qty": "", "haz": "", "weight": "", "cube": "",
        "pct_truck": "", "case": "", "exp_pallets": "", "rec_pallets": ""
    }
    _fill_detail_row(table_row, empty)


def fill_manifest_table_rows(doc, po_rows_list):
    detail_table = doc.tables[-1]
    data_rows = detail_table.rows[1:]
    n_pos = len(po_rows_list)
    n_template = len(data_rows)

    for i, tr in enumerate(data_rows):
        if i < n_pos:
            _fill_detail_row(tr, po_rows_list[i])
        else:
            _clear_detail_row(tr)

    if n_pos > n_template:
        last_tr = data_rows[-1]
        for i in range(n_template, n_pos):
            new_tr = deepcopy(last_tr._tr)
            detail_table._tbl.append(new_tr)
            new_row = detail_table.rows[-1]
            _fill_detail_row(new_row, po_rows_list[i])


def create_manifest_doc(manifest_template_path, sid, shipment_rows, upload_merged, truck=None):
    carrier_name = shipment_rows["Carrier_mapped"].iloc[0]
    scac = shipment_rows["SCAC Code"].iloc[0]

    po_summary = build_po_summary(sid, upload_merged, truck)
    replacements = build_manifest_replacements(sid, po_summary, carrier_name, scac, shipment_rows)

    doc = fill_template_doc(manifest_template_path, replacements)

    po_rows_list = [
        _get_row_data_for_po(po_row)
        for _, po_row in po_summary.iterrows()
    ]

    fill_manifest_table_rows(doc, po_rows_list)
    return doc


# =========================================================
# FILE READING HELPERS
# =========================================================
def read_multiple_planex_csv(planex_files, accepted_only=True):
    cols = {
        "Shipment ID": str,
        "Purchase Order(s)": str,
        "Destination ID": str,
        "Carrier Code": str,
        "Weight": float,
        "Cube": float,
        "Quantity": float
    }

    dfs = []
    for f in planex_files:
        f.seek(0)
        temp = pd.read_csv(f, dtype=cols)
        temp["source_file"] = f.name
        dfs.append(temp)

    df_upload_planex = pd.concat(dfs, ignore_index=True)
    if accepted_only:
        df_upload_planex = df_upload_planex[df_upload_planex["Shipment Status"] == "Accepted"].copy()
    for column in ["Shipment ID", "Destination ID", "Carrier Code"]:
        df_upload_planex[column] = df_upload_planex[column].map(
            _normalize_open_order_value
        )
    df_upload_planex["Weight"] = df_upload_planex["Weight"].fillna(0).astype(int)
    df_upload_planex["Cube"] = df_upload_planex["Cube"].fillna(0).astype(int)
    df_upload_planex["Quantity"] = df_upload_planex["Quantity"].fillna(0).astype(int)

    return df_upload_planex


def read_multiple_open_orders_excel(order_files, bol_rows_only=True):
    cols_as_str = {
        "Purchase order no.": str,
        "Customer": str,
        "SID": str,
        "DN#": str,
        "Material": str,
        "Carrier": str
    }

    dfs = []
    for f in order_files:
        f.seek(0)
        temp = pd.read_excel(f, dtype=cols_as_str)
        temp.columns = [
            str(column).strip().lstrip("\ufeff")
            for column in temp.columns
        ]
        for col in cols_as_str.keys():
            if col in temp.columns:
                temp[col] = temp[col].map(_normalize_open_order_value)
        temp["source_file"] = f.name
        dfs.append(temp)

    df_upload = pd.concat(dfs, ignore_index=True)

    cols_to_drop = [c for c in ["Unnamed: 31", "Unnamed: 32", "Unnamed: 33", "source_file"] if c in df_upload.columns]
    df_upload = df_upload.drop(columns=cols_to_drop, errors="ignore")
    df_upload = df_upload.drop_duplicates()

    if bol_rows_only:
        if "DN#" not in df_upload.columns:
            raise ValueError("The RDC Open Orders file is missing the required 'DN#' column.")
        df_upload = df_upload[df_upload["DN#"].notna()]
    for column in ["Purchase order no.", "Customer", "Material", "SID", "DN#"]:
        if column in df_upload.columns:
            df_upload[column] = df_upload[column].map(_normalize_open_order_value)
    df_upload = df_upload.reset_index(drop=True)

    return df_upload


def read_completed_truck_split_excel(uploaded_file):
    uploaded_file.seek(0)
    df_orders = pd.read_excel(uploaded_file)
    df_orders.columns = [
        str(column).strip().lstrip("\ufeff")
        for column in df_orders.columns
    ]
    df_orders = df_orders.rename(columns={
        "Purchase Order No": "Purchase order no.",
        "Gross weight (kg)": "Gross weight"
    })

    required_columns = [
        "SID",
        "DN#",
        "Truck",
        "Purchase order no.",
        "Customer",
        "Material",
        "Order Quantity",
        "Gross weight"
    ]
    missing_columns = [
        column for column in required_columns
        if column not in df_orders.columns
    ]
    if missing_columns:
        raise ValueError(
            "The completed truck split workbook is missing required columns: "
            + ", ".join(missing_columns)
        )

    df_orders["Purchase order no."] = df_orders["Purchase order no."].map(
        _normalize_open_order_value
    )
    df_orders = df_orders[
        df_orders["Purchase order no."].notna()
        & df_orders["Purchase order no."].astype(str).str.strip().ne("")
    ].copy().reset_index(drop=True)
    if df_orders.empty:
        raise ValueError("The completed truck split workbook has no order lines.")

    for column in ["SID", "DN#", "Truck"]:
        df_orders[column] = df_orders[column].map(_normalize_open_order_value)
        df_orders[column] = df_orders[column].mask(
            df_orders[column].astype("string").str.strip().str.lower().isin(
                {"", "nan", "none"}
            )
        )
    for column in ["Customer", "Material"]:
        df_orders[column] = df_orders[column].map(_normalize_open_order_value)

    missing_sid = df_orders["SID"].isna()
    missing_truck = (
        df_orders["Truck"].isna()
        | df_orders["Truck"].str.lower().eq("unassigned")
    )
    missing_dn = df_orders["DN#"].isna()
    incomplete_rows = missing_sid | missing_truck | missing_dn
    issues = []
    for row_index in df_orders.index[incomplete_rows]:
        missing_fields = []
        if missing_sid.loc[row_index]:
            missing_fields.append("SID")
        if missing_truck.loc[row_index]:
            missing_fields.append("assigned Truck")
        if missing_dn.loc[row_index]:
            missing_fields.append("DN#")
        issues.append({
            "Order line": int(row_index) + 1,
            "Purchase Order No": df_orders.loc[row_index, "Purchase order no."],
            "SID": df_orders.loc[row_index, "SID"],
            "Truck": df_orders.loc[row_index, "Truck"],
            "Issue": "Missing " + ", ".join(missing_fields)
        })

    valid_orders = df_orders.loc[~incomplete_rows].copy()

    valid_orders["Gross weight"] = pd.to_numeric(
        valid_orders["Gross weight"],
        errors="coerce"
    )
    valid_orders["Order Quantity"] = pd.to_numeric(
        valid_orders["Order Quantity"],
        errors="coerce"
    )
    if (
        valid_orders["Gross weight"].isna().any()
        or valid_orders["Order Quantity"].isna().any()
    ):
        raise ValueError(
            "The completed truck split workbook contains invalid gross weight or order quantity values."
        )
    return (
        valid_orders.reset_index(drop=True),
        pd.DataFrame(
            issues,
            columns=["Order line", "Purchase Order No", "SID", "Truck", "Issue"]
        )
    )


def _normalize_open_order_value(value):
    if pd.isna(value):
        return pd.NA
    if isinstance(value, (int, np.integer)):
        return str(value)
    if isinstance(value, (float, np.floating)) and np.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def build_planex_po_pairs(df_planex):
    pairs = []
    for shipment_id, purchase_orders in df_planex[
        ["Shipment ID", "Purchase Order(s)"]
    ].itertuples(index=False, name=None):
        shipment_id = "" if pd.isna(shipment_id) else str(shipment_id).strip()
        purchase_orders = "" if pd.isna(purchase_orders) else str(purchase_orders)
        if not shipment_id or shipment_id.lower() in {"nan", "none"}:
            continue

        for purchase_order in purchase_orders.split(","):
            purchase_order = purchase_order.strip()
            if purchase_order and purchase_order.lower() not in {"nan", "none"}:
                pairs.append({
                    "Shipment ID": shipment_id,
                    "Purchase order no.": purchase_order
                })

    return pd.DataFrame(
        pairs,
        columns=["Shipment ID", "Purchase order no."]
    ).drop_duplicates()


def _pack_order_lines(line_weights, weight_limit):
    line_weights = [int(np.ceil(round(weight, 6))) for weight in line_weights]
    reachable = {0: None}
    nodes = []
    selected_lines = {
        line_index for line_index, line_weight in enumerate(line_weights)
        if line_weight == 0
    }

    for line_index, line_weight in enumerate(line_weights):
        if line_weight <= 0 or line_weight > weight_limit:
            continue

        for current_weight, previous_node in list(reachable.items()):
            new_weight = current_weight + line_weight
            if new_weight <= weight_limit and new_weight not in reachable:
                nodes.append((line_index, previous_node))
                reachable[new_weight] = len(nodes) - 1

    best_weight = max(reachable)
    node_index = reachable[best_weight]
    while node_index is not None:
        line_index, node_index = nodes[node_index]
        selected_lines.add(line_index)

    if not selected_lines and line_weights:
        selected_lines.add(0)

    return selected_lines


def build_truck_split_output(order_weights, planex_pairs, ambiguous_pos, truck_summary):
    missing_columns = [
        column for column in TRUCK_SPLIT_INPUT_COLUMNS
        if column not in order_weights.columns
    ]
    if missing_columns:
        raise ValueError(
            "The RDC Open Orders file is missing required truck-split columns: "
            + ", ".join(missing_columns)
        )

    valid_pairs = planex_pairs[
        ~planex_pairs["Purchase order no."].isin(ambiguous_pos)
    ].drop_duplicates()
    split_rows = order_weights.merge(
        valid_pairs,
        on="Purchase order no.",
        how="left",
        sort=False
    )

    summary_by_sid = truck_summary.set_index("Shipment ID")
    truck_assignments = pd.Series(index=split_rows.index, dtype="object")

    for shipment_id, sid_rows in split_rows[
        split_rows["Shipment ID"].notna()
    ].groupby("Shipment ID", sort=False):
        truck_names = {row_index: "Truck 1" for row_index in sid_rows.index}
        if shipment_id in summary_by_sid.index:
            shipment_summary = summary_by_sid.loc[shipment_id]
            if (
                shipment_summary["Status"] == "Overweight"
                and pd.notna(shipment_summary["Weight limit (lbs)"])
            ):
                remaining_rows = list(sid_rows.index)
                truck_number = 1
                weight_limit = int(shipment_summary["Weight limit (lbs)"])

                while remaining_rows:
                    weights = split_rows.loc[remaining_rows, "Weight (lbs)"]
                    if weights.isna().any():
                        break

                    selected_positions = _pack_order_lines(weights.tolist(), weight_limit)
                    selected_rows = [
                        remaining_rows[position]
                        for position in selected_positions
                    ]
                    for row_index in selected_rows:
                        truck_names[row_index] = f"Truck {truck_number}"
                    remaining_rows = [
                        row_index for row_index in remaining_rows
                        if row_index not in selected_rows
                    ]
                    truck_number += 1

        for row_index, truck_name in truck_names.items():
            truck_assignments.loc[row_index] = truck_name

    split_rows["Truck"] = truck_assignments
    split_rows["SID"] = split_rows["Shipment ID"].fillna("")
    split_rows["DN#"] = ""
    split_rows.loc[split_rows["SID"].eq(""), "Truck"] = "Unassigned"
    split_rows["Gross weight (kg)"] = split_rows["Gross weight"]
    split_rows["Gross weight (lbs)"] = split_rows["Weight (lbs)"].round().astype("Int64")
    shipment_order = {
        shipment_id: order
        for order, shipment_id in enumerate(split_rows["SID"].drop_duplicates())
    }
    split_rows["_shipment_order"] = split_rows["SID"].map(shipment_order)
    split_rows["_truck_order"] = (
        split_rows["Truck"].str.extract(r"(\d+)", expand=False).fillna("0").astype(int)
    )
    split_rows = split_rows.sort_values(
        ["_shipment_order", "_truck_order"],
        kind="stable"
    )
    split_rows = split_rows.rename(columns={"Purchase order no.": "Purchase Order No"})
    return split_rows.reindex(columns=TRUCK_SPLIT_COLUMNS)


def build_truck_load_summary(df_planex, df_orders, ltl_qty_df, carrier_df):
    planex_pairs = build_planex_po_pairs(df_planex)

    orders = df_orders.copy()
    if "Material" in orders.columns:
        orders["Material"] = orders["Material"].map(_normalize_open_order_value)
    orders["Purchase order no."] = orders["Purchase order no."].fillna("").astype(str).str.strip()
    orders = orders[
        ~orders["Purchase order no."].str.lower().isin({"", "nan", "none"})
    ].copy()

    order_weights = pd.merge(
        orders,
        ltl_qty_df[["SAP Code", "Case_Pallet"]],
        left_on="Material",
        right_on="SAP Code",
        how="left"
    )
    order_weights["Gross weight"] = pd.to_numeric(order_weights["Gross weight"], errors="coerce")
    order_weights["Order Quantity"] = pd.to_numeric(order_weights["Order Quantity"], errors="coerce")
    order_weights["Case_Pallet"] = pd.to_numeric(order_weights["Case_Pallet"], errors="coerce")
    order_weights["Pallet_qty"] = np.ceil(
        order_weights["Order Quantity"] / order_weights["Case_Pallet"]
    )
    order_weights["Weight data issue"] = (
        order_weights["Gross weight"].isna()
        | order_weights["Order Quantity"].isna()
        | order_weights["Case_Pallet"].isna()
        | ~np.isfinite(order_weights["Gross weight"])
        | ~np.isfinite(order_weights["Order Quantity"])
        | ~np.isfinite(order_weights["Case_Pallet"])
        | order_weights["Case_Pallet"].le(0)
        | order_weights["Gross weight"].lt(0)
        | order_weights["Order Quantity"].lt(0)
    )
    order_weights["Weight (lbs)"] = (
        order_weights["Gross weight"] * 2.205 + order_weights["Pallet_qty"] * 46
    )
    order_weights.loc[order_weights["Weight data issue"], "Weight (lbs)"] = np.nan

    po_weights = (
        order_weights
        .groupby("Purchase order no.", as_index=False)
        .agg({
            "Weight (lbs)": lambda weights: weights.sum() if weights.notna().all() else np.nan,
            "Weight data issue": "any"
        })
    )

    sid_counts = planex_pairs.groupby("Purchase order no.")["Shipment ID"].nunique()
    ambiguous_pos = set(sid_counts[sid_counts > 1].index)
    mapped_pos = set(planex_pairs["Purchase order no."])
    open_order_pos = set(orders["Purchase order no."])
    unmatched_pos = sorted(open_order_pos - mapped_pos)

    unique_pairs = planex_pairs[
        ~planex_pairs["Purchase order no."].isin(ambiguous_pos)
    ]
    assignments = (
        unique_pairs[unique_pairs["Purchase order no."].isin(open_order_pos)]
        .merge(po_weights, on="Purchase order no.", how="left")
        .groupby(["Shipment ID", "Purchase order no."], as_index=False)
        .agg({
            "Weight (lbs)": "first",
            "Weight data issue": "max"
        })
    )

    shipment_carriers = df_planex[["Shipment ID", "Carrier Code"]].copy()
    shipment_carriers["Shipment ID"] = shipment_carriers["Shipment ID"].astype(str).str.strip()
    shipment_carriers["Carrier Code"] = shipment_carriers["Carrier Code"].astype(str).str.strip()
    shipment_carriers = shipment_carriers[
        ~shipment_carriers["Shipment ID"].str.lower().isin({"", "nan", "none"})
        & ~shipment_carriers["Carrier Code"].str.lower().isin({"", "nan", "none"})
    ].drop_duplicates()

    carrier_types = carrier_df[["SCAC Code", "Type of Carr"]].copy()
    carrier_types["SCAC Code"] = carrier_types["SCAC Code"].astype(str).str.strip()
    carrier_types["Type of Carr"] = carrier_types["Type of Carr"].fillna("").astype(str).str.strip()
    shipment_carriers = shipment_carriers.merge(
        carrier_types.drop_duplicates(),
        left_on="Carrier Code",
        right_on="SCAC Code",
        how="left"
    )

    def format_values(values, fallback):
        unique_values = sorted({
            str(value).strip()
            for value in values
            if pd.notna(value) and str(value).strip() and str(value).strip().lower() != "nan"
        })
        if not unique_values:
            return fallback
        return ", ".join(unique_values)

    shipment_carrier_summary = (
        shipment_carriers
        .groupby("Shipment ID", as_index=False)
        .agg({
            "Carrier Code": lambda values: format_values(values, "Unmapped"),
            "Type of Carr": lambda values: format_values(values, "Unmapped")
        })
    )

    truck_summary = (
        assignments
        .groupby("Shipment ID", as_index=False)
        .agg({
            "Purchase order no.": "nunique",
            "Weight (lbs)": lambda weights: weights.sum() if weights.notna().all() else np.nan
        })
        .rename(columns={
            "Purchase order no.": "PO count",
            "Weight (lbs)": "Total weight (lbs)"
        })
    )
    truck_summary = truck_summary.merge(
        shipment_carrier_summary,
        on="Shipment ID",
        how="left"
    )
    truck_summary["Carrier Code"] = truck_summary["Carrier Code"].fillna("Unmapped")
    truck_summary["Type of Carr"] = truck_summary["Type of Carr"].fillna("Unmapped")

    both_tl_int = truck_summary["Type of Carr"].str.upper().str.contains(
        "BOTH TL/INT",
        regex=False
    )
    truck_summary["Weight limit (lbs)"] = MAX_TRUCK_WEIGHT_LBS
    truck_summary.loc[
        truck_summary["Type of Carr"].str.upper().eq("INT"),
        "Weight limit (lbs)"
    ] = INT_TRUCK_WEIGHT_LBS
    truck_summary.loc[both_tl_int, "Weight limit (lbs)"] = pd.NA
    truck_summary["Weight limit (lbs)"] = truck_summary["Weight limit (lbs)"].astype("Int64")
    overweight_int = (
        truck_summary["Type of Carr"].str.upper().eq("INT")
        & truck_summary["Total weight (lbs)"].gt(INT_TRUCK_WEIGHT_LBS)
    )
    review_both = (
        both_tl_int
        & truck_summary["Total weight (lbs)"].gt(INT_TRUCK_WEIGHT_LBS)
    )
    overweight_other = (
        ~truck_summary["Type of Carr"].str.upper().eq("INT")
        & ~both_tl_int
        & truck_summary["Total weight (lbs)"].gt(MAX_TRUCK_WEIGHT_LBS)
    )
    truck_summary["Status"] = np.select(
        [
            truck_summary["Total weight (lbs)"].isna(),
            review_both,
            overweight_int | overweight_other
        ],
        [
            "Incomplete weight data",
            "Customer Service review: carrier type",
            "Overweight"
        ],
        default="Within limit"
    )
    truck_summary.insert(
        0,
        "Approved",
        np.where(truck_summary["Status"].eq("Within limit"), "Yes", "No")
    )

    truck_split_output = build_truck_split_output(
        order_weights,
        planex_pairs,
        ambiguous_pos,
        truck_summary
    )

    issues = [
        {
            "Purchase order no.": purchase_order,
            "Shipment ID": "",
            "Issue": "Open-order PO not found in Vendor Planex"
        }
        for purchase_order in unmatched_pos
    ]
    for purchase_order in sorted(ambiguous_pos):
        shipment_ids = sorted(
            planex_pairs.loc[
                planex_pairs["Purchase order no."] == purchase_order,
                "Shipment ID"
            ].unique()
        )
        issues.append({
            "Purchase order no.": purchase_order,
            "Shipment ID": ", ".join(shipment_ids),
            "Issue": "PO maps to multiple Shipment IDs"
        })
    for _, row in assignments[assignments["Weight data issue"]].iterrows():
        issues.append({
            "Purchase order no.": row["Purchase order no."],
            "Shipment ID": row["Shipment ID"],
            "Issue": "Missing or invalid gross weight, order quantity, or pallet mapping"
        })

    issues_df = pd.DataFrame(
        issues,
        columns=["Purchase order no.", "Shipment ID", "Issue"]
    )
    assignments = assignments.drop(columns=["Weight data issue"])
    assignments["Weight (lbs)"] = assignments["Weight (lbs)"].round().astype("Int64")
    truck_summary["Total weight (lbs)"] = (
        truck_summary["Total weight (lbs)"].round().astype("Int64")
    )
    return truck_summary, issues_df, truck_split_output


def uploaded_files_signature(files):
    return tuple(
        (uploaded_file.name, hashlib.sha256(uploaded_file.getvalue()).hexdigest())
        for uploaded_file in files
    )


def load_reference_files():
    ltl_qty_df = pd.read_excel(LTL_QTY_PATH)
    ltl_qty_df["SAP Code"] = ltl_qty_df["SAP Code"].map(_normalize_open_order_value)

    rdc_list_df = pd.read_excel(
        RDC_LIST_PATH,
        dtype={"Zip Code": str, "Cust.(StP)": str, "Codice RDC": str},
        engine="openpyxl"
    )
    rdc_list_df["Cust.(StP)"] = rdc_list_df["Cust.(StP)"].map(
        _normalize_open_order_value
    )
    for c in ["RDC list", "RDC names", "Adress", "City", "State"]:
        if c in rdc_list_df.columns:
            rdc_list_df[c] = rdc_list_df[c].astype(str).str.strip()

    carrier_df = pd.read_excel(CARRIER_MAP_PATH, dtype={"Code": str}, engine="openpyxl")
    carrier_df.columns = carrier_df.columns.str.strip()
    carrier_df = carrier_df[["SCAC Code", "Carrier", "Code", "Type of Carr"]]
    carrier_df["SCAC Code"] = carrier_df["SCAC Code"].map(_normalize_open_order_value)
    carrier_df["Carrier"] = carrier_df["Carrier"].astype(str).str.strip()
    carrier_df["Code"] = carrier_df["Code"].map(_normalize_open_order_value)
    carrier_df["Type of Carr"] = carrier_df["Type of Carr"].fillna("").astype(str).str.strip()
    carrier_df = carrier_df.dropna(subset=["SCAC Code"])

    df_cube = pd.read_excel(CUBE_PATH, dtype={"Cod_articolo": str}, sheet_name="Mapping")
    df_cube["Cod_articolo"] = df_cube["Cod_articolo"].map(_normalize_open_order_value)
    df_cube["gross lbs + pallet"] = df_cube["gross lbs + pallet"].round().astype(int)
    df_cube["Cube Vendor PLX"] = df_cube["Cube Vendor PLX"].round().astype(int)
    df_cube = df_cube.rename(columns={"Cube Vendor PLX": "Cube_sku"})

    return ltl_qty_df, rdc_list_df, carrier_df, df_cube


# =========================================================
# CORE PROCESSING
# =========================================================
def process_bol_files(planex_files, order_files=None, completed_orders=None):
    ltl_qty_df, rdc_list_df, carrier_df, df_cube = load_reference_files()

    df_upload_planex = read_multiple_planex_csv(planex_files)
    if completed_orders is not None:
        df_upload = completed_orders.copy()
    else:
        df_upload = read_multiple_open_orders_excel(order_files)
    for column in ["Purchase order no.", "Customer", "Material"]:
        df_upload[column] = df_upload[column].map(_normalize_open_order_value)

    if "SID" not in df_upload.columns:
        df_upload["SID"] = pd.NA
    df_upload["SID"] = df_upload["SID"].astype("string").str.strip()
    df_upload["SID"] = df_upload["SID"].mask(
        df_upload["SID"].str.lower().isin({"", "nan", "none"})
    )

    planex_pairs = build_planex_po_pairs(df_upload_planex)
    sid_counts = planex_pairs.groupby("Purchase order no.")["Shipment ID"].nunique()
    ambiguous_pos = set(sid_counts[sid_counts > 1].index)
    sid_by_po = (
        planex_pairs[
            ~planex_pairs["Purchase order no."].isin(ambiguous_pos)
        ]
        .drop_duplicates(subset=["Purchase order no."])
        .rename(columns={"Shipment ID": "Planex SID"})
    )
    df_upload = df_upload.merge(
        sid_by_po,
        on="Purchase order no.",
        how="left"
    )
    df_upload["SID"] = df_upload["Planex SID"].fillna(df_upload["SID"])
    df_upload.loc[
        df_upload["Purchase order no."].isin(ambiguous_pos),
        "SID"
    ] = pd.NA
    df_upload = df_upload.drop(columns=["Planex SID"])
    df_upload = df_upload[df_upload["SID"].notna()].reset_index(drop=True)

    df_copy = pd.merge(
        df_upload.copy(),
        ltl_qty_df[["SAP Code", "LTL Qty", "Case_Pallet"]],
        left_on="Material",
        right_on="SAP Code",
        how="inner"
    )
    # pallet quantity
    df_copy["Pallet_qty"] = np.ceil(df_copy["Order Quantity"] / df_copy["Case_Pallet"])

    merged_df = pd.merge(
        df_copy,
        rdc_list_df,
        left_on="Customer",
        right_on="Cust.(StP)",
        how="left"
    )

    df_dest = df_upload_planex[["Shipment ID", "Destination ID", "Pickup Date", "Carrier Code", "Weight", "Cube", "Quantity"]].copy()
    df_dest = df_dest.rename(columns={"Carrier Code": "SCAC Code"})
    df_dest["SCAC Code"] = df_dest["SCAC Code"].map(_normalize_open_order_value)

    # Deduplicate to one row per SID — Weight/Cube/Carrier are shipment-level, not PO-level
    df_dest = df_dest.drop_duplicates(subset=["Shipment ID", "Destination ID"])

    upload_merged = pd.merge(
        merged_df,
        df_dest,
        left_on="SID",
        right_on="Shipment ID",
        how="left"
    )
    
    upload_merged = upload_merged.drop(columns=["Carrier"], errors="ignore")

    upload_merged = pd.merge(
        upload_merged,
        carrier_df[["SCAC Code", "Carrier"]],
        on="SCAC Code",
        how="left"
    )
    upload_merged = upload_merged.rename(columns={"Carrier": "Carrier_mapped"})

    upload_merged = pd.merge(
        upload_merged,
        df_cube[["Cod_articolo", "Cube_sku"]],
        left_on="Material",
        right_on="Cod_articolo",
        how="left"
    )
    upload_merged = upload_merged.drop(columns=["Cod_articolo"], errors="ignore")

    upload_merged["Cube_po"] = upload_merged["Cube_sku"] * upload_merged["Pallet_qty"]
    upload_merged["Cube_po"] = upload_merged["Cube_po"].fillna(0).round().astype(int)

    upload_merged = upload_merged.drop_duplicates()

    bol_group_columns = ["SID", "RDC names", "Destination ID"]
    if "Truck" in upload_merged.columns:
        bol_group_columns.append("Truck")

    df_bol = (
        upload_merged
        .groupby(bol_group_columns, as_index=False)
        .agg({
            "Carrier_mapped": "first",
            "Adress": "first",
            "City": "first",
            "State": "first",
            "Zip Code": "first",
            "SCAC Code": "first",
            "Order Quantity": "sum",
            "Gross weight": "sum", # in kg
            "Pallet_qty": "sum",
            "Weight": "first",
            "Cube": "first",
            "Pickup Date": "first",
            "Quantity": "first",
            "Purchase order no.": lambda s: ", ".join(pd.unique(s.astype(str))),
            "DN#": lambda s: ", ".join([x for x in pd.unique(s.astype(str)) if x and x.lower() != "nan"]),
        })
    )
    
    # convert weight in lbs - i do it here and not before the grouping to avoid rounding issues that might slightly change the actual total
    df_bol["Gross weight"] = df_bol["Gross weight"] * 2.205 + df_bol["Pallet_qty"]*46 ## pallet weight times count of pallets
    df_bol["Gross weight"] = df_bol["Gross weight"].fillna(0).round().astype(int)
    
    df_bol["Order Quantity"] = df_bol["Order Quantity"].astype(int)
    df_bol["Pallet_qty"] = df_bol["Pallet_qty"].astype(int)

    return df_bol, upload_merged, df_upload, df_upload_planex, df_copy, merged_df


# =========================================================
# DOCX/ZIP OUTPUT
# =========================================================
def generate_documents(df_bol, upload_merged):
    output_buffers = {}

    group_columns = ["SID"]
    if "Truck" in df_bol.columns:
        group_columns.append("Truck")

    for group_key, shipment_rows in df_bol.groupby(group_columns, sort=False):
        if len(group_columns) == 1:
            shipment_id = group_key
            truck = None
        else:
            shipment_id, truck = group_key
        sid = str(shipment_id).strip()
        bol_doc = None

        for _, row in shipment_rows.iterrows():
            replacements = {
                "{{SHIPMENT ID}}": sid,
                "{{PICKUP}}": row.get("Pickup Date", ""),
                "{{CARRIER NAME}}": row.get("Carrier_mapped", ""),
                "{{RDC NAME}}": row.get("RDC names", ""),
                "{{ADRESS}}": row.get("Adress", ""),
                "{{CITY}}": row.get("City", ""),
                "{{STATE}}": row.get("State", ""),
                "{{ZIP CODE}}": row.get("Zip Code", ""),
                "{{SCAC}}": row.get("SCAC Code", ""),
                "{{PO_LIST}}": row.get("Purchase order no.", ""),
                "{{DN_LIST}}": row.get("DN#", ""),
                "{{NUM_PACKAGES}}": row.get("Order Quantity", ""),
                "{{WEIGHT}}": row.get("Gross weight", ""),
                "{{QTY_1}}": row.get("Pallet_qty", ""),
                "{{QTY_PACK}}": row.get("Order Quantity", ""),
            }

            filled_doc = fill_template_doc(BOL_TEMPLATE_PATH, replacements)

            if bol_doc is None:
                bol_doc = filled_doc
            else:
                append_doc_content(bol_doc, filled_doc)

        if shipment_rows["Destination ID"].iloc[0] == "XD16":
            manifest_doc = create_manifest_doc(
                MANIFEST_TEMPLATE_PATH,
                sid,
                shipment_rows,
                upload_merged,
                truck
            )
            prepend_doc_content(bol_doc, manifest_doc)

        bio = BytesIO()
        bol_doc.save(bio)
        bio.seek(0)

        truck_suffix = f"_{str(truck).replace(' ', '_')}" if truck is not None else ""
        filename = f"Shipment_{sid}{truck_suffix}.docx"
        output_buffers[filename] = bio.getvalue()

    return output_buffers


def create_zip_from_docs(doc_buffers):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for filename, content in doc_buffers.items():
            zipf.writestr(filename, content)
    zip_buffer.seek(0)
    return zip_buffer


def create_excel_report(df_bol, upload_merged):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_bol.to_excel(writer, index=False, sheet_name="BOL Summary")
        upload_merged.to_excel(writer, index=False, sheet_name="Merged Detail")
    output.seek(0)
    return output


def create_truck_split_excel(truck_split_output):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        truck_split_output.to_excel(writer, index=False, sheet_name="Truck Splits")
        worksheet = writer.sheets["Truck Splits"]
        group_keys = truck_split_output[["SID", "Truck"]].reset_index(drop=True)
        spacer_rows = [
            row_number + 3
            for row_number in range(len(group_keys) - 1)
            if tuple(group_keys.iloc[row_number]) != tuple(group_keys.iloc[row_number + 1])
        ]
        for row_number in reversed(spacer_rows):
            worksheet.insert_rows(row_number)
    output.seek(0)
    return output


def build_missing_bol_reason_table(
    df_upload_planex,
    df_upload,
    df_copy,
    merged_df,
    upload_merged,
    df_bol
):
    accepted_sids = set(df_upload_planex["Shipment ID"].dropna().astype(str).str.strip().unique())
    generated_sids = set(df_bol["SID"].dropna().astype(str).str.strip().unique())
    missing_sids = sorted(accepted_sids - generated_sids)

    open_order_sids = set(df_upload["SID"].dropna().astype(str).str.strip().unique())
    after_ltl_merge_sids = set(df_copy["SID"].dropna().astype(str).str.strip().unique())
    after_planex_merge_sids = set(upload_merged["SID"].dropna().astype(str).str.strip().unique())

    rows = []

    for sid in missing_sids:
        if sid not in open_order_sids:
            reason = "No matching SID in RDC Open Orders"
        elif sid not in after_ltl_merge_sids:
            reason = "Dropped in LTL qty merge (missing material mapping)"
        elif sid not in after_planex_merge_sids:
            reason = "Missing in Planex merge"
        elif sid in after_planex_merge_sids:
            sid_subset = upload_merged[upload_merged["SID"].astype(str).str.strip() == sid]
            if sid_subset["Carrier_mapped"].isna().all():
                reason = "Missing carrier mapping (SCAC not found)"
            else:
                reason = "Present in merged data but no BOL created"
        else:
            reason = "Present in merged data but no BOL created"

        planex_subset = df_upload_planex[
            df_upload_planex["Shipment ID"].astype(str).str.strip() == sid
        ]

        destination_id = ""
        shipment_status = ""
        pickup_date = ""

        if not planex_subset.empty:
            destination_id = planex_subset["Destination ID"].iloc[0] if "Destination ID" in planex_subset.columns else ""
            shipment_status = planex_subset["Shipment Status"].iloc[0] if "Shipment Status" in planex_subset.columns else ""
            pickup_date = planex_subset["Pickup Date"].iloc[0] if "Pickup Date" in planex_subset.columns else ""

        rows.append({
            "Shipment ID": sid,
            "Reason BOL not created": reason,
            "Destination ID": destination_id,
            "Shipment Status": shipment_status,
            "Pickup Date": pickup_date
        })

    return pd.DataFrame(rows)


# =========================================================
# STREAMLIT UI
# =========================================================
st.title("📦 BOL Generator - DEV")
st.caption("Generate BOL Word documents and Master Manifests from uploaded files.")

with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    **Upload these files:**
    1. Vendor Planex export CSV file(s)
    2. RDC Open Orders Excel file(s)

    **Bundled in the app:**
    - LTL qty mapping
    - RDC list
    - Carrier mapping
    - Cube mapping
    - Word templates
    """)

st.subheader("Upload input files")

col1, col2 = st.columns(2)

with col1:
    planex_files = st.file_uploader(
        "Vendor Planex CSV files",
        type=["csv"],
        accept_multiple_files=True
    )

with col2:
    order_files = st.file_uploader(
        "RDC Open Orders Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

required_assets = [
    LTL_QTY_PATH,
    RDC_LIST_PATH,
    CARRIER_MAP_PATH,
    CUBE_PATH,
    BOL_TEMPLATE_PATH,
    MANIFEST_TEMPLATE_PATH
]

missing_assets = [p.name for p in required_assets if not p.exists()]

if missing_assets:
    st.error("Missing required files in the app folder:")
    for f in missing_assets:
        st.write(f"- {f}")
    st.stop()

st.subheader("Step 1: Analyze truck loads")
st.caption(
    "Match each open-order PO to its Vendor Planex Shipment ID and check the total "
    "against the mapped carrier type's weight rule. Shipment status is not used."
)

input_signature = (
    TRUCK_ANALYSIS_VERSION,
    uploaded_files_signature(planex_files or []),
    uploaded_files_signature(order_files or [])
)
if st.session_state.get("truck_analysis_signature") != input_signature:
    st.session_state.pop("truck_analysis_signature", None)
    st.session_state.pop("truck_analysis_result", None)

analyze = st.button("Analyze Truck Loads", type="primary")
if analyze:
    if not planex_files:
        st.warning("Please upload at least one Vendor Planex CSV.")
    elif not order_files:
        st.warning("Please upload at least one RDC Open Orders Excel file.")
    else:
        try:
            with st.spinner("Matching POs and calculating truck weights..."):
                ltl_qty_df, _, carrier_df, _ = load_reference_files()
                df_all_planex = read_multiple_planex_csv(planex_files, accepted_only=False)
                df_all_orders = read_multiple_open_orders_excel(
                    order_files,
                    bol_rows_only=False
                )
                (
                    truck_summary,
                    truck_issues,
                    truck_split_output
                ) = build_truck_load_summary(
                    df_all_planex,
                    df_all_orders,
                    ltl_qty_df,
                    carrier_df
                )

            st.session_state["truck_analysis_signature"] = input_signature
            st.session_state["truck_analysis_result"] = (
                truck_summary,
                truck_issues,
                truck_split_output
            )
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.exception(e)

truck_analysis_result = st.session_state.get("truck_analysis_result")
has_current_analysis = (
    st.session_state.get("truck_analysis_signature") == input_signature
    and isinstance(truck_analysis_result, tuple)
    and len(truck_analysis_result) == 3
)
if (
    st.session_state.get("truck_analysis_signature") == input_signature
    and not has_current_analysis
):
    st.session_state.pop("truck_analysis_signature", None)
    st.session_state.pop("truck_analysis_result", None)
    st.info("Please click Analyze Truck Loads to refresh the truck analysis.")

if has_current_analysis:
    (
        truck_summary,
        truck_issues,
        truck_split_output
    ) = truck_analysis_result
    if truck_summary.empty:
        st.info("No open-order POs could be assigned to a truck.")
    else:
        overweight_count = truck_summary["Status"].eq("Overweight").sum()
        carrier_review_count = truck_summary["Status"].eq(
            "Customer Service review: carrier type"
        ).sum()
        incomplete_count = truck_summary["Status"].eq("Incomplete weight data").sum()
        if overweight_count:
            st.warning(f"{overweight_count} truck(s) exceed their carrier weight limit.")
        if carrier_review_count:
            st.warning(
                f"{carrier_review_count} truck(s) exceed {INT_TRUCK_WEIGHT_LBS:,} lbs "
                "with a BOTH TL/INT carrier; Customer Service should review the carrier type."
            )
        if incomplete_count:
            st.warning(
                f"{incomplete_count} truck(s) have incomplete weight data; review the PO issues."
            )
        if not overweight_count and not carrier_review_count and not incomplete_count:
            if truck_issues.empty:
                st.success("All assigned trucks are within the weight limit.")
            else:
                st.info(
                    "Calculated truck loads are within the limit; review the PO issues "
                    "before generating BOLs."
                )
        st.dataframe(truck_summary, use_container_width=True)

    st.subheader("Truck split order lines")
    st.caption(
        "The check table above shows the original loads. Below, overweight loads are "
        "repacked from whole order lines as close to each truck's limit as possible "
        "without exceeding it. BOTH TL/INT loads remain assigned for Customer Service review."
    )
    if not truck_split_output.empty:
        unassigned_rows = truck_split_output["Truck"].eq("Unassigned").sum()
        if unassigned_rows:
            st.warning(
                f"{unassigned_rows} order line(s) could not be assigned to a Shipment ID; "
                "review the PO matching issues."
            )
        truck_weight_limits = truck_summary.set_index("Shipment ID")["Weight limit (lbs)"]
        split_truck_weights = (
            truck_split_output[truck_split_output["Truck"] != "Unassigned"]
            .groupby(["SID", "Truck"], as_index=False)
            .agg({
                "Gross weight (lbs)": lambda weights: (
                    weights.sum() if weights.notna().all() else np.nan
                )
            })
            .rename(columns={"SID": "Shipment ID"})
        )
        split_truck_weights["Weight limit (lbs)"] = split_truck_weights[
            "Shipment ID"
        ].map(truck_weight_limits)
        remaining_overweight = split_truck_weights[
            split_truck_weights["Gross weight (lbs)"].notna()
            & split_truck_weights["Weight limit (lbs)"].notna()
            & split_truck_weights["Gross weight (lbs)"].gt(
                split_truck_weights["Weight limit (lbs)"]
            )
        ]
        if not remaining_overweight.empty:
            st.warning(
                "Some split trucks still exceed the limit because at least one order line "
                "is too heavy to fit on a compliant truck."
            )
            st.dataframe(remaining_overweight, use_container_width=True)
        st.dataframe(truck_split_output, use_container_width=True)
        truck_split_excel = create_truck_split_excel(truck_split_output)
        st.download_button(
            label="Download truck split table",
            data=truck_split_excel,
            file_name="Truck_splits.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_truck_splits"
        )
    else:
        st.info("No open-order lines were available for the truck split table.")

    if not truck_issues.empty:
        st.warning(f"{len(truck_issues)} PO-level issue(s) need review.")
        with st.expander("PO matching and weight issues"):
            st.dataframe(truck_issues, use_container_width=True)

st.subheader("Step 2: Complete and upload DN# values")
st.caption(
    "Download the truck split workbook, fill in DN# for every order line, then upload "
    "the completed workbook here. The completed workbook will be used to generate BOLs."
)
completed_truck_file = st.file_uploader(
    "Completed truck split workbook",
    type=["xlsx"],
    accept_multiple_files=False,
    key="completed_truck_split_file"
)

completed_orders = None
completed_order_issues = pd.DataFrame()
completed_workbook_ready = False
if completed_truck_file is not None:
    try:
        completed_orders, completed_order_issues = read_completed_truck_split_excel(
            completed_truck_file
        )
        completed_workbook_ready = not completed_orders.empty
        if not completed_order_issues.empty:
            st.warning(
                f"{len(completed_order_issues)} incomplete order line(s) will be skipped. "
                "BOLs will still be generated for eligible lines."
            )
            st.dataframe(completed_order_issues, use_container_width=True)
        if completed_workbook_ready:
            st.success(
                f"{len(completed_orders)} complete order line(s) are ready for BOL generation."
            )
        else:
            st.warning(
                "No complete order lines are available. Add a valid SID, Truck, and DN# "
                "to at least one line to enable BOL generation."
            )
    except ValueError as e:
        st.warning(str(e))
    except Exception as e:
        st.exception(e)

st.subheader("Step 3: Generate BOL files")
run = st.button(
    "Generate BOL Files",
    disabled=not has_current_analysis or not completed_workbook_ready
)

if run:
    if not planex_files:
        st.warning("Please upload at least one Vendor Planex CSV.")
        st.stop()

    if not order_files:
        st.warning("Please upload at least one RDC Open Orders Excel file.")
        st.stop()

    try:
        with st.spinner("Processing files and generating BOLs..."):
            if not completed_order_issues.empty:
                st.warning(
                    f"Skipping {len(completed_order_issues)} incomplete order line(s); "
                    "generating BOLs for the remaining complete lines."
                )
            df_bol, upload_merged, df_upload, df_upload_planex, df_copy, merged_df = process_bol_files(
                planex_files,
                completed_orders=completed_orders
            )
            doc_buffers = generate_documents(df_bol, upload_merged)
            zip_buffer = create_zip_from_docs(doc_buffers)
            excel_report = create_excel_report(df_bol, upload_merged)

        st.success(f"Done. Generated {len(doc_buffers)} shipment document(s).")

        docx_generated = len(doc_buffers)
        unique_sids = df_upload["SID"].dropna().nunique()
        unique_pos = df_upload["Purchase order no."].dropna().nunique()
        unique_dns = df_upload["DN#"].dropna().nunique()

        ifc_pos = (
            upload_merged.loc[upload_merged["Destination ID"] == "XD16", "Purchase order no."]
            .dropna()
            .nunique()
        )

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("DOCX generated", docx_generated)
        k2.metric("Unique SIDs", unique_sids)
        k3.metric("Unique POs", unique_pos)
        k4.metric("Unique Deliveries", unique_dns)
        k5.metric("IFC POs", ifc_pos)

        warning_df = build_missing_bol_reason_table(
            df_upload_planex=df_upload_planex,
            df_upload=df_upload,
            df_copy=df_copy,
            merged_df=merged_df,
            upload_merged=upload_merged,
            df_bol=df_bol
        )

        if not warning_df.empty:
            st.warning(f"{len(warning_df)} accepted Shipment ID(s) did not generate a BOL.")
            st.dataframe(warning_df, use_container_width=True)
        else:
            st.info("All accepted Shipment IDs generated a BOL.")

        ## Warning if Destination ID is different than known ones
        valid_dest = df_bol["Destination ID"].astype(str).str.strip()
        unexpected_dest = df_bol[
            ~valid_dest.eq("XD16") & ~valid_dest.str.match(r"^\d+$")
        ][["SID", "Destination ID", "Carrier_mapped", "RDC names", "Purchase order no."]]

        if not unexpected_dest.empty:
            st.warning(
                f"⚠️ {len(unexpected_dest)} shipment(s) have an unexpected Destination ID "
                f"(neither 'XD16' nor a numeric RDC code):"
            )
            st.dataframe(unexpected_dest, use_container_width=True)

        with st.expander("BOL Summary"):
            st.dataframe(df_bol, use_container_width=True)

        st.download_button(
            label="Download ZIP of BOL files",
            data=zip_buffer,
            file_name="BOL_created.zip",
            mime="application/zip"
        )

        st.download_button(
            label="Download Excel report",
            data=excel_report,
            file_name="BOL_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.exception(e)
