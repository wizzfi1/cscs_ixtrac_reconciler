# reconcile.py

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from core.loader import load_excel
from core.normalizer import normalize_name, first_two_names
from core.matcher import build_cscs_index, build_cscs_index_2name
from core.mapping import load_mappings, validate_mapping, resolve_columns
from core.engine import match_row
from core.duplicates import detect_duplicates

from config.rules import (
    STATUS_PRIORITY,
    STATUS_POSITION_CSCS,
    STATUS_MORE_THAN_5,
)


# =================================================
# Helper: write DataFrame to openpyxl workbook safely
# =================================================
def write_df_to_sheet(wb, sheet_name, df):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)


# =================================================
# Helper: decide what shows in MATCH_STATUS column
# =================================================
def resolve_display_status(decision):
    """
    Certain validation failures must be visible directly
    in the MATCH_STATUS column.
    """
    if decision.reason == STATUS_POSITION_CSCS:
        return STATUS_POSITION_CSCS

    if decision.reason == STATUS_MORE_THAN_5:
        return STATUS_MORE_THAN_5

    return decision.status


# =================================================
# Main reconciliation entry point
# =================================================
def run_reconciliation(file_path: str, mapping_name: str):
    os.makedirs("output", exist_ok=True)
    output_path = "output/IXTRAC_RECONCILED.xlsx"

    mappings = load_mappings()
    if mapping_name not in mappings:
        raise ValueError(f"Unknown mapping: {mapping_name}")

    mapping = mappings[mapping_name]

    # =================================================
    # LOAD CSCS
    # =================================================
    cscs = load_excel(file_path, mapping["cscs_sheet"])
    cscs_name_col = mapping.get("cscs_name", "NAME")

    if cscs_name_col not in cscs.columns:
        raise ValueError("CSCS name column missing")

    cscs["NORM_NAME"] = cscs[cscs_name_col].apply(normalize_name)
    cscs["FIRST2"] = cscs[cscs_name_col].apply(first_two_names)

    exact_index = build_cscs_index(cscs)
    two_name_index = build_cscs_index_2name(cscs)

    # =================================================
    # CSCS DUPLICATE DETECTION
    # =================================================
    duplicates_df = detect_duplicates(
        cscs,
        ["NORM_NAME", "CHN"]
    )

    # =================================================
    # LOAD IX TRAC (OPEN ONCE)
    # =================================================
    wb = load_workbook(file_path)
    sheet = wb[mapping["ixtrac_sheet"]]

    validate_mapping(sheet, mapping)
    cols = resolve_columns(sheet, mapping)

    decisions = []

    # =================================================
    # RECONCILIATION LOOP
    # =================================================
    for r in range(2, sheet.max_row + 1):
        name = sheet.cell(r, cols["name"]).value
        chn = sheet.cell(r, cols["chn"]).value

        decision = match_row(
            name,
            chn,
            exact_index,
            two_name_index
        )

        sheet.cell(r, cols["membercode"]).value = decision.membercode or ""
        sheet.cell(r, cols["status"]).value = resolve_display_status(decision)

        decisions.append({
            "ROW": r,
            "NAME": name,
            "CHN": chn,
            "STATUS": decision.status,
            "DISPLAY_STATUS": resolve_display_status(decision),
            "MEMBERCODE": decision.membercode,
            "REASON": decision.reason,
        })

    # SAVE ENRICHED IX TRAC FIRST
    wb.save(output_path)

    # =================================================
    # BUILD REVIEW / SUMMARY DATAFRAMES
    # =================================================
    ix_df = pd.read_excel(output_path, sheet_name=mapping["ixtrac_sheet"])
    status_col = mapping["status_out"]

    review_df = ix_df.copy()
    review_df["__rank"] = review_df[status_col].map(STATUS_PRIORITY)
    review_df = review_df.sort_values("__rank").drop(columns="__rank")

    summary_df = (
        review_df[status_col]
        .value_counts()
        .reset_index()
        .rename(columns={"index": "STATUS", status_col: "COUNT"})
    )

    summary_df.loc[len(summary_df)] = {
        "STATUS": "TOTAL_ROWS",
        "COUNT": len(review_df),
    }

    decision_df = pd.DataFrame(decisions)

    # =================================================
    # WRITE EXTRA SHEETS (NO pandas.ExcelWriter)
    # =================================================
    write_df_to_sheet(wb, "IX_TRAC_REVIEW", review_df)
    write_df_to_sheet(wb, "RECONCILIATION_SUMMARY", summary_df)
    write_df_to_sheet(wb, "DECISION_LOG", decision_df)
    write_df_to_sheet(wb, "CSCS_DUPLICATES", duplicates_df)

    wb.save(output_path)

    print("✔ Reconciliation complete")
    print(f"✔ Output written to {output_path}")
