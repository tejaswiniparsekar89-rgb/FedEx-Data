import io
import sys
import traceback
import pandas as pd
import streamlit as st

APP_TITLE = "FedEx report automation"
PHASE = "Phase 1: Shipment Data + Criteria (+ LOC test data pivots helper)"

REQUIRED_RAW_COLUMNS = [
    "Carrier Name",
    "Order Number",
    "Active Equipment ID",
    "Historical Equipment ID",
]

# ---------- Helpers ----------

def read_any_table(uploaded_file, expected_sheet_name=None):
    """
    Reads CSV or Excel. If Excel and expected_sheet_name is provided and present, loads that sheet;
    otherwise loads the first sheet. Returns (df, error_message or None).
    """
    if uploaded_file is None:
        return None, "No file uploaded."
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            return df, None
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            xls = pd.ExcelFile(uploaded_file)
            sheet_to_read = expected_sheet_name if (expected_sheet_name and expected_sheet_name in xls.sheet_names) else xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=sheet_to_read)
            return df, None
        else:
            return None, "Unsupported file type. Please upload .csv or .xlsx"
    except Exception as e:
        return None, f"Failed to read file: {e}"

def validate_columns(df, needed, label):
    missing = [c for c in needed if c not in df.columns]
    if missing:
        return f"{label} is missing required column(s): {', '.join(missing)}"
    return None

def fill_loc_test_data(loc_test_df: pd.DataFrame) -> pd.DataFrame:
    """
    LOC test data sheet structure:
      - A:E contains main mapping table:
        A Carrier Name, B LOC, C TRUE, D FALSE, E Grand Total
      - I:N contains region sub tables where I & J are populated and K:N must be filled:
        I Carrier Name, J LOC, K Tracked, L Not Tracked, M Grand Total, N Tracked%
    This function matches (Carrier Name, LOC) from I:J to A:B and fills K:N.
    """
    df = loc_test_df.copy()

    # Expect columns by position (since your sheet is structured like Excel columns)
    # But we’ll also support cases where headers are present.
    # If headers are present and not exactly as expected, we’ll still use positions.

    # --- Build base map from A:E (0..4) ---
    if df.shape[1] < 14:
        # Need at least up to column N (index 13)
        raise ValueError("LOC test data sheet does not have enough columns (needs up to column N).")

    base = df.iloc[:, 0:5].copy()
    base.columns = ["Carrier Name", "LOC", "TRUE", "FALSE", "Grand Total"]

    # Clean keys for matching
    base["Carrier Name_key"] = base["Carrier Name"].astype(str).str.strip().str.upper()
    base["LOC_key"] = base["LOC"].astype(str).str.strip().str.upper()

    # Keep only rows that look like real data
    base = base[
        base["Carrier Name"].notna()
        & base["LOC"].notna()
        & (base["Carrier Name"].astype(str).str.strip() != "")
        & (base["LOC"].astype(str).str.strip() != "")
    ].copy()

    # If TRUE/FALSE/Grand Total are empty strings, coerce safely to numeric
    for col in ["TRUE", "FALSE", "Grand Total"]:
        base[col] = pd.to_numeric(base[col], errors="coerce")

    base_map = base[["Carrier Name_key", "LOC_key", "TRUE", "FALSE", "Grand Total"]].drop_duplicates()

    # --- Subtable keys from I:J (8..9) ---
    sub = df.iloc[:, 8:10].copy()
    sub.columns = ["Carrier Name_sub", "LOC_sub"]
    sub["Carrier Name_key"] = sub["Carrier Name_sub"].astype(str).str.strip().str.upper()
    sub["LOC_key"] = sub["LOC_sub"].astype(str).str.strip().str.upper()

    # Identify rows where subtable has both keys
    valid_mask = (
        sub["Carrier Name_sub"].notna()
        & sub["LOC_sub"].notna()
        & (sub["Carrier Name_sub"].astype(str).str.strip() != "")
        & (sub["LOC_sub"].astype(str).str.strip() != "")
    )

    # Merge to get TRUE/FALSE/Total
    merged = sub.loc[valid_mask, ["Carrier Name_key", "LOC_key"]].merge(
        base_map,
        on=["Carrier Name_key", "LOC_key"],
        how="left"
    )

    # Write into K:L:M:N (10..13)
    # Default blanks if no match
    df.iloc[:, 10] = df.iloc[:, 10]  # ensure exists
    df.iloc[:, 11] = df.iloc[:, 11]
    df.iloc[:, 12] = df.iloc[:, 12]
    df.iloc[:, 13] = df.iloc[:, 13]

    # Assign matched values back to the same row indexes
    target_idx = df.index[valid_mask]
    df.loc[target_idx, df.columns[10]] = merged["TRUE"].values                 # K Tracked
    df.loc[target_idx, df.columns[11]] = merged["FALSE"].values                # L Not Tracked
    df.loc[target_idx, df.columns[12]] = merged["Grand Total"].values          # M Grand Total

    # N Tracked% = TRUE / Grand Total (safe)
    tracked_pct = merged["TRUE"] / merged["Grand Total"]
    tracked_pct = tracked_pct.replace([pd.NA, pd.NaT, float("inf"), float("-inf")], pd.NA)

    df.loc[target_idx, df.columns[13]] = tracked_pct.values

    return df

def build_workbook(raw_df: pd.DataFrame, criteria_df: pd.DataFrame, loc_test_df: pd.DataFrame | None = None) -> bytes:
    """
    Create the Excel workbook in-memory with:
      - Sheet 'Shipment Data' as an Excel Table with formulas
      - Sheet 'Criteria' as provided
      - Optional: Sheet 'LOC test data' filled for K:N using A:E lookup
    """
    out_df = raw_df.copy()
    for col in ["Trailer", "Network", "LOC"]:
        if col in out_df.columns:
            out_df.drop(columns=[col], inplace=True)

    out_df["Trailer"] = ""
    out_df["Network"] = ""
    out_df["LOC"] = ""

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Criteria
        criteria_df.to_excel(writer, sheet_name="Criteria", index=False)

        # Shipment Data
        out_df.to_excel(writer, sheet_name="Shipment Data", index=False, startrow=0, startcol=0)

        wb = writer.book
        ws = writer.sheets["Shipment Data"]

        n_rows, n_cols = out_df.shape
        first_row, first_col = 0, 0
        last_row = n_rows
        last_col = n_cols - 1

        cols_meta = []
        for c in out_df.columns[:-3]:
            cols_meta.append({"header": c})

        trailer_formula = (
            '=IF(OR(LEFT([@[Active Equipment ID]],6)="861861",'
            'LEFT([@[Active Equipment ID]],5)="86355"),'
            '"FedEx Trailer",'
            'IF(OR(UPPER(TRIM([@[Active Equipment ID]]))="UNKNOWN",'
            'UPPER(TRIM([@[Active Equipment ID]]))="UNKOWN"),'
            'IF(OR(LEFT([@[Historical Equipment ID]],6)="861861",'
            'LEFT([@[Historical Equipment ID]],5)="86355"),'
            '"FedEx Trailer","Subco Trailer"),'
            '"Subco Trailer"))'
        )

        network_formula = '=IFERROR(LEFT([@[Order Number]],3),"")'
        loc_formula = '=IFERROR(VLOOKUP([@[Carrier Name]],Criteria!$A:$D,4,0),"")'

        cols_meta.append({"header": "Trailer", "formula": trailer_formula})
        cols_meta.append({"header": "Network", "formula": network_formula})
        cols_meta.append({"header": "LOC", "formula": loc_formula})

        ws.add_table(
            first_row,
            first_col,
            last_row,
            last_col,
            {
                "name": "ShipmentDataTable",
                "columns": cols_meta,
                "style": "Table Style Medium 2",
                "banded_rows": True,
                "banded_columns": False,
                "first_column": False,
                "last_column": False,
            },
        )

        # Widths (best effort)
        try:
            width_map = {
                "Carrier Name": 24,
                "Order Number": 18,
                "Active Equipment ID": 20,
                "Historical Equipment ID": 22,
                "Trailer": 16,
                "Network": 10,
                "LOC": 10,
            }
            headers = list(out_df.columns)
            for idx, h in enumerate(headers):
                ws.set_column(idx, idx, width_map.get(h, 14))
        except Exception:
            pass

        # -------- NEW: LOC test data sheet filled --------
        if loc_test_df is not None:
            filled_loc_test = fill_loc_test_data(loc_test_df)

            # Write as-is to a dedicated sheet
            # (If you need to keep original formatting exactly, we’d switch to openpyxl templating,
            # but for pivot prep this table output usually works fine.)
            filled_loc_test.to_excel(writer, sheet_name="LOC test data", index=False)

    output.seek(0)
    return output.getvalue()

# ---------- App ----------

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="centered")
    st.title(APP_TITLE)
    st.caption(PHASE)

    st.markdown("**Inputs required:**")
    col1, col2 = st.columns(2)
    with col1:
        raw_file = st.file_uploader(
            "Upload Raw Export (A:AW) — CSV or Excel",
            type=["csv", "xlsx", "xls"],
            key="raw_upload",
            accept_multiple_files=False
        )
    with col2:
        criteria_file = st.file_uploader(
            "Upload Criteria (with columns A:D)",
            type=["csv", "xlsx", "xls"],
            key="criteria_upload",
            accept_multiple_files=False
        )

    st.markdown("**Optional:**")
    loc_test_file = st.file_uploader(
        "Upload LOC test data template (to fill region subtables K:N based on A:E)",
        type=["csv", "xlsx", "xls"],
        key="loc_test_upload",
        accept_multiple_files=False
    )

    if st.button("Generate Excel", type="primary", use_container_width=True):
        try:
            raw_df, err1 = read_any_table(raw_file)
            criteria_df, err2 = read_any_table(criteria_file, expected_sheet_name="Criteria")

            loc_test_df = None
            if loc_test_file is not None:
                loc_test_df, err3 = read_any_table(loc_test_file, expected_sheet_name="LOC test data")
                if err3:
                    st.warning(
                        "LOC test data note: If your workbook doesn't have a 'LOC test data' sheet, "
                        "the first sheet was used automatically."
                    )

            if err1:
                st.error(f"Raw file error: {err1}")
                return
            if err2:
                st.warning(
                    "Criteria note: If your workbook does not have a 'Criteria' sheet, "
                    "the first sheet was used automatically."
                )

            if raw_df is None:
                st.error("Raw file could not be read.")
                return
            if criteria_df is None:
                st.error("Criteria file could not be read.")
                return

            err = validate_columns(raw_df, REQUIRED_RAW_COLUMNS, "Raw export")
            if err:
                st.error(err)
                with st.expander("Columns found in Raw export"):
                    st.write(list(raw_df.columns))
                return

            if len(raw_df.columns) > 49:
                st.info(
                    "Heads up: Raw export contains more than 49 columns (A:AW). "
                    "That's okay; Trailer/Network/LOC will still be appended at the end."
                )

            excel_bytes = build_workbook(raw_df, criteria_df, loc_test_df=loc_test_df)

            st.success("Excel built successfully.")
            st.download_button(
                label="Download Final Excel",
                data=excel_bytes,
                file_name="FedEx_Report_Automation_Phase1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception:
            st.error("Something went wrong while generating the Excel.")
            st.code("".join(traceback.format_exception(*sys.exc_info())))

if __name__ == "__main__":
    main()
