from __future__ import annotations

from io import BytesIO
import zipfile
import re
import copy
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Client Audit Reporting", page_icon="📊", layout="wide")

REPORT_SHEETS = {
    "2026 Indy": {
        "description": "E-Cig - No ID audits for IND, BAT, C0 and WHSmith sites.",
        "item_to_order": "E-Cig - No ID",
        "filter": "indy",
    },
    "2026 Qcomm": {
        "description": "Rapid Delivery audits at sites already listed in the Qcomm sheet.",
        "item_to_order": "Rapid Delivery",
        "filter": "site_match",
    },
    "2026 16-17 YOs": {
        "description": "Home Delivery audits at sites already listed in the 16-17 YOs sheet.",
        "item_to_order": "Home Delivery",
        "filter": "site_match",
    },
}

REQUIRED_SOURCE_COLUMNS = {"internal_id", "site_name", "item_to_order", "site_code"}


def norm(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def get_headers(ws):
    return [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]


def get_site_names(ws) -> set[str]:
    # Existing report sheets use site_name in column B.
    return {
        norm(ws.cell(row, 2).value)
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, 2).value not in (None, "")
    }


def is_whsmith(row) -> bool:
    site_name = str(row.get("site_name", ""))
    site_code = str(row.get("site_code", ""))
    return bool(re.search(r"wh\s*smith", site_name, re.I) or
                re.search(r"wh\s*smith", site_code, re.I))


def filter_for_sheet(df: pd.DataFrame, sheet_name: str, ws):
    cfg = REPORT_SHEETS[sheet_name]
    mask = df["item_to_order"].astype(str).str.strip().eq(cfg["item_to_order"])

    if cfg["filter"] == "indy":
        code = df["site_code"].astype(str).str.strip().str.upper()
        prefixes = code.str.startswith(("IND", "BAT", "C0"))
        whsmith = df.apply(is_whsmith, axis=1)
        mask = mask & (prefixes | whsmith)
        return df.loc[mask].copy(), df.loc[~mask].copy()

    sites = get_site_names(ws)
    names = df["site_name"].map(norm)
    selected = mask & names.isin(sites)

    # For 16-17 YOs Home Delivery, include sibling audits sharing the same
    # order_internal_id as an audit already eligible by site name. This captures
    # multi-audit orders (e.g. Totally Wicked) without broadening other reports.
    if sheet_name == "2026 16-17 YOs" and "order_internal_id" in df.columns:
        order_ids = set(
            df.loc[selected, "order_internal_id"].astype(str).str.strip()
            .loc[lambda values: values.ne("")]
        )
        order_match = df["order_internal_id"].astype(str).str.strip().isin(order_ids)
        selected = selected | (mask & order_match)

    return df.loc[selected].copy(), df.loc[mask & ~selected].copy()


def set_cell_value(cell, value, header):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        cell.value = None
        return
    if header == "date_of_visit":
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
        cell.value = parsed.to_pydatetime() if pd.notna(parsed) else str(value)
    elif header == "time_of_visit":
        parsed = pd.to_datetime(str(value), errors="coerce")
        cell.value = parsed.time() if pd.notna(parsed) else str(value)
    else:
        cell.value = str(value) if not isinstance(value, (int, float, bool)) else value


def append_rows(ws, records: pd.DataFrame) -> tuple[int, int]:
    headers = get_headers(ws)
    existing_ids = {
        str(ws.cell(row, 1).value).strip()
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, 1).value not in (None, "")
    }
    # Use source internal_id to avoid adding the same audit twice.
    pending = records.loc[~records["internal_id"].astype(str).str.strip().isin(existing_ids)]
    added = 0

    for _, record in pending.iterrows():
        row_num = ws.max_row + 1
        template_row = row_num - 1 if row_num > 2 else 2

        for col_num, header in enumerate(headers, start=1):
            if not header:
                continue
            key = str(header).strip()
            value = record.get(key, "")

            # Handle legacy workbook question wording while keeping workbook headers unchanged.
            legacy_map = {
                "Was your transaction allowed to go through?":
                    "Was your transaction allowed to go through initially?",
                "Please detail how your purchase was refused:":
                    "Please detail how your purchase was refused during checkout:",
                "Please upload a photo of your full receipt showing any visible codes:":
                    "Please upload a screenshot of your receipt image which includes your order number:",
            }
            if (value == "" or value is None) and key in legacy_map:
                value = record.get(legacy_map[key], "")

            cell = ws.cell(row=row_num, column=col_num)
            set_cell_value(cell, value, key)

            if template_row >= 2 and template_row <= ws.max_row:
                source_cell = ws.cell(row=template_row, column=col_num)
                if source_cell.has_style:
                    cell._style = copy.copy(source_cell._style)
                if source_cell.number_format:
                    cell.number_format = source_cell.number_format
                if source_cell.alignment:
                    cell.alignment = copy.copy(source_cell.alignment)
                if source_cell.protection:
                    cell.protection = copy.copy(source_cell.protection)
        added += 1

    return added, len(pending)



def update_pass_percent_sheet(wb, df: pd.DataFrame):
    """Update monthly Pass % metrics for months represented in the current export."""
    ws = wb["Pass %"]
    # Each section maps to the row containing Visits, Pass, Fail, Abort, Pass Rate.
    sections = {
        "2026 16-17 YOs": {"start_row": 5, "item": "Home Delivery", "kind": "site_match"},
        "2026 Qcomm": {"start_row": 13, "item": "Rapid Delivery", "kind": "site_match"},
        "2026 Indy": {"start_row": 21, "item": "E-Cig - No ID", "kind": "indy"},
    }

    if "date_of_visit" not in df.columns or "primary_result" not in df.columns:
        raise ValueError("The audit export must contain date_of_visit and primary_result to update Pass %.")

    dates = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce")
    month_columns = {}
    for col in range(2, ws.max_column + 1):
        month = str(ws.cell(4, col).value or "").strip().upper()
        if month:
            month_columns[month[:3]] = col

    summary = {}
    for sheet_name, cfg in sections.items():
        if sheet_name not in wb.sheetnames:
            continue
        selected, _ = filter_for_sheet(df, sheet_name, wb[sheet_name])
        selected = selected.copy()
        selected["_parsed_date"] = pd.to_datetime(
            selected["date_of_visit"], dayfirst=True, errors="coerce"
        )
        selected["_month"] = selected["_parsed_date"].dt.strftime("%b").str.upper()
        selected["_result"] = selected["primary_result"].astype(str).str.strip().str.casefold()

        # Only touch month columns for which this export has relevant records.
        months_in_export = sorted(
            set(selected.loc[selected["_parsed_date"].notna(), "_month"]),
            key=lambda m: list(month_columns).index(m) if m in month_columns else 99
        )
        for month in months_in_export:
            if month not in month_columns:
                continue
            col = month_columns[month]
            month_rows = selected[selected["_month"] == month]
            passes = int(month_rows["_result"].isin(["pass", "passed"]).sum())
            fails = int(month_rows["_result"].isin(["fail", "failed"]).sum())
            aborts = int(month_rows["_result"].isin(["abort", "aborted"]).sum())
            visits = passes + fails + aborts
            base = cfg["start_row"]
            ws.cell(base, col).value = visits
            ws.cell(base + 1, col).value = passes
            ws.cell(base + 2, col).value = fails
            ws.cell(base + 3, col).value = aborts
            letter = ws.cell(1, col).column_letter
            ws.cell(base + 4, col).value = (passes / (passes + fails)) if (passes + fails) else 0
            summary[f"{sheet_name} - {month}"] = {
                "Visits": visits, "Pass": passes, "Fail": fails, "Abort": aborts
            }
    return summary


def workbook_bytes(template_file, df, sainsbury_template=None, tesco_template=None, coop_template=None, waitrose_template=None):
    template_file.seek(0)
    wb = load_workbook(template_file)
    results = {}
    unmatched = {}

    for sheet_name in REPORT_SHEETS:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Template is missing the required worksheet: {sheet_name}")
        selected, unmatched_rows = filter_for_sheet(df, sheet_name, wb[sheet_name])
        # Indy E-Cig - No ID fallback: only add rows not eligible for another client.
        if sheet_name == "2026 Indy":
            year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
            item = df["item_to_order"].astype(str).str.strip().str.casefold()
            other_masks = []
            if sainsbury_template is not None:
                other_masks.extend([sainsbury_online_mask(df) & year_2026, sainsbury_retail_mask(df)])
            if tesco_template is not None:
                other_masks.extend([tesco_online_mask(df), tesco_retail_mask(df)])
            if coop_template is not None:
                other_masks.append(coop_mask(df))
            if waitrose_template is not None:
                other_masks.extend([waitrose_online_mask(df), waitrose_retail_mask(df)])
            matched_elsewhere = pd.concat(other_masks, axis=1).any(axis=1) if other_masks else pd.Series(False, index=df.index)
            fallback = year_2026 & item.eq("e-cig - no id") & ~matched_elsewhere
            selected = df.loc[selected.index.isin(df.index) & (selected.index.to_series().isin(selected.index))].copy() if False else selected
            selected = pd.concat([selected, df.loc[fallback]], axis=0).drop_duplicates(subset=["internal_id"])
            unmatched_rows = unmatched_rows.loc[
                ~unmatched_rows["internal_id"].astype(str).isin(selected["internal_id"].astype(str))
            ]
        added, duplicates_removed = append_rows(wb[sheet_name], selected)
        results[sheet_name] = {
            "matching_source_records": len(selected),
            "new_rows_added": added,
            "already_present": len(selected) - added,
        }
        unmatched[sheet_name] = unmatched_rows

    pass_summary = update_pass_percent_sheet(wb, df)

    output = BytesIO()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(output)
    output.seek(0)
    return output, results, unmatched, pass_summary



def sainsbury_online_mask(df: pd.DataFrame):
    site_names = df["site_name"].astype(str)
    sainsbury_name = site_names.str.contains(r"sainsbur(?:y|ies)", case=False, regex=True, na=False)
    return df["item_to_order"].astype(str).str.strip().eq("Home Delivery") & sainsbury_name


def sainsbury_retail_mask(df: pd.DataFrame):
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    retail_code = df["site_code"].astype(str).str.strip().str.upper().str.startswith("SAI")
    online_items = df["item_to_order"].astype(str).str.strip().str.casefold().isin(
        {"home delivery", "rapid delivery", "compliance audit"}
    )
    return year_2026 & retail_code & ~online_items


def append_sainsbury_rows(ws, records: pd.DataFrame, retail=False):
    headers = get_headers(ws)
    id_header = "ID" if retail else "internal_id"
    existing_ids = {
        str(ws.cell(row, 1).value).strip()
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, 1).value not in (None, "")
    }
    pending = records.loc[~records["internal_id"].astype(str).str.strip().isin(existing_ids)]
    for _, record in pending.iterrows():
        row_num = ws.max_row + 1
        template_row = row_num - 1
        for col_num, header in enumerate(headers, start=1):
            if not header:
                continue
            key = str(header).strip()
            source_key = "internal_id" if key == id_header else key
            value = record.get(source_key, "")
            cell = ws.cell(row=row_num, column=col_num)
            set_cell_value(cell, value, key)
            if template_row >= 2 and template_row < row_num:
                source_cell = ws.cell(row=template_row, column=col_num)
                if source_cell.has_style:
                    cell._style = copy.copy(source_cell._style)
                if source_cell.number_format:
                    cell.number_format = source_cell.number_format
                if source_cell.alignment:
                    cell.alignment = copy.copy(source_cell.alignment)
                if source_cell.protection:
                    cell.protection = copy.copy(source_cell.protection)
    return len(pending)


def update_sainsbury_summary(wb, retail_rows: pd.DataFrame, online_rows: pd.DataFrame):
    ws = wb["Summary"]
    # Existing summary structure: Retail rows 4-8; Online rows 11-15.
    sections = [
        (retail_rows, 4),
        (online_rows, 11),
    ]
    month_columns = {}
    for col in range(2, 14):  # B:M = Jan:Dec
        month = str(ws.cell(3, col).value or "").strip().upper()[:3]
        if month:
            month_columns[month] = col

    updates = {}
    for rows, start_row in sections:
        if "date_of_visit" not in rows.columns:
            continue
        dated = rows.copy()
        dated["_date"] = pd.to_datetime(dated["date_of_visit"], dayfirst=True, errors="coerce")
        dated["_month"] = dated["_date"].dt.strftime("%b").str.upper()
        dated["_result"] = dated["primary_result"].astype(str).str.strip().str.casefold()
        for month in set(dated.loc[dated["_date"].notna(), "_month"]):
            if month not in month_columns:
                continue
            col = month_columns[month]
            mrows = dated[dated["_month"] == month]
            passed = int(mrows["_result"].isin(["pass", "passed"]).sum())
            failed = int(mrows["_result"].isin(["fail", "failed"]).sum())
            aborted = int(mrows["_result"].isin(["abort", "aborted"]).sum())
            ws.cell(start_row, col).value = passed + failed + aborted
            ws.cell(start_row + 1, col).value = passed
            ws.cell(start_row + 2, col).value = failed
            ws.cell(start_row + 3, col).value = aborted
            letter = ws.cell(1, col).column_letter
            ws.cell(start_row + 4, col).value = (passed / (passed + failed)) if (passed + failed) else 0
            updates[f"{'Retail' if start_row == 4 else 'Online'} - {month}"] = {
                "Visits": passed + failed + aborted, "Pass": passed,
                "Fail": failed, "Abort": aborted
            }
    return updates


def build_sainsbury_workbook(template_file, df):
    template_file.seek(0)
    wb = load_workbook(template_file)
    for required in ["Summary", "Online 26", "2026"]:
        if required not in wb.sheetnames:
            raise ValueError(f"Sainsbury's workbook is missing worksheet: {required}")

    online = df.loc[sainsbury_online_mask(df)].copy()
    retail = df.loc[sainsbury_retail_mask(df)].copy()
    online_added = append_sainsbury_rows(wb["Online 26"], online, retail=False)
    retail_added = append_sainsbury_rows(wb["2026"], retail, retail=True)
    summary = update_sainsbury_summary(wb, retail, online)

    output = BytesIO()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(output)
    output.seek(0)
    return output, {
        "Online 26 matching records": len(online),
        "Online 26 new rows": online_added,
        "2026 retail matching records": len(retail),
        "2026 retail new rows": retail_added,
    }, summary



def tesco_retail_mask(df: pd.DataFrame):
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    tesco_code = df["site_code"].astype(str).str.strip().str.upper().str.startswith("TES")
    online_items = df["item_to_order"].astype(str).str.strip().str.casefold().isin(
        {"home delivery", "rapid delivery", "compliance audit"}
    )
    return year_2026 & tesco_code & ~online_items


def tesco_online_mask(df: pd.DataFrame):
    site_names = df["site_name"].astype(str).str.strip().str.lower()
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    return (
        year_2026
        & df["item_to_order"].astype(str).str.strip().eq("Home Delivery")
        & site_names.str.contains(r"tesco\.co\.uk", regex=True, na=False)
    )


def append_tesco_rows(ws, records: pd.DataFrame, retail=False):
    headers = get_headers(ws)
    id_header = "ID" if retail else "internal_id"
    id_col = headers.index(id_header) + 1 if id_header in headers else 1
    existing_ids = {
        str(ws.cell(row, id_col).value).strip()
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, id_col).value not in (None, "")
    }
    pending = records.loc[~records["internal_id"].astype(str).str.strip().isin(existing_ids)]
    for _, record in pending.iterrows():
        row_num = ws.max_row + 1
        template_row = row_num - 1
        for col_num, header in enumerate(headers, start=1):
            if not header:
                continue
            key = str(header).strip()
            source_key = "internal_id" if key == id_header else key
            value = record.get(source_key, "")
            cell = ws.cell(row=row_num, column=col_num)
            set_cell_value(cell, value, key)
            if template_row >= 2 and template_row < row_num:
                source_cell = ws.cell(row=template_row, column=col_num)
                if source_cell.has_style:
                    cell._style = copy.copy(source_cell._style)
                if source_cell.number_format:
                    cell.number_format = source_cell.number_format
                if source_cell.alignment:
                    cell.alignment = copy.copy(source_cell.alignment)
                if source_cell.protection:
                    cell.protection = copy.copy(source_cell.protection)
    return len(pending)


def update_tesco_summary(wb, retail_rows: pd.DataFrame, online_rows: pd.DataFrame):
    ws = wb["Summary"]
    sections = [(retail_rows, 4), (online_rows, 12)]
    month_columns = {}
    for col in range(2, 14):  # B:M = Jan:Dec
        month = str(ws.cell(3 if col >= 2 else 3, col).value or "").strip().upper()[:3]
        # Online section has its own month header on row 11.
        if month:
            month_columns[month] = col
    online_month_columns = {
        str(ws.cell(11, col).value or "").strip().upper()[:3]: col
        for col in range(2, 14)
        if ws.cell(11, col).value
    }
    updates = {}
    for rows, start_row in sections:
        cols = month_columns if start_row == 4 else online_month_columns
        if "date_of_visit" not in rows.columns:
            continue
        dated = rows.copy()
        dated["_date"] = pd.to_datetime(dated["date_of_visit"], dayfirst=True, errors="coerce")
        dated["_month"] = dated["_date"].dt.strftime("%b").str.upper()
        dated["_result"] = dated["primary_result"].astype(str).str.strip().str.casefold()
        for month in sorted(set(dated.loc[dated["_date"].notna(), "_month"])):
            if month not in cols:
                continue
            col = cols[month]
            mrows = dated[dated["_month"] == month]
            passed = int(mrows["_result"].isin(["pass", "passed"]).sum())
            failed = int(mrows["_result"].isin(["fail", "failed"]).sum())
            aborted = int(mrows["_result"].isin(["abort", "aborted"]).sum())
            ws.cell(start_row, col).value = passed + failed + aborted
            ws.cell(start_row + 1, col).value = passed
            ws.cell(start_row + 2, col).value = failed
            ws.cell(start_row + 3, col).value = aborted
            letter = ws.cell(1, col).column_letter
            ws.cell(start_row + 4, col).value = (passed / (passed + failed)) if (passed + failed) else 0
            updates[f"{'Retail' if start_row == 4 else 'Online'} - {month}"] = {
                "Visits": passed + failed + aborted, "Pass": passed,
                "Fail": failed, "Abort": aborted
            }
    return updates


def build_tesco_workbook(template_file, df):
    template_file.seek(0)
    wb = load_workbook(template_file)
    for required in ["Summary", "Retail 2026", "Online 2026"]:
        if required not in wb.sheetnames:
            raise ValueError(f"Tesco workbook is missing worksheet: {required}")

    retail = df.loc[tesco_retail_mask(df)].copy()
    online = df.loc[tesco_online_mask(df)].copy()
    retail_added = append_tesco_rows(wb["Retail 2026"], retail, retail=True)
    online_added = append_tesco_rows(wb["Online 2026"], online, retail=False)
    summary = update_tesco_summary(wb, retail, online)

    output = BytesIO()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(output)
    output.seek(0)
    return output, {
        "Retail 2026 matching records": len(retail),
        "Retail 2026 new rows": retail_added,
        "Online 2026 matching records": len(online),
        "Online 2026 new rows": online_added,
    }, summary



def coop_mask(df: pd.DataFrame):
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    coop_code = df["site_code"].astype(str).str.strip().str.casefold().str.startswith("cooptcg")
    online_items = df["item_to_order"].astype(str).str.strip().str.casefold().isin(
        {"home delivery", "rapid delivery", "compliance audit"}
    )
    return year_2026 & coop_code & ~online_items


def append_coop_rows(ws, records: pd.DataFrame):
    headers = get_headers(ws)
    id_col = headers.index("ID") + 1 if "ID" in headers else 1
    existing_ids = {
        str(ws.cell(row, id_col).value).strip()
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, id_col).value not in (None, "")
    }
    pending = records.loc[~records["internal_id"].astype(str).str.strip().isin(existing_ids)]
    aliases = {
        "ID": "internal_id",
        "Retailer": "retailer",
        "Premises Name": "site_name",
        "Address": "site_address_1",
        "Post Code": "site_post_code",
        "Date of visit": "date_of_visit",
        "Time of visit": "time_of_visit",
        "Site Code": "site_code",
        "Pass/Fail": "primary_result",
    }
    for _, record in pending.iterrows():
        row_num = ws.max_row + 1
        template_row = row_num - 1 if row_num > 2 else 2
        for col_num, header in enumerate(headers, start=1):
            if not header:
                continue
            key = str(header).strip()
            source_key = aliases.get(key, key)
            value = record.get(source_key, "")
            if key == "Pass/Fail":
                result = str(value).strip().casefold()
                value = "Pass" if result in ("pass", "passed") else "Fail" if result in ("fail", "failed") else "Abort" if result in ("abort", "aborted") else value
            cell = ws.cell(row=row_num, column=col_num)
            set_cell_value(cell, value, key)
            if template_row >= 2 and template_row < row_num:
                src = ws.cell(row=template_row, column=col_num)
                if src.has_style:
                    cell._style = copy.copy(src._style)
                if src.number_format:
                    cell.number_format = src.number_format
                if src.alignment:
                    cell.alignment = copy.copy(src.alignment)
                if src.protection:
                    cell.protection = copy.copy(src.protection)
    return len(pending)


def update_coop_summary(wb, rows: pd.DataFrame):
    ws = wb["Summary"]
    month_columns = {
        str(ws.cell(3, col).value or "").strip().upper()[:3]: col
        for col in range(2, 14)
        if ws.cell(3, col).value
    }
    dated = rows.copy()
    dated["_date"] = pd.to_datetime(dated["date_of_visit"], dayfirst=True, errors="coerce")
    dated = dated[dated["_date"].dt.year.eq(2026)].copy()
    dated["_month"] = dated["_date"].dt.strftime("%b").str.upper()
    dated["_result"] = dated["primary_result"].astype(str).str.strip().str.casefold()
    updates = {}
    for month in set(dated["_month"].dropna()):
        if month not in month_columns:
            continue
        col = month_columns[month]
        mr = dated[dated["_month"] == month]
        passed = int(mr["_result"].isin(["pass", "passed"]).sum())
        failed = int(mr["_result"].isin(["fail", "failed"]).sum())
        aborted = int(mr["_result"].isin(["abort", "aborted"]).sum())
        ws.cell(4, col).value = passed + failed + aborted
        ws.cell(5, col).value = passed
        ws.cell(6, col).value = failed
        ws.cell(7, col).value = aborted
        letter = ws.cell(1, col).column_letter
        ws.cell(8, col).value = (passed / (passed + failed)) if (passed + failed) else 0
        updates[month] = {"Visits": passed + failed + aborted, "Pass": passed, "Fail": failed, "Abort": aborted}
    return updates


def build_coop_workbook(template_file, df):
    template_file.seek(0)
    wb = load_workbook(template_file)
    for required in ["Summary", "2026"]:
        if required not in wb.sheetnames:
            raise ValueError(f"The Co-operative workbook is missing worksheet: {required}")
    selected = df.loc[coop_mask(df)].copy()
    added = append_coop_rows(wb["2026"], selected)
    summary = update_coop_summary(wb, selected)
    output = BytesIO()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(output)
    output.seek(0)
    return output, {"2026 matching records": len(selected), "2026 new rows": added}, summary



def waitrose_retail_mask(df: pd.DataFrame):
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    retail_code = df["site_code"].astype(str).str.strip().str.upper().str.startswith("WAI")
    online_items = df["item_to_order"].astype(str).str.strip().str.casefold().isin(
        {"home delivery", "rapid delivery", "compliance audit"}
    )
    return year_2026 & retail_code & ~online_items


def waitrose_online_mask(df: pd.DataFrame):
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    return year_2026 & df["site_name"].map(norm).eq(norm("Waitrose.co.uk"))


def append_waitrose_rows(ws, records: pd.DataFrame):
    headers = get_headers(ws)
    id_col = headers.index("internal_id") + 1 if "internal_id" in headers else 1
    existing_ids = {
        str(ws.cell(row, id_col).value).strip()
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, id_col).value not in (None, "")
    }
    pending = records.loc[~records["internal_id"].astype(str).str.strip().isin(existing_ids)]
    for _, record in pending.iterrows():
        row_num = ws.max_row + 1
        template_row = row_num - 1 if row_num > 2 else 2
        for col_num, header in enumerate(headers, start=1):
            if not header:
                continue
            key = str(header).strip()
            value = record.get(key, "")
            cell = ws.cell(row=row_num, column=col_num)
            set_cell_value(cell, value, key)
            if 2 <= template_row < row_num:
                src = ws.cell(row=template_row, column=col_num)
                if src.has_style:
                    cell._style = copy.copy(src._style)
                if src.number_format:
                    cell.number_format = src.number_format
                if src.alignment:
                    cell.alignment = copy.copy(src.alignment)
                if src.protection:
                    cell.protection = copy.copy(src.protection)
    return len(pending)


def update_waitrose_summary(wb, retail_rows: pd.DataFrame, online_rows: pd.DataFrame):
    ws = wb["Pass %"]
    # Retail metrics occupy rows 5-9; online metrics occupy rows 13-17.
    sections = [(retail_rows, 5), (online_rows, 13)]
    updates = {}
    for rows, visits_row in sections:
        dated = rows.copy()
        dated["_date"] = pd.to_datetime(dated["date_of_visit"], dayfirst=True, errors="coerce")
        dated = dated[dated["_date"].dt.year.eq(2026)].copy()
        dated["_month"] = dated["_date"].dt.strftime("%b").str.upper()
        dated["_result"] = dated["primary_result"].astype(str).str.strip().str.casefold()
        for month in dated["_month"].dropna().unique():
            col = next(
                (c for c in range(2, 14)
                 if str(ws.cell(4 if visits_row == 5 else 12, c).value or "").strip().upper() == month),
                None,
            )
            if col is None:
                continue
            mr = dated[dated["_month"] == month]
            passed = int(mr["_result"].isin(["pass", "passed"]).sum())
            failed = int(mr["_result"].isin(["fail", "failed"]).sum())
            aborted = int(mr["_result"].isin(["abort", "aborted"]).sum())
            ws.cell(visits_row, col).value = passed + failed + aborted
            ws.cell(visits_row + 1, col).value = passed
            ws.cell(visits_row + 2, col).value = failed
            ws.cell(visits_row + 3, col).value = aborted
            letter = ws.cell(1, col).column_letter
            ws.cell(visits_row + 4, col).value = (passed / (passed + failed)) if (passed + failed) else 0
            updates[f"{'Retail' if visits_row == 5 else 'Online'} - {month.title()}"] = {
                "Visits": passed + failed + aborted,
                "Pass": passed, "Fail": failed, "Abort": aborted,
            }
        # Write concrete YTD values so the updated rates are visible immediately,
        # even in previews that do not recalculate Excel formulas.
        all_passed = int(dated["_result"].isin(["pass", "passed"]).sum())
        all_failed = int(dated["_result"].isin(["fail", "failed"]).sum())
        all_aborted = int(dated["_result"].isin(["abort", "aborted"]).sum())
        ws.cell(visits_row, 15).value = all_passed + all_failed + all_aborted
        ws.cell(visits_row + 1, 15).value = all_passed
        ws.cell(visits_row + 2, 15).value = all_failed
        ws.cell(visits_row + 3, 15).value = all_aborted
        ws.cell(visits_row + 4, 15).value = (
            all_passed / (all_passed + all_failed)
            if (all_passed + all_failed) else 0
        )
    return updates


def build_waitrose_workbook(template_file, df):
    template_file.seek(0)
    wb = load_workbook(template_file)
    for required in ["Pass %", "2026 Retail", "2026 Online"]:
        if required not in wb.sheetnames:
            raise ValueError(f"Waitrose workbook is missing worksheet: {required}")
    retail = df.loc[waitrose_retail_mask(df)].copy()
    online = df.loc[waitrose_online_mask(df)].copy()
    retail_added = append_waitrose_rows(wb["2026 Retail"], retail)
    online_added = append_waitrose_rows(wb["2026 Online"], online)
    summary = update_waitrose_summary(wb, retail, online)
    output = BytesIO()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(output)
    output.seek(0)
    return output, {
        "2026 Retail matching records": len(retail),
        "2026 Retail new rows": retail_added,
        "2026 Online matching records": len(online),
        "2026 Online new rows": online_added,
    }, summary



def allocation_audit(df, file1_template=None, sainsbury_template=None,
                      tesco_template=None, coop_template=None, waitrose_template=None):
    """Classify each audit using the same masks used by workbook generation."""
    assigned = {}
    year_2026 = pd.to_datetime(df["date_of_visit"], dayfirst=True, errors="coerce").dt.year.eq(2026)
    item = df["item_to_order"].astype(str).str.strip().str.casefold()
    code = df["site_code"].astype(str).str.strip().str.upper()
    names = df["site_name"].map(norm)

    if file1_template is not None:
        file1_template.seek(0)
        wb = load_workbook(file1_template, read_only=True, data_only=True)
        for sheet, item_name in [("2026 Indy", "e-cig - no id"),
                                 ("2026 Qcomm", "rapid delivery"),
                                 ("2026 16-17 YOs", "home delivery")]:
            if sheet not in wb.sheetnames:
                continue
            if sheet == "2026 Indy":
                whsmith = df.apply(is_whsmith, axis=1)
                assigned[sheet] = year_2026 & item.eq(item_name) & (
                    code.str.startswith(("IND", "BAT", "C0")) | whsmith
                )
            else:
                sites = get_site_names(wb[sheet])
                selected = year_2026 & item.eq(item_name) & names.isin(sites)
                if sheet == "2026 16-17 YOs" and "order_internal_id" in df.columns:
                    order_values = df["order_internal_id"].astype(str).str.strip()
                    eligible_orders = set(order_values.loc[selected]) - {""}
                    selected = selected | (year_2026 & item.eq(item_name) & order_values.isin(eligible_orders))
                assigned[sheet] = selected

    if sainsbury_template is not None:
        assigned["Sainsbury Online"] = sainsbury_online_mask(df) & year_2026
        assigned["Sainsbury Retail"] = sainsbury_retail_mask(df)
    if tesco_template is not None:
        assigned["Tesco Online"] = tesco_online_mask(df)
        assigned["Tesco Retail"] = tesco_retail_mask(df)
    if coop_template is not None:
        assigned["The Co-operative"] = coop_mask(df)
    if waitrose_template is not None:
        assigned["Waitrose Online"] = waitrose_online_mask(df)
        assigned["Waitrose Retail"] = waitrose_retail_mask(df)

    if not assigned:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # Compliance Audit is intentionally excluded from all client reports.
    compliance = item.eq("compliance audit")
    assigned = {name: (mask & ~compliance) for name, mask in assigned.items()}

    # E-Cig - No ID is a fallback to Indy only if no other client rule matches.
    # Use the exact same client-rule masks as the review and report builders.
    e_cig = year_2026 & item.eq("e-cig - no id")
    other_masks = [mask for name, mask in assigned.items() if name != "2026 Indy"]
    matched_elsewhere = (
        pd.concat(other_masks, axis=1).any(axis=1)
        if other_masks else pd.Series(False, index=df.index)
    )
    if file1_template is not None and "2026 Indy" in assigned:
        assigned["2026 Indy"] = assigned["2026 Indy"] | (e_cig & ~matched_elsewhere)

    matrix = pd.DataFrame({name: mask.reindex(df.index, fill_value=False) for name, mask in assigned.items()})
    count = matrix.sum(axis=1)
    unallocated = df.loc[count.eq(0) & ~compliance].copy()
    overlapping = df.loc[count.gt(1)].copy()
    if not unallocated.empty:
        unallocated.insert(0, "allocation_status", "Unallocated")
        def reason(row):
            val = norm(row.get("item_to_order", ""))
            if val == "compliance audit":
                return "Intentionally excluded"
            if val == "e-cig - no id":
                return "E-Cig - No ID did not match another client rule and was not eligible for Indy under current year/site rules"
            if val == "home delivery" and norm(row.get("site_name", "")) == norm("Totally Wicked"):
                return "Totally Wicked Home Delivery: no matching order_internal_id linked to an eligible 16-17 YOs audit"
            if val in {"rapid delivery", "home delivery"}:
                return "Audit type found, but site/order did not match a report tab's eligibility rules"
            return "No configured client report rule matched this audit"
        unallocated.insert(1, "allocation_reason", unallocated.apply(reason, axis=1))
    if not overlapping.empty:
        overlapping.insert(0, "allocation_status", "Multiple report matches")
        overlapping.insert(1, "matched_reports", matrix.loc[count.gt(1)].apply(
            lambda row: ", ".join(row.index[row.values].tolist()), axis=1
        ).values)
    excluded = df.loc[compliance].copy()
    if not excluded.empty:
        excluded.insert(0, "allocation_status", "Intentionally excluded — Compliance Audit")
    return unallocated, overlapping, excluded


st.title("Client Audit Reporting")
st.write(
    "Upload the latest audit export and the existing client workbook. "
    "The app applies report-specific filters and keeps the workbook's existing columns unchanged."
)

with st.sidebar:
    st.header("Report rules")
    st.markdown(
        "- **2026 Indy:** `E-Cig - No ID` plus site codes beginning `IND`, `BAT`, or `C0`, "
        "and WHSmith sites.\n"
        "- **2026 Qcomm:** `Rapid Delivery` at site names already present in the Qcomm sheet.\n"
        "- **2026 16-17 YOs:** `Home Delivery` at site names already present in that sheet, plus Home Delivery audits sharing the same `order_internal_id` as an eligible audit.\n\n"
        "Records not selected for these sheets are not deleted from the source export."
    )

csv_file = st.file_uploader("1. Audit export (CSV)", type=["csv"])
left, right = st.columns(2)
with left:
    template_file = st.file_uploader("2. Qcomm, 16-17 YOs, Indy template (XLSX)", type=["xlsx"])
with right:
    sainsbury_template = st.file_uploader("3. File 2 template — Sainsbury's (XLSX)", type=["xlsx"])
tesco_template = st.file_uploader("4. File 3 template — Tesco (XLSX)", type=["xlsx"])
coop_template = st.file_uploader("5. File 4 template — The Co-operative (XLSX)", type=["xlsx"])
waitrose_template = st.file_uploader("6. File 5 template — Waitrose (XLSX)", type=["xlsx"])

if all([csv_file, template_file, sainsbury_template, tesco_template, coop_template, waitrose_template]):
    try:
        df = pd.read_csv(csv_file, dtype=str, keep_default_na=False)
        missing = REQUIRED_SOURCE_COLUMNS - set(df.columns)
        if missing:
            st.error("The audit export is missing required columns: " + ", ".join(sorted(missing)))
            st.stop()

        if template_file:
            template_file.seek(0)
            wb_check = load_workbook(template_file, read_only=True, data_only=True)
            missing_sheets = set(REPORT_SHEETS) - set(wb_check.sheetnames)
            if missing_sheets:
                st.error("The Qcomm, 16-17 YOs, Indy workbook is missing required sheets: " + ", ".join(sorted(missing_sheets)))
                st.stop()
        if sainsbury_template:
            sainsbury_template.seek(0)
            sains_check = load_workbook(sainsbury_template, read_only=True, data_only=True)
            missing_sains = {"Summary", "Online 26", "2026"} - set(sains_check.sheetnames)
            if missing_sains:
                st.error("The Sainsbury's workbook is missing required sheets: " + ", ".join(sorted(missing_sains)))
                st.stop()

        if tesco_template:
            tesco_template.seek(0)
            tesco_check = load_workbook(tesco_template, read_only=True, data_only=True)
            missing_tesco = {"Summary", "Retail 2026", "Online 2026"} - set(tesco_check.sheetnames)
            if missing_tesco:
                st.error("The Tesco workbook is missing required sheets: " + ", ".join(sorted(missing_tesco)))
                st.stop()

        if coop_template:
            coop_template.seek(0)
            coop_check = load_workbook(coop_template, read_only=True, data_only=True)
            missing_coop = {"Summary", "2026"} - set(coop_check.sheetnames)
            if missing_coop:
                st.error("The Co-operative workbook is missing required sheets: " + ", ".join(sorted(missing_coop)))
                st.stop()

        if waitrose_template:
            waitrose_template.seek(0)
            waitrose_check = load_workbook(waitrose_template, read_only=True, data_only=True)
            missing_waitrose = {"Pass %", "2026 Retail", "2026 Online"} - set(waitrose_check.sheetnames)
            if missing_waitrose:
                st.error("The Waitrose workbook is missing required sheets: " + ", ".join(sorted(missing_waitrose)))
                st.stop()

        st.subheader("Download all reports")
        st.caption("Builds all five client workbooks using the uploaded audit export and templates, then bundles them into one ZIP file.")
        if st.button("Build and download all 5 reports", type="primary", key="build_all_reports"):
            with st.spinner("Building all five client reports..."):
                all_reports = {}
                # Build each workbook using the same report-generation functions as the individual buttons.
                indy_output, indy_results, indy_unmatched, indy_summary = workbook_bytes(
                    df=df, template_file=template_file,
                    sainsbury_template=sainsbury_template, tesco_template=tesco_template,
                    coop_template=coop_template, waitrose_template=waitrose_template
                )
                all_reports["Qcomm, 16-17 YOs, Indy_updated.xlsx"] = indy_output.getvalue()

                sains_output, sains_results, sains_summary = build_sainsbury_workbook(sainsbury_template, df)
                all_reports["Sainsburys_report_updated.xlsx"] = sains_output.getvalue()

                tesco_output, tesco_results, tesco_summary = build_tesco_workbook(tesco_template, df)
                all_reports["Tesco_report_updated.xlsx"] = tesco_output.getvalue()

                coop_output, coop_results, coop_summary = build_coop_workbook(coop_template, df)
                all_reports["Cooperative_report_updated.xlsx"] = coop_output.getvalue()

                waitrose_output, waitrose_results, waitrose_summary = build_waitrose_workbook(waitrose_template, df)
                all_reports["Waitrose_report_updated.xlsx"] = waitrose_output.getvalue()

                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                    for filename, content in all_reports.items():
                        archive.writestr(filename, content)
                st.session_state["all_reports_zip"] = zip_buffer.getvalue()
                st.success("All five reports have been built.")

        if "all_reports_zip" in st.session_state:
            st.download_button(
                "Download all 5 reports (ZIP)",
                data=st.session_state["all_reports_zip"],
                file_name="client_audit_reports.zip",
                mime="application/zip",
                key="download_all_reports",
            )

        st.subheader("Source data")
        a, b, c = st.columns(3)
        a.metric("Audit records in export", f"{len(df):,}")
        b.metric("Columns in export", f"{len(df.columns):,}")
        c.metric("Distinct audit IDs", f"{df['internal_id'].nunique():,}")

        if "order_internal_id" not in df.columns or not df["order_internal_id"].astype(str).str.strip().ne("").any():
            st.warning(
                "The main audit export has no usable `order_internal_id` values. The app cannot "
                "match sibling Home Delivery audits by order number; those audits will be highlighted for review."
            )

        if template_file:
            if st.button("Build Qcomm, 16-17 YOs, Indy report", type="primary"):
                output, results, unmatched, pass_summary = workbook_bytes(df=df, template_file=template_file, sainsbury_template=sainsbury_template, tesco_template=tesco_template, coop_template=coop_template, waitrose_template=waitrose_template)
                st.session_state["report_bytes"] = output.getvalue()
                st.session_state["report_results"] = results
                st.session_state["report_unmatched"] = unmatched
                st.session_state["pass_summary"] = pass_summary
                st.session_state["source_filename"] = csv_file.name

        if "report_results" in st.session_state:
            st.subheader("Processing summary")
            result_df = pd.DataFrame(st.session_state["report_results"]).T
            result_df.index.name = "Worksheet"
            st.dataframe(result_df, use_container_width=True)

            st.subheader("Pass % tab updates")
            pass_summary = st.session_state.get("pass_summary", {})
            if pass_summary:
                pass_df = pd.DataFrame(pass_summary).T
                pass_df.index.name = "Report section / month"
                st.dataframe(pass_df, use_container_width=True)
            else:
                st.info("No relevant dated records were found to update the Pass % tab.")

            st.download_button(
                "Download populated client workbook",
                data=st.session_state["report_bytes"],
                file_name="Qcomm, 16-17 YOs, Indy_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

            st.subheader("Site names to review")
            st.caption(
                "These records match the audit type but their site names are not currently listed "
                "in the corresponding sheet. They are not added automatically."
            )
            for sheet_name, rows in st.session_state["report_unmatched"].items():
                if len(rows):
                    with st.expander(f"{sheet_name}: {len(rows)} unmatched records"):
                        cols = [c for c in ["internal_id", "site_code", "site_name", "item_to_order", "date_of_visit", "primary_result"] if c in rows.columns]
                        st.dataframe(rows[cols], use_container_width=True, hide_index=True)
                else:
                    st.success(f"{sheet_name}: no unmatched site names for the relevant audit type.")
        if sainsbury_template:
            if st.button("Build Sainsbury's report"):
                sains_output, sains_results, sains_summary = build_sainsbury_workbook(sainsbury_template, df)
                st.session_state["sains_report_bytes"] = sains_output.getvalue()
                st.session_state["sains_report_results"] = sains_results
                st.session_state["sains_summary"] = sains_summary
            if "sains_report_results" in st.session_state:
                st.subheader("Sainsbury's processing summary")
                st.dataframe(pd.DataFrame(st.session_state["sains_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("sains_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["sains_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated Sainsbury's workbook",
                    data=st.session_state["sains_report_bytes"],
                    file_name="sainsburys_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_sainsbury",
                )

        if waitrose_template:
            if st.button("Build Waitrose report"):
                waitrose_output, waitrose_results, waitrose_summary = build_waitrose_workbook(waitrose_template, df)
                st.session_state["waitrose_report_bytes"] = waitrose_output.getvalue()
                st.session_state["waitrose_report_results"] = waitrose_results
                st.session_state["waitrose_summary"] = waitrose_summary
            if "waitrose_report_results" in st.session_state:
                st.subheader("Waitrose processing summary")
                st.dataframe(pd.DataFrame(st.session_state["waitrose_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("waitrose_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["waitrose_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated Waitrose workbook",
                    data=st.session_state["waitrose_report_bytes"],
                    file_name="waitrose_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_waitrose",
                )

        if coop_template:
            if st.button("Build The Co-operative report"):
                coop_output, coop_results, coop_summary = build_coop_workbook(coop_template, df)
                st.session_state["coop_report_bytes"] = coop_output.getvalue()
                st.session_state["coop_report_results"] = coop_results
                st.session_state["coop_summary"] = coop_summary
            if "coop_report_results" in st.session_state:
                st.subheader("The Co-operative processing summary")
                st.dataframe(pd.DataFrame(st.session_state["coop_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("coop_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["coop_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated The Co-operative workbook",
                    data=st.session_state["coop_report_bytes"],
                    file_name="cooperative_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_coop",
                )

        if tesco_template:
            if st.button("Build Tesco report"):
                tesco_output, tesco_results, tesco_summary = build_tesco_workbook(tesco_template, df)
                st.session_state["tesco_report_bytes"] = tesco_output.getvalue()
                st.session_state["tesco_report_results"] = tesco_results
                st.session_state["tesco_summary"] = tesco_summary
            if "tesco_report_results" in st.session_state:
                st.subheader("Tesco processing summary")
                st.dataframe(pd.DataFrame(st.session_state["tesco_report_results"], index=["Count"]).T,
                             use_container_width=True)
                if st.session_state.get("tesco_summary"):
                    st.dataframe(pd.DataFrame(st.session_state["tesco_summary"]).T,
                                 use_container_width=True)
                st.download_button(
                    "Download populated Tesco workbook",
                    data=st.session_state["tesco_report_bytes"],
                    file_name="tesco_report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_tesco",
                )

        st.divider()
        st.subheader("Post-build audit checks")
        st.caption("Highlighted rows are source audits not assigned by the active report rules. Compliance Audit records are intentionally excluded. E-Cig - No ID records still need the Indy fallback applied after checking all other client rules; they must not be treated as allocated until they appear in the generated Indy workbook.")
        unallocated, overlapping, compliance_excluded = allocation_audit(
            df, template_file, sainsbury_template, tesco_template, coop_template, waitrose_template
        )
        m1, m2, m3 = st.columns(3)
        m1.metric("Audits not added", len(unallocated))
        m2.metric("Audits matching multiple reports", len(overlapping))
        m3.metric("Compliance Audit intentionally excluded", len(compliance_excluded))
        if not unallocated.empty:
            st.warning("Review these audits: they did not match any of the uploaded report rules and will not be added to a workbook.")
            preview_cols = [c for c in ["internal_id", "allocation_reason", "date_of_visit", "item_to_order", "site_name", "site_code", "primary_result"] if c in unallocated.columns]
            highlighted = unallocated[preview_cols].style.apply(lambda row: ["background-color: #ffe0e0; color: #7f1d1d"] * len(row), axis=1)
            st.dataframe(highlighted, use_container_width=True, hide_index=True)
            st.download_button(
                "Download unallocated audits CSV",
                data=unallocated.to_csv(index=False).encode("utf-8-sig"),
                file_name="unallocated_audits.csv",
                mime="text/csv",
                key="download_unallocated",
            )
        if not compliance_excluded.empty:
            with st.expander(f"Intentionally excluded Compliance Audit records ({len(compliance_excluded)})"):
                cols = [c for c in ["internal_id", "date_of_visit", "item_to_order", "site_name", "site_code", "primary_result"] if c in compliance_excluded.columns]
                st.dataframe(compliance_excluded[cols], use_container_width=True, hide_index=True)
        if not overlapping.empty:
            st.caption("Any remaining overlaps are shown for review. Retail rules exclude online audit types and Compliance Audit.")
            overlap_cols = [c for c in ["internal_id", "date_of_visit", "item_to_order", "site_name", "site_code", "matched_reports"] if c in overlapping.columns]
            st.dataframe(overlapping[overlap_cols], use_container_width=True)
            st.download_button(
                "Download multi-match audits CSV",
                data=overlapping.to_csv(index=False).encode("utf-8-sig"),
                file_name="multi_match_audits.csv",
                mime="text/csv",
                key="download_overlapping",
            )

    except Exception as exc:
        st.error(f"Could not process these files: {exc}")
else:
    st.info("Report processing is locked until all six required files have been uploaded.")
    required_uploads = [
        ("Audit export CSV", csv_file),
        ("Qcomm, 16-17 YOs, Indy workbook", template_file),
        ("Sainsbury's workbook", sainsbury_template),
        ("Tesco workbook", tesco_template),
        ("The Co-operative workbook", coop_template),
        ("Waitrose workbook", waitrose_template),
    ]
    for label, uploaded in required_uploads:
        st.write(("✅ " if uploaded else "⬜ ") + label)
    st.caption("Upload the missing file(s) above. No report building or allocation checks will run until every required file is present.")
