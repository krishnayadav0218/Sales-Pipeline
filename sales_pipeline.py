"""
Automated Reporting Pipeline -- Auto-Adapting Edition
---------------------------------------------------------------------------
Works with ANY .xlsx, .xls, or .csv file, whatever its header names or
layout are. It does NOT rely on a fixed list of expected column names.

How the auto-detection works:
  1. HEADER ROW DETECTION -- scans the first ~15 rows of each file to find
     which row is actually the header (skips title rows, blank rows, merged
     banner cells that real-world files often have above the real table).
  2. COLUMN TYPE DETECTION -- for every column, decides if it's:
       - a DATE column   (parses as a date for most of its values)
       - a NUMBER column (parses as numeric for most of its values --
         handles "1,234", "₹1,234.50", "1234" etc.)
       - a CATEGORY/label column (everything else -- names, regions,
         products, dealer names, etc.)
  3. COMBINE -- every cleaned file is stacked into one master table, with
     a "Source File" column so you can always trace a row back to its
     original file.
  4. REPORT -- builds a formula-driven Excel report:
       - Raw Data sheet: the full combined, cleaned table
       - Summary sheet: auto-picks the best category column to group by,
         and SUMIFS-totals every detected number column against it,
         with KPI cards and a chart -- all live formulas, not hardcoded.

Required environment variables (set as Render/GitHub secrets):
  GOOGLE_SERVICE_ACCOUNT_JSON   full JSON key content of your service account
  RAW_FOLDER_ID                 Drive folder ID with the manager-uploaded files
  OUTPUT_FOLDER_ID              Drive folder ID where the report should be saved
"""
import os
import re
import glob
import tempfile
import logging
from datetime import datetime

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, PieChart, LineChart, DoughnutChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

from gdrive_utils import (
    get_drive_service, list_data_files_in_folder, download_file, upload_or_replace_file,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("pipeline")

REPORT_FILENAME = "Master_Report.xlsx"

# ---------------------------------------------------------------------------
# STEP 1: SMART FILE READING -- auto-detects the real header row
# ---------------------------------------------------------------------------

def smart_parse_dates(series):
    """
    Parses a column to dates without the month/day-swap bug that a blanket
    dayfirst=True causes on ISO-style strings. Unambiguous ISO dates
    (YYYY-MM-DD) are parsed as-is; anything else falls back to day-first
    parsing, since that's the convention most Indian business files use.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    s = series.astype(str)
    iso_mask = s.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if iso_mask.any():
        parsed.loc[iso_mask] = pd.to_datetime(series[iso_mask], errors="coerce")
    if (~iso_mask).any():
        parsed.loc[~iso_mask] = pd.to_datetime(series[~iso_mask], errors="coerce", dayfirst=True)
    return parsed


def _looks_numeric(value):
    if value is None:
        return False
    s = str(value).strip().replace(",", "").replace("₹", "").replace("$", "")
    return bool(re.match(r"^-?\d+(\.\d+)?%?$", s))


def detect_header_row(raw_df, max_scan=15):
    """Scores each of the first `max_scan` rows on how header-like it looks,
    and how data-like the row right after it looks, then returns the best."""
    best_row, best_score = 0, -1e9
    limit = min(max_scan, len(raw_df))
    for i in range(limit):
        row = raw_df.iloc[i]
        filled = row.notna().sum()
        if filled == 0:
            continue
        text_like = sum(
            1 for v in row if isinstance(v, str) and v.strip() and not _looks_numeric(v)
        )
        score = text_like - (row.isna().sum() * 0.5)
        if i + 1 < len(raw_df):
            next_row = raw_df.iloc[i + 1]
            data_like = sum(1 for v in next_row if _looks_numeric(v) or v is not None)
            score += data_like * 0.2
        if score > best_score:
            best_score = score
            best_row = i
    return best_row


def smart_read(filepath):
    """Reads a csv/xls/xlsx file, auto-detecting which row is the real header."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        raw = pd.read_csv(filepath, header=None, dtype=object)
        header_row = detect_header_row(raw)
        df = pd.read_csv(filepath, header=header_row, dtype=object)
    else:
        raw = pd.read_excel(filepath, header=None, dtype=object)
        header_row = detect_header_row(raw)
        df = pd.read_excel(filepath, header=header_row, dtype=object)

    df.columns = [str(c).strip() if not str(c).startswith("Unnamed") else f"Column_{i+1}"
                  for i, c in enumerate(df.columns)]
    # drop fully blank rows/columns that sneak in from title/footer bands
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return df


# ---------------------------------------------------------------------------
# STEP 2: COLUMN TYPE DETECTION
# ---------------------------------------------------------------------------

def classify_columns(df):
    """Returns (date_col, numeric_cols, category_cols) for a dataframe,
    without assuming any particular column names."""
    date_col = None
    numeric_cols = []
    category_cols = []

    # pass 1: prefer a column whose name hints it's a date
    for col in df.columns:
        name_lower = str(col).lower()
        if date_col is None and any(k in name_lower for k in ["date", "month", "period"]):
            parsed = smart_parse_dates(df[col])
            if parsed.notna().mean() > 0.4:
                date_col = col

    for col in df.columns:
        if col == date_col:
            continue
        series = df[col]
        cleaned = series.astype(str).str.replace(",", "", regex=False)\
                                     .str.replace("₹", "", regex=False)\
                                     .str.replace("$", "", regex=False)\
                                     .str.replace("%", "", regex=False).str.strip()
        numeric = pd.to_numeric(cleaned, errors="coerce")
        if numeric.notna().mean() > 0.6:
            numeric_cols.append(col)
        elif date_col is None:
            parsed = smart_parse_dates(series)
            if parsed.notna().mean() > 0.6:
                date_col = col
            else:
                category_cols.append(col)
        else:
            category_cols.append(col)

    return date_col, numeric_cols, category_cols


def pick_grouping_column(category_cols, df, n_rows):
    """Picks the best column to group the summary by: prefers an obviously
    label-like name, otherwise the category column with a sensible number
    of distinct values (not near-unique, not a single repeated value)."""
    preferred_keywords = ["region", "product", "item", "name", "party", "dealer",
                          "branch", "category", "zone", "state", "city", "sku"]
    for col in category_cols:
        if any(k in str(col).lower() for k in preferred_keywords):
            return col
    best_col, best_score = None, -1
    for col in category_cols:
        n_unique = df[col].nunique(dropna=True)
        if 1 < n_unique <= max(50, n_rows * 0.5):
            score = -abs(n_unique - 8)  # sweet spot around ~8 groups
            if score > best_score:
                best_score = score
                best_col = col
    return best_col or (category_cols[0] if category_cols else None)


# ---------------------------------------------------------------------------
# STEP 3: CLEAN + COMBINE
# ---------------------------------------------------------------------------

def clean_one_file(filepath):
    df = smart_read(filepath)
    date_col, numeric_cols, category_cols = classify_columns(df)

    for c in numeric_cols:
        cleaned = df[c].astype(str).str.replace(",", "", regex=False)\
                                    .str.replace("₹", "", regex=False)\
                                    .str.replace("$", "", regex=False)\
                                    .str.replace("%", "", regex=False).str.strip()
        df[c] = pd.to_numeric(cleaned, errors="coerce")
    if date_col:
        df[date_col] = smart_parse_dates(df[date_col])

    df["Source File"] = os.path.basename(filepath)
    df.attrs["date_col"] = date_col
    df.attrs["numeric_cols"] = numeric_cols
    df.attrs["category_cols"] = category_cols
    return df


def build_master_dataset(local_dir):
    files = (glob.glob(os.path.join(local_dir, "*.xlsx"))
             + glob.glob(os.path.join(local_dir, "*.xls"))
             + glob.glob(os.path.join(local_dir, "*.csv")))
    if not files:
        raise RuntimeError("No .xlsx/.xls/.csv files found in the Drive raw folder.")

    frames = []
    all_numeric, all_category = [], []
    common_date_col = None

    for f in files:
        try:
            df = clean_one_file(f)
            frames.append(df)
            for c in df.attrs["numeric_cols"]:
                if c not in all_numeric:
                    all_numeric.append(c)
            for c in df.attrs["category_cols"]:
                if c not in all_category:
                    all_category.append(c)
            if not common_date_col and df.attrs["date_col"]:
                common_date_col = df.attrs["date_col"]
            log.info(f"Cleaned {os.path.basename(f)}: {len(df)} rows, "
                      f"date_col={df.attrs['date_col']}, numeric={df.attrs['numeric_cols']}")
        except Exception as e:
            log.warning(f"Skipped {os.path.basename(f)} due to error: {e}")

    if not frames:
        raise RuntimeError("Every file failed cleaning -- check file format.")

    master = pd.concat(frames, ignore_index=True, sort=False)
    return master, all_numeric, all_category, common_date_col


# ---------------------------------------------------------------------------
# STEP 4: REPORT BUILDING (openpyxl, formula-driven)
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="Arial", bold=True, size=16, color="1F4E78")
LABEL_FONT = Font(name="Arial", bold=True, size=11)
NORMAL_FONT = Font(name="Arial", size=10)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_raw_sheet(wb, master, date_col):
    ws = wb.active
    ws.title = "Raw Data"
    headers = list(master.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row in enumerate(master.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, 1):
            col_name = headers[col_idx - 1]
            if isinstance(value, float) and pd.isna(value):
                value = None
            if pd.isna(value):
                value = None
            c = ws.cell(row=row_idx, column=col_idx, value=value)
            c.font = NORMAL_FONT
            c.border = BORDER
            if col_name == date_col and value is not None:
                c.number_format = "DD-MMM-YYYY"

    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18
    ws.freeze_panes = "A2"
    return len(master) + 1, {h: i + 1 for i, h in enumerate(headers)}


def write_summary_sheet(wb, master, last_raw_row, col_index, numeric_cols, group_col):
    ws = wb.create_sheet("Summary")
    ws["A1"] = "Executive Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Auto-generated {datetime.now().strftime('%d-%b-%Y %H:%M')} from {last_raw_row - 1} rows"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="666666")

    # KPI cards: total of each numeric column
    kpi_row = 4
    col_cursor = 1
    for metric in numeric_cols[:4]:
        letter = get_column_letter(col_index[metric])
        label_cell = ws.cell(row=kpi_row, column=col_cursor, value=f"Total {metric}")
        label_cell.font = LABEL_FONT
        value_cell = ws.cell(row=kpi_row, column=col_cursor + 1,
                              value=f"=SUM('Raw Data'!{letter}2:{letter}{last_raw_row})")
        value_cell.number_format = "#,##0"
        value_cell.font = Font(name="Arial", bold=True, size=14, color="1F4E78")
        col_cursor += 3

    if not group_col or not numeric_cols:
        ws["A6"] = "No suitable category column or numeric metric detected for a breakdown table."
        return

    groups = sorted([g for g in master[group_col].dropna().unique()])
    group_letter = get_column_letter(col_index[group_col])

    start_row = 7
    ws.cell(row=start_row, column=1, value=group_col).font = HEADER_FONT
    ws.cell(row=start_row, column=1).fill = HEADER_FILL
    for col_idx, metric in enumerate(numeric_cols, start=2):
        c = ws.cell(row=start_row, column=col_idx, value=metric)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for r_idx, g in enumerate(groups, start=start_row + 1):
        ws.cell(row=r_idx, column=1, value=g).font = LABEL_FONT
        for c_idx, metric in enumerate(numeric_cols, start=2):
            metric_letter = get_column_letter(col_index[metric])
            formula = (
                f"=SUMIFS('Raw Data'!${metric_letter}$2:${metric_letter}${last_raw_row},"
                f"'Raw Data'!${group_letter}$2:${group_letter}${last_raw_row},$A{r_idx})"
            )
            cell = ws.cell(row=r_idx, column=c_idx, value=formula)
            cell.number_format = "#,##0"
            cell.border = BORDER

    total_row = start_row + len(groups) + 1
    ws.cell(row=total_row, column=1, value="Total").font = Font(name="Arial", bold=True)
    for c_idx in range(2, len(numeric_cols) + 2):
        col_letter = get_column_letter(c_idx)
        cell = ws.cell(row=total_row, column=c_idx,
                        value=f"=SUM({col_letter}{start_row+1}:{col_letter}{total_row-1})")
        cell.number_format = "#,##0"
        cell.font = Font(name="Arial", bold=True)

    if numeric_cols:
        data_range = f"B{start_row+1}:{get_column_letter(len(numeric_cols)+1)}{total_row-1}"
        median_est = master[numeric_cols[0]].median()
        threshold = median_est * 0.5 if pd.notna(median_est) and median_est > 0 else 0
        if threshold > 0:
            ws.conditional_formatting.add(
                data_range,
                CellIsRule(operator="lessThan", formula=[str(threshold)],
                           fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"))
            )

        chart = BarChart()
        chart.title = f"{numeric_cols[0]} by {group_col}"
        chart.y_axis.title = numeric_cols[0]
        chart.x_axis.title = group_col
        data_ref = Reference(ws, min_col=2, min_row=start_row, max_row=start_row + len(groups))
        cats_ref = Reference(ws, min_col=1, min_row=start_row + 1, max_row=start_row + len(groups))
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.width, chart.height = 16, 8
        ws.add_chart(chart, f"A{total_row + 3}")

    for col_idx in range(1, len(numeric_cols) + 2):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16


def write_dashboard_sheet(wb, master, last_raw_row, col_index, numeric_cols, group_col, date_col, category_cols):
    """
    Builds a polished, dark-theme 'Dashboard' sheet styled like a modern BI
    dashboard: a title banner, colored KPI cards, a filter panel with
    dropdown(s) that drive live formulas, a donut chart, horizontal bar
    charts, a top-N breakdown, and a trend line -- all built from whatever
    columns were actually detected in the source files (nothing hardcoded).

    Interactivity note: openpyxl cannot create native Excel PivotTables or
    Slicers, so the "click to filter" behaviour here comes from a dropdown
    (data validation) cell wired to IF/SUMIFS formulas -- genuinely live,
    just not built on PivotTable machinery.
    """
    ws = wb.create_sheet("Dashboard")

    # ---- palette ----
    PAGE_BG = "0B1F3A"
    BANNER_BG = "081530"
    CARD_COLORS = ["2E75B6", "2E8B57", "7030A0", "C55A11", "1F9C9C"]
    PANEL_BG = "13294B"
    SECTION_BG = "13294B"
    WHITE = "FFFFFF"
    LIGHT_GREY = "C9D6E8"

    last_col = 24  # page background spans columns A..X
    last_row = 62

    # ---- page background ----
    dark_fill = PatternFill(start_color=PAGE_BG, end_color=PAGE_BG, fill_type="solid")
    for r in range(1, last_row):
        for c in range(1, last_col):
            ws.cell(row=r, column=c).fill = dark_fill
    for c in range(1, last_col):
        ws.column_dimensions[get_column_letter(c)].width = 12

    # ---- title banner ----
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=last_col - 1)
    banner = ws.cell(row=1, column=1, value="DASHBOARD")
    banner.font = Font(name="Arial", bold=True, size=22, color=WHITE)
    banner.alignment = Alignment(horizontal="center", vertical="center")
    for r in (1, 2):
        for c in range(1, last_col):
            ws.cell(row=r, column=c).fill = PatternFill(start_color=BANNER_BG, end_color=BANNER_BG, fill_type="solid")
    sub = ws.cell(row=3, column=1,
                   value=f"Auto-generated {datetime.now().strftime('%d-%b-%Y %H:%M')}")
    sub.font = Font(name="Arial", italic=True, size=9, color=LIGHT_GREY)

    if not numeric_cols:
        ws.cell(row=4, column=1, value="No numeric metrics detected.").font = Font(color=WHITE)
        return

    # ==================================================================
    # KPI CARDS
    # ==================================================================
    kpi_row = 4
    card_width = 4  # columns per card
    filter_value_cell = None
    raw_group_letter = get_column_letter(col_index[group_col]) if group_col else None
    if group_col:
        filter_value_cell = "$B$21"  # set below, in the filter panel

    for i, metric in enumerate(numeric_cols[:5]):
        col_start = 1 + i * card_width
        color = CARD_COLORS[i % len(CARD_COLORS)]
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        col_end = col_start + card_width - 2
        ws.merge_cells(start_row=kpi_row, start_column=col_start, end_row=kpi_row, end_column=col_end)
        ws.merge_cells(start_row=kpi_row + 1, start_column=col_start, end_row=kpi_row + 2, end_column=col_end)
        for r in range(kpi_row, kpi_row + 3):
            for c in range(col_start, col_start + card_width - 1):
                ws.cell(row=r, column=c).fill = fill

        label_cell = ws.cell(row=kpi_row, column=col_start, value=metric.upper())
        label_cell.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        label_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        metric_letter = get_column_letter(col_index[metric])
        if filter_value_cell:
            formula = (
                f'=IF({filter_value_cell}="All",'
                f"SUM('Raw Data'!${metric_letter}$2:${metric_letter}${last_raw_row}),"
                f"SUMIFS('Raw Data'!${metric_letter}$2:${metric_letter}${last_raw_row},"
                f"'Raw Data'!${raw_group_letter}$2:${raw_group_letter}${last_raw_row},{filter_value_cell}))"
            )
        else:
            formula = f"=SUM('Raw Data'!${metric_letter}$2:${metric_letter}${last_raw_row})"
        val_cell = ws.cell(row=kpi_row + 1, column=col_start, value=formula)
        val_cell.font = Font(name="Arial", bold=True, size=18, color=WHITE)
        val_cell.number_format = "#,##0"
        val_cell.alignment = Alignment(horizontal="left", vertical="center")

    # ==================================================================
    # FILTER PANEL (styled card, dropdown-driven)
    # ==================================================================
    groups = sorted([g for g in master[group_col].dropna().unique()]) if group_col else []
    panel_row = 21
    if group_col and groups:
        ws.merge_cells(start_row=panel_row - 2, start_column=1, end_row=panel_row - 2, end_column=3)
        hdr = ws.cell(row=panel_row - 2, column=1, value="FILTERS")
        hdr.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        for c in range(1, 4):
            ws.cell(row=panel_row - 2, column=c).fill = PatternFill(start_color=SECTION_BG, end_color=SECTION_BG, fill_type="solid")

        lbl = ws.cell(row=panel_row - 1, column=1, value=f"{group_col}:")
        lbl.font = Font(name="Arial", bold=True, size=10, color=LIGHT_GREY)

        filter_cell = ws.cell(row=panel_row, column=2, value="All")
        filter_cell.font = Font(name="Arial", bold=True, size=11, color="FFD966")
        filter_cell.fill = PatternFill(start_color=PANEL_BG, end_color=PANEL_BG, fill_type="solid")
        filter_cell.border = BORDER

        list_col = 26
        ws.cell(row=3, column=list_col, value="All")
        for i, g in enumerate(groups, start=4):
            ws.cell(row=i, column=list_col, value=g)
        list_letter = get_column_letter(list_col)
        dv = DataValidation(type="list",
                             formula1=f"=${list_letter}$3:${list_letter}${3 + len(groups)}",
                             allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(filter_cell)
        ws.column_dimensions[list_letter].hidden = True

        hint = ws.cell(row=panel_row + 1, column=1, value="(click cell, use dropdown arrow)")
        hint.font = Font(name="Arial", italic=True, size=8, color=LIGHT_GREY)

    # ==================================================================
    # STATIC HELPER DATA (hidden) for the charts
    # ==================================================================
    helper_col = 28
    cat_start_row = None
    date_start_row = None

    if group_col:
        agg = master.groupby(group_col)[numeric_cols].sum(numeric_only=True).reset_index()
        agg = agg.sort_values(numeric_cols[0], ascending=False)
        cat_start_row = 3
        ws.cell(row=cat_start_row, column=helper_col, value=group_col)
        for i, m in enumerate(numeric_cols):
            ws.cell(row=cat_start_row, column=helper_col + 1 + i, value=m)
        for r_idx, record in enumerate(agg.itertuples(index=False), start=cat_start_row + 1):
            for c_idx, val in enumerate(record):
                ws.cell(row=r_idx, column=helper_col + c_idx, value=val)
        cat_end_row = cat_start_row + len(agg)

    secondary_col = next((c for c in category_cols if c != group_col), None)
    top_start_row = None
    if secondary_col:
        top_agg = (master.groupby(secondary_col)[numeric_cols[0]].sum()
                   .sort_values(ascending=False).head(10).reset_index())
        top_start_row = (cat_end_row if cat_start_row else 3) + 3
        ws.cell(row=top_start_row, column=helper_col, value=secondary_col)
        ws.cell(row=top_start_row, column=helper_col + 1, value=numeric_cols[0])
        for r_idx, rec in enumerate(top_agg.itertuples(index=False), start=top_start_row + 1):
            ws.cell(row=r_idx, column=helper_col, value=rec[0])
            ws.cell(row=r_idx, column=helper_col + 1, value=rec[1])
        top_end_row = top_start_row + len(top_agg)

    if date_col and master[date_col].notna().any():
        by_date = (master.dropna(subset=[date_col])
                         .groupby(date_col)[numeric_cols[0]]
                         .sum().sort_index().reset_index())
        if len(by_date) > 1:
            date_start_row = (top_end_row if top_start_row else (cat_end_row if cat_start_row else 3)) + 3
            ws.cell(row=date_start_row, column=helper_col, value=date_col)
            ws.cell(row=date_start_row, column=helper_col + 1, value=numeric_cols[0])
            for r_idx, rec in enumerate(by_date.itertuples(index=False), start=date_start_row + 1):
                dcell = ws.cell(row=r_idx, column=helper_col, value=rec[0])
                dcell.number_format = "DD-MMM"
                ws.cell(row=r_idx, column=helper_col + 1, value=rec[1])
            date_end_row = date_start_row + len(by_date)

    for c in range(helper_col, helper_col + len(numeric_cols) + 2):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    def section_header(row, col_start, col_end, text):
        ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
        cell = ws.cell(row=row, column=col_start, value=text.upper())
        cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        for c in range(col_start, col_end + 1):
            ws.cell(row=row, column=c).fill = PatternFill(start_color=SECTION_BG, end_color=SECTION_BG, fill_type="solid")

    # ==================================================================
    # DONUT CHART -- share by category
    # ==================================================================
    if group_col and cat_start_row:
        section_header(26, 1, 8, f"{numeric_cols[0]} Share by {group_col}")
        cats_ref = Reference(ws, min_col=helper_col, min_row=cat_start_row + 1, max_row=cat_end_row)
        donut = DoughnutChart()
        donut.visible_cells_only = False
        donut_data_ref = Reference(ws, min_col=helper_col + 1, min_row=cat_start_row, max_row=cat_end_row)
        donut.add_data(donut_data_ref, titles_from_data=True)
        donut.set_categories(cats_ref)
        donut.width, donut.height = 16, 9
        ws.add_chart(donut, "A27")

    # ==================================================================
    # HORIZONTAL BAR -- totals by category
    # ==================================================================
    if group_col and cat_start_row:
        section_header(26, 10, 17, f"{numeric_cols[0]} by {group_col}")
        bar = BarChart()
        bar.type = "bar"  # horizontal
        bar.visible_cells_only = False
        bar_data_ref = Reference(ws, min_col=helper_col + 1, min_row=cat_start_row, max_row=cat_end_row)
        bar.add_data(bar_data_ref, titles_from_data=True)
        bar.set_categories(cats_ref)
        bar.width, bar.height = 16, 9
        ws.add_chart(bar, "J27")

    # ==================================================================
    # TOP-N -- secondary category breakdown
    # ==================================================================
    if secondary_col and top_start_row:
        section_header(45, 1, 8, f"Top {secondary_col} by {numeric_cols[0]}")
        top_cats_ref = Reference(ws, min_col=helper_col, min_row=top_start_row + 1, max_row=top_end_row)
        top_bar = BarChart()
        top_bar.type = "bar"
        top_bar.visible_cells_only = False
        top_data_ref = Reference(ws, min_col=helper_col + 1, min_row=top_start_row, max_row=top_end_row)
        top_bar.add_data(top_data_ref, titles_from_data=True)
        top_bar.set_categories(top_cats_ref)
        top_bar.width, top_bar.height = 16, 9
        ws.add_chart(top_bar, "A46")

    # ==================================================================
    # TREND LINE
    # ==================================================================
    if date_col and date_start_row:
        section_header(45, 10, 17, f"{numeric_cols[0]} Trend Over Time")
        line = LineChart()
        line.visible_cells_only = False
        line_data_ref = Reference(ws, min_col=helper_col + 1, min_row=date_start_row, max_row=date_end_row)
        line_cats_ref = Reference(ws, min_col=helper_col, min_row=date_start_row + 1, max_row=date_end_row)
        line.add_data(line_data_ref, titles_from_data=True)
        line.set_categories(line_cats_ref)
        line.width, line.height = 16, 9
        ws.add_chart(line, "J46")

    ws.sheet_view.showGridLines = False


def main():
    raw_folder_id = os.environ.get("RAW_FOLDER_ID")
    output_folder_id = os.environ.get("OUTPUT_FOLDER_ID")
    if not raw_folder_id or not output_folder_id:
        raise RuntimeError("RAW_FOLDER_ID and OUTPUT_FOLDER_ID environment variables must be set.")
    if raw_folder_id == output_folder_id:
        raise RuntimeError(
            "RAW_FOLDER_ID and OUTPUT_FOLDER_ID are set to the SAME folder ID -- "
            "these must be two different Drive folders. Check your GitHub Secrets."
        )
    log.info(f"Raw folder ID:    {raw_folder_id}")
    log.info(f"Output folder ID: {output_folder_id}")

    with tempfile.TemporaryDirectory() as tmp_dir:
        log.info("Connecting to Google Drive...")
        service = get_drive_service()

        log.info(f"Listing files in raw folder {raw_folder_id}...")
        files = list_data_files_in_folder(service, raw_folder_id)
        files = [f for f in files if f["name"] != REPORT_FILENAME]
        if not files:
            raise RuntimeError("No supported files (.xlsx/.xls/.csv) found in the Drive raw folder.")

        for f in files:
            dest = os.path.join(tmp_dir, f["name"])
            download_file(service, f["id"], dest)
            log.info(f"Downloaded: {f['name']}")

        log.info("Cleaning and auto-detecting structure...")
        master, numeric_cols, category_cols, date_col = build_master_dataset(tmp_dir)
        log.info(f"Total rows: {len(master)} | numeric: {numeric_cols} | "
                  f"category: {category_cols} | date_col: {date_col}")

        group_col = pick_grouping_column(category_cols, master, len(master))
        log.info(f"Grouping summary by: {group_col}")

        log.info("Building formatted Excel report...")
        wb = Workbook()
        last_raw_row, col_index = write_raw_sheet(wb, master, date_col)
        write_summary_sheet(wb, master, last_raw_row, col_index, numeric_cols, group_col)
        write_dashboard_sheet(wb, master, last_raw_row, col_index, numeric_cols, group_col, date_col, category_cols)

        report_path = os.path.join(tmp_dir, REPORT_FILENAME)
        wb.save(report_path)

        log.info("Uploading report back to Drive output folder...")
        file_id = upload_or_replace_file(service, report_path, output_folder_id, REPORT_FILENAME)
        log.info(f"Done. Report file ID: {file_id}")


if __name__ == "__main__":
    main()
