"""
Automated Sales & Financial Reporting Pipeline -- Render + Google Drive edition
---------------------------------------------------------------------------
Runs on Render as a scheduled Cron Job. No local Drive sync needed --
talks to Google Drive directly via API using a Service Account.

Flow:
  1. Download every .xlsx from the Drive "raw" folder into a temp folder
  2. Clean & standardize each region's data (handles different column
     names/order/date formats managers actually send)
  3. Combine into one master dataset
  4. Build a formatted Excel report (formulas, KPI cards, conditional
     formatting, chart -- not hardcoded numbers)
  5. Upload the report back to the Drive "output" folder (replacing the
     previous run's file so the link stays the same)

Required environment variables (set these in Render's dashboard):
  GOOGLE_SERVICE_ACCOUNT_JSON   full JSON key content of your service account
  RAW_FOLDER_ID                 Drive folder ID containing manager-uploaded files
  OUTPUT_FOLDER_ID              Drive folder ID where the report should be saved
"""
import os
import glob
import tempfile
import logging
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, Reference

from gdrive_utils import (
    get_drive_service, list_excel_files_in_folder, download_file, upload_or_replace_file,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("sales_pipeline")

REPORT_FILENAME = "Sales_Master_Report.xlsx"

# ---------------------------------------------------------------------------
# COLUMN MAPPING -- every column-name variation a manager's file might use
# ---------------------------------------------------------------------------
COLUMN_MAP = {
    "date": "Date", "txn date": "Date", "sale date": "Date",
    "product name": "Product", "item": "Product", "product": "Product",
    "units sold": "Units", "qty": "Units", "quantity": "Units",
    "unit price (inr)": "UnitPrice", "price": "UnitPrice",
    "revenue (inr)": "Revenue", "total revenue": "Revenue", "revenue": "Revenue",
}


def region_from_filename(filename):
    """Derives a region label from the filename, e.g. Sales_North.xlsx -> North."""
    stem = os.path.splitext(filename)[0]
    return stem.replace("Sales_", "").replace("sales_", "").strip() or stem


def clean_one_file(filepath):
    filename = os.path.basename(filepath)
    region = region_from_filename(filename)

    df = pd.read_excel(filepath)
    df.columns = [c.strip().lower() for c in df.columns]
    df = df.rename(columns={c: COLUMN_MAP.get(c, c) for c in df.columns})

    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")

    if {"Units", "Revenue", "UnitPrice"}.issubset(df.columns):
        mask = df["Units"].isna() & df["UnitPrice"].notna() & df["Revenue"].notna()
        df.loc[mask, "Units"] = (df.loc[mask, "Revenue"] / df.loc[mask, "UnitPrice"]).round()

        mask = df["Revenue"].isna() & df["Units"].notna() & df["UnitPrice"].notna()
        df.loc[mask, "Revenue"] = df.loc[mask, "Units"] * df.loc[mask, "UnitPrice"]

    df["Region"] = region
    required = [c for c in ["Date", "Product", "Units", "Revenue"] if c in df.columns]
    df = df.dropna(subset=required)
    df["Units"] = df["Units"].astype(int)
    return df[["Date", "Region", "Product", "Units", "UnitPrice", "Revenue"]]


def build_master_dataset(local_raw_dir):
    files = glob.glob(os.path.join(local_raw_dir, "*.xlsx"))
    if not files:
        raise RuntimeError("No .xlsx files found in the Drive raw folder.")
    frames = []
    for f in files:
        try:
            frames.append(clean_one_file(f))
            log.info(f"Cleaned {os.path.basename(f)}: OK")
        except Exception as e:
            log.warning(f"Skipped {os.path.basename(f)} due to error: {e}")
    if not frames:
        raise RuntimeError("Every file failed cleaning -- check column names/format.")
    master = pd.concat(frames, ignore_index=True)
    return master.sort_values(["Date", "Region", "Product"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# REPORT BUILDING (openpyxl, formula-driven, same logic as before)
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name="Arial", bold=True, size=16, color="1F4E78")
LABEL_FONT = Font(name="Arial", bold=True, size=11)
NORMAL_FONT = Font(name="Arial", size=10)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_raw_sheet(wb, master):
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
            c = ws.cell(row=row_idx, column=col_idx, value=value)
            c.font = NORMAL_FONT
            c.border = BORDER
            if headers[col_idx - 1] == "Date":
                c.number_format = "DD-MMM-YYYY"
            if headers[col_idx - 1] in ("UnitPrice", "Revenue"):
                c.number_format = "#,##0"

    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
    ws.freeze_panes = "A2"
    return len(master) + 1


def write_summary_sheet(wb, master, last_raw_row):
    ws = wb.create_sheet("Regional Summary")
    ws["A1"] = "Executive Sales Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Auto-generated {datetime.now().strftime('%d-%b-%Y %H:%M')} from {last_raw_row - 1} transactions"
    ws["A2"].font = Font(name="Arial", italic=True, size=9, color="666666")

    regions = sorted(master["Region"].unique())

    ws["A4"] = "Total Revenue (INR)"
    ws["A4"].font = LABEL_FONT
    ws["B4"] = f"=SUM('Raw Data'!F2:F{last_raw_row})"
    ws["B4"].number_format = "#,##0"
    ws["B4"].font = Font(name="Arial", bold=True, size=14, color="1F4E78")

    ws["D4"] = "Total Units Sold"
    ws["D4"].font = LABEL_FONT
    ws["E4"] = f"=SUM('Raw Data'!D2:D{last_raw_row})"
    ws["E4"].number_format = "#,##0"
    ws["E4"].font = Font(name="Arial", bold=True, size=14, color="1F4E78")

    ws["G4"] = "Avg Order Value"
    ws["G4"].font = LABEL_FONT
    ws["H4"] = f"=B4/COUNTA('Raw Data'!A2:A{last_raw_row})"
    ws["H4"].number_format = "#,##0"
    ws["H4"].font = Font(name="Arial", bold=True, size=14, color="1F4E78")

    start_row = 7
    ws.cell(row=start_row, column=1, value="Region").font = HEADER_FONT
    ws.cell(row=start_row, column=1).fill = HEADER_FILL
    products = sorted(master["Product"].unique())
    for col_idx, p in enumerate(products, start=2):
        c = ws.cell(row=start_row, column=col_idx, value=p)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    total_col = len(products) + 2
    c = ws.cell(row=start_row, column=total_col, value="Region Total")
    c.font = HEADER_FONT
    c.fill = HEADER_FILL

    for r_idx, region in enumerate(regions, start=start_row + 1):
        ws.cell(row=r_idx, column=1, value=region).font = LABEL_FONT
        for c_idx, product in enumerate(products, start=2):
            formula = (
                f"=SUMIFS('Raw Data'!$F$2:$F${last_raw_row},"
                f"'Raw Data'!$B$2:$B${last_raw_row},$A{r_idx},"
                f"'Raw Data'!$C$2:$C${last_raw_row},{get_column_letter(c_idx)}${start_row})"
            )
            cell = ws.cell(row=r_idx, column=c_idx, value=formula)
            cell.number_format = "#,##0"
            cell.border = BORDER
        cell = ws.cell(row=r_idx, column=total_col,
                        value=f"=SUM(B{r_idx}:{get_column_letter(total_col-1)}{r_idx})")
        cell.number_format = "#,##0"
        cell.font = Font(name="Arial", bold=True)

    total_row = start_row + len(regions) + 1
    ws.cell(row=total_row, column=1, value="Product Total").font = Font(name="Arial", bold=True)
    for c_idx in range(2, total_col + 1):
        col_letter = get_column_letter(c_idx)
        cell = ws.cell(row=total_row, column=c_idx,
                        value=f"=SUM({col_letter}{start_row+1}:{col_letter}{total_row-1})")
        cell.number_format = "#,##0"
        cell.font = Font(name="Arial", bold=True)

    data_range = f"B{start_row+1}:{get_column_letter(total_col-1)}{total_row-1}"
    ws.conditional_formatting.add(
        data_range,
        CellIsRule(operator="lessThan", formula=["50000"],
                   fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"))
    )

    chart = BarChart()
    chart.title = "Revenue by Region"
    chart.y_axis.title = "Revenue (INR)"
    chart.x_axis.title = "Region"
    data_ref = Reference(ws, min_col=total_col, min_row=start_row, max_row=start_row + len(regions))
    cats_ref = Reference(ws, min_col=1, min_row=start_row + 1, max_row=start_row + len(regions))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width, chart.height = 16, 8
    ws.add_chart(chart, f"A{total_row + 3}")

    for col_idx in range(1, total_col + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 15


def main():
    raw_folder_id = os.environ.get("RAW_FOLDER_ID")
    output_folder_id = os.environ.get("OUTPUT_FOLDER_ID")
    if not raw_folder_id or not output_folder_id:
        raise RuntimeError("RAW_FOLDER_ID and OUTPUT_FOLDER_ID environment variables must be set.")

    with tempfile.TemporaryDirectory() as tmp_dir:
        log.info("Connecting to Google Drive...")
        service = get_drive_service()

        log.info(f"Listing files in raw folder {raw_folder_id}...")
        files = list_excel_files_in_folder(service, raw_folder_id)
        if not files:
            raise RuntimeError("No .xlsx files found in the Drive raw folder.")

        for f in files:
            dest = os.path.join(tmp_dir, f["name"])
            download_file(service, f["id"], dest)
            log.info(f"Downloaded: {f['name']}")

        log.info("Cleaning and aggregating data...")
        master = build_master_dataset(tmp_dir)
        log.info(f"Total clean rows: {len(master)}")

        log.info("Building formatted Excel report...")
        wb = Workbook()
        last_raw_row = write_raw_sheet(wb, master)
        write_summary_sheet(wb, master, last_raw_row)

        report_path = os.path.join(tmp_dir, REPORT_FILENAME)
        wb.save(report_path)

        log.info("Uploading report back to Drive output folder...")
        file_id = upload_or_replace_file(service, report_path, output_folder_id, REPORT_FILENAME)
        log.info(f"Done. Report file ID: {file_id}")


if __name__ == "__main__":
    main()
