"""
excel_writer.py
---------------
Writes the optimization workbook (Summary + Detail sheets) and applies the
red highlighting required by the specification.

The workbook is built directly with openpyxl so we have full control over fills,
fonts, column widths and freeze panes.
"""

import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# --- Style constants --------------------------------------------------------

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
LIGHT_RED_FILL = PatternFill(start_color="FFE3E6", end_color="FFE3E6", fill_type="solid")
RED_FONT = Font(color="9C0006")
DARK_RED_FONT = Font(color="9C0006", bold=True)

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(color="FFFFFF", bold=True, size=12)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


# Detail sheet column order (logical names = output headers).
DETAIL_COLUMNS = [
    "Site", "Site Name", "Item", "Description", "Site Zip", "Site Country",
    "Pur/Man", "Item type", "Type Trim", "Group",
    "Current Safety Time", "Current Safety Stock", "Current Transport Days",
    "Current Total MRP", "Current Daily Usage", "Current Inco Terms", "KM",
    "Proposed Safety Time", "Proposed Safety Stock", "Proposed Transport Days",
    "Issue Type", "Issue Description", "Recommended Action", "Review Required",
    "Triggered Rules",
    "Local STD Cost", "Euro XC Rate", "Inv Value", "DOH", "MRP 49 Days",
]

SUMMARY_COLUMNS = [
    "Plant", "Site", "Site Name", "Issue Type", "Issue Count",
    "Recommended Action", "Status",
]


def _style_header_row(ws, row_idx, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER


def _auto_width(ws, headers, max_width=55):
    for idx, header in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        # Base width on the header text; data is clamped to a max.
        width = min(max(len(str(header)) + 2, 12), max_width)
        ws.column_dimensions[letter].width = width


def _write_summary_sheet(ws, summary_df):
    ws.append(["ASL Inventory Parameter Optimization - Summary"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1,
                   end_column=len(SUMMARY_COLUMNS))
    title_cell = ws.cell(row=1, column=1)
    title_cell.fill = HEADER_FILL
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.append(SUMMARY_COLUMNS)
    _style_header_row(ws, 2, len(SUMMARY_COLUMNS))

    for _, r in summary_df.iterrows():
        ws.append([r.get(col, "") for col in SUMMARY_COLUMNS])

    # Totals row per issue type is implicit via grouping; add a grand total.
    if not summary_df.empty:
        total = int(summary_df["Issue Count"].sum())
        total_row = ws.max_row + 1
        ws.cell(row=total_row, column=1, value="Total")
        ws.cell(row=total_row, column=4, value="All Issue Types")
        ws.cell(row=total_row, column=5, value=total)
        for c in range(1, len(SUMMARY_COLUMNS) + 1):
            ws.cell(row=total_row, column=c).font = Font(bold=True)

    _auto_width(ws, SUMMARY_COLUMNS)
    ws.column_dimensions["F"].width = 55
    ws.freeze_panes = "A3"


def _write_detail_sheet(ws, detail_df):
    ws.append(["ASL Inventory Parameter Optimization - Detail"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1,
                   end_column=len(DETAIL_COLUMNS))
    title_cell = ws.cell(row=1, column=1)
    title_cell.fill = HEADER_FILL
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.append(DETAIL_COLUMNS)
    _style_header_row(ws, 2, len(DETAIL_COLUMNS))

    # Column index lookup for targeted highlighting.
    col_index = {name: i + 1 for i, name in enumerate(DETAIL_COLUMNS)}

    for _, r in detail_df.iterrows():
        values = [r.get(col, "") for col in DETAIL_COLUMNS]
        # Replace None with blank for a clean look.
        values = ["" if v is None else v for v in values]
        ws.append(values)
        excel_row = ws.max_row

        highlight = r.get("_highlight", {}) or {}

        # Always highlight Site + Site Name when the row has any issue.
        for name in ("Site", "Site Name"):
            cell = ws.cell(row=excel_row, column=col_index[name])
            cell.fill = RED_FILL
            cell.font = RED_FONT

        # Highlight the specific problematic current-parameter cell(s).
        if highlight.get("safety_stock"):
            _mark_red(ws, excel_row, col_index["Current Safety Stock"])
        if highlight.get("safety_time"):
            _mark_red(ws, excel_row, col_index["Current Safety Time"])
        if highlight.get("transport_days"):
            _mark_red(ws, excel_row, col_index["Current Transport Days"])

        # Light-red the issue / action columns.
        for name in ("Issue Type", "Issue Description", "Recommended Action"):
            cell = ws.cell(row=excel_row, column=col_index[name])
            cell.fill = LIGHT_RED_FILL

        # Manual-review rows get an amber-ish dark red font on the flag.
        if str(r.get("Review Required", "")).strip().lower() == "yes":
            cell = ws.cell(row=excel_row, column=col_index["Review Required"])
            cell.font = DARK_RED_FONT

    _auto_width(ws, DETAIL_COLUMNS)
    # Widen the free-text columns.
    for name, width in (("Description", 40), ("Issue Description", 50),
                        ("Recommended Action", 40), ("Site Name", 20)):
        ws.column_dimensions[get_column_letter(col_index[name])].width = width

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = "A2:{}{}".format(
        get_column_letter(len(DETAIL_COLUMNS)), max(ws.max_row, 2)
    )


def _mark_red(ws, row, col):
    cell = ws.cell(row=row, column=col)
    cell.fill = RED_FILL
    cell.font = DARK_RED_FONT


def write_output_excel(summary_df, detail_df, output_path):
    """
    Create the workbook with a Summary sheet and a Detail sheet and save it to
    ``output_path``.  Red highlighting is applied inline while writing.

    Raises ``PermissionError`` with a clear message when the destination file is
    open in Excel (common user mistake).
    """
    wb = Workbook()

    summary_ws = wb.active
    summary_ws.title = "Summary"
    _write_summary_sheet(summary_ws, summary_df)

    detail_ws = wb.create_sheet("Detail")
    _write_detail_sheet(detail_ws, detail_df)

    try:
        wb.save(output_path)
    except PermissionError:
        raise PermissionError(
            "The output file could not be saved because it is already open.\n"
            "Please close '{}' in Excel and try again."
            .format(os.path.basename(output_path))
        )

    return output_path


def build_output_path(input_file_path, output_folder=None):
    """
    Build the output file path.

    Defaults to the input file's folder.  If the standard output file already
    exists, a timestamp suffix is appended so nothing is overwritten.
    """
    folder = output_folder or os.path.dirname(os.path.abspath(input_file_path))
    base = "ASL_Inventory_Parameter_Optimization_Output"
    path = os.path.join(folder, base + ".xlsx")

    if os.path.exists(path):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(folder, "{}_{}.xlsx".format(base, stamp))

    return path
