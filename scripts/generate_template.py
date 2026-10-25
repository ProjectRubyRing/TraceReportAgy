#!/usr/bin/env python3
"""
generate_template.py
Generate a professional Excel template (trace_report_template.xlsx)
modeled after the AWS CloudWatch / X-Ray trace details console.
Designed for RHEL 9.8 / Python 3.9+ with openpyxl.
"""

import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Monotone Palette (Refined Executive Grayscale)
HEADER_DARK = "212529"          # Deep Charcoal / Black for Titles & Primary Banners
HEADER_MID = "343A40"           # Dark Slate Charcoal for Table Headers
HEADER_SCALE = "495057"         # Slate Gray for Timeline Scale Headers
WHITE = "FFFFFF"                # Crisp Pure White
TEXT_MAIN = "212529"            # Primary Text Color
TEXT_MUTED = "6C757D"           # Secondary / Muted Gray Text
CARD_BG = "F8F9FA"              # Card Panel Soft Light Gray
CARD_BORDER = "DEE2E6"          # Card Border Neutral Gray
BORDER_LIGHT = "E9ECEF"         # Grid & Table Cell Inner Border
BORDER_DARK = "CED4DA"          # Table Outer / Header Separator Border

# Backward-compatibility aliases
NAVY_HEADER = HEADER_DARK
GRAY_TEXT = TEXT_MUTED

# Status Colors (Monotone Base)
SUCCESS_BG = "E9ECEF"           # Light Silver Gray
SUCCESS_FG = "212529"           # Dark Charcoal
WARNING_BG = "6C757D"           # Medium Charcoal Gray
WARNING_FG = "FFFFFF"           # Pure White
DANGER_BG = "212529"            # Dark Charcoal Black
DANGER_FG = "FFFFFF"            # Pure White
INFO_BG = "DEE2E6"              # Subtle Neutral Gray
INFO_FG = "212529"              # Dark Charcoal

TIMELINE_BG = "F8F9FA"          # Empty Timeline Slot
ROW_ALT = "F8F9FA"              # Alternating Row Tint
ROW_WHITE = "FFFFFF"            # White Row

FONT_FAMILY = "Meiryo UI"


def create_thin_border(color=BORDER_LIGHT):
    side = Side(border_style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def create_template(output_path: str = "templates/trace_report_template.xlsx"):
    wb = openpyxl.Workbook()

    # ----------------------------------------------------
    # Sheet 1: Summary (Trace Execution Index & Overview)
    # ----------------------------------------------------
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.views.sheetView[0].showGridLines = True

    # Title Banner (Row 2 to 3, Cols B to M)
    ws_sum.merge_cells("B2:M3")
    title_cell = ws_sum["B2"]
    title_cell.value = "AWS CloudWatch Transaction Search - Trace Report"
    title_cell.font = Font(name=FONT_FAMILY, size=15, bold=True, color=WHITE)
    title_cell.fill = PatternFill(start_color=HEADER_DARK, end_color=HEADER_DARK, fill_type="solid")
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Subtitle / Metadata (Row 4, 5)
    ws_sum["B4"] = "Log Group:"
    ws_sum["B4"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=TEXT_MUTED)
    ws_sum["C4"] = "aws/spans"
    ws_sum["C4"].font = Font(name=FONT_FAMILY, size=9, bold=False, color=TEXT_MAIN)

    ws_sum["F4"] = "Query Window (UTC):"
    ws_sum["F4"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=TEXT_MUTED)
    ws_sum["G4"] = "{{QUERY_START}} ~ {{QUERY_END}}"
    ws_sum["G4"].font = Font(name=FONT_FAMILY, size=9, bold=False, color=TEXT_MAIN)

    ws_sum["B5"] = "Generated At:"
    ws_sum["B5"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=TEXT_MUTED)
    ws_sum["C5"] = "{{GENERATED_AT}}"
    ws_sum["C5"].font = Font(name=FONT_FAMILY, size=9, bold=False, color=TEXT_MAIN)

    ws_sum["F5"] = "Status Filter:"
    ws_sum["F5"].font = Font(name=FONT_FAMILY, size=9, bold=True, color=TEXT_MUTED)
    ws_sum["G5"] = "{{FILTER_CRITERIA}}"
    ws_sum["G5"].font = Font(name=FONT_FAMILY, size=9, bold=False, color=TEXT_MAIN)

    # KPI Summary Cards (Row 7 to 9)
    # Refined Monotone KPI Cards
    kpi_configs = [
        ("B", "D", "TOTAL TRACES", "{{TOTAL_TRACES}}", CARD_BG, TEXT_MUTED, TEXT_MAIN),
        ("E", "G", "TOTAL SPANS", "{{TOTAL_SPANS}}", CARD_BG, TEXT_MUTED, TEXT_MAIN),
        ("H", "I", "OK (2xx)", "{{OK_TRACES}}", CARD_BG, TEXT_MUTED, TEXT_MAIN),
        ("J", "K", "ERROR / FAULT", "{{ERROR_TRACES}}", CARD_BG, TEXT_MUTED, TEXT_MAIN),
        ("L", "M", "MAX DURATION", "{{MAX_DURATION}}", CARD_BG, TEXT_MUTED, TEXT_MAIN),
    ]

    for start_col, end_col, label, val_placeholder, bg_color, lbl_color, val_color in kpi_configs:
        cell_lbl = f"{start_col}7"
        cell_val = f"{start_col}8"
        ws_sum.merge_cells(f"{start_col}7:{end_col}7")
        ws_sum.merge_cells(f"{start_col}8:{end_col}9")

        c_lbl = ws_sum[cell_lbl]
        c_lbl.value = label
        c_lbl.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=lbl_color)
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        c_lbl.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        c_val = ws_sum[cell_val]
        c_val.value = val_placeholder
        c_val.font = Font(name=FONT_FAMILY, size=15, bold=True, color=val_color)
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        # Set borders around card
        for row in range(7, 10):
            for col_letter in [chr(c) for c in range(ord(start_col), ord(end_col) + 1)]:
                ws_sum[f"{col_letter}{row}"].border = create_thin_border(CARD_BORDER)

    # Section Title for Table
    ws_sum["B11"] = "Trace Summary Table"
    ws_sum["B11"].font = Font(name=FONT_FAMILY, size=11, bold=True, color=HEADER_DARK)

    # Trace Table Headers (Row 12)
    sum_headers = [
        ("Trace ID", 38),
        ("Start Time", 24),
        ("Root Service", 22),
        ("Root Operation / URL", 32),
        ("Duration (ms)", 16),
        ("Spans", 11),
        ("Services", 11),
        ("Status", 14),
        ("Errors", 11),
        ("Details Link", 16),
    ]

    for col_idx, (header_name, col_width) in enumerate(sum_headers, start=2):
        col_letter = get_column_letter(col_idx)
        cell = ws_sum.cell(row=12, column=col_idx, value=header_name)
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=HEADER_MID, end_color=HEADER_MID, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = create_thin_border(BORDER_DARK)
        ws_sum.column_dimensions[col_letter].width = col_width

    ws_sum.row_dimensions[2].height = 26
    ws_sum.row_dimensions[3].height = 22
    ws_sum.row_dimensions[4].height = 20
    ws_sum.row_dimensions[5].height = 20
    ws_sum.row_dimensions[7].height = 18
    ws_sum.row_dimensions[8].height = 24
    ws_sum.row_dimensions[9].height = 18
    ws_sum.row_dimensions[11].height = 24
    ws_sum.row_dimensions[12].height = 26
    ws_sum.column_dimensions["A"].width = 3
    ws_sum.freeze_panes = "B13"

    # ----------------------------------------------------
    # Sheet 2: Trace_Waterfall (X-Ray Timeline View)
    # ----------------------------------------------------
    ws_wat = wb.create_sheet(title="Trace_Waterfall")
    ws_wat.views.sheetView[0].showGridLines = True

    # Title Banner (Row 2-3)
    ws_wat.merge_cells("B2:U3")
    t_cell = ws_wat["B2"]
    t_cell.value = "AWS X-Ray Trace Timeline & Waterfall View"
    t_cell.font = Font(name=FONT_FAMILY, size=15, bold=True, color=WHITE)
    t_cell.fill = PatternFill(start_color=HEADER_DARK, end_color=HEADER_DARK, fill_type="solid")
    t_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Trace KPI Overview Bar (Row 5 to 7)
    # Refined Monotone KPI Cards
    trace_kpis = [
        ("B", "E", "TRACE ID", "{{CURRENT_TRACE_ID}}", CARD_BG, TEXT_MUTED, TEXT_MAIN, 10),
        ("F", "G", "STATUS", "{{CURRENT_STATUS}}", SUCCESS_BG, TEXT_MUTED, SUCCESS_FG, 13),
        ("H", "L", "ROOT SERVICE / OPERATION", "{{CURRENT_ROOT}}", CARD_BG, TEXT_MUTED, TEXT_MAIN, 10),
        ("M", "N", "TOTAL DURATION", "{{CURRENT_DURATION}}", CARD_BG, TEXT_MUTED, TEXT_MAIN, 13),
        ("O", "P", "TOTAL SPANS", "{{CURRENT_SPANS_COUNT}}", CARD_BG, TEXT_MUTED, TEXT_MAIN, 13),
        ("Q", "U", "START TIMESTAMP (UTC)", "{{CURRENT_START_TIME}}", CARD_BG, TEXT_MUTED, TEXT_MAIN, 10),
    ]

    for start_col, end_col, label, val_placeholder, bg_color, lbl_color, val_color, val_size in trace_kpis:
        ws_wat.merge_cells(f"{start_col}5:{end_col}5")
        ws_wat.merge_cells(f"{start_col}6:{end_col}7")

        lbl_c = ws_wat[f"{start_col}5"]
        lbl_c.value = label
        lbl_c.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=lbl_color)
        lbl_c.alignment = Alignment(horizontal="center", vertical="center")
        lbl_c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        val_c = ws_wat[f"{start_col}6"]
        val_c.value = val_placeholder
        val_c.font = Font(name=FONT_FAMILY, size=val_size, bold=True, color=val_color)
        val_c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        val_c.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        for row in range(5, 8):
            for col_letter in [chr(c) for c in range(ord(start_col[0]), ord(end_col[0]) + 1)]:
                ws_wat[f"{col_letter}{row}"].border = create_thin_border(CARD_BORDER)

    # Section Title
    ws_wat["B9"] = "Timeline Segments & Waterfall"
    ws_wat["B9"].font = Font(name=FONT_FAMILY, size=11, bold=True, color=HEADER_DARK)

    # Timeline Table Headers (Row 10)
    # Left columns: Span attributes (B to I)
    # Right columns: Timeline bar visualization (J to AC: 20 buckets)
    span_headers = [
        ("Status", 12),                # B
        ("Hierarchy & Operation", 44), # C
        ("Service", 22),               # D
        ("Kind", 13),                  # E
        ("HTTP", 10),                  # F
        ("Start Offset", 16),          # G
        ("Duration (ms)", 16),         # H
        ("% of Trace", 12),            # I
    ]

    for col_idx, (header_name, col_width) in enumerate(span_headers, start=2):
        col_letter = get_column_letter(col_idx)
        cell = ws_wat.cell(row=10, column=col_idx, value=header_name)
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=HEADER_MID, end_color=HEADER_MID, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = create_thin_border(BORDER_DARK)
        ws_wat.column_dimensions[col_letter].width = col_width

    # Timeline Bar Buckets (Cols J to AC, 20 intervals: 0% to 100%)
    timeline_cols_count = 20
    timeline_start_col = 10  # Col J

    ws_wat.merge_cells(start_row=9, start_column=timeline_start_col, end_row=9, end_column=timeline_start_col + timeline_cols_count - 1)
    timeline_label = ws_wat.cell(row=9, column=timeline_start_col, value="Waterfall Timeline (0% ────────────────────────> 100%)")
    timeline_label.font = Font(name=FONT_FAMILY, size=8.5, bold=True, color=TEXT_MUTED)
    timeline_label.alignment = Alignment(horizontal="center", vertical="center")

    for i in range(timeline_cols_count):
        col_idx = timeline_start_col + i
        col_letter = get_column_letter(col_idx)
        pct_label = f"{int(i * 100 / timeline_cols_count)}%"
        cell = ws_wat.cell(row=10, column=col_idx, value=pct_label)
        cell.font = Font(name=FONT_FAMILY, size=8, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=HEADER_SCALE, end_color=HEADER_SCALE, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = create_thin_border(BORDER_DARK)
        ws_wat.column_dimensions[col_letter].width = 4.5

    ws_wat.row_dimensions[2].height = 26
    ws_wat.row_dimensions[3].height = 22
    ws_wat.row_dimensions[5].height = 18
    ws_wat.row_dimensions[6].height = 24
    ws_wat.row_dimensions[7].height = 18
    ws_wat.row_dimensions[9].height = 24
    ws_wat.row_dimensions[10].height = 26
    ws_wat.column_dimensions["A"].width = 3
    ws_wat.freeze_panes = "B11"

    # ----------------------------------------------------
    # Sheet 3: Raw_Data (Spans Flat Data)
    # ----------------------------------------------------
    ws_raw = wb.create_sheet(title="Raw_Data")
    ws_raw.views.sheetView[0].showGridLines = True

    # Title Banner (Row 2)
    ws_raw.merge_cells("A2:N2")
    raw_title = ws_raw["A2"]
    raw_title.value = "CloudWatch Logs Insights - Exported Spans (aws/spans)"
    raw_title.font = Font(name=FONT_FAMILY, size=14, bold=True, color=WHITE)
    raw_title.fill = PatternFill(start_color=HEADER_DARK, end_color=HEADER_DARK, fill_type="solid")
    raw_title.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    raw_headers = [
        ("timestamp", 24),
        ("trace_id", 38),
        ("span_id", 20),
        ("parent_span_id", 20),
        ("service_name", 22),
        ("operation_name", 32),
        ("duration_ms", 16),
        ("duration_nano", 18),
        ("http_status", 12),
        ("http_method", 12),
        ("http_url", 40),
        ("status_code", 14),
        ("status_message", 30),
        ("kind", 14),
    ]

    for col_idx, (h_name, h_width) in enumerate(raw_headers, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws_raw.cell(row=4, column=col_idx, value=h_name)
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=HEADER_MID, end_color=HEADER_MID, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = create_thin_border(BORDER_DARK)
        ws_raw.column_dimensions[col_letter].width = h_width

    ws_raw.row_dimensions[2].height = 28
    ws_raw.row_dimensions[4].height = 24
    ws_raw.freeze_panes = "A5"

    # Save template
    wb.save(output_path)
    print(f"Created Excel template successfully: {output_path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "templates/trace_report_template.xlsx"
    create_template(out)
