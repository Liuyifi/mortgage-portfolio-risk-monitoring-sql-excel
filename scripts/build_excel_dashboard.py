#!/usr/bin/env python3
"""Build the Excel dashboard from local aggregate DuckDB exports."""

from __future__ import annotations

import argparse
import csv
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = (
    PROJECT_ROOT
    / "outputs"
    / "mortgage_portfolio_risk_monitoring_sql_excel_dashboard.xlsx"
)

FONT = "Arial"
COLORS = {
    "navy": "#17365D",
    "blue": "#2F75B5",
    "teal": "#2A9D8F",
    "orange": "#F4A261",
    "deep_orange": "#E76F51",
    "red": "#C0392B",
    "gray": "#6B7280",
    "light_blue": "#DDEBF7",
    "light_gray": "#F3F4F6",
    "light_amber": "#FFF2CC",
    "white": "#FFFFFF",
    "dark": "#1F2937",
    "border": "#D1D5DB",
}

SHEET_NAMES = [
    "Portfolio Overview",
    "Risk Segments",
    "Transitions",
    "Monthly Data",
    "Segment Data",
    "Transition Data",
    "ReadMe",
]

MONTHLY_FIELDS = [
    "as_of_month",
    "on_book_loan_count",
    "on_book_upb",
    "delinquency_eligible_loan_count",
    "delinquency_eligible_upb",
    "excluded_pre_due_loan_count",
    "excluded_pre_due_upb",
    "excluded_loan_age_zero_loan_count",
    "excluded_loan_age_zero_upb",
    "excluded_invalid_loan_age_loan_count",
    "excluded_invalid_loan_age_upb",
    "excluded_non_numeric_status_loan_count",
    "excluded_non_numeric_status_upb",
    "dq30_loan_count",
    "dq60_loan_count",
    "dq90_loan_count",
    "dq30_upb",
    "dq60_upb",
    "dq90_upb",
]

SEGMENT_FIELDS = [
    "as_of_month",
    "segment_name",
    "segment_value",
    "privacy_merged",
    "source_mart_row_count",
    "on_book_loan_count",
    "on_book_upb",
    "delinquency_eligible_loan_count",
    "delinquency_eligible_upb",
    "dq30_loan_count",
    "dq60_loan_count",
    "dq90_loan_count",
    "dq30_upb",
    "dq60_upb",
    "dq90_upb",
    "segment_sort_order",
]

TRANSITION_FIELDS = [
    "from_month",
    "to_month",
    "from_state",
    "to_state",
    "privacy_merged",
    "source_mart_row_count",
    "transition_loan_count",
    "transition_from_upb",
    "from_state_sort_order",
    "to_state_sort_order",
]

INTEGER_FIELDS = {
    "on_book_loan_count",
    "delinquency_eligible_loan_count",
    "excluded_pre_due_loan_count",
    "excluded_loan_age_zero_loan_count",
    "excluded_invalid_loan_age_loan_count",
    "excluded_non_numeric_status_loan_count",
    "dq30_loan_count",
    "dq60_loan_count",
    "dq90_loan_count",
    "source_mart_row_count",
    "segment_sort_order",
    "transition_loan_count",
    "from_state_sort_order",
    "to_state_sort_order",
}

FLOAT_FIELDS = {
    "on_book_upb",
    "delinquency_eligible_upb",
    "excluded_pre_due_upb",
    "excluded_loan_age_zero_upb",
    "excluded_invalid_loan_age_upb",
    "excluded_non_numeric_status_upb",
    "dq30_upb",
    "dq60_upb",
    "dq90_upb",
    "transition_from_upb",
}

BOOLEAN_FIELDS = {"privacy_merged"}

MONTHLY_HEADERS = [
    "As of month",
    "On-book loans",
    "On-book UPB",
    "Eligible loans",
    "Eligible UPB",
    "Excluded PRE_DUE loans",
    "Excluded PRE_DUE UPB",
    "Excluded Loan Age 0 loans",
    "Excluded Loan Age 0 UPB",
    "Excluded invalid Loan Age loans",
    "Excluded invalid Loan Age UPB",
    "Excluded nonnumeric status loans",
    "Excluded nonnumeric status UPB",
    "30+ loans",
    "60+ loans",
    "90+ loans",
    "30+ UPB",
    "60+ UPB",
    "90+ UPB",
]

SEGMENT_HEADERS = [
    "As of month",
    "Segment name",
    "Segment value",
    "Privacy merged",
    "Source mart rows",
    "On-book loans",
    "On-book UPB",
    "Eligible loans",
    "Eligible UPB",
    "30+ loans",
    "60+ loans",
    "90+ loans",
    "30+ UPB",
    "60+ UPB",
    "90+ UPB",
    "Sort order",
]

TRANSITION_HEADERS = [
    "From month",
    "To month",
    "From state",
    "To state",
    "Privacy merged",
    "Source mart rows",
    "Transition loans",
    "From-month UPB",
    "From sort",
    "To sort",
]

FROM_STATES = [
    "PRE_DUE",
    "CURRENT",
    "30",
    "60",
    "90+",
    "REO_ACQUISITION",
    "UNKNOWN_LOAN_AGE",
    "UNKNOWN",
]

TO_STATES = [
    "PRE_DUE",
    "CURRENT",
    "30",
    "60",
    "90+",
    "REO_ACQUISITION",
    "TERMINATED:01",
    "TERMINATED:02",
    "TERMINATED:03",
    "TERMINATED:09",
    "TERMINATED:15",
    "TERMINATED:16",
    "TERMINATED:96",
    "UNKNOWN_LOAN_AGE",
    "UNKNOWN",
    "OTHER_OR_LOW_VOLUME",
]

@dataclass
class DashboardData:
    monthly: list[dict[str, Any]]
    segments: list[dict[str, Any]]
    transitions: list[dict[str, Any]]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--database",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "mortgage_risk.duckdb",
    )
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def cast_value(field: str, value: str) -> Any:
    if value == "":
        return None
    if field in INTEGER_FIELDS:
        return int(value)
    if field in FLOAT_FIELDS:
        return float(value)
    if field in BOOLEAN_FIELDS:
        return value.lower() == "true"
    return value


def read_projected_csv(path: Path, fields: list[str]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        missing = [field for field in fields if field not in (reader.fieldnames or [])]
        if missing:
            raise RuntimeError(f"Missing columns in {path}: {missing}")
        for source_row in reader:
            rows.append({field: cast_value(field, source_row[field]) for field in fields})
    return rows


def refresh_dashboard_exports(database: Path) -> Path:
    export_dir = PROJECT_ROOT / "data" / "processed" / "dashboard"
    subprocess.run(
        [
            sys.executable,
            str(PROJECT_ROOT / "scripts" / "export_dashboard_data.py"),
            "--database",
            str(database),
            "--output-dir",
            str(export_dir),
        ],
        cwd=PROJECT_ROOT,
        check=True,
    )
    return export_dir


def load_dashboard_data(database: Path) -> DashboardData:
    export_dir = refresh_dashboard_exports(database)
    monthly = read_projected_csv(export_dir / "portfolio_monthly.csv", MONTHLY_FIELDS)
    segments = read_projected_csv(
        export_dir / "portfolio_monthly_by_segment.csv", SEGMENT_FIELDS
    )
    transitions = read_projected_csv(
        export_dir / "delinquency_transition_monthly.csv", TRANSITION_FIELDS
    )
    if not monthly or not segments or not transitions:
        raise RuntimeError("One or more aggregate dashboard exports are empty")
    return DashboardData(monthly=monthly, segments=segments, transitions=transitions)


def excel_date(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%d")


def month_label(value: str) -> str:
    return value[:7]


def create_formats(workbook: xlsxwriter.Workbook) -> dict[str, Any]:
    base = {"font_name": FONT, "font_color": COLORS["dark"], "valign": "vcenter"}
    formats = {
        "title": workbook.add_format(
            {**base, "font_size": 15, "bold": True, "font_color": COLORS["navy"]}
        ),
        "subtitle": workbook.add_format(
            {**base, "font_size": 10, "italic": True, "font_color": COLORS["gray"]}
        ),
        "source_note": workbook.add_format(
            {**base, "font_size": 9, "italic": True, "font_color": COLORS["gray"]}
        ),
        "body": workbook.add_format({**base, "font_size": 9}),
        "body_text": workbook.add_format({**base, "font_size": 10}),
        "body_wrap": workbook.add_format({**base, "font_size": 10, "text_wrap": True}),
        "small_note": workbook.add_format(
            {**base, "font_size": 8, "italic": True, "font_color": COLORS["gray"]}
        ),
        "section_label": workbook.add_format(
            {
                **base,
                "font_size": 10,
                "bold": True,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["light_blue"],
            }
        ),
        "link_text": workbook.add_format(
            {**base, "font_size": 9, "font_color": COLORS["blue"]}
        ),
        "input": workbook.add_format(
            {
                **base,
                "font_size": 10,
                "bold": True,
                "bg_color": COLORS["light_amber"],
                "border": 1,
                "border_color": COLORS["orange"],
                "align": "center",
            }
        ),
        "header": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 9,
                "bold": True,
                "font_color": COLORS["white"],
                "bg_color": COLORS["navy"],
                "border": 1,
                "border_color": COLORS["white"],
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "card_label": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 9,
                "bold": True,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["light_blue"],
                "border": 1,
                "border_color": COLORS["border"],
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "card_integer": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 13,
                "bold": True,
                "font_color": COLORS["dark"],
                "bg_color": COLORS["white"],
                "border": 1,
                "border_color": COLORS["border"],
                "align": "center",
                "valign": "vcenter",
                "num_format": "#,##0",
            }
        ),
        "card_upb": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 13,
                "bold": True,
                "font_color": COLORS["dark"],
                "bg_color": COLORS["white"],
                "border": 1,
                "border_color": COLORS["border"],
                "align": "center",
                "valign": "vcenter",
                "num_format": '$0.00,,,"B"',
            }
        ),
        "card_percent": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 13,
                "bold": True,
                "font_color": COLORS["dark"],
                "bg_color": COLORS["white"],
                "border": 1,
                "border_color": COLORS["border"],
                "align": "center",
                "valign": "vcenter",
                "num_format": "0.00%",
            }
        ),
        "date": workbook.add_format({**base, "font_size": 9, "num_format": "yyyy-mm"}),
        "integer": workbook.add_format({**base, "font_size": 9, "num_format": "#,##0"}),
        "currency": workbook.add_format(
            {**base, "font_size": 9, "num_format": "$#,##0.00"}
        ),
        "billions": workbook.add_format(
            {**base, "font_size": 9, "num_format": '$0.00,,,"B"'}
        ),
        "percent": workbook.add_format({**base, "font_size": 9, "num_format": "0.00%"}),
        "percent_1": workbook.add_format({**base, "font_size": 9, "num_format": "0.0%"}),
        "text": workbook.add_format({**base, "font_size": 9}),
        "text_center": workbook.add_format({**base, "font_size": 9, "align": "center"}),
        "note_box": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 9,
                "italic": True,
                "font_color": COLORS["gray"],
                "bg_color": COLORS["light_gray"],
                "text_wrap": True,
                "valign": "vcenter",
            }
        ),
    }
    return formats


def configure_sheet(worksheet: Any, tab_color: str | None = None) -> None:
    worksheet.hide_gridlines(2)
    if tab_color:
        worksheet.set_tab_color(tab_color)


def write_title(
    worksheet: Any, cell_range: str, title: str, formats: dict[str, Any]
) -> None:
    worksheet.merge_range(cell_range, title, formats["title"])
    first_row = xlsxwriter.utility.xl_cell_to_rowcol(cell_range.split(":")[0])[0]
    worksheet.set_row(first_row, 24)


def write_subtitle(
    worksheet: Any, cell_range: str, text: str, formats: dict[str, Any]
) -> None:
    worksheet.merge_range(cell_range, text, formats["subtitle"])


def write_header_row(
    worksheet: Any, row: int, start_col: int, headers: list[str], header_format: Any
) -> None:
    for offset, header in enumerate(headers):
        worksheet.write(row, start_col + offset, header, header_format)


def write_card(
    worksheet: Any,
    label_range: str,
    value_range: str,
    label: str,
    formula: str,
    value_format: Any,
    formats: dict[str, Any],
) -> None:
    worksheet.merge_range(label_range, label, formats["card_label"])
    worksheet.merge_range(value_range, "", value_format)
    worksheet.write_formula(value_range.split(":")[0], formula, value_format)


def write_cell_value(worksheet: Any, row: int, col: int, value: Any) -> None:
    if isinstance(value, datetime):
        worksheet.write_datetime(row, col, value)
    elif isinstance(value, bool):
        worksheet.write_boolean(row, col, value)
    elif value is None:
        worksheet.write_blank(row, col, None)
    elif isinstance(value, (int, float)):
        worksheet.write_number(row, col, value)
    else:
        worksheet.write(row, col, value)


def add_data_table(
    worksheet: Any,
    headers: list[str],
    rows: list[list[Any]],
    table_name: str,
    column_formats: list[Any],
) -> int:
    first_row = 3
    for row_index, row_values in enumerate(rows, start=first_row + 1):
        for col_index, value in enumerate(row_values):
            write_cell_value(worksheet, row_index, col_index, value)
    last_row = first_row + len(rows)
    worksheet.add_table(
        first_row,
        0,
        last_row,
        len(headers) - 1,
        {
            "name": table_name,
            "style": "Table Style Medium 2",
            "columns": [
                {"header": header, "format": column_formats[index]}
                for index, header in enumerate(headers)
            ],
        },
    )
    return last_row + 1


def write_supporting_tables(
    sheets: dict[str, Any], data: DashboardData, formats: dict[str, Any]
) -> dict[str, int]:
    monthly_sheet = sheets["Monthly Data"]
    write_title(
        monthly_sheet,
        "A1:S1",
        "Monthly portfolio data",
        formats,
    )
    monthly_sheet.write(
        "A2",
        "Source: DuckDB mart.portfolio_monthly; Freddie Mac 2019 and 2020 samples.",
        formats["source_note"],
    )
    monthly_rows = [
        [
            excel_date(row["as_of_month"]),
            row["on_book_loan_count"],
            row["on_book_upb"],
            row["delinquency_eligible_loan_count"],
            row["delinquency_eligible_upb"],
            row["excluded_pre_due_loan_count"],
            row["excluded_pre_due_upb"],
            row["excluded_loan_age_zero_loan_count"],
            row["excluded_loan_age_zero_upb"],
            row["excluded_invalid_loan_age_loan_count"],
            row["excluded_invalid_loan_age_upb"],
            row["excluded_non_numeric_status_loan_count"],
            row["excluded_non_numeric_status_upb"],
            row["dq30_loan_count"],
            row["dq60_loan_count"],
            row["dq90_loan_count"],
            row["dq30_upb"],
            row["dq60_upb"],
            row["dq90_upb"],
        ]
        for row in data.monthly
    ]
    monthly_formats = [
        formats["date"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["integer"],
        formats["integer"],
        formats["currency"],
        formats["currency"],
        formats["currency"],
    ]
    monthly_end = add_data_table(
        monthly_sheet,
        MONTHLY_HEADERS,
        monthly_rows,
        "MonthlyAggregates",
        monthly_formats,
    )
    monthly_sheet.freeze_panes(4, 0)
    monthly_sheet.set_column("A:A", 12)
    monthly_sheet.set_column("B:S", 16)

    segment_sheet = sheets["Segment Data"]
    write_title(
        segment_sheet,
        "A1:P1",
        "Monthly segment data",
        formats,
    )
    segment_sheet.write(
        "A2",
        "Source: mart.portfolio_monthly_by_segment; displayed cells contain at least 20 loans.",
        formats["source_note"],
    )
    segment_rows = [
        [
            excel_date(row["as_of_month"]),
            row["segment_name"],
            row["segment_value"],
            row["privacy_merged"],
            row["source_mart_row_count"],
            row["on_book_loan_count"],
            row["on_book_upb"],
            row["delinquency_eligible_loan_count"],
            row["delinquency_eligible_upb"],
            row["dq30_loan_count"],
            row["dq60_loan_count"],
            row["dq90_loan_count"],
            row["dq30_upb"],
            row["dq60_upb"],
            row["dq90_upb"],
            row["segment_sort_order"],
        ]
        for row in data.segments
    ]
    segment_formats = [
        formats["date"],
        formats["text"],
        formats["text"],
        formats["text_center"],
        formats["integer"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["integer"],
        formats["integer"],
        formats["currency"],
        formats["currency"],
        formats["currency"],
        formats["integer"],
    ]
    segment_end = add_data_table(
        segment_sheet,
        SEGMENT_HEADERS,
        segment_rows,
        "SegmentAggregates",
        segment_formats,
    )
    segment_sheet.freeze_panes(4, 0)
    segment_sheet.set_column("A:A", 15)
    segment_sheet.set_column("B:C", 20)
    segment_sheet.set_column("D:P", 15)

    transition_sheet = sheets["Transition Data"]
    write_title(
        transition_sheet,
        "A1:J1",
        "Monthly transition data",
        formats,
    )
    transition_sheet.write(
        "A2",
        "Source: mart.delinquency_transition_monthly; exact next-month pairs only.",
        formats["source_note"],
    )
    transition_rows = [
        [
            excel_date(row["from_month"]),
            excel_date(row["to_month"]),
            row["from_state"],
            row["to_state"],
            row["privacy_merged"],
            row["source_mart_row_count"],
            row["transition_loan_count"],
            row["transition_from_upb"],
            row["from_state_sort_order"],
            row["to_state_sort_order"],
        ]
        for row in data.transitions
    ]
    transition_formats = [
        formats["date"],
        formats["date"],
        formats["text"],
        formats["text"],
        formats["text_center"],
        formats["integer"],
        formats["integer"],
        formats["currency"],
        formats["integer"],
        formats["integer"],
    ]
    transition_end = add_data_table(
        transition_sheet,
        TRANSITION_HEADERS,
        transition_rows,
        "TransitionAggregates",
        transition_formats,
    )
    transition_sheet.freeze_panes(4, 0)
    transition_sheet.set_column("A:B", 17)
    transition_sheet.set_column("C:D", 23)
    transition_sheet.set_column("E:J", 17)

    return {
        "monthly_end": monthly_end,
        "segment_end": segment_end,
        "transition_end": transition_end,
    }


def write_readme_sheet(
    worksheet: Any, formats: dict[str, Any]
) -> None:
    write_title(
        worksheet,
        "A1:F1",
        "Mortgage Portfolio Risk Monitoring: SQL & Excel Dashboard",
        formats,
    )
    rows = [
        ("Workbook scope", "Excel dashboard built from monthly aggregate tables in DuckDB."),
        ("Data source", "Freddie Mac SFLLD Release 47, 2019 and 2020 Standard Dataset annual samples."),
        ("Performance window", "2019-01 through 2026-03."),
        ("Population", "Sample month-end on-book loans; balances do not represent Freddie Mac's actual asset scale."),
        ("Delinquency denominator", "On-book, Loan Age >= 1, numeric 00-99 status."),
        ("30+/60+/90+", "Numeric delinquency status >= 1 / 2 / 3; 90+ means severe delinquency, not default."),
        ("Status coverage", "Eligible count or UPB divided by on-book count or UPB."),
        ("Transitions", "Exact next natural month only; balance rates use the origin month's UPB."),
        ("Low-volume cells", "No loan IDs. Segment cells below 20 loans are merged; origin-state transition cohorts below 20 are omitted."),
        ("Refresh", "Run .venv/bin/python scripts/build_excel_dashboard.py from the project root."),
        ("Workbook design", "The dashboard uses native Excel formulas, charts and dropdowns."),
        ("File use", "The workbook is designed for local review in Excel."),
    ]
    for row_index, (label, value) in enumerate(rows, start=2):
        worksheet.write(row_index, 0, label, formats["section_label"])
        worksheet.write(row_index, 1, value, formats["body_wrap"])

    sources = [
        ("Official source", "https://www.freddiemac.com/research/datasets/sf-loanlevel-dataset"),
        ("Release 47 guide", "https://www.freddiemac.com/fmac-resources/research/pdf/general_user_guide_july_2026.pdf"),
        ("Terms", "https://capitalmarkets.freddiemac.com/crt/docs/pdfs/fre_terms_conditions_sflld.pdf"),
        ("Data flow", "DuckDB mart tables -> aggregate CSV files -> workbook data tabs"),
    ]
    for row_index, (label, value) in enumerate(sources, start=15):
        worksheet.write(row_index, 0, label, formats["body_text"])
        worksheet.write(row_index, 1, value, formats["link_text"])

    mappings = [
        ["Order", "Origination Year", "Classic FICO", "Original CLTV", "Original DTI"],
        [1, "2019", "<620", "<=60", "<=20"],
        [2, "2020", "620-659", "61-70", "21-30"],
        [3, None, "660-699", "71-80", "31-40"],
        [4, None, "700-739", "81-90", "41-45"],
        [5, None, "740-779", "91-95", "46-50"],
        [6, None, "780+", ">95", ">50"],
        [7, None, "Unknown", "Unknown", "Unknown"],
        [8, None, "Other / low volume", "Other / low volume", "Other / low volume"],
    ]
    write_header_row(worksheet, 0, 7, mappings[0], formats["header"])
    for row_index, row_values in enumerate(mappings[1:], start=1):
        for col_index, value in enumerate(row_values, start=7):
            worksheet.write(row_index, col_index, value, formats["body"])
    worksheet.set_column("A:A", 24)
    worksheet.set_column("B:B", 95)
    worksheet.set_column("H:L", 20)


def style_line_chart(chart: Any, title: str, y_format: str) -> None:
    chart.set_title(
        {"name": title, "name_font": {"name": FONT, "size": 12, "bold": False}}
    )
    chart.set_legend(
        {"position": "top", "font": {"name": FONT, "size": 9}}
    )
    chart.set_x_axis(
        {
            "date_axis": True,
            "num_format": "yyyy-mm",
            "interval_unit": 12,
            "num_font": {"name": FONT, "size": 9},
            "line": {"color": COLORS["border"]},
        }
    )
    chart.set_y_axis(
        {
            "num_format": y_format,
            "num_font": {"name": FONT, "size": 9},
            "line": {"color": COLORS["border"]},
            "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}},
        }
    )
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    chart.set_size({"width": 640, "height": 300})


def style_bar_chart(chart: Any, title: str, value_format: str, width: int = 640) -> None:
    chart.set_title(
        {"name": title, "name_font": {"name": FONT, "size": 12, "bold": False}}
    )
    chart.set_legend(
        {"position": "top", "font": {"name": FONT, "size": 9}}
    )
    chart.set_x_axis(
        {
            "num_format": value_format,
            "num_font": {"name": FONT, "size": 9},
            "line": {"color": COLORS["border"]},
            "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}},
        }
    )
    chart.set_y_axis(
        {
            "num_font": {"name": FONT, "size": 9},
            "line": {"color": COLORS["border"]},
            "reverse": True,
        }
    )
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    chart.set_size({"width": width, "height": 300})


def add_overview_charts(
    workbook: xlsxwriter.Workbook, worksheet: Any, overview_end: int
) -> None:
    specs = [
        (
            "O3",
            "Monthly on-book loans",
            "#,##0",
            [("On-book loans", 1, COLORS["blue"])],
            False,
        ),
        (
            "W3",
            "Monthly sample UPB ($bn)",
            '$0.0,,,"B"',
            [("Sample UPB", 2, COLORS["teal"])],
            False,
        ),
        (
            "O19",
            "Delinquency rates by loan count",
            "0.0%",
            [
                ("30+ count rate", 7, COLORS["orange"]),
                ("60+ count rate", 8, COLORS["deep_orange"]),
                ("90+ count rate", 9, COLORS["red"]),
            ],
            True,
        ),
        (
            "W19",
            "Delinquency rates by sample UPB",
            "0.0%",
            [
                ("30+ balance rate", 10, COLORS["orange"]),
                ("60+ balance rate", 11, COLORS["deep_orange"]),
                ("90+ balance rate", 12, COLORS["red"]),
            ],
            True,
        ),
    ]
    for anchor, title, number_format, series_specs, show_legend in specs:
        chart = workbook.add_chart({"type": "line"})
        for series_name, value_col, color in series_specs:
            chart.add_series(
                {
                    "name": series_name,
                    "categories": ["Portfolio Overview", 12, 0, overview_end - 1, 0],
                    "values": [
                        "Portfolio Overview",
                        12,
                        value_col,
                        overview_end - 1,
                        value_col,
                    ],
                    "line": {"color": color, "width": 2.25},
                }
            )
        style_line_chart(chart, title, number_format)
        if not show_legend:
            chart.set_legend({"none": True})
        worksheet.insert_chart(anchor, chart)


def write_portfolio_overview(
    workbook: xlsxwriter.Workbook,
    worksheet: Any,
    data: DashboardData,
    formats: dict[str, Any],
    monthly_end: int,
) -> None:
    write_title(worksheet, "A2:M2", "Portfolio Overview", formats)
    write_subtitle(
        worksheet,
        "A3:M3",
        "2019/2020 are origination vintages. Monthly performance continues through 2026-03; trends reflect entries, exits and changing sample composition.",
        formats,
    )
    worksheet.write("A4", "Snapshot month", formats["body_text"])
    worksheet.write_string("B4", month_label(data.monthly[-1]["as_of_month"]), formats["input"])
    worksheet.write_formula("AZ1", "=DATE(VALUE(LEFT(B4,4)),VALUE(RIGHT(B4,2)),1)")
    for row_index, row in enumerate(data.monthly, start=1):
        worksheet.write_string(row_index, 51, month_label(row["as_of_month"]))
    worksheet.set_column("AZ:AZ", 12)
    worksheet.data_validation(
        "B4", {"validate": "list", "source": "=$AZ$2:$AZ$88"}
    )

    write_card(
        worksheet,
        "A6:B6",
        "A7:B8",
        "On-book loans",
        f"=SUMIFS('Monthly Data'!$B$5:$B${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$AZ$1)",
        formats["card_integer"],
        formats,
    )
    write_card(
        worksheet,
        "C6:D6",
        "C7:D8",
        "Sample UPB",
        f"=SUMIFS('Monthly Data'!$C$5:$C${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$AZ$1)",
        formats["card_upb"],
        formats,
    )
    write_card(
        worksheet,
        "E6:F6",
        "E7:F8",
        "Eligible loans",
        f"=SUMIFS('Monthly Data'!$D$5:$D${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$AZ$1)",
        formats["card_integer"],
        formats,
    )
    write_card(
        worksheet,
        "G6:H6",
        "G7:H8",
        "Eligible UPB",
        f"=SUMIFS('Monthly Data'!$E$5:$E${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$AZ$1)",
        formats["card_upb"],
        formats,
    )
    write_card(
        worksheet,
        "I6:J6",
        "I7:J8",
        "Count status coverage",
        '=IF(A7=0,"",E7/A7)',
        formats["card_percent"],
        formats,
    )
    write_card(
        worksheet,
        "K6:L6",
        "K7:L8",
        "Balance status coverage",
        '=IF(C7=0,"",G7/C7)',
        formats["card_percent"],
        formats,
    )

    headers = [
        "Month",
        "On-book loans",
        "Sample UPB",
        "Eligible loans",
        "Eligible UPB",
        "Count coverage",
        "Balance coverage",
        "30+ count rate",
        "60+ count rate",
        "90+ count rate",
        "30+ balance rate",
        "60+ balance rate",
        "90+ balance rate",
    ]
    write_header_row(worksheet, 11, 0, headers, formats["header"])
    for index in range(len(data.monthly)):
        row = 13 + index
        source_row = 5 + index
        formulas = [
            f"='Monthly Data'!A{source_row}",
            f"='Monthly Data'!B{source_row}",
            f"='Monthly Data'!C{source_row}",
            f"='Monthly Data'!D{source_row}",
            f"='Monthly Data'!E{source_row}",
            f'=IF(B{row}=0,"",D{row}/B{row})',
            f'=IF(C{row}=0,"",E{row}/C{row})',
            f'=IF(D{row}=0,"",\'Monthly Data\'!N{source_row}/D{row})',
            f'=IF(D{row}=0,"",\'Monthly Data\'!O{source_row}/D{row})',
            f'=IF(D{row}=0,"",\'Monthly Data\'!P{source_row}/D{row})',
            f'=IF(E{row}=0,"",\'Monthly Data\'!Q{source_row}/E{row})',
            f'=IF(E{row}=0,"",\'Monthly Data\'!R{source_row}/E{row})',
            f'=IF(E{row}=0,"",\'Monthly Data\'!S{source_row}/E{row})',
        ]
        row_formats = [
            formats["date"],
            formats["integer"],
            formats["billions"],
            formats["integer"],
            formats["billions"],
        ] + [formats["percent"]] * 8
        for col_index, formula in enumerate(formulas):
            worksheet.write_formula(row - 1, col_index, formula, row_formats[col_index])

    overview_end = 12 + len(data.monthly)
    worksheet.add_table(
        11,
        0,
        overview_end - 1,
        12,
        {
            "name": "PortfolioTrendTable",
            "style": "Table Style Medium 2",
            "columns": [{"header": header} for header in headers],
        },
    )
    worksheet.freeze_panes(4, 0)
    worksheet.set_column("A:A", 12)
    worksheet.set_column("B:M", 14)
    add_overview_charts(workbook, worksheet, overview_end)


def add_segment_charts(workbook: xlsxwriter.Workbook, worksheet: Any) -> None:
    specs = [
        (
            "P4",
            "On-book loans by segment",
            "#,##0",
            [("On-book loans", 1, COLORS["blue"])],
            False,
        ),
        (
            "X4",
            "Sample UPB by segment ($bn)",
            '$0.0,,,"B"',
            [("Sample UPB", 2, COLORS["teal"])],
            False,
        ),
        (
            "P19",
            "Delinquency rates by loan count",
            "0.0%",
            [
                ("30+ count", 7, COLORS["orange"]),
                ("60+ count", 8, COLORS["deep_orange"]),
                ("90+ count", 9, COLORS["red"]),
            ],
            True,
        ),
        (
            "X19",
            "Delinquency rates by sample UPB",
            "0.0%",
            [
                ("30+ balance", 10, COLORS["orange"]),
                ("60+ balance", 11, COLORS["deep_orange"]),
                ("90+ balance", 12, COLORS["red"]),
            ],
            True,
        ),
    ]
    for anchor, title, number_format, series_specs, show_legend in specs:
        chart = workbook.add_chart({"type": "bar"})
        for series_name, value_col, color in series_specs:
            chart.add_series(
                {
                    "name": series_name,
                    "categories": ["Risk Segments", 9, 0, 16, 0],
                    "values": ["Risk Segments", 9, value_col, 16, value_col],
                    "fill": {"color": color},
                    "border": {"color": color},
                }
            )
        style_bar_chart(chart, title, number_format)
        if not show_legend:
            chart.set_legend({"none": True})
        worksheet.insert_chart(anchor, chart)


def write_risk_segments(
    workbook: xlsxwriter.Workbook,
    worksheet: Any,
    data: DashboardData,
    formats: dict[str, Any],
    segment_end: int,
) -> None:
    write_title(worksheet, "A2:N2", "Risk Segments", formats)
    write_subtitle(
        worksheet,
        "A3:N3",
        "Rates are recalculated as summed numerators divided by summed denominators. Unknown values remain explicit; low-volume cells are combined.",
        formats,
    )
    worksheet.write_row("A4", ["Observation month", None, "Dimension", None])
    worksheet.write_string("B4", month_label(data.monthly[-1]["as_of_month"]), formats["input"])
    worksheet.write_string("D4", "Origination Year", formats["input"])
    worksheet.write_formula("AZ1", "=DATE(VALUE(LEFT(B4,4)),VALUE(RIGHT(B4,2)),1)")
    for row_index, row in enumerate(data.monthly, start=1):
        worksheet.write_string(row_index, 51, month_label(row["as_of_month"]))
    dimensions = ["Origination Year", "Classic FICO", "Original CLTV", "Original DTI"]
    for row_index, value in enumerate(dimensions, start=1):
        worksheet.write_string(row_index, 52, value)
    worksheet.data_validation(
        "B4", {"validate": "list", "source": "=$AZ$2:$AZ$88"}
    )
    worksheet.data_validation(
        "D4", {"validate": "list", "source": "=$BA$2:$BA$5"}
    )
    worksheet.write("C5", "Internal key", formats["small_note"])
    worksheet.write_formula(
        "D5",
        '=IF(D4="Origination Year","origination_year",IF(D4="Classic FICO","classic_fico",IF(D4="Original CLTV","original_cltv","original_dti")))',
        formats["small_note"],
    )

    headers = [
        "Segment",
        "On-book loans",
        "Sample UPB",
        "Eligible loans",
        "Eligible UPB",
        "Count coverage",
        "Balance coverage",
        "30+ count",
        "60+ count",
        "90+ count",
        "30+ balance",
        "60+ balance",
        "90+ balance",
        "Low-volume note",
    ]
    write_header_row(worksheet, 8, 0, headers, formats["header"])
    for index in range(8):
        row = 10 + index
        formulas = [
            f'=IF($D$4="Origination Year",INDEX(\'ReadMe\'!$I$2:$I$9,{index + 1}),IF($D$4="Classic FICO",INDEX(\'ReadMe\'!$J$2:$J$9,{index + 1}),IF($D$4="Original CLTV",INDEX(\'ReadMe\'!$K$2:$K$9,{index + 1}),INDEX(\'ReadMe\'!$L$2:$L$9,{index + 1}))))',
            f'=IF(A{row}="","",SUMIFS(\'Segment Data\'!$F$5:$F${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row}))',
            f'=IF(A{row}="","",SUMIFS(\'Segment Data\'!$G$5:$G${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row}))',
            f'=IF(A{row}="","",SUMIFS(\'Segment Data\'!$H$5:$H${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row}))',
            f'=IF(A{row}="","",SUMIFS(\'Segment Data\'!$I$5:$I${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row}))',
            f'=IF(OR(A{row}="",B{row}=0),"",D{row}/B{row})',
            f'=IF(OR(A{row}="",C{row}=0),"",E{row}/C{row})',
            f'=IF(OR(A{row}="",D{row}=0),"",SUMIFS(\'Segment Data\'!$J$5:$J${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row})/D{row})',
            f'=IF(OR(A{row}="",D{row}=0),"",SUMIFS(\'Segment Data\'!$K$5:$K${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row})/D{row})',
            f'=IF(OR(A{row}="",D{row}=0),"",SUMIFS(\'Segment Data\'!$L$5:$L${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row})/D{row})',
            f'=IF(OR(A{row}="",E{row}=0),"",SUMIFS(\'Segment Data\'!$M$5:$M${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row})/E{row})',
            f'=IF(OR(A{row}="",E{row}=0),"",SUMIFS(\'Segment Data\'!$N$5:$N${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row})/E{row})',
            f'=IF(OR(A{row}="",E{row}=0),"",SUMIFS(\'Segment Data\'!$O$5:$O${segment_end},\'Segment Data\'!$A$5:$A${segment_end},$AZ$1,\'Segment Data\'!$B$5:$B${segment_end},$D$5,\'Segment Data\'!$C$5:$C${segment_end},A{row})/E{row})',
            f'=IF(A{row}="Other / low volume","Merged to minimum 20","")',
        ]
        row_formats = [
            formats["text"],
            formats["integer"],
            formats["billions"],
            formats["integer"],
            formats["billions"],
        ] + [formats["percent"]] * 8 + [formats["small_note"]]
        for col_index, formula in enumerate(formulas):
            worksheet.write_formula(row - 1, col_index, formula, row_formats[col_index])
    worksheet.set_column("A:A", 20)
    worksheet.set_column("B:C", 13)
    worksheet.set_column("D:D", 20)
    worksheet.set_column("E:M", 13)
    worksheet.set_column("N:N", 20)
    add_segment_charts(workbook, worksheet)


def add_transition_chart(workbook: xlsxwriter.Workbook, worksheet: Any) -> None:
    chart = workbook.add_chart({"type": "bar"})
    for name, col, color in [
        ("Count rate", 19, COLORS["blue"]),
        ("Balance rate", 20, COLORS["teal"]),
    ]:
        chart.add_series(
            {
                "name": name,
                "categories": ["Transitions", 22, 18, 37, 18],
                "values": ["Transitions", 22, col, 37, col],
                "fill": {"color": color},
                "border": {"color": color},
            }
        )
    style_bar_chart(
        chart,
        "Destination mix for selected origin state",
        "0.0%",
        width=1000,
    )
    worksheet.insert_chart("S4", chart)


def write_transitions(
    workbook: xlsxwriter.Workbook,
    worksheet: Any,
    data: DashboardData,
    formats: dict[str, Any],
    transition_end: int,
) -> None:
    write_title(worksheet, "A2:Q2", "Delinquency Transitions", formats)
    write_subtitle(
        worksheet,
        "A3:Q3",
        "Only exact next-natural-month pairs are included. Rates use loans observable in the following month; balance rates use origin-month UPB.",
        formats,
    )
    worksheet.write_row("A4", ["From month", None, "To month", None, "Focus state", None])
    latest_from = max(row["from_month"] for row in data.transitions)
    worksheet.write_string("B4", month_label(latest_from), formats["input"])
    worksheet.write_formula("AZ1", "=DATE(VALUE(LEFT(B4,4)),VALUE(RIGHT(B4,2)),1)")
    worksheet.write_formula("D4", "=EDATE($AZ$1,1)", formats["date"])
    worksheet.write_string("F4", "CURRENT", formats["input"])
    for row_index, row in enumerate(data.monthly[:-1], start=1):
        worksheet.write_string(row_index, 51, month_label(row["as_of_month"]))
    for row_index, state in enumerate(FROM_STATES, start=1):
        worksheet.write_string(row_index, 52, state)
    worksheet.data_validation(
        "B4", {"validate": "list", "source": "=$AZ$2:$AZ$87"}
    )
    worksheet.data_validation(
        "F4", {"validate": "list", "source": "=$BA$2:$BA$9"}
    )
    write_card(
        worksheet,
        "A6:B6",
        "A7:B8",
        "Observable loans",
        f"=SUMIFS('Transition Data'!$G$5:$G${transition_end},'Transition Data'!$A$5:$A${transition_end},$AZ$1,'Transition Data'!$C$5:$C${transition_end},$F$4)",
        formats["card_integer"],
        formats,
    )
    write_card(
        worksheet,
        "C6:D6",
        "C7:D8",
        "Origin-month sample UPB",
        f"=SUMIFS('Transition Data'!$H$5:$H${transition_end},'Transition Data'!$A$5:$A${transition_end},$AZ$1,'Transition Data'!$C$5:$C${transition_end},$F$4)",
        formats["card_upb"],
        formats,
    )
    write_card(
        worksheet,
        "E6:F6",
        "E7:F8",
        "Visible destinations",
        f"=COUNTIFS('Transition Data'!$A$5:$A${transition_end},$AZ$1,'Transition Data'!$C$5:$C${transition_end},$F$4)",
        formats["card_integer"],
        formats,
    )
    worksheet.merge_range(
        "A9:Q10",
        "Privacy note: low-volume destinations are combined as OTHER_OR_LOW_VOLUME. Entire from-month × from-state cohorts below 20 loans are suppressed, so a blank cell is not automatically zero.",
        formats["note_box"],
    )

    worksheet.write("A12", "Count migration rate", formats["header"])
    write_header_row(worksheet, 11, 1, TO_STATES, formats["header"])
    for index, state in enumerate(FROM_STATES, start=12):
        worksheet.write(index, 0, state, formats["text"])
    for state_index in range(len(FROM_STATES)):
        row = 13 + state_index
        for destination_index in range(len(TO_STATES)):
            col_name = xl_col_to_name(1 + destination_index)
            formula = (
                f'=IF(SUMIFS(\'Transition Data\'!$G$5:$G${transition_end},'
                f'\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,'
                f'\'Transition Data\'!$C$5:$C${transition_end},$A{row})=0,"",'
                f'SUMIFS(\'Transition Data\'!$G$5:$G${transition_end},'
                f'\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,'
                f'\'Transition Data\'!$C$5:$C${transition_end},$A{row},'
                f'\'Transition Data\'!$D$5:$D${transition_end},{col_name}$12)/'
                f'SUMIFS(\'Transition Data\'!$G$5:$G${transition_end},'
                f'\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,'
                f'\'Transition Data\'!$C$5:$C${transition_end},$A{row}))'
            )
            worksheet.write_formula(row - 1, 1 + destination_index, formula, formats["percent_1"])
    worksheet.conditional_format(
        "B13:Q20",
        {
            "type": "3_color_scale",
            "min_color": "#FFFFFF",
            "mid_color": "#FDE68A",
            "max_color": COLORS["red"],
            "mid_type": "percentile",
            "mid_value": 50,
        },
    )

    worksheet.write("A24", "Balance-weighted migration rate", formats["header"])
    write_header_row(worksheet, 23, 1, TO_STATES, formats["header"])
    for index, state in enumerate(FROM_STATES, start=24):
        worksheet.write(index, 0, state, formats["text"])
    for state_index in range(len(FROM_STATES)):
        row = 25 + state_index
        for destination_index in range(len(TO_STATES)):
            col_name = xl_col_to_name(1 + destination_index)
            formula = (
                f'=IF(SUMIFS(\'Transition Data\'!$H$5:$H${transition_end},'
                f'\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,'
                f'\'Transition Data\'!$C$5:$C${transition_end},$A{row})=0,"",'
                f'SUMIFS(\'Transition Data\'!$H$5:$H${transition_end},'
                f'\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,'
                f'\'Transition Data\'!$C$5:$C${transition_end},$A{row},'
                f'\'Transition Data\'!$D$5:$D${transition_end},{col_name}$24)/'
                f'SUMIFS(\'Transition Data\'!$H$5:$H${transition_end},'
                f'\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,'
                f'\'Transition Data\'!$C$5:$C${transition_end},$A{row}))'
            )
            worksheet.write_formula(row - 1, 1 + destination_index, formula, formats["percent_1"])
    worksheet.conditional_format(
        "B25:Q32",
        {
            "type": "3_color_scale",
            "min_color": "#FFFFFF",
            "mid_color": "#FDE68A",
            "max_color": COLORS["red"],
            "mid_type": "percentile",
            "mid_value": 50,
        },
    )

    write_header_row(
        worksheet,
        34,
        0,
        ["From state", "Observable loans", "Observable origin UPB"],
        formats["header"],
    )
    for index, state in enumerate(FROM_STATES):
        row = 36 + index
        worksheet.write(row - 1, 0, state, formats["text"])
        worksheet.write_formula(
            row - 1,
            1,
            f"=SUMIFS('Transition Data'!$G$5:$G${transition_end},'Transition Data'!$A$5:$A${transition_end},$AZ$1,'Transition Data'!$C$5:$C${transition_end},$A{row})",
            formats["integer"],
        )
        worksheet.write_formula(
            row - 1,
            2,
            f"=SUMIFS('Transition Data'!$H$5:$H${transition_end},'Transition Data'!$A$5:$A${transition_end},$AZ$1,'Transition Data'!$C$5:$C${transition_end},$A{row})",
            formats["billions"],
        )

    write_header_row(
        worksheet,
        21,
        18,
        ["Destination", "Count rate", "Balance rate"],
        formats["header"],
    )
    for index, state in enumerate(TO_STATES):
        row = 23 + index
        worksheet.write(row - 1, 18, state, formats["text"])
        worksheet.write_formula(
            row - 1,
            19,
            f'=IF(SUMIFS(\'Transition Data\'!$G$5:$G${transition_end},\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,\'Transition Data\'!$C$5:$C${transition_end},$F$4)=0,"",SUMIFS(\'Transition Data\'!$G$5:$G${transition_end},\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,\'Transition Data\'!$C$5:$C${transition_end},$F$4,\'Transition Data\'!$D$5:$D${transition_end},S{row})/SUMIFS(\'Transition Data\'!$G$5:$G${transition_end},\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,\'Transition Data\'!$C$5:$C${transition_end},$F$4))',
            formats["percent_1"],
        )
        worksheet.write_formula(
            row - 1,
            20,
            f'=IF(SUMIFS(\'Transition Data\'!$H$5:$H${transition_end},\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,\'Transition Data\'!$C$5:$C${transition_end},$F$4)=0,"",SUMIFS(\'Transition Data\'!$H$5:$H${transition_end},\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,\'Transition Data\'!$C$5:$C${transition_end},$F$4,\'Transition Data\'!$D$5:$D${transition_end},S{row})/SUMIFS(\'Transition Data\'!$H$5:$H${transition_end},\'Transition Data\'!$A$5:$A${transition_end},$AZ$1,\'Transition Data\'!$C$5:$C${transition_end},$F$4))',
            formats["percent_1"],
        )
    worksheet.set_column("A:A", 20)
    worksheet.set_column("B:Q", 12)
    worksheet.set_column("S:S", 24)
    worksheet.set_column("T:U", 14)
    worksheet.set_row(11, 44)
    worksheet.set_row(23, 44)
    add_transition_chart(workbook, worksheet)


def build_workbook(output: Path, data: DashboardData) -> None:
    workbook = xlsxwriter.Workbook(output)
    workbook.set_properties(
        {
            "title": "Mortgage Portfolio Risk Monitoring with SQL and Excel",
            "subject": "Aggregate mortgage portfolio risk monitoring dashboard",
            "keywords": "mortgage risk, DuckDB, SQL, Python, Excel",
            "comments": "Built from aggregate mortgage portfolio tables without loan-level rows.",
        }
    )
    workbook.set_calc_mode("auto")
    formats = create_formats(workbook)
    sheets = {name: workbook.add_worksheet(name) for name in SHEET_NAMES}
    for name, worksheet in sheets.items():
        tab_color = None
        if name in {"Portfolio Overview", "Risk Segments", "Transitions"}:
            tab_color = COLORS["navy"]
        elif name == "ReadMe":
            tab_color = "#9CA3AF"
        configure_sheet(worksheet, tab_color)

    endpoints = write_supporting_tables(sheets, data, formats)
    write_readme_sheet(sheets["ReadMe"], formats)
    write_portfolio_overview(
        workbook,
        sheets["Portfolio Overview"],
        data,
        formats,
        endpoints["monthly_end"],
    )
    write_risk_segments(
        workbook,
        sheets["Risk Segments"],
        data,
        formats,
        endpoints["segment_end"],
    )
    write_transitions(
        workbook,
        sheets["Transitions"],
        data,
        formats,
        endpoints["transition_end"],
    )
    workbook.close()


def validate_dashboard_data(data: DashboardData) -> None:
    """Check the small set of assumptions used by the workbook formulas."""
    datasets = {
        "monthly": (data.monthly, lambda row: row["as_of_month"]),
        "segments": (
            data.segments,
            lambda row: (row["as_of_month"], row["segment_name"], row["segment_value"]),
        ),
        "transitions": (
            data.transitions,
            lambda row: (row["from_month"], row["from_state"], row["to_state"]),
        ),
    }
    for name, (rows, key_function) in datasets.items():
        if not rows:
            raise RuntimeError(f"{name} dashboard data is empty")
        keys = [key_function(row) for row in rows]
        if len(keys) != len(set(keys)):
            raise RuntimeError(f"{name} dashboard data contains duplicate keys")
        if any("loan_id" in field.lower() for row in rows for field in row):
            raise RuntimeError(f"{name} dashboard data contains a loan identifier column")

    if len(data.monthly) != 87:
        raise RuntimeError(f"Expected 87 monthly rows, found {len(data.monthly)}")
    if data.monthly[0]["as_of_month"] != "2019-01-01":
        raise RuntimeError("Monthly data does not start at 2019-01")
    if data.monthly[-1]["as_of_month"] != "2026-03-01":
        raise RuntimeError("Monthly data does not end at 2026-03")


def main() -> int:
    args = parse_args()
    database = args.database.resolve()
    output = args.output.resolve()
    if not database.is_file():
        raise FileNotFoundError(f"DuckDB database not found: {database}")

    data = load_dashboard_data(database)
    validate_dashboard_data(data)
    output.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(output, data)
    if not output.is_file() or output.stat().st_size == 0:
        raise RuntimeError("XlsxWriter did not create the workbook")

    print(f"Workbook: {output}")
    print(
        f"Data rows: {len(data.monthly):,} monthly, "
        f"{len(data.segments):,} segment, {len(data.transitions):,} transition"
    )
    print(f"Worksheets: {len(SHEET_NAMES)}; charts: 9; dropdowns: 5")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
