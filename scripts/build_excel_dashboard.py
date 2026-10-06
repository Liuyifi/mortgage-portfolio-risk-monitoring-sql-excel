#!/usr/bin/env python3
"""Build and validate the offline Excel dashboard with XlsxWriter.

The script refreshes the existing governed aggregate exports, reads the
staging-independent DuckDB reconciliation samples, writes only aggregate data
to the workbook, and validates the saved XLSX package with Python's standard
library. The reference workbook is optional and is never modified.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import posixpath
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import duckdb
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


PROJECT_ROOT = Path(__file__).resolve().parents[1]
REFERENCE_WORKBOOK = (
    PROJECT_ROOT / "outputs" / "mortgage_portfolio_risk_monitoring_sql_excel_dashboard.xlsx"
)
DEFAULT_OUTPUT = (
    PROJECT_ROOT
    / "outputs"
    / "mortgage_portfolio_risk_monitoring_sql_excel_dashboard_xlsxwriter_candidate.xlsx"
)
DEFAULT_SUMMARY = (
    PROJECT_ROOT
    / "data"
    / "processed"
    / "validation"
    / "excel_dashboard_xlsxwriter_build_summary.json"
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
    "Validation",
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

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_NS = {
    "m": MAIN_NS,
    "r": REL_NS,
    "pr": PKG_REL_NS,
    "c": CHART_NS,
    "a": DRAWING_NS,
}


@dataclass
class DashboardData:
    monthly: list[dict[str, Any]]
    segments: list[dict[str, Any]]
    transitions: list[dict[str, Any]]
    validation_samples: list[dict[str, Any]]


@dataclass
class WorkbookSnapshot:
    sheet_names: list[str]
    sheet_states: dict[str, str]
    cells: dict[str, dict[str, Any]]
    formulas: dict[str, dict[str, str]]
    merges: dict[str, list[str]]
    validations: dict[str, list[dict[str, str]]]
    panes: dict[str, list[dict[str, str]]]
    tables: dict[str, list[dict[str, str]]]
    charts: list[dict[str, Any]]
    defined_names: list[dict[str, Any]]
    external_links: list[str]
    vba_parts: list[str]
    power_query_parts: list[str]
    formula_error_tokens: list[str]
    chinese_character_count: int
    calc_mode: str | None
    full_calc_on_load: str | None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--database",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "mortgage_risk.duckdb",
    )
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--summary", type=Path, default=DEFAULT_SUMMARY)
    parser.add_argument("--reference", type=Path, default=REFERENCE_WORKBOOK)
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


def read_validation_samples(database: Path) -> list[dict[str, Any]]:
    sql = """
        SELECT
            strftime(as_of_month, '%Y-%m-%d') AS as_of_month,
            independent_on_book_loan_count AS on_book_loan_count,
            CAST(independent_on_book_upb AS DOUBLE) AS on_book_upb,
            independent_eligible_loan_count AS eligible_loan_count,
            CAST(independent_eligible_upb AS DOUBLE) AS eligible_upb,
            independent_dq30_loan_count / NULLIF(independent_eligible_loan_count, 0)::DOUBLE AS dq30_count_rate,
            independent_dq60_loan_count / NULLIF(independent_eligible_loan_count, 0)::DOUBLE AS dq60_count_rate,
            independent_dq90_loan_count / NULLIF(independent_eligible_loan_count, 0)::DOUBLE AS dq90_count_rate,
            CAST(independent_dq30_upb AS DOUBLE) / NULLIF(CAST(independent_eligible_upb AS DOUBLE), 0) AS dq30_balance_rate,
            CAST(independent_dq60_upb AS DOUBLE) / NULLIF(CAST(independent_eligible_upb AS DOUBLE), 0) AS dq60_balance_rate,
            CAST(independent_dq90_upb AS DOUBLE) / NULLIF(CAST(independent_eligible_upb AS DOUBLE), 0) AS dq90_balance_rate,
            reconciled
        FROM quality.metric_reconciliation_samples
        ORDER BY as_of_month
    """
    with duckdb.connect(str(database), read_only=True) as connection:
        cursor = connection.execute(sql)
        columns = [item[0] for item in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
    expected_months = ["2020-06-01", "2021-06-01", "2026-03-01"]
    if [row["as_of_month"] for row in rows] != expected_months:
        raise RuntimeError(f"Unexpected validation months: {rows}")
    if not all(row["reconciled"] for row in rows):
        raise RuntimeError("DuckDB staging-independent sample reconciliation is not PASS")
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
    return DashboardData(
        monthly=monthly,
        segments=segments,
        transitions=transitions,
        validation_samples=read_validation_samples(database),
    )


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
        "percent_4": workbook.add_format({**base, "font_size": 9, "num_format": "0.0000%"}),
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
        "validation_status": workbook.add_format(
            {
                "font_name": FONT,
                "font_size": 11,
                "bold": True,
                "font_color": COLORS["navy"],
                "bg_color": COLORS["light_blue"],
                "border": 1,
                "border_color": COLORS["border"],
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
        "Embedded monthly portfolio aggregates",
        formats,
    )
    monthly_sheet.write(
        "A2",
        "Source: validated DuckDB mart.portfolio_monthly; Freddie Mac SFLLD Release 47 samples.",
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
        "Embedded monthly segment aggregates",
        formats,
    )
    segment_sheet.write(
        "A2",
        "Source: local display export of mart.portfolio_monthly_by_segment; minimum displayed cell is 20 loans.",
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
        "Embedded delinquency transition aggregates",
        formats,
    )
    transition_sheet.write(
        "A2",
        "Source: local overall-scope display export of mart.delinquency_transition_monthly; exact next-month pairs only.",
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
        ("Workbook scope", "Offline Excel dashboard built from validated aggregate DuckDB marts."),
        ("Data source", "Freddie Mac SFLLD Release 47, 2019 and 2020 Standard Dataset annual samples."),
        ("Performance window", "2019-01 through 2026-03."),
        ("Population", "Sample month-end on-book loans; balances do not represent Freddie Mac's actual asset scale."),
        ("Delinquency denominator", "On-book, Loan Age >= 1, numeric 00-99 status."),
        ("30+/60+/90+", "Numeric delinquency status >= 1 / 2 / 3; 90+ means severe delinquency, not default."),
        ("Status coverage", "Eligible count or UPB divided by on-book count or UPB."),
        ("Transitions", "Exact next natural month only; balance rates use the origin month's UPB."),
        ("Privacy", "No loan IDs. Segment cells below 20 loans are merged; origin-state transition cohorts below 20 are suppressed."),
        ("Refresh", "Run .venv/bin/python scripts/build_excel_dashboard.py from the project root."),
        ("Workbook design", "No macros, external data connections, or cloud account."),
        ("Delivery", "Offline Excel workbook for local use."),
    ]
    for row_index, (label, value) in enumerate(rows, start=2):
        worksheet.write(row_index, 0, label, formats["section_label"])
        worksheet.write(row_index, 1, value, formats["body_wrap"])

    sources = [
        ("Official source", "https://www.freddiemac.com/research/datasets/sf-loanlevel-dataset"),
        ("Release 47 guide", "https://www.freddiemac.com/fmac-resources/research/pdf/general_user_guide_july_2026.pdf"),
        ("Terms", "https://capitalmarkets.freddiemac.com/crt/docs/pdfs/fre_terms_conditions_sflld.pdf"),
        ("Local lineage", "DuckDB mart tables -> local aggregate CSV -> embedded workbook data tabs"),
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


def write_validation_sheet(
    worksheet: Any,
    data: DashboardData,
    formats: dict[str, Any],
    monthly_end: int,
) -> None:
    write_title(
        worksheet,
        "A1:AB1",
        "Validation against DuckDB sample months",
        formats,
    )
    write_subtitle(
        worksheet,
        "A2:AB2",
        "All rates are recomputed from embedded additive components. A zero in each Check column means the workbook matches the DuckDB benchmark.",
        formats,
    )
    headers = [
        "Month",
        "On-book",
        "Expected",
        "Check",
        "On-book UPB",
        "Expected",
        "Check",
        "Eligible",
        "Expected",
        "Check",
        "30+ count",
        "Expected",
        "Check",
        "60+ count",
        "Expected",
        "Check",
        "90+ count",
        "Expected",
        "Check",
        "30+ balance",
        "Expected",
        "Check",
        "60+ balance",
        "Expected",
        "Check",
        "90+ balance",
        "Expected",
        "Check",
    ]
    write_header_row(worksheet, 3, 0, headers, formats["header"])
    for index, sample in enumerate(data.validation_samples):
        row = 5 + index
        worksheet.write_datetime(row - 1, 0, excel_date(sample["as_of_month"]), formats["date"])
        worksheet.write_formula(
            row - 1,
            1,
            f"=SUMIFS('Monthly Data'!$B$5:$B${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$A{row})",
            formats["integer"],
        )
        worksheet.write_formula(row - 1, 2, f'={sample["on_book_loan_count"]}', formats["integer"])
        worksheet.write_formula(row - 1, 3, f"=IF(B{row}=C{row},0,1)", formats["integer"])
        worksheet.write_formula(
            row - 1,
            4,
            f"=SUMIFS('Monthly Data'!$C$5:$C${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$A{row})",
            formats["currency"],
        )
        worksheet.write_formula(row - 1, 5, f'={sample["on_book_upb"]}', formats["currency"])
        worksheet.write_formula(row - 1, 6, f"=IF(ABS(E{row}-F{row})<0.01,0,1)", formats["integer"])
        worksheet.write_formula(
            row - 1,
            7,
            f"=SUMIFS('Monthly Data'!$D$5:$D${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$A{row})",
            formats["integer"],
        )
        worksheet.write_formula(row - 1, 8, f'={sample["eligible_loan_count"]}', formats["integer"])
        worksheet.write_formula(row - 1, 9, f"=IF(H{row}=I{row},0,1)", formats["integer"])

        count_numerator_cols = ["N", "O", "P"]
        count_triples = [(10, 11, 12), (13, 14, 15), (16, 17, 18)]
        expected_count = [
            sample["dq30_count_rate"],
            sample["dq60_count_rate"],
            sample["dq90_count_rate"],
        ]
        balance_numerator_cols = ["Q", "R", "S"]
        balance_triples = [(19, 20, 21), (22, 23, 24), (25, 26, 27)]
        expected_balance = [
            sample["dq30_balance_rate"],
            sample["dq60_balance_rate"],
            sample["dq90_balance_rate"],
        ]
        for metric_index in range(3):
            start, expected_col, check_col = count_triples[metric_index]
            start_name = xl_col_to_name(start)
            expected_name = xl_col_to_name(expected_col)
            worksheet.write_formula(
                row - 1,
                start,
                f"=SUMIFS('Monthly Data'!${count_numerator_cols[metric_index]}$5:${count_numerator_cols[metric_index]}${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$A{row})/H{row}",
                formats["percent_4"],
            )
            worksheet.write_formula(
                row - 1,
                expected_col,
                f"={expected_count[metric_index]}",
                formats["percent_4"],
            )
            worksheet.write_formula(
                row - 1,
                check_col,
                f"=IF(ABS({start_name}{row}-{expected_name}{row})<0.0000000001,0,1)",
                formats["percent_4"],
            )
            b_start, b_expected, b_check = balance_triples[metric_index]
            b_start_name = xl_col_to_name(b_start)
            b_expected_name = xl_col_to_name(b_expected)
            worksheet.write_formula(
                row - 1,
                b_start,
                f"=SUMIFS('Monthly Data'!${balance_numerator_cols[metric_index]}$5:${balance_numerator_cols[metric_index]}${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$A{row})/SUMIFS('Monthly Data'!$E$5:$E${monthly_end},'Monthly Data'!$A$5:$A${monthly_end},$A{row})",
                formats["percent_4"],
            )
            worksheet.write_formula(
                row - 1,
                b_expected,
                f"={expected_balance[metric_index]}",
                formats["percent_4"],
            )
            worksheet.write_formula(
                row - 1,
                b_check,
                f"=IF(ABS({b_start_name}{row}-{b_expected_name}{row})<0.0000000001,0,1)",
                formats["percent_4"],
            )
    worksheet.write("A9", "Overall status", formats["validation_status"])
    worksheet.write_formula(
        "B9",
        '=IF(SUM(D5:D7,G5:G7,J5:J7,M5:M7,P5:P7,S5:S7,V5:V7,Y5:Y7,AB5:AB7)=0,"PASS","CHECK")',
        formats["validation_status"],
    )
    worksheet.set_column("A:D", 12)
    worksheet.set_column("E:F", 18)
    worksheet.set_column("G:AB", 12)


def build_workbook(output: Path, data: DashboardData) -> None:
    workbook = xlsxwriter.Workbook(output)
    workbook.set_properties(
        {
            "title": "Mortgage Portfolio Risk Monitoring with SQL and Excel",
            "subject": "Aggregate mortgage portfolio risk monitoring dashboard",
            "keywords": "mortgage risk, DuckDB, SQL, Python, Excel",
            "comments": "Built from validated aggregate marts without loan-level records.",
        }
    )
    workbook.set_calc_mode("auto")
    formats = create_formats(workbook)
    sheets = {name: workbook.add_worksheet(name) for name in SHEET_NAMES}
    for name, worksheet in sheets.items():
        tab_color = None
        if name in {"Portfolio Overview", "Risk Segments", "Transitions"}:
            tab_color = COLORS["navy"]
        elif name == "Validation":
            tab_color = COLORS["gray"]
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
    write_validation_sheet(
        sheets["Validation"], data, formats, endpoints["monthly_end"]
    )
    workbook.close()


def normalize_theme_fonts(path: Path) -> None:
    """Replace locale-specific default theme font names with Arial.

    XlsxWriter's bundled Office theme includes East Asian typeface names even
    when every workbook cell uses Arial. Normalizing those metadata-only names
    keeps the generated package language-neutral without changing workbook
    content, formulas, formatting, or calculations.
    """
    normalized = path.with_name(f"{path.stem}.theme-normalized.xlsx")
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
            normalized, "w", compression=zipfile.ZIP_DEFLATED
        ) as destination:
            for member in source.infolist():
                payload = source.read(member.filename)
                if member.filename == "xl/theme/theme1.xml":
                    theme = payload.decode("utf-8")
                    theme = theme.replace(
                        'typeface="\u5b8b\u4f53"', 'typeface="Arial"'
                    )
                    theme = theme.replace(
                        'typeface="\u65b0\u7d30\u660e\u9ad4"',
                        'typeface="Arial"',
                    )
                    payload = theme.encode("utf-8")
                destination.writestr(member, payload)
        os.replace(normalized, path)
    finally:
        if normalized.exists():
            normalized.unlink()


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def normalize_part_path(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def relationships(
    zf: zipfile.ZipFile, source: str
) -> dict[str, tuple[str, str]]:
    rel_name = posixpath.join(
        posixpath.dirname(source), "_rels", posixpath.basename(source) + ".rels"
    )
    if rel_name not in zf.namelist():
        return {}
    root = read_xml(zf, rel_name)
    return {
        relationship.attrib["Id"]: (
            relationship.attrib["Type"],
            normalize_part_path(source, relationship.attrib["Target"]),
        )
        for relationship in root.findall("pr:Relationship", XML_NS)
    }


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = read_xml(zf, "xl/sharedStrings.xml")
    return [
        "".join(text.text or "" for text in item.iter(f"{{{MAIN_NS}}}t"))
        for item in root.findall("m:si", XML_NS)
    ]


def parsed_cell_value(cell: ET.Element, strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(
            text.text or "" for text in cell.iter(f"{{{MAIN_NS}}}t")
        )
    value = cell.find("m:v", XML_NS)
    if value is None:
        return None
    raw = value.text or ""
    if cell_type == "s":
        return strings[int(raw)]
    if cell_type == "b":
        return raw == "1"
    if cell_type in {"str", "e"}:
        return raw
    try:
        number = float(raw)
        return int(number) if number.is_integer() else number
    except ValueError:
        return raw


def inspect_workbook_structure(path: Path) -> WorkbookSnapshot:
    with zipfile.ZipFile(path) as zf:
        part_names = set(zf.namelist())
        strings = shared_strings(zf)
        workbook = read_xml(zf, "xl/workbook.xml")
        workbook_rels = relationships(zf, "xl/workbook.xml")
        calc = workbook.find("m:calcPr", XML_NS)
        sheet_names: list[str] = []
        sheet_states: dict[str, str] = {}
        all_cells: dict[str, dict[str, Any]] = {}
        all_formulas: dict[str, dict[str, str]] = {}
        all_merges: dict[str, list[str]] = {}
        all_validations: dict[str, list[dict[str, str]]] = {}
        all_panes: dict[str, list[dict[str, str]]] = {}
        all_tables: dict[str, list[dict[str, str]]] = {}
        charts: list[dict[str, Any]] = []
        chart_paths_seen: set[str] = set()

        for sheet_node in workbook.findall("m:sheets/m:sheet", XML_NS):
            sheet_name = sheet_node.attrib["name"]
            sheet_names.append(sheet_name)
            sheet_states[sheet_name] = sheet_node.attrib.get("state", "visible")
            relation_id = sheet_node.attrib[f"{{{REL_NS}}}id"]
            sheet_path = workbook_rels[relation_id][1]
            sheet = read_xml(zf, sheet_path)
            sheet_rels = relationships(zf, sheet_path)
            cells: dict[str, Any] = {}
            formulas: dict[str, str] = {}
            for cell in sheet.findall(".//m:sheetData/m:row/m:c", XML_NS):
                address = cell.attrib["r"]
                cells[address] = parsed_cell_value(cell, strings)
                formula = cell.find("m:f", XML_NS)
                if formula is not None:
                    formulas[address] = formula.text or ""
            all_cells[sheet_name] = cells
            all_formulas[sheet_name] = formulas
            all_merges[sheet_name] = sorted(
                item.attrib["ref"]
                for item in sheet.findall("m:mergeCells/m:mergeCell", XML_NS)
            )
            validations = []
            for item in sheet.findall("m:dataValidations/m:dataValidation", XML_NS):
                validations.append(
                    {
                        "sqref": item.attrib.get("sqref", ""),
                        "type": item.attrib.get("type", ""),
                        "formula1": item.findtext(
                            "m:formula1", default="", namespaces=XML_NS
                        ).lstrip("="),
                    }
                )
            all_validations[sheet_name] = sorted(
                validations, key=lambda value: (value["sqref"], value["formula1"])
            )
            panes = []
            for item in sheet.findall("m:sheetViews/m:sheetView/m:pane", XML_NS):
                panes.append(
                    {
                        key: item.attrib[key]
                        for key in ["state", "xSplit", "ySplit", "topLeftCell"]
                        if key in item.attrib
                    }
                )
            all_panes[sheet_name] = panes
            tables = []
            for relation_type, target in sheet_rels.values():
                if relation_type.endswith("/table"):
                    table = read_xml(zf, target)
                    tables.append(
                        {
                            "name": table.attrib.get("name", ""),
                            "ref": table.attrib.get("ref", ""),
                        }
                    )
                if relation_type.endswith("/drawing"):
                    for drawing_type, chart_path in relationships(zf, target).values():
                        if not drawing_type.endswith("/chart") or chart_path in chart_paths_seen:
                            continue
                        chart_paths_seen.add(chart_path)
                        chart = read_xml(zf, chart_path)
                        title = "".join(
                            text.text or ""
                            for text in chart.findall(".//c:title//a:t", XML_NS)
                        )
                        series = []
                        for series_node in chart.findall(".//c:ser", XML_NS):
                            refs = [
                                reference.text or ""
                                for reference in series_node.findall(".//c:f", XML_NS)
                            ]
                            series.append(refs)
                        charts.append(
                            {"sheet": sheet_name, "title": title, "series": series}
                        )
            all_tables[sheet_name] = sorted(tables, key=lambda value: value["name"])

        defined_names = [
            {
                "name": item.attrib.get("name"),
                "local_sheet_id": item.attrib.get("localSheetId"),
                "value": item.text,
            }
            for item in workbook.findall("m:definedNames/m:definedName", XML_NS)
        ]
        xml_text = "\n".join(
            zf.read(name).decode("utf-8", errors="ignore")
            for name in part_names
            if name.endswith((".xml", ".rels"))
        )
        error_markers = sorted(
            {
                marker
                for marker in [
                    "#REF!",
                    "#DIV/0!",
                    "#VALUE!",
                    "#NAME?",
                    "#N/A",
                    "#NUM!",
                    "#NULL!",
                    "#SPILL!",
                    "#CALC!",
                ]
                if marker in xml_text
            }
        )
        return WorkbookSnapshot(
            sheet_names=sheet_names,
            sheet_states=sheet_states,
            cells=all_cells,
            formulas=all_formulas,
            merges=all_merges,
            validations=all_validations,
            panes=all_panes,
            tables=all_tables,
            charts=charts,
            defined_names=defined_names,
            external_links=sorted(
                name
                for name in part_names
                if "externalLink" in name
                or "connections" in name
                or "queryTable" in name
            ),
            vba_parts=sorted(
                name for name in part_names if name.lower().endswith("vbaproject.bin")
            ),
            power_query_parts=sorted(
                name
                for name in part_names
                if "customXml" in name
                or "dataMashup" in name
                or "powerquery" in name.lower()
            ),
            formula_error_tokens=error_markers,
            chinese_character_count=len(
                re.findall(r"[\u3400-\u9fff\uf900-\ufaff]", xml_text)
            ),
            calc_mode=calc.attrib.get("calcMode") if calc is not None else None,
            full_calc_on_load=(
                calc.attrib.get("fullCalcOnLoad") if calc is not None else None
            ),
        )


def parse_cell_address(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", address)
    if not match:
        raise ValueError(f"Unsupported cell address: {address}")
    letters, row_text = match.groups()
    column = 0
    for letter in letters:
        column = column * 26 + ord(letter) - ord("A") + 1
    return int(row_text), column


def iter_range(address_range: str):
    start, end = address_range.split(":")
    start_row, start_col = parse_cell_address(start)
    end_row, end_col = parse_cell_address(end)
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            yield f"{xl_col_to_name(col - 1)}{row}"


def values_equal(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(float(left), float(right), rel_tol=1e-15, abs_tol=1e-9)
    return left == right


def canonical_chart_reference(reference: str) -> str:
    """Normalize optional sheet-name quoting in chart formulas."""
    return re.sub(r"^'([^']+)'!", r"\1!", reference)


def comparable_charts(charts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "sheet": item["sheet"],
            "title": item["title"],
            "series": [
                [canonical_chart_reference(reference) for reference in series]
                for series in item["series"]
            ],
        }
        for item in charts
    ]


def compare_reference_and_candidate(
    reference: WorkbookSnapshot, candidate: WorkbookSnapshot
) -> dict[str, Any]:
    data_table_names = {
        "Monthly Data": "MonthlyAggregates",
        "Segment Data": "SegmentAggregates",
        "Transition Data": "TransitionAggregates",
    }
    data_mismatches = []
    data_cells_compared = 0
    for sheet_name, table_name in data_table_names.items():
        reference_table = next(
            table for table in reference.tables[sheet_name] if table["name"] == table_name
        )
        candidate_table = next(
            table for table in candidate.tables[sheet_name] if table["name"] == table_name
        )
        if reference_table["ref"] != candidate_table["ref"]:
            data_mismatches.append(
                {
                    "sheet": sheet_name,
                    "cell": "table_range",
                    "reference": reference_table["ref"],
                    "candidate": candidate_table["ref"],
                }
            )
            continue
        for address in iter_range(reference_table["ref"]):
            data_cells_compared += 1
            reference_value = reference.cells[sheet_name].get(address)
            candidate_value = candidate.cells[sheet_name].get(address)
            if not values_equal(reference_value, candidate_value):
                data_mismatches.append(
                    {
                        "sheet": sheet_name,
                        "cell": address,
                        "reference": reference_value,
                        "candidate": candidate_value,
                    }
                )
                if len(data_mismatches) >= 25:
                    break

    formula_mismatches = []
    formula_cells_compared = 0
    for sheet_name in SHEET_NAMES:
        all_addresses = sorted(
            set(reference.formulas[sheet_name]) | set(candidate.formulas[sheet_name])
        )
        for address in all_addresses:
            formula_cells_compared += 1
            reference_formula = reference.formulas[sheet_name].get(address)
            candidate_formula = candidate.formulas[sheet_name].get(address)
            if reference_formula != candidate_formula:
                formula_mismatches.append(
                    {
                        "sheet": sheet_name,
                        "cell": address,
                        "reference": reference_formula,
                        "candidate": candidate_formula,
                    }
                )
                if len(formula_mismatches) >= 25:
                    break

    structural_fields = {
        "sheet_names": reference.sheet_names == candidate.sheet_names,
        "sheet_states": reference.sheet_states == candidate.sheet_states,
        "merges": reference.merges == candidate.merges,
        "validations": reference.validations == candidate.validations,
        "panes": reference.panes == candidate.panes,
        "tables": reference.tables == candidate.tables,
        "defined_names": reference.defined_names == candidate.defined_names,
        "chart_titles_and_ranges": comparable_charts(reference.charts)
        == comparable_charts(candidate.charts),
    }
    return {
        "structural_fields": structural_fields,
        "structure_match": all(structural_fields.values()),
        "data_cells_compared": data_cells_compared,
        "data_mismatch_count": len(data_mismatches),
        "data_mismatch_examples": data_mismatches,
        "formula_cells_compared": formula_cells_compared,
        "formula_mismatch_count": len(formula_mismatches),
        "formula_mismatch_examples": formula_mismatches,
    }


def validate_workbook_structure(
    candidate_path: Path,
    data: DashboardData,
    reference_path: Path | None,
) -> dict[str, Any]:
    candidate = inspect_workbook_structure(candidate_path)
    expected_formula_counts = {
        "Portfolio Overview": 1138,
        "Risk Segments": 114,
        "Transitions": 309,
        "Validation": 82,
    }
    table_summary = {
        sheet_name: tables for sheet_name, tables in candidate.tables.items() if tables
    }
    validation_count = sum(
        len(items) for items in candidate.validations.values()
    )
    loan_identifier_headers = []
    for sheet_name in ["Monthly Data", "Segment Data", "Transition Data"]:
        for table in candidate.tables[sheet_name]:
            start, end = table["ref"].split(":")
            header_row, start_col = parse_cell_address(start)
            _, end_col = parse_cell_address(end)
            for col in range(start_col, end_col + 1):
                address = f"{xl_col_to_name(col - 1)}{header_row}"
                value = str(candidate.cells[sheet_name].get(address, ""))
                normalized = re.sub(r"[^a-z0-9]", "", value.lower())
                if "loanid" in normalized or "loanidentifier" in normalized:
                    loan_identifier_headers.append(
                        {"sheet": sheet_name, "cell": address, "value": value}
                    )
    external_formula_refs = [
        {"sheet": sheet_name, "cell": address, "formula": formula}
        for sheet_name, formulas in candidate.formulas.items()
        for address, formula in formulas.items()
        if "[" in formula or "]" in formula
    ]
    core_checks = {
        "sheet_names_and_order": candidate.sheet_names == SHEET_NAMES,
        "all_sheets_visible": all(
            state == "visible" for state in candidate.sheet_states.values()
        ),
        "chart_count": len(candidate.charts) == 9,
        "chart_series_count": sum(len(chart["series"]) for chart in candidate.charts)
        == 18,
        "table_count": sum(len(tables) for tables in candidate.tables.values()) == 4,
        "data_validation_count": validation_count == 5,
        "formula_counts": all(
            len(candidate.formulas[sheet_name]) == count
            for sheet_name, count in expected_formula_counts.items()
        ),
        "validation_month_count": len(data.validation_samples) == 3,
        "external_links": not candidate.external_links and not external_formula_refs,
        "macros": not candidate.vba_parts,
        "power_query": not candidate.power_query_parts,
        "loan_identifier_headers": not loan_identifier_headers,
        "formula_error_tokens": not candidate.formula_error_tokens,
        "chinese_characters": candidate.chinese_character_count == 0,
        "automatic_calculation": candidate.calc_mode in {None, "auto"},
        "full_calculation_on_load": candidate.full_calc_on_load in {None, "1"},
    }
    if not all(core_checks.values()):
        failures = [name for name, passed in core_checks.items() if not passed]
        raise RuntimeError(f"Candidate workbook structure checks failed: {failures}")

    comparison = None
    if reference_path and reference_path.is_file():
        reference = inspect_workbook_structure(reference_path)
        comparison = compare_reference_and_candidate(reference, candidate)
        if not comparison["structure_match"]:
            failures = [
                name
                for name, passed in comparison["structural_fields"].items()
                if not passed
            ]
            raise RuntimeError(f"Reference structure comparison failed: {failures}")
        if comparison["data_mismatch_count"]:
            raise RuntimeError(
                "Reference data comparison failed: "
                f"{comparison['data_mismatch_examples'][:3]}"
            )
        if comparison["formula_mismatch_count"]:
            raise RuntimeError(
                "Reference formula comparison failed: "
                f"{comparison['formula_mismatch_examples'][:3]}"
            )

    return {
        "core_checks": core_checks,
        "sheet_names": candidate.sheet_names,
        "chart_count": len(candidate.charts),
        "chart_series_count": sum(len(chart["series"]) for chart in candidate.charts),
        "chart_details": candidate.charts,
        "table_count": sum(len(tables) for tables in candidate.tables.values()),
        "tables": table_summary,
        "data_validation_count": validation_count,
        "data_validations": candidate.validations,
        "merges": candidate.merges,
        "panes": candidate.panes,
        "defined_names": candidate.defined_names,
        "formula_counts": {
            sheet_name: len(formulas)
            for sheet_name, formulas in candidate.formulas.items()
            if formulas
        },
        "validation_months": [
            row["as_of_month"][:7] for row in data.validation_samples
        ],
        "recalculation": {
            "calc_mode": candidate.calc_mode,
            "full_calc_on_load": candidate.full_calc_on_load,
            "validation_formula_cell": "Validation!B9",
            "validation_requires_excel_recalculation": True,
        },
        "external_formula_references": external_formula_refs,
        "loan_identifier_headers": loan_identifier_headers,
        "comparison_with_reference": comparison,
    }


def write_summary(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    database = args.database.resolve()
    output = args.output.resolve()
    summary = args.summary.resolve()
    reference = args.reference.resolve() if args.reference else None
    if not database.is_file():
        raise FileNotFoundError(f"DuckDB database not found: {database}")
    if reference and output == reference:
        raise ValueError("Candidate output must not overwrite the reference workbook")

    data = load_dashboard_data(database)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary_output = output.with_name(f"{output.stem}.tmp.xlsx")
    if temporary_output.exists():
        temporary_output.unlink()
    build_workbook(temporary_output, data)
    normalize_theme_fonts(temporary_output)
    validation = validate_workbook_structure(temporary_output, data, reference)
    os.replace(temporary_output, output)

    summary_payload = {
        "generated_by": "scripts/build_excel_dashboard.py",
        "xlsxwriter_version": xlsxwriter.__version__,
        "output_path": str(output.relative_to(PROJECT_ROOT)),
        "output_size_bytes": output.stat().st_size,
        "reference_path": (
            str(reference.relative_to(PROJECT_ROOT))
            if reference and reference.is_relative_to(PROJECT_ROOT)
            else None
        ),
        "embedded_rows": {
            "monthly": len(data.monthly),
            "segments": len(data.segments),
            "transitions": len(data.transitions),
        },
        "duckdb_validation_samples": data.validation_samples,
        "workbook_validation": validation,
        "manual_excel_validation_required": True,
    }
    write_summary(summary, summary_payload)
    print(f"Workbook: {output}")
    print(f"Summary: {summary}")
    print(
        "Structure: "
        f"{len(validation['sheet_names'])} sheets, "
        f"{validation['chart_count']} charts, "
        f"{validation['table_count']} tables, "
        f"{validation['data_validation_count']} validations"
    )
    print("Mac Excel recalculation and manual acceptance remain required")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
