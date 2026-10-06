#!/usr/bin/env python3
"""Export aggregate DuckDB tables used by the Excel dashboard."""

from __future__ import annotations

import argparse
from pathlib import Path

import duckdb


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DATABASE = PROJECT_ROOT / "data" / "processed" / "mortgage_risk.duckdb"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "data" / "processed" / "dashboard"
EXPORT_SQL = PROJECT_ROOT / "sql" / "40_dashboard_exports.sql"
MIN_CELL_COUNT = 20
EXPORTS = {
    "portfolio_monthly.csv": "dashboard_portfolio_monthly",
    "portfolio_monthly_by_segment.csv": "dashboard_portfolio_monthly_by_segment",
    "delinquency_transition_monthly.csv": "dashboard_delinquency_transition_monthly",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", type=Path, default=DEFAULT_DATABASE)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--min-cell-count", type=int, default=MIN_CELL_COUNT)
    return parser.parse_args()


def copy_table(connection: duckdb.DuckDBPyConnection, table_name: str, path: Path) -> None:
    output_path = str(path.resolve()).replace("'", "''")
    connection.execute(
        f"COPY (SELECT * FROM {table_name}) TO '{output_path}' "
        "(FORMAT CSV, HEADER TRUE, DELIMITER ',', NULL '')"
    )


def check_exports(connection: duckdb.DuckDBPyConnection, min_cell_count: int) -> None:
    checks = {
        "portfolio row count": """
            SELECT abs(
                (SELECT count(*) FROM dashboard_portfolio_monthly)
                - (SELECT count(*) FROM mart.portfolio_monthly)
            )
        """,
        "segment key uniqueness": """
            SELECT count(*) FROM (
                SELECT as_of_month, segment_name, segment_value, count(*) AS rows
                FROM dashboard_portfolio_monthly_by_segment
                GROUP BY ALL HAVING rows > 1
            )
        """,
        "segment totals": """
            SELECT count(*) FROM (
                SELECT s.as_of_month, s.segment_name
                FROM (
                    SELECT as_of_month, segment_name,
                           sum(on_book_loan_count) AS on_book_loan_count,
                           sum(on_book_upb) AS on_book_upb,
                           sum(delinquency_eligible_loan_count) AS eligible_count,
                           sum(delinquency_eligible_upb) AS eligible_upb,
                           sum(dq30_loan_count) AS dq30_count,
                           sum(dq60_loan_count) AS dq60_count,
                           sum(dq90_loan_count) AS dq90_count,
                           sum(dq30_upb) AS dq30_upb,
                           sum(dq60_upb) AS dq60_upb,
                           sum(dq90_upb) AS dq90_upb
                    FROM dashboard_portfolio_monthly_by_segment
                    GROUP BY as_of_month, segment_name
                ) s
                JOIN dashboard_portfolio_monthly p USING (as_of_month)
                WHERE s.on_book_loan_count <> p.on_book_loan_count
                   OR s.on_book_upb <> p.on_book_upb
                   OR s.eligible_count <> p.delinquency_eligible_loan_count
                   OR s.eligible_upb <> p.delinquency_eligible_upb
                   OR s.dq30_count <> p.dq30_loan_count
                   OR s.dq60_count <> p.dq60_loan_count
                   OR s.dq90_count <> p.dq90_loan_count
                   OR s.dq30_upb <> p.dq30_upb
                   OR s.dq60_upb <> p.dq60_upb
                   OR s.dq90_upb <> p.dq90_upb
            )
        """,
        "transition key uniqueness": """
            SELECT count(*) FROM (
                SELECT from_month, from_state, to_state, count(*) AS rows
                FROM dashboard_delinquency_transition_monthly
                GROUP BY ALL HAVING rows > 1
            )
        """,
        "transition cohort totals": f"""
            WITH expected AS (
                SELECT from_month, from_state,
                       sum(transition_loan_count) AS loan_count,
                       sum(transition_from_upb) AS upb
                FROM mart.delinquency_transition_monthly
                WHERE segment_name = 'all'
                GROUP BY from_month, from_state
                HAVING loan_count >= {min_cell_count}
            ), exported AS (
                SELECT from_month, from_state,
                       sum(transition_loan_count) AS loan_count,
                       sum(transition_from_upb) AS upb
                FROM dashboard_delinquency_transition_monthly
                GROUP BY from_month, from_state
            )
            SELECT count(*) FROM expected e
            FULL OUTER JOIN exported x USING (from_month, from_state)
            WHERE e.loan_count IS DISTINCT FROM x.loan_count
               OR e.upb IS DISTINCT FROM x.upb
        """,
        "loan identifier columns": """
            SELECT count(*) FROM information_schema.columns
            WHERE table_name IN (
                'dashboard_portfolio_monthly',
                'dashboard_portfolio_monthly_by_segment',
                'dashboard_delinquency_transition_monthly'
            ) AND lower(column_name) LIKE '%loan_id%'
        """,
    }
    failures = []
    for name, query in checks.items():
        failed_rows = connection.execute(query).fetchone()[0]
        if failed_rows:
            failures.append(f"{name}: {failed_rows} failed rows")
    if failures:
        raise RuntimeError("Dashboard export checks failed:\n- " + "\n- ".join(failures))


def main() -> int:
    args = parse_args()
    database = args.database.resolve()
    output_dir = args.output_dir.resolve()
    if args.min_cell_count < 2:
        raise ValueError("--min-cell-count must be at least 2")
    if not database.is_file():
        raise FileNotFoundError(f"DuckDB database not found: {database}")
    output_dir.mkdir(parents=True, exist_ok=True)

    with duckdb.connect(str(database), read_only=True) as connection:
        sql = EXPORT_SQL.read_text(encoding="utf-8").replace(
            "{{MIN_CELL_COUNT}}", str(args.min_cell_count)
        )
        connection.execute(sql)
        check_exports(connection, args.min_cell_count)
        for filename, table_name in EXPORTS.items():
            path = output_dir / filename
            copy_table(connection, table_name, path)
            row_count = connection.execute(f"SELECT count(*) FROM {table_name}").fetchone()[0]
            print(f"Exported {filename}: {row_count:,} rows")

    print(f"Dashboard exports completed: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
