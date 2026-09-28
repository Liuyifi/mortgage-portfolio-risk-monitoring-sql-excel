#!/usr/bin/env python3
"""Export local display aggregate files for dashboard generation.

The persisted DuckDB database is opened read-only. Only validated mart tables are
queried; no loan identifiers or loan-month detail are exported. Low-volume cells
are rolled up or suppressed according to a configurable minimum cell count.
The outputs are still data-bearing derived artifacts and are intended to stay local.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

import duckdb


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DATABASE = PROJECT_ROOT / "data" / "processed" / "mortgage_risk.duckdb"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "data" / "processed" / "dashboard"
DEFAULT_SUMMARY = (
    PROJECT_ROOT / "data" / "processed" / "validation" / "dashboard_export_summary.json"
)
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
    parser.add_argument("--summary", type=Path, default=DEFAULT_SUMMARY)
    parser.add_argument("--min-cell-count", type=int, default=MIN_CELL_COUNT)
    return parser.parse_args()


def fetch_one(connection: duckdb.DuckDBPyConnection, query: str, params=None):
    return connection.execute(query, params or []).fetchone()[0]


def fetch_dicts(connection: duckdb.DuckDBPyConnection, query: str) -> list[dict]:
    cursor = connection.execute(query)
    columns = [item[0] for item in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def copy_table(connection: duckdb.DuckDBPyConnection, table_name: str, path: Path) -> None:
    escaped = str(path.resolve()).replace("'", "''")
    connection.execute(
        f"COPY (SELECT * FROM {table_name}) TO '{escaped}' "
        "(FORMAT CSV, HEADER TRUE, DELIMITER ',', NULL '')"
    )


def run_checks(
    connection: duckdb.DuckDBPyConnection, min_cell_count: int
) -> tuple[list[dict], dict]:
    checks = [
        (
            "upstream_error_gates_passed",
            "SELECT count(*) FROM quality.dq_test_results "
            "WHERE severity = 'ERROR' AND NOT passed",
        ),
        (
            "portfolio_row_count_matches_mart",
            "SELECT abs((SELECT count(*) FROM dashboard_portfolio_monthly) - "
            "(SELECT count(*) FROM mart.portfolio_monthly))",
        ),
        (
            "portfolio_values_match_mart",
            """
            SELECT count(*) FROM (
                SELECT * FROM dashboard_portfolio_monthly
                EXCEPT SELECT * FROM mart.portfolio_monthly
            )
            """,
        ),
        (
            "segment_key_unique",
            """
            SELECT count(*) FROM (
                SELECT as_of_month, segment_name, segment_value, count(*) n
                FROM dashboard_portfolio_monthly_by_segment
                GROUP BY ALL HAVING n > 1
            )
            """,
        ),
        (
            "segment_minimum_cell_count",
            f"SELECT count(*) FROM dashboard_portfolio_monthly_by_segment "
            f"WHERE on_book_loan_count < {min_cell_count}",
        ),
        (
            "segment_month_dimension_complete",
            """
            SELECT count(*) FROM (
                SELECT m.as_of_month, d.segment_name
                FROM dashboard_portfolio_monthly m
                CROSS JOIN (VALUES ('origination_year'), ('classic_fico'),
                                   ('original_cltv'), ('original_dti')) d(segment_name)
                ANTI JOIN (
                    SELECT DISTINCT as_of_month, segment_name
                    FROM dashboard_portfolio_monthly_by_segment
                ) s USING (as_of_month, segment_name)
            )
            """,
        ),
        (
            "segment_additive_metrics_reconcile",
            """
            SELECT count(*) FROM (
                SELECT
                    s.as_of_month, s.segment_name
                FROM (
                    SELECT as_of_month, segment_name,
                           sum(on_book_loan_count) on_book_loan_count,
                           sum(on_book_upb) on_book_upb,
                           sum(delinquency_eligible_loan_count) delinquency_eligible_loan_count,
                           sum(delinquency_eligible_upb) delinquency_eligible_upb,
                           sum(dq30_loan_count) dq30_loan_count,
                           sum(dq60_loan_count) dq60_loan_count,
                           sum(dq90_loan_count) dq90_loan_count,
                           sum(dq30_upb) dq30_upb,
                           sum(dq60_upb) dq60_upb,
                           sum(dq90_upb) dq90_upb
                    FROM dashboard_portfolio_monthly_by_segment GROUP BY 1, 2
                ) s
                JOIN dashboard_portfolio_monthly p USING (as_of_month)
                WHERE s.on_book_loan_count <> p.on_book_loan_count
                   OR s.on_book_upb <> p.on_book_upb
                   OR s.delinquency_eligible_loan_count <> p.delinquency_eligible_loan_count
                   OR s.delinquency_eligible_upb <> p.delinquency_eligible_upb
                   OR s.dq30_loan_count <> p.dq30_loan_count
                   OR s.dq60_loan_count <> p.dq60_loan_count
                   OR s.dq90_loan_count <> p.dq90_loan_count
                   OR s.dq30_upb <> p.dq30_upb
                   OR s.dq60_upb <> p.dq60_upb
                   OR s.dq90_upb <> p.dq90_upb
            )
            """,
        ),
        (
            "segment_rate_null_semantics",
            """
            SELECT count(*) FROM dashboard_portfolio_monthly_by_segment
            WHERE (delinquency_eligible_loan_count > 0 AND
                   (dq30_count_rate IS NULL OR dq60_count_rate IS NULL OR dq90_count_rate IS NULL))
               OR (delinquency_eligible_loan_count = 0 AND
                   (dq30_count_rate IS NOT NULL OR dq60_count_rate IS NOT NULL OR dq90_count_rate IS NOT NULL))
               OR (delinquency_eligible_upb > 0 AND
                   (dq30_balance_rate IS NULL OR dq60_balance_rate IS NULL OR dq90_balance_rate IS NULL))
               OR (delinquency_eligible_upb = 0 AND
                   (dq30_balance_rate IS NOT NULL OR dq60_balance_rate IS NOT NULL OR dq90_balance_rate IS NOT NULL))
            """,
        ),
        (
            "transition_key_unique",
            """
            SELECT count(*) FROM (
                SELECT from_month, from_state, to_state, count(*) n
                FROM dashboard_delinquency_transition_monthly
                GROUP BY ALL HAVING n > 1
            )
            """,
        ),
        (
            "transition_exact_next_month",
            """
            SELECT count(*) FROM dashboard_delinquency_transition_monthly
            WHERE to_month <> from_month + INTERVAL '1 month'
            """,
        ),
        (
            "transition_minimum_cell_count",
            f"SELECT count(*) FROM dashboard_delinquency_transition_monthly "
            f"WHERE transition_loan_count < {min_cell_count}",
        ),
        (
            "transition_retained_cohorts_reconcile",
            f"""
            WITH source AS (
                SELECT from_month, from_state,
                       sum(transition_loan_count) transition_loan_count,
                       sum(transition_from_upb) transition_from_upb
                FROM mart.delinquency_transition_monthly
                WHERE segment_name = 'all'
                GROUP BY 1, 2
                HAVING sum(transition_loan_count) >= {min_cell_count}
            ), exported AS (
                SELECT from_month, from_state,
                       sum(transition_loan_count) transition_loan_count,
                       sum(transition_from_upb) transition_from_upb
                FROM dashboard_delinquency_transition_monthly
                GROUP BY 1, 2
            )
            SELECT count(*) FROM source s
            FULL OUTER JOIN exported e USING (from_month, from_state)
            WHERE s.transition_loan_count IS DISTINCT FROM e.transition_loan_count
               OR s.transition_from_upb IS DISTINCT FROM e.transition_from_upb
            """,
        ),
        (
            "transition_rates_reconcile",
            """
            SELECT count(*) FROM dashboard_delinquency_transition_monthly
            WHERE abs(transition_count_rate -
                      transition_loan_count::DOUBLE / observable_from_state_loan_count) > 1e-12
               OR abs(transition_balance_rate -
                      transition_from_upb::DOUBLE / observable_from_state_upb) > 1e-12
            """,
        ),
        (
            "no_loan_identifier_columns",
            """
            SELECT count(*) FROM information_schema.columns
            WHERE table_name IN ('dashboard_portfolio_monthly',
                                 'dashboard_portfolio_monthly_by_segment',
                                 'dashboard_delinquency_transition_monthly')
              AND lower(column_name) LIKE '%loan_id%'
            """,
        ),
    ]

    results = []
    for name, query in checks:
        failed_rows = int(fetch_one(connection, query))
        results.append(
            {"check_name": name, "passed": failed_rows == 0, "failed_row_count": failed_rows}
        )

    suppression = {
        "min_cell_count": min_cell_count,
        "segment_source_rows_rolled_up": int(
            fetch_one(
                connection,
                "SELECT coalesce(sum(source_mart_row_count), 0) "
                "FROM dashboard_portfolio_monthly_by_segment WHERE privacy_merged",
            )
        ),
        "segment_export_rows_marked_merged": int(
            fetch_one(
                connection,
                "SELECT count(*) FROM dashboard_portfolio_monthly_by_segment WHERE privacy_merged",
            )
        ),
        "transition_source_rows_rolled_up": int(
            fetch_one(
                connection,
                "SELECT coalesce(sum(source_mart_row_count), 0) "
                "FROM dashboard_delinquency_transition_monthly WHERE privacy_merged",
            )
        ),
        "transition_export_rows_marked_merged": int(
            fetch_one(
                connection,
                "SELECT count(*) FROM dashboard_delinquency_transition_monthly WHERE privacy_merged",
            )
        ),
        "transition_suppressed_from_state_cohorts": int(
            fetch_one(
                connection,
                f"""
                SELECT count(*) FROM (
                    SELECT from_month, from_state
                    FROM mart.delinquency_transition_monthly
                    WHERE segment_name = 'all'
                    GROUP BY 1, 2
                    HAVING sum(transition_loan_count) < {min_cell_count}
                )
                """,
            )
        ),
        "transition_suppressed_loan_observations": int(
            fetch_one(
                connection,
                f"""
                SELECT coalesce(sum(cohort_count), 0) FROM (
                    SELECT sum(transition_loan_count) cohort_count
                    FROM mart.delinquency_transition_monthly
                    WHERE segment_name = 'all'
                    GROUP BY from_month, from_state
                    HAVING cohort_count < {min_cell_count}
                )
                """,
            )
        ),
    }
    return results, suppression


def json_default(value: object) -> str:
    if hasattr(value, "isoformat"):
        return value.isoformat()  # type: ignore[union-attr]
    return str(value)


def main() -> int:
    args = parse_args()
    if args.min_cell_count < 2:
        raise ValueError("--min-cell-count must be at least 2")

    database_path = args.database.resolve()
    output_dir = args.output_dir.resolve()
    summary_path = args.summary.resolve()
    if not database_path.is_file():
        raise FileNotFoundError(f"DuckDB database not found: {database_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    summary_path.parent.mkdir(parents=True, exist_ok=True)

    connection = duckdb.connect(str(database_path), read_only=True)
    try:
        export_sql = EXPORT_SQL.read_text(encoding="utf-8").replace(
            "{{MIN_CELL_COUNT}}", str(args.min_cell_count)
        )
        connection.execute(export_sql)
        checks, suppression = run_checks(connection, args.min_cell_count)
        failures = [item for item in checks if not item["passed"]]
        if failures:
            print(json.dumps(failures, indent=2), file=sys.stderr)
            raise RuntimeError(f"{len(failures)} dashboard export checks failed")

        files = []
        for filename, table_name in EXPORTS.items():
            path = output_dir / filename
            copy_table(connection, table_name, path)
            files.append(
                {
                    "path": str(path.relative_to(PROJECT_ROOT)),
                    "source_temp_table": table_name,
                    "row_count": int(fetch_one(connection, f"SELECT count(*) FROM {table_name}")),
                    "size_bytes": path.stat().st_size,
                    "sha256": sha256(path),
                }
            )

        summary = {
            "generated_by": "scripts/export_dashboard_data.py",
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "duckdb_version": duckdb.__version__,
            "database_path": str(database_path.relative_to(PROJECT_ROOT)),
            "source_tables": [
                "mart.portfolio_monthly",
                "mart.portfolio_monthly_by_segment",
                "mart.delinquency_transition_monthly",
            ],
            "privacy_policy": suppression,
            "date_ranges": fetch_dicts(
                connection,
                """
                SELECT 'portfolio' dataset, min(as_of_month) min_month, max(as_of_month) max_month
                FROM dashboard_portfolio_monthly
                UNION ALL
                SELECT 'transitions', min(from_month), max(to_month)
                FROM dashboard_delinquency_transition_monthly
                """,
            ),
            "files": files,
            "checks": checks,
            "sample_metrics": fetch_dicts(
                connection,
                """
                SELECT as_of_month, on_book_loan_count, on_book_upb,
                       delinquency_eligible_loan_count, delinquency_eligible_upb,
                       dq30_count_rate, dq60_count_rate, dq90_count_rate,
                       dq30_balance_rate, dq60_balance_rate, dq90_balance_rate
                FROM dashboard_portfolio_monthly
                WHERE as_of_month IN (DATE '2020-06-01', DATE '2021-06-01', DATE '2026-03-01')
                ORDER BY as_of_month
                """,
            ),
        }
        summary_path.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2, default=json_default) + "\n",
            encoding="utf-8",
        )

        print(f"Exported {len(files)} files to {output_dir}")
        for item in files:
            print(f"- {item['path']}: {item['row_count']:,} rows, {item['size_bytes']:,} bytes")
        print(f"Checks: {len(checks)} passed, 0 failed")
        print(f"Summary: {summary_path}")
    finally:
        connection.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
