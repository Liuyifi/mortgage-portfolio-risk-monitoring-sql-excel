#!/usr/bin/env python3
"""Build the local DuckDB analysis tables and run SQL validation checks."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import duckdb


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SQL_FILES = (
    "00_init.sql",
    "10_staging.sql",
    "20_curated.sql",
    "30_marts.sql",
    "90_validation_checks.sql",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--database",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "mortgage_risk.duckdb",
    )
    parser.add_argument(
        "--raw-dir",
        type=Path,
        default=PROJECT_ROOT / "data" / "raw",
    )
    return parser.parse_args()


def check_source_files(raw_dir: Path) -> None:
    expected = [
        raw_dir / f"sample_{year}" / f"sample_{file_type}_{year}.txt"
        for year in (2019, 2020)
        for file_type in ("orig", "perf")
    ]
    missing = [str(path) for path in expected if not path.is_file()]
    if missing:
        raise FileNotFoundError("Missing source files:\n- " + "\n- ".join(missing))


def run_sql_steps(connection: duckdb.DuckDBPyConnection, raw_dir: Path) -> None:
    raw_path = str(raw_dir.resolve()).replace("'", "''")
    for filename in SQL_FILES:
        sql = (PROJECT_ROOT / "sql" / filename).read_text(encoding="utf-8")
        connection.execute(sql.replace("{{RAW_DIR}}", raw_path))
        print(f"Completed {filename}", flush=True)


def print_results(connection: duckdb.DuckDBPyConnection) -> None:
    counts = connection.execute(
        """
        SELECT 'staging.origination', count(*) FROM staging.origination
        UNION ALL SELECT 'staging.monthly_performance', count(*) FROM staging.monthly_performance
        UNION ALL SELECT 'curated.fact_loan_month', count(*) FROM curated.fact_loan_month
        UNION ALL SELECT 'mart.portfolio_monthly', count(*) FROM mart.portfolio_monthly
        UNION ALL SELECT 'mart.portfolio_monthly_by_segment', count(*)
            FROM mart.portfolio_monthly_by_segment
        UNION ALL SELECT 'mart.delinquency_transition_monthly', count(*)
            FROM mart.delinquency_transition_monthly
        ORDER BY 1
        """
    ).fetchall()
    print("\nTable row counts")
    for table_name, row_count in counts:
        print(f"- {table_name}: {row_count:,}")

    summary = connection.execute(
        """
        SELECT count(*) AS checks, count(*) FILTER (WHERE passed) AS passed
        FROM quality.validation_results
        """
    ).fetchone()
    print("\nValidation summary")
    print(f"- {summary[1]}/{summary[0]} checks passed")


def main() -> int:
    args = parse_args()
    database = args.database.resolve()
    raw_dir = args.raw_dir.resolve()
    check_source_files(raw_dir)
    database.parent.mkdir(parents=True, exist_ok=True)

    connection = duckdb.connect(str(database))
    connection.execute("SET preserve_insertion_order = false")
    connection.execute("SET threads = 4")
    try:
        connection.execute("BEGIN TRANSACTION")
        run_sql_steps(connection, raw_dir)
        failures = connection.execute(
            """
            SELECT check_name, failed_row_count, details
            FROM quality.validation_results
            WHERE NOT passed
            ORDER BY check_name
            """
        ).fetchall()
        if failures:
            messages = [f"{name}: {count} failed rows ({details})"
                        for name, count, details in failures]
            raise RuntimeError("SQL validation failed:\n- " + "\n- ".join(messages))
        connection.execute("COMMIT")
        connection.execute("CHECKPOINT")
        print_results(connection)
        print(f"\nDuckDB build completed: {database}")
        return 0
    except Exception:
        try:
            connection.execute("ROLLBACK")
        except duckdb.Error:
            pass
        raise
    finally:
        connection.close()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Build failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
