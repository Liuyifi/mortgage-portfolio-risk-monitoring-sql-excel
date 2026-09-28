#!/usr/bin/env python3
"""Rebuild and validate the Freddie Mac sample DuckDB data layer.

The script verifies the raw files against the Phase 1 manifest, executes the
versioned SQL pipeline in one transaction, and exits nonzero if an ERROR-level
SQL quality test fails. Raw source files are opened read-only and never changed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import duckdb


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SQL_FILES = (
    "00_init.sql",
    "10_staging.sql",
    "20_curated.sql",
    "30_marts.sql",
    "90_quality_checks.sql",
)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_raw_manifest(raw_dir: Path, profile_path: Path) -> list[dict]:
    """Validate exact files, sizes, and hashes captured during Phase 1."""
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    verified: list[dict] = []
    mismatches: list[str] = []

    for year, year_profile in sorted(profile["years"].items()):
        for dataset_name, file_profile in year_profile.items():
            if dataset_name not in {"origination", "performance"}:
                continue
            prefix = "orig" if dataset_name == "origination" else "perf"
            path = raw_dir / f"sample_{year}" / f"sample_{prefix}_{year}.txt"
            if not path.is_file():
                mismatches.append(f"missing: {path}")
                continue
            observed_size = path.stat().st_size
            observed_hash = file_sha256(path)
            if observed_size != file_profile["size_bytes"]:
                mismatches.append(
                    f"size: {path} expected={file_profile['size_bytes']} observed={observed_size}"
                )
            if observed_hash != file_profile["sha256"]:
                mismatches.append(
                    f"sha256: {path} expected={file_profile['sha256']} observed={observed_hash}"
                )
            verified.append(
                {
                    "path": str(path.relative_to(PROJECT_ROOT)),
                    "size_bytes": observed_size,
                    "sha256": observed_hash,
                }
            )

    if mismatches:
        raise RuntimeError("Raw manifest verification failed:\n- " + "\n- ".join(mismatches))
    return verified


def execute_pipeline(connection: duckdb.DuckDBPyConnection, raw_dir: Path) -> dict[str, float]:
    timings: dict[str, float] = {}
    raw_sql_path = str(raw_dir.resolve()).replace("'", "''")
    for filename in SQL_FILES:
        sql_path = PROJECT_ROOT / "sql" / filename
        sql = sql_path.read_text(encoding="utf-8").replace("{{RAW_DIR}}", raw_sql_path)
        started = time.perf_counter()
        connection.execute(sql)
        timings[filename] = round(time.perf_counter() - started, 3)
        print(f"[sql] {filename}: {timings[filename]:.3f}s", flush=True)
    return timings


def fetch_dicts(connection: duckdb.DuckDBPyConnection, query: str) -> list[dict]:
    cursor = connection.execute(query)
    columns = [item[0] for item in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def json_default(value: object) -> str:
    if hasattr(value, "isoformat"):
        return value.isoformat()  # type: ignore[union-attr]
    return str(value)


def build_summary(
    connection: duckdb.DuckDBPyConnection,
    database_path: Path,
    verified_files: list[dict],
    timings: dict[str, float],
) -> dict:
    table_counts = fetch_dicts(
        connection,
        """
        SELECT 'staging.origination' AS table_name, count(*) AS row_count FROM staging.origination
        UNION ALL SELECT 'staging.monthly_performance', count(*) FROM staging.monthly_performance
        UNION ALL SELECT 'curated.dim_loan', count(*) FROM curated.dim_loan
        UNION ALL SELECT 'curated.fact_loan_month', count(*) FROM curated.fact_loan_month
        UNION ALL SELECT 'mart.portfolio_monthly', count(*) FROM mart.portfolio_monthly
        UNION ALL SELECT 'mart.portfolio_monthly_by_segment', count(*) FROM mart.portfolio_monthly_by_segment
        UNION ALL SELECT 'mart.delinquency_transition_monthly', count(*) FROM mart.delinquency_transition_monthly
        UNION ALL SELECT 'quality.loan_age_audit_summary', count(*) FROM quality.loan_age_audit_summary
        UNION ALL SELECT 'quality.loan_age_zero_monthly', count(*) FROM quality.loan_age_zero_monthly
        UNION ALL SELECT 'quality.metric_reconciliation_samples', count(*) FROM quality.metric_reconciliation_samples
        UNION ALL SELECT 'quality.dq_test_results', count(*) FROM quality.dq_test_results
        ORDER BY table_name
        """,
    )
    ranges = fetch_dicts(
        connection,
        """
        SELECT source_year AS origination_year, count(*) AS row_count,
               min(as_of_month) AS min_month, max(as_of_month) AS max_month,
               count(DISTINCT loan_id) AS distinct_loans
        FROM staging.monthly_performance
        GROUP BY source_year ORDER BY source_year
        """,
    )
    quality = fetch_dicts(
        connection,
        """
        SELECT severity, count(*) AS test_count,
               count(*) FILTER (WHERE NOT passed) AS non_passing_count,
               sum(failed_row_count) FILTER (WHERE NOT passed) AS affected_rows
        FROM quality.dq_test_results
        GROUP BY severity ORDER BY severity
        """,
    )
    warnings = fetch_dicts(
        connection,
        """
        SELECT test_name, failed_row_count, details
        FROM quality.dq_test_results
        WHERE severity = 'WARNING' AND NOT passed
        ORDER BY test_name
        """,
    )
    loan_age_audit = fetch_dicts(
        connection,
        """
        SELECT * FROM quality.loan_age_audit_summary
        ORDER BY CASE age_class
            WHEN 'NULL_OR_BLANK' THEN 1
            WHEN 'UNPARSEABLE' THEN 2
            WHEN 'NEGATIVE' THEN 3
            WHEN 'ZERO' THEN 4
        END
        """,
    )
    loan_age_zero_by_year_status = fetch_dicts(
        connection,
        """
        SELECT source_year, delinquency_status_raw,
               sum(record_count) AS record_count,
               coalesce(sum(record_count) FILTER (WHERE is_terminal_record), 0) AS terminal_count,
               coalesce(sum(record_count) FILTER (WHERE modification_status = 'Y'), 0) AS current_modification_count,
               min(as_of_month) AS min_month, max(as_of_month) AS max_month
        FROM quality.loan_age_zero_monthly
        GROUP BY source_year, delinquency_status_raw
        ORDER BY source_year, delinquency_status_raw
        """,
    )
    loan_age_zero_top_months = fetch_dicts(
        connection,
        """
        SELECT as_of_month, sum(record_count) AS record_count,
               coalesce(sum(record_count) FILTER (WHERE delinquency_status_raw = '01'), 0) AS status_01_count,
               coalesce(sum(record_count) FILTER (WHERE modification_status = 'Y'), 0) AS current_modification_count,
               coalesce(sum(record_count) FILTER (WHERE is_terminal_record), 0) AS terminal_count
        FROM quality.loan_age_zero_monthly
        GROUP BY as_of_month
        ORDER BY record_count DESC, as_of_month
        LIMIT 15
        """,
    )
    independent_reconciliation = fetch_dicts(
        connection,
        """
        SELECT * FROM quality.metric_reconciliation_samples
        ORDER BY as_of_month
        """,
    )
    sample_metrics = fetch_dicts(
        connection,
        """
        SELECT as_of_month, on_book_loan_count, on_book_upb,
               delinquency_eligible_loan_count,
               dq30_count_rate, dq60_count_rate, dq90_count_rate,
               dq30_balance_rate, dq60_balance_rate, dq90_balance_rate
        FROM mart.portfolio_monthly
        WHERE as_of_month IN (DATE '2020-06-01', DATE '2021-06-01', DATE '2026-03-01')
        ORDER BY as_of_month
        """,
    )
    transition_example = fetch_dicts(
        connection,
        """
        SELECT from_month, to_month, from_state, to_state,
               transition_loan_count, transition_from_upb,
               transition_count_rate, transition_balance_rate
        FROM mart.delinquency_transition_monthly
        WHERE segment_name = 'all' AND from_month = DATE '2020-06-01'
        ORDER BY transition_loan_count DESC, from_state, to_state
        LIMIT 8
        """,
    )
    return {
        "generated_by": "scripts/build_data_layer.py",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "duckdb_version": duckdb.__version__,
        "database_path": str(database_path.relative_to(PROJECT_ROOT)),
        "database_size_bytes": database_path.stat().st_size if database_path.exists() else None,
        "raw_manifest_verified": verified_files,
        "sql_stage_seconds": timings,
        "table_counts": table_counts,
        "performance_ranges": ranges,
        "quality_summary": quality,
        "quality_warnings": warnings,
        "loan_age_audit": loan_age_audit,
        "loan_age_zero_by_year_status": loan_age_zero_by_year_status,
        "loan_age_zero_top_months": loan_age_zero_top_months,
        "independent_metric_reconciliation": independent_reconciliation,
        "sample_portfolio_metrics": sample_metrics,
        "sample_transition_flows": transition_example,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--database",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "mortgage_risk.duckdb",
        help="Generated DuckDB database path.",
    )
    parser.add_argument(
        "--raw-dir",
        type=Path,
        default=PROJECT_ROOT / "data" / "raw",
        help="Read-only directory containing sample_2019/ and sample_2020/.",
    )
    parser.add_argument(
        "--profile",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "validation" / "raw_data_profile.json",
        help="Phase 1 raw-file manifest/profile.",
    )
    parser.add_argument(
        "--summary",
        type=Path,
        default=PROJECT_ROOT / "data" / "processed" / "validation" / "data_layer_build_summary.json",
        help="Compact generated build evidence.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    database_path = args.database.resolve()
    raw_dir = args.raw_dir.resolve()
    profile_path = args.profile.resolve()
    summary_path = args.summary.resolve()

    print("[preflight] validating Phase 1 raw-file manifest", flush=True)
    verified_files = validate_raw_manifest(raw_dir, profile_path)
    database_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.parent.mkdir(parents=True, exist_ok=True)

    connection = duckdb.connect(str(database_path))
    connection.execute("SET preserve_insertion_order = false")
    connection.execute("SET threads = 4")
    try:
        connection.execute("BEGIN TRANSACTION")
        timings = execute_pipeline(connection, raw_dir)
        failures = fetch_dicts(
            connection,
            """
            SELECT test_name, failed_row_count, details
            FROM quality.dq_test_results
            WHERE severity = 'ERROR' AND NOT passed
            ORDER BY test_name
            """,
        )
        if failures:
            lines = [
                f"{row['test_name']}: failed_rows={row['failed_row_count']} ({row['details']})"
                for row in failures
            ]
            raise RuntimeError("SQL quality gate failed:\n- " + "\n- ".join(lines))
        connection.execute("COMMIT")
        connection.execute("CHECKPOINT")

        summary = build_summary(connection, database_path, verified_files, timings)
        summary_path.write_text(
            json.dumps(summary, indent=2, ensure_ascii=False, default=json_default) + "\n",
            encoding="utf-8",
        )
        warning_count = sum(item["non_passing_count"] for item in summary["quality_summary"] if item["severity"] == "WARNING")
        print(
            f"[ok] all ERROR quality gates passed; warnings={warning_count}; "
            f"summary={summary_path.relative_to(PROJECT_ROOT)}",
            flush=True,
        )
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
        sys.exit(main())
    except Exception as exc:
        print(f"[failed] {exc}", file=sys.stderr)
        sys.exit(1)
