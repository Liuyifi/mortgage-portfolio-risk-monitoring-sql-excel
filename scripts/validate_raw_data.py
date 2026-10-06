#!/usr/bin/env python3
"""Run practical checks on the four local Freddie Mac sample files."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path


ORIG_COLUMNS = [
    "classic_fico", "first_payment_date", "first_time_homebuyer_indicator",
    "maturity_date", "msa_or_metropolitan_division", "mi_percentage",
    "number_of_units", "occupancy_status", "original_cltv", "original_dti",
    "original_upb", "original_ltv", "original_interest_rate", "channel",
    "prepayment_penalty_indicator", "amortization_type", "property_state",
    "property_type", "postal_code", "loan_id", "loan_purpose",
    "original_loan_term", "number_of_borrowers", "seller_name",
    "super_conforming_flag", "pre_harp_loan_id", "special_eligibility_program",
    "harp_indicator", "property_valuation_method", "interest_only_indicator",
    "vantage_score_4",
]

PERF_COLUMNS = [
    "loan_id", "period", "current_actual_upb", "current_loan_delinquency_status",
    "loan_age", "remaining_months_to_legal_maturity", "defect_settlement_date",
    "modification_flag", "zero_balance_code", "zero_balance_effective_date",
    "current_interest_rate", "current_non_interest_bearing_upb", "ddlpi",
    "mi_recoveries", "net_sales_proceeds", "non_mi_recoveries", "total_expenses",
    "legal_costs", "maintenance_and_preservation_costs", "taxes_and_insurance",
    "miscellaneous_expenses", "actual_loss", "cumulative_modification_costs",
    "interest_rate_step_indicator", "payment_deferral_flag", "estimated_ltv",
    "zero_balance_removal_upb", "delinquent_accrued_interest",
    "delinquency_due_to_disaster", "borrower_assistance_plan",
    "current_period_modification_costs", "current_interest_bearing_upb",
    "mi_cancellation_indicator", "servicer_name", "bankruptcy_cramdown_costs",
]

EXPECTED_ROWS = {
    "sample_orig_2019.txt": 50_000,
    "sample_perf_2019.txt": 1_934_614,
    "sample_orig_2020.txt": 50_000,
    "sample_perf_2020.txt": 2_517_857,
}
VALID_DELINQUENCY = re.compile(r"^(\d{2}|RA|XX)$")
VALID_ZERO_BALANCE_CODES = {"01", "02", "03", "09", "15", "16", "96"}


def valid_month(value: str) -> bool:
    if not re.fullmatch(r"\d{6}", value):
        return False
    try:
        datetime.strptime(value, "%Y%m")
        return True
    except ValueError:
        return False


def is_number(value: str) -> bool:
    if value == "":
        return False
    try:
        float(value)
        return True
    except ValueError:
        return False


def origination_checks(path: Path) -> tuple[dict, set[str]]:
    row_count = 0
    bad_field_count = 0
    duplicate_ids = 0
    invalid_dates = 0
    invalid_numbers = 0
    loan_ids: set[str] = set()
    missing = Counter()
    sentinels = Counter()
    important_fields = {
        "classic_fico": 0, "first_payment_date": 1, "maturity_date": 3,
        "original_cltv": 8, "original_dti": 9, "original_upb": 10, "loan_id": 19,
    }

    with path.open("r", encoding="ascii", newline="") as handle:
        for row in csv.reader(handle, delimiter="|"):
            row_count += 1
            if len(row) != len(ORIG_COLUMNS):
                bad_field_count += 1
                continue
            for name, index in important_fields.items():
                if row[index] == "":
                    missing[name] += 1

            loan_id = row[19]
            if loan_id in loan_ids:
                duplicate_ids += 1
            loan_ids.add(loan_id)
            invalid_dates += sum(not valid_month(row[index]) for index in (1, 3))
            for index in (0, 8, 9, 10, 11, 12, 21, 30):
                if not is_number(row[index]):
                    invalid_numbers += 1
            for index, sentinel in ((0, "9999"), (30, "9999"), (5, "999"),
                                    (8, "999"), (9, "999"), (11, "999")):
                if row[index] == sentinel:
                    sentinels[f"{ORIG_COLUMNS[index]}={sentinel}"] += 1

    return {
        "file": str(path),
        "row_count": row_count,
        "expected_fields": len(ORIG_COLUMNS),
        "rows_with_wrong_field_count": bad_field_count,
        "duplicate_loan_ids": duplicate_ids,
        "invalid_required_dates": invalid_dates,
        "nonnumeric_values_in_numeric_fields": invalid_numbers,
        "missing_important_fields": dict(missing),
        "official_sentinel_counts": dict(sentinels),
    }, loan_ids


def performance_checks(path: Path, origination_ids: set[str]) -> tuple[dict, set[str]]:
    row_count = 0
    bad_field_count = 0
    invalid_periods = 0
    invalid_numbers = 0
    invalid_statuses = Counter()
    invalid_zero_balance_codes = Counter()
    missing = Counter()
    performance_ids: set[str] = set()
    min_period = None
    max_period = None
    important_fields = {
        "loan_id": 0, "period": 1, "current_actual_upb": 2,
        "current_loan_delinquency_status": 3, "loan_age": 4,
    }

    with path.open("r", encoding="ascii", newline="") as handle:
        for row in csv.reader(handle, delimiter="|"):
            row_count += 1
            if len(row) != len(PERF_COLUMNS):
                bad_field_count += 1
                continue
            for name, index in important_fields.items():
                if row[index] == "":
                    missing[name] += 1

            performance_ids.add(row[0])
            if valid_month(row[1]):
                min_period = row[1] if min_period is None else min(min_period, row[1])
                max_period = row[1] if max_period is None else max(max_period, row[1])
            else:
                invalid_periods += 1
            for index in (2, 4):
                if not is_number(row[index]):
                    invalid_numbers += 1
            status = row[3]
            if not VALID_DELINQUENCY.fullmatch(status or ""):
                invalid_statuses[status or "<blank>"] += 1
            zero_balance_code = row[8]
            if zero_balance_code and zero_balance_code not in VALID_ZERO_BALANCE_CODES:
                invalid_zero_balance_codes[zero_balance_code] += 1

    return {
        "file": str(path),
        "row_count": row_count,
        "expected_fields": len(PERF_COLUMNS),
        "rows_with_wrong_field_count": bad_field_count,
        "invalid_periods": invalid_periods,
        "nonnumeric_upb_or_loan_age": invalid_numbers,
        "invalid_delinquency_statuses": dict(invalid_statuses),
        "invalid_zero_balance_codes": dict(invalid_zero_balance_codes),
        "missing_important_fields": dict(missing),
        "min_performance_month": min_period,
        "max_performance_month": max_period,
        "performance_loans_without_origination": len(performance_ids - origination_ids),
        "originations_without_performance": len(origination_ids - performance_ids),
    }, performance_ids


def find_failures(result: dict) -> list[str]:
    failures = []
    for year, checks in result["years"].items():
        for dataset, summary in checks.items():
            expected = EXPECTED_ROWS[Path(summary["file"]).name]
            if summary["row_count"] != expected:
                failures.append(f"{year} {dataset}: expected {expected:,} rows")
            for key, value in summary.items():
                if key.startswith(("rows_with_wrong", "duplicate_", "invalid_", "nonnumeric_")) and value:
                    failures.append(f"{year} {dataset}: {key}={value}")
            if dataset == "performance":
                for key in ("performance_loans_without_origination", "originations_without_performance"):
                    if summary[key]:
                        failures.append(f"{year} performance: {key}={summary[key]}")
    if result["cross_vintage_loan_id_overlap"]:
        failures.append("loan IDs overlap between the 2019 and 2020 samples")
    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--raw-dir", type=Path, default=Path("data/raw"))
    parser.add_argument("--output", type=Path,
                        default=Path("data/processed/validation/raw_data_validation.json"))
    args = parser.parse_args()

    paths = {
        year: (args.raw_dir / f"sample_{year}" / f"sample_orig_{year}.txt",
               args.raw_dir / f"sample_{year}" / f"sample_perf_{year}.txt")
        for year in (2019, 2020)
    }
    missing_files = [str(path) for pair in paths.values() for path in pair if not path.is_file()]
    if missing_files:
        raise FileNotFoundError("Missing source files:\n- " + "\n- ".join(missing_files))

    result = {"years": {}}
    origination_ids = {}
    for year, (orig_path, perf_path) in paths.items():
        orig_summary, orig_ids = origination_checks(orig_path)
        perf_summary, _ = performance_checks(perf_path, orig_ids)
        result["years"][str(year)] = {
            "origination": orig_summary,
            "performance": perf_summary,
        }
        origination_ids[year] = orig_ids
        print(f"{year}: {orig_summary['row_count']:,} originations, "
              f"{perf_summary['row_count']:,} performance rows, "
              f"{perf_summary['min_performance_month']}-{perf_summary['max_performance_month']}")

    result["cross_vintage_loan_id_overlap"] = len(
        origination_ids[2019] & origination_ids[2020]
    )
    failures = find_failures(result)
    result["passed"] = not failures
    result["failures"] = failures
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    if failures:
        print("Raw data validation failed:\n- " + "\n- ".join(failures), file=sys.stderr)
        return 1
    print(f"Raw data checks passed. Summary: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
