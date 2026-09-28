#!/usr/bin/env python3
"""Stream-validate Freddie Mac SFLLD Release 47 annual sample files.

The raw files are never modified.  The script writes only a compact JSON profile.
Column positions follow Freddie Mac's Release 47 (July 2026) user guide:
https://www.freddiemac.com/fmac-resources/research/pdf/general_user_guide_july_2026.pdf
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from collections import Counter
from datetime import datetime
from pathlib import Path


GUIDE_VERSION = "Freddie Mac SFLLD General User Guide, Release 47, July 2026"
GUIDE_URL = (
    "https://www.freddiemac.com/fmac-resources/research/pdf/"
    "general_user_guide_july_2026.pdf"
)

ORIG_COLUMNS = [
    "classic_fico",
    "first_payment_date",
    "first_time_homebuyer_indicator",
    "maturity_date",
    "msa_or_metropolitan_division",
    "mi_percentage",
    "number_of_units",
    "occupancy_status",
    "original_cltv",
    "original_dti",
    "original_upb",
    "original_ltv",
    "original_interest_rate",
    "channel",
    "prepayment_penalty_indicator",
    "amortization_type",
    "property_state",
    "property_type",
    "postal_code",
    "loan_id",
    "loan_purpose",
    "original_loan_term",
    "number_of_borrowers",
    "seller_name",
    "super_conforming_flag",
    "pre_harp_loan_id",
    "special_eligibility_program",
    "harp_indicator",
    "property_valuation_method",
    "interest_only_indicator",
    "vantage_score_4",
]

PERF_COLUMNS = [
    "loan_id",
    "period",
    "current_actual_upb",
    "current_loan_delinquency_status",
    "loan_age",
    "remaining_months_to_legal_maturity",
    "defect_settlement_date",
    "modification_flag",
    "zero_balance_code",
    "zero_balance_effective_date",
    "current_interest_rate",
    "current_non_interest_bearing_upb",
    "ddlpi",
    "mi_recoveries",
    "net_sales_proceeds",
    "non_mi_recoveries",
    "total_expenses",
    "legal_costs",
    "maintenance_and_preservation_costs",
    "taxes_and_insurance",
    "miscellaneous_expenses",
    "actual_loss",
    "cumulative_modification_costs",
    "interest_rate_step_indicator",
    "payment_deferral_flag",
    "estimated_ltv",
    "zero_balance_removal_upb",
    "delinquent_accrued_interest",
    "delinquency_due_to_disaster",
    "borrower_assistance_plan",
    "current_period_modification_costs",
    "current_interest_bearing_upb",
    "mi_cancellation_indicator",
    "servicer_name",
    "bankruptcy_cramdown_costs",
]

LOAN_ID_RE = re.compile(r"^[FA]\d{2}Q[1-4]\d{7}$")
NUMERIC_DELINQUENCY_RE = re.compile(r"^\d{2}$")
VALID_ZERO_BALANCE_CODES = {"01", "02", "03", "09", "15", "16", "96"}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def valid_yyyymm(value: str) -> bool:
    if not re.fullmatch(r"\d{6}", value):
        return False
    try:
        datetime.strptime(value, "%Y%m")
    except ValueError:
        return False
    return True


def month_index(value: str) -> int:
    return int(value[:4]) * 12 + int(value[4:]) - 1


def yyyymm_from_index(value: int) -> str:
    year, month_zero_based = divmod(value, 12)
    return f"{year:04d}{month_zero_based + 1:02d}"


def parse_float(value: str) -> float | None:
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def base_profile(path: Path) -> dict:
    return {
        "path": str(path),
        "size_bytes": path.stat().st_size,
        "sha256": sha256(path),
        "encoding_observed": "ASCII-compatible text",
        "delimiter": "|",
        "header_row": False,
    }


def profile_origination(path: Path, vintage: int) -> tuple[dict, set[str]]:
    profile = base_profile(path)
    field_counts: Counter[int] = Counter()
    blank_counts: Counter[str] = Counter()
    sentinel_counts: Counter[str] = Counter()
    invalid_counts: Counter[str] = Counter()
    loan_ids: set[str] = set()
    duplicate_ids = 0
    rows = 0
    first_payment_dates: list[str] = []
    maturity_dates: list[str] = []
    loan_id_quarters: Counter[str] = Counter()
    selected_value_counts = {
        field: Counter()
        for field in (
            "amortization_type",
            "prepayment_penalty_indicator",
            "interest_only_indicator",
            "harp_indicator",
            "property_valuation_method",
        )
    }

    with path.open("r", encoding="ascii", newline="") as handle:
        reader = csv.reader(handle, delimiter="|")
        for row in reader:
            rows += 1
            field_counts[len(row)] += 1
            if len(row) != len(ORIG_COLUMNS):
                invalid_counts["wrong_column_count"] += 1
                continue
            record = dict(zip(ORIG_COLUMNS, row))
            blank_counts.update(name for name, value in record.items() if value == "")

            loan_id = record["loan_id"]
            if loan_id in loan_ids:
                duplicate_ids += 1
            loan_ids.add(loan_id)
            if not LOAN_ID_RE.fullmatch(loan_id):
                invalid_counts["malformed_loan_id"] += 1
            elif int(loan_id[1:3]) != vintage % 100:
                invalid_counts["loan_id_vintage_mismatch"] += 1
            else:
                loan_id_quarters[loan_id[3:5]] += 1

            for field in ("first_payment_date", "maturity_date"):
                if not valid_yyyymm(record[field]):
                    invalid_counts[f"invalid_{field}"] += 1
            if valid_yyyymm(record["first_payment_date"]):
                first_payment_dates.append(record["first_payment_date"])
            if valid_yyyymm(record["maturity_date"]):
                maturity_dates.append(record["maturity_date"])
            for field, counts in selected_value_counts.items():
                counts[record[field] if record[field] else "<blank>"] += 1

            for field, sentinel in (
                ("classic_fico", "9999"),
                ("vantage_score_4", "9999"),
                ("original_cltv", "999"),
                ("original_dti", "999"),
                ("original_ltv", "999"),
                ("mi_percentage", "999"),
            ):
                if record[field] == sentinel:
                    sentinel_counts[f"{field}={sentinel}"] += 1

            fico = parse_float(record["classic_fico"])
            if fico is None or (record["classic_fico"] != "9999" and not 300 <= fico <= 850):
                invalid_counts["invalid_classic_fico"] += 1
            vantage = parse_float(record["vantage_score_4"])
            if vantage is None or (
                record["vantage_score_4"] != "9999" and not 300 <= vantage <= 850
            ):
                invalid_counts["invalid_vantage_score_4"] += 1
            dti = parse_float(record["original_dti"])
            if dti is None or (record["original_dti"] != "999" and not 0 <= dti <= 65):
                invalid_counts["invalid_original_dti"] += 1
            for field in ("original_ltv", "original_cltv"):
                value = parse_float(record[field])
                if value is None or (record[field] != "999" and not 1 <= value <= 998):
                    invalid_counts[f"invalid_{field}"] += 1
            upb = parse_float(record["original_upb"])
            if upb is None or upb <= 0:
                invalid_counts["invalid_original_upb"] += 1
            rate = parse_float(record["original_interest_rate"])
            if rate is None or rate <= 0:
                invalid_counts["invalid_original_interest_rate"] += 1
            term = parse_float(record["original_loan_term"])
            if term is None or term <= 0:
                invalid_counts["invalid_original_loan_term"] += 1

    profile.update(
        {
            "record_count": rows,
            "expected_column_count": len(ORIG_COLUMNS),
            "observed_column_count_distribution": dict(sorted(field_counts.items())),
            "column_positions": {str(i + 1): name for i, name in enumerate(ORIG_COLUMNS)},
            "distinct_loan_ids": len(loan_ids),
            "duplicate_loan_id_rows": duplicate_ids,
            "loan_id_quarter_counts": dict(sorted(loan_id_quarters.items())),
            "first_payment_date_min": min(first_payment_dates, default=None),
            "first_payment_date_max": max(first_payment_dates, default=None),
            "maturity_date_min": min(maturity_dates, default=None),
            "maturity_date_max": max(maturity_dates, default=None),
            "selected_value_counts": {
                field: dict(sorted(counts.items()))
                for field, counts in selected_value_counts.items()
            },
            "blank_counts": dict(sorted(blank_counts.items())),
            "documented_sentinel_counts": dict(sorted(sentinel_counts.items())),
            "invalid_counts": dict(sorted(invalid_counts.items())),
        }
    )
    return profile, loan_ids


def profile_performance(path: Path, orig_ids: set[str]) -> tuple[dict, set[str]]:
    profile = base_profile(path)
    field_counts: Counter[int] = Counter()
    blank_counts: Counter[str] = Counter()
    invalid_counts: Counter[str] = Counter()
    delinquency_counts: Counter[str] = Counter()
    zero_balance_counts: Counter[str] = Counter()
    monthly_record_counts: Counter[str] = Counter()
    perf_ids: set[str] = set()
    completed_ids: set[str] = set()

    current_loan: str | None = None
    months_for_current_loan: set[str] = set()
    previous_period: str | None = None
    terminal_rows_for_current_loan = 0
    terminal_seen = False

    rows = 0
    duplicate_loan_period_rows = 0
    non_contiguous_loan_blocks = 0
    adjacent_one_month_pairs = 0
    gap_pairs = 0
    reverse_period_pairs = 0
    rows_after_terminal_event = 0
    loans_with_multiple_terminal_rows = 0
    min_period: str | None = None
    max_period: str | None = None
    current_upb_zero_rows = 0
    current_upb_positive_rows = 0
    loan_age_zero_rows = 0
    cramdown_populated_rows = 0
    cramdown_non_terminal_rows = 0
    cramdown_nonzero_rows = 0
    cramdown_periods: list[str] = []

    def finish_loan() -> None:
        nonlocal loans_with_multiple_terminal_rows
        if terminal_rows_for_current_loan > 1:
            loans_with_multiple_terminal_rows += 1

    with path.open("r", encoding="ascii", newline="") as handle:
        reader = csv.reader(handle, delimiter="|")
        for row in reader:
            rows += 1
            field_counts[len(row)] += 1
            if len(row) != len(PERF_COLUMNS):
                invalid_counts["wrong_column_count"] += 1
                continue
            record = dict(zip(PERF_COLUMNS, row))
            blank_counts.update(name for name, value in record.items() if value == "")
            loan_id = record["loan_id"]
            period = record["period"]
            perf_ids.add(loan_id)

            if current_loan != loan_id:
                if current_loan is not None:
                    finish_loan()
                    completed_ids.add(current_loan)
                if loan_id in completed_ids:
                    non_contiguous_loan_blocks += 1
                current_loan = loan_id
                months_for_current_loan = set()
                previous_period = None
                terminal_rows_for_current_loan = 0
                terminal_seen = False

            if period in months_for_current_loan:
                duplicate_loan_period_rows += 1
            months_for_current_loan.add(period)

            if not LOAN_ID_RE.fullmatch(loan_id):
                invalid_counts["malformed_loan_id"] += 1
            if not valid_yyyymm(period):
                invalid_counts["invalid_period"] += 1
            else:
                monthly_record_counts[period] += 1
                min_period = period if min_period is None or period < min_period else min_period
                max_period = period if max_period is None or period > max_period else max_period
                if previous_period and valid_yyyymm(previous_period):
                    distance = month_index(period) - month_index(previous_period)
                    if distance == 1:
                        adjacent_one_month_pairs += 1
                    elif distance > 1:
                        gap_pairs += 1
                    elif distance < 0:
                        reverse_period_pairs += 1
                previous_period = period

            status = record["current_loan_delinquency_status"]
            delinquency_counts[status if status else "<blank>"] += 1
            if not (NUMERIC_DELINQUENCY_RE.fullmatch(status) or status in {"RA", "XX"}):
                invalid_counts["invalid_delinquency_status"] += 1

            current_upb = parse_float(record["current_actual_upb"])
            if current_upb is None:
                invalid_counts["missing_or_non_numeric_current_actual_upb"] += 1
            elif current_upb < 0:
                invalid_counts["negative_current_actual_upb"] += 1
            elif current_upb == 0:
                current_upb_zero_rows += 1
            else:
                current_upb_positive_rows += 1
            if record["loan_age"] == "0":
                loan_age_zero_rows += 1

            for field in (
                "defect_settlement_date",
                "zero_balance_effective_date",
                "ddlpi",
            ):
                if record[field] and not valid_yyyymm(record[field]):
                    invalid_counts[f"invalid_{field}"] += 1

            zero_code = record["zero_balance_code"]
            if zero_code:
                zero_balance_counts[zero_code] += 1
                terminal_rows_for_current_loan += 1
                terminal_seen = True
                if zero_code not in VALID_ZERO_BALANCE_CODES:
                    invalid_counts["invalid_zero_balance_code"] += 1
                if not record["zero_balance_effective_date"]:
                    invalid_counts["zero_balance_code_without_effective_date"] += 1
            elif record["zero_balance_effective_date"]:
                invalid_counts["effective_date_without_zero_balance_code"] += 1
            elif terminal_seen:
                rows_after_terminal_event += 1

            cramdown = record["bankruptcy_cramdown_costs"]
            if cramdown:
                cramdown_populated_rows += 1
                if not zero_code:
                    cramdown_non_terminal_rows += 1
                cramdown_value = parse_float(cramdown)
                if cramdown_value is None:
                    invalid_counts["non_numeric_bankruptcy_cramdown_costs"] += 1
                elif cramdown_value != 0:
                    cramdown_nonzero_rows += 1
                if valid_yyyymm(period):
                    cramdown_periods.append(period)

    if current_loan is not None:
        finish_loan()

    unmatched_perf_ids = perf_ids - orig_ids
    orig_without_perf = orig_ids - perf_ids
    uniqueness_is_exact = non_contiguous_loan_blocks == 0
    missing_global_months: list[str] = []
    if min_period and max_period:
        for value in range(month_index(min_period), month_index(max_period) + 1):
            period = yyyymm_from_index(value)
            if period not in monthly_record_counts:
                missing_global_months.append(period)
    profile.update(
        {
            "record_count": rows,
            "expected_column_count": len(PERF_COLUMNS),
            "observed_column_count_distribution": dict(sorted(field_counts.items())),
            "column_positions": {str(i + 1): name for i, name in enumerate(PERF_COLUMNS)},
            "distinct_loan_ids": len(perf_ids),
            "loan_period_duplicate_rows": duplicate_loan_period_rows,
            "loan_period_uniqueness_check_exact": uniqueness_is_exact,
            "non_contiguous_loan_blocks": non_contiguous_loan_blocks,
            "period_min": min_period,
            "period_max": max_period,
            "missing_calendar_months_in_global_range": missing_global_months,
            "monthly_record_counts": dict(sorted(monthly_record_counts.items())),
            "adjacent_natural_month_pairs": adjacent_one_month_pairs,
            "pairs_with_month_gaps": gap_pairs,
            "reverse_period_pairs": reverse_period_pairs,
            "rows_after_terminal_event": rows_after_terminal_event,
            "loans_with_multiple_terminal_rows": loans_with_multiple_terminal_rows,
            "current_actual_upb_positive_rows": current_upb_positive_rows,
            "current_actual_upb_zero_rows": current_upb_zero_rows,
            "loan_age_zero_rows": loan_age_zero_rows,
            "bankruptcy_cramdown_populated_rows": cramdown_populated_rows,
            "bankruptcy_cramdown_non_terminal_rows": cramdown_non_terminal_rows,
            "bankruptcy_cramdown_nonzero_rows": cramdown_nonzero_rows,
            "bankruptcy_cramdown_period_min": min(cramdown_periods, default=None),
            "bankruptcy_cramdown_period_max": max(cramdown_periods, default=None),
            "origination_ids_with_performance": len(orig_ids & perf_ids),
            "origination_ids_without_performance": len(orig_without_perf),
            "performance_ids_without_origination": len(unmatched_perf_ids),
            "origination_to_performance_match_rate": (
                len(orig_ids & perf_ids) / len(orig_ids) if orig_ids else None
            ),
            "performance_to_origination_match_rate": (
                len(orig_ids & perf_ids) / len(perf_ids) if perf_ids else None
            ),
            "blank_counts": dict(sorted(blank_counts.items())),
            "delinquency_status_counts": dict(sorted(delinquency_counts.items())),
            "zero_balance_code_counts": dict(sorted(zero_balance_counts.items())),
            "invalid_counts": dict(sorted(invalid_counts.items())),
        }
    )
    return profile, perf_ids


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--raw-dir", type=Path, default=Path("data/raw"))
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("data/processed/validation/raw_data_profile.json"),
    )
    args = parser.parse_args()

    result = {
        "generated_by": "scripts/validate_raw_data.py",
        "official_layout_reference": {"version": GUIDE_VERSION, "url": GUIDE_URL},
        "years": {},
    }
    orig_ids_by_year: dict[int, set[str]] = {}
    perf_ids_by_year: dict[int, set[str]] = {}

    for year in (2019, 2020):
        year_dir = args.raw_dir / f"sample_{year}"
        orig_path = year_dir / f"sample_orig_{year}.txt"
        perf_path = year_dir / f"sample_perf_{year}.txt"
        orig_profile, orig_ids = profile_origination(orig_path, year)
        perf_profile, perf_ids = profile_performance(perf_path, orig_ids)
        orig_ids_by_year[year] = orig_ids
        perf_ids_by_year[year] = perf_ids
        result["years"][str(year)] = {
            "origination": orig_profile,
            "performance": perf_profile,
        }

    result["cross_year_checks"] = {
        "origination_loan_id_overlap": len(orig_ids_by_year[2019] & orig_ids_by_year[2020]),
        "performance_loan_id_overlap": len(perf_ids_by_year[2019] & perf_ids_by_year[2020]),
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"output": str(args.output), "years": list(result["years"])}))


if __name__ == "__main__":
    main()
