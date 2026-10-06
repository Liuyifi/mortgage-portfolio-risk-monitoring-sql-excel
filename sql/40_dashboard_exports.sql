-- Local aggregate extracts used by the Excel dashboard.
-- This file creates TEMP tables: it never changes the persisted DuckDB database.
-- {{MIN_CELL_COUNT}} is replaced by scripts/export_dashboard_data.py.

CREATE OR REPLACE TEMP TABLE dashboard_portfolio_monthly AS
SELECT
    as_of_month,
    on_book_loan_count,
    on_book_upb,
    delinquency_eligible_loan_count,
    delinquency_eligible_upb,
    excluded_pre_due_loan_count,
    excluded_pre_due_upb,
    excluded_loan_age_zero_loan_count,
    excluded_loan_age_zero_upb,
    excluded_invalid_loan_age_loan_count,
    excluded_invalid_loan_age_upb,
    excluded_non_numeric_status_loan_count,
    excluded_non_numeric_status_upb,
    dq30_loan_count,
    dq60_loan_count,
    dq90_loan_count,
    dq30_upb,
    dq60_upb,
    dq90_upb,
    numeric_status_coverage_count_rate,
    numeric_status_coverage_balance_rate,
    dq30_count_rate,
    dq60_count_rate,
    dq90_count_rate,
    dq30_balance_rate,
    dq60_balance_rate,
    dq90_balance_rate
FROM mart.portfolio_monthly
ORDER BY as_of_month;

CREATE OR REPLACE TEMP TABLE dashboard_portfolio_monthly_by_segment AS
WITH source AS (
    SELECT *
    FROM mart.portfolio_monthly_by_segment
), partition_stats AS (
    SELECT
        *,
        sum(CASE WHEN on_book_loan_count < {{MIN_CELL_COUNT}} THEN on_book_loan_count ELSE 0 END)
            OVER (PARTITION BY as_of_month, segment_name) AS low_volume_loan_count,
        row_number() OVER (
            PARTITION BY as_of_month, segment_name
            ORDER BY
                CASE WHEN on_book_loan_count >= {{MIN_CELL_COUNT}} THEN 0 ELSE 1 END,
                on_book_loan_count,
                segment_value
        ) AS smallest_safe_rank
    FROM source
), tagged AS (
    SELECT
        *,
        (
            on_book_loan_count < {{MIN_CELL_COUNT}}
            OR (
                low_volume_loan_count BETWEEN 1 AND {{MIN_CELL_COUNT}} - 1
                AND smallest_safe_rank = 1
            )
        ) AS privacy_merge
    FROM partition_stats
), rolled AS (
    SELECT
        as_of_month,
        segment_name,
        CASE WHEN privacy_merge THEN 'Other / low volume' ELSE segment_value END AS segment_value,
        bool_or(privacy_merge) AS privacy_merged,
        count(*) AS source_mart_row_count,
        sum(on_book_loan_count) AS on_book_loan_count,
        sum(on_book_upb) AS on_book_upb,
        sum(delinquency_eligible_loan_count) AS delinquency_eligible_loan_count,
        sum(delinquency_eligible_upb) AS delinquency_eligible_upb,
        sum(dq30_loan_count) AS dq30_loan_count,
        sum(dq60_loan_count) AS dq60_loan_count,
        sum(dq90_loan_count) AS dq90_loan_count,
        sum(dq30_upb) AS dq30_upb,
        sum(dq60_upb) AS dq60_upb,
        sum(dq90_upb) AS dq90_upb
    FROM tagged
    GROUP BY as_of_month, segment_name,
             CASE WHEN privacy_merge THEN 'Other / low volume' ELSE segment_value END
), rated AS (
    SELECT
        *,
        delinquency_eligible_loan_count::DOUBLE / NULLIF(on_book_loan_count, 0)
            AS numeric_status_coverage_count_rate,
        delinquency_eligible_upb::DOUBLE / NULLIF(on_book_upb, 0)
            AS numeric_status_coverage_balance_rate,
        dq30_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq30_count_rate,
        dq60_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq60_count_rate,
        dq90_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq90_count_rate,
        dq30_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq30_balance_rate,
        dq60_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq60_balance_rate,
        dq90_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq90_balance_rate
    FROM rolled
)
SELECT
    *,
    CASE segment_name
        WHEN 'origination_year' THEN try_cast(segment_value AS INTEGER)
        WHEN 'classic_fico' THEN CASE segment_value
            WHEN '<620' THEN 1 WHEN '620-659' THEN 2 WHEN '660-699' THEN 3
            WHEN '700-739' THEN 4 WHEN '740-779' THEN 5 WHEN '780+' THEN 6
            WHEN 'Unknown' THEN 7 ELSE 99 END
        WHEN 'original_cltv' THEN CASE segment_value
            WHEN '<=60' THEN 1 WHEN '61-70' THEN 2 WHEN '71-80' THEN 3
            WHEN '81-90' THEN 4 WHEN '91-95' THEN 5 WHEN '>95' THEN 6
            WHEN 'Unknown' THEN 7 ELSE 99 END
        WHEN 'original_dti' THEN CASE segment_value
            WHEN '<=20' THEN 1 WHEN '21-30' THEN 2 WHEN '31-40' THEN 3
            WHEN '41-45' THEN 4 WHEN '46-50' THEN 5 WHEN '>50' THEN 6
            WHEN 'Unknown' THEN 7 ELSE 99 END
        ELSE 99
    END AS segment_sort_order
FROM rated
ORDER BY as_of_month, segment_name, segment_sort_order, segment_value;

-- Only the all-portfolio transition scope is exported. Vintage transition cuts
-- are omitted because month x vintage x state paths produce many
-- very small cells. Entire from-state cohorts below the threshold are suppressed;
-- rare destinations inside retained cohorts are rolled into one safe bucket.
CREATE OR REPLACE TEMP TABLE dashboard_delinquency_transition_monthly AS
WITH source AS (
    SELECT *
    FROM mart.delinquency_transition_monthly
    WHERE segment_name = 'all'
), partition_stats AS (
    SELECT
        *,
        sum(transition_loan_count) OVER (PARTITION BY from_month, from_state)
            AS cohort_loan_count,
        sum(transition_from_upb) OVER (PARTITION BY from_month, from_state)
            AS cohort_upb,
        sum(CASE WHEN transition_loan_count < {{MIN_CELL_COUNT}} THEN transition_loan_count ELSE 0 END)
            OVER (PARTITION BY from_month, from_state) AS low_volume_loan_count,
        row_number() OVER (
            PARTITION BY from_month, from_state
            ORDER BY
                CASE WHEN transition_loan_count >= {{MIN_CELL_COUNT}} THEN 0 ELSE 1 END,
                transition_loan_count,
                to_state
        ) AS smallest_safe_rank
    FROM source
), retained AS (
    SELECT
        *,
        (
            transition_loan_count < {{MIN_CELL_COUNT}}
            OR (
                low_volume_loan_count BETWEEN 1 AND {{MIN_CELL_COUNT}} - 1
                AND smallest_safe_rank = 1
            )
        ) AS privacy_merge
    FROM partition_stats
    WHERE cohort_loan_count >= {{MIN_CELL_COUNT}}
), rolled AS (
    SELECT
        from_month,
        to_month,
        from_state,
        CASE WHEN privacy_merge THEN 'OTHER_OR_LOW_VOLUME' ELSE to_state END AS to_state,
        bool_or(privacy_merge) AS privacy_merged,
        count(*) AS source_mart_row_count,
        sum(transition_loan_count) AS transition_loan_count,
        sum(transition_from_upb) AS transition_from_upb,
        max(cohort_loan_count) AS observable_from_state_loan_count,
        max(cohort_upb) AS observable_from_state_upb
    FROM retained
    GROUP BY from_month, to_month, from_state,
             CASE WHEN privacy_merge THEN 'OTHER_OR_LOW_VOLUME' ELSE to_state END
)
SELECT
    *,
    transition_loan_count::DOUBLE / NULLIF(observable_from_state_loan_count, 0)
        AS transition_count_rate,
    transition_from_upb::DOUBLE / NULLIF(observable_from_state_upb, 0)
        AS transition_balance_rate,
    CASE from_state
        WHEN 'PRE_DUE' THEN 0 WHEN 'CURRENT' THEN 1 WHEN '30' THEN 2
        WHEN '60' THEN 3 WHEN '90+' THEN 4 WHEN 'REO_ACQUISITION' THEN 5
        WHEN 'UNKNOWN_LOAN_AGE' THEN 97 WHEN 'UNKNOWN' THEN 98 ELSE 6
    END AS from_state_sort_order,
    CASE to_state
        WHEN 'PRE_DUE' THEN 0 WHEN 'CURRENT' THEN 1 WHEN '30' THEN 2
        WHEN '60' THEN 3 WHEN '90+' THEN 4 WHEN 'REO_ACQUISITION' THEN 5
        WHEN 'UNKNOWN_LOAN_AGE' THEN 97 WHEN 'UNKNOWN' THEN 98
        WHEN 'OTHER_OR_LOW_VOLUME' THEN 99 ELSE 6
    END AS to_state_sort_order
FROM rolled
ORDER BY from_month, from_state_sort_order, to_state_sort_order, to_state;
