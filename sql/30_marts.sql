-- Portfolio metrics use the end-of-month on-book population. Delinquency-rate
-- denominators further require loan_age >= 1 and a numeric 00-99 status.

CREATE TABLE mart.portfolio_monthly AS
WITH aggregates AS (
    SELECT
        as_of_month,
        count(*) FILTER (WHERE is_on_book_eom) AS on_book_loan_count,
        coalesce(sum(current_actual_upb) FILTER (WHERE is_on_book_eom), 0) AS on_book_upb,
        count(*) FILTER (WHERE is_delinquency_eligible) AS delinquency_eligible_loan_count,
        coalesce(sum(current_actual_upb) FILTER (WHERE is_delinquency_eligible), 0) AS delinquency_eligible_upb,
        count(*) FILTER (
            WHERE is_on_book_eom AND monitoring_state = 'PRE_DUE'
        ) AS excluded_pre_due_loan_count,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_on_book_eom AND monitoring_state = 'PRE_DUE'
        ), 0) AS excluded_pre_due_upb,
        count(*) FILTER (
            WHERE is_on_book_eom AND loan_age = 0
        ) AS excluded_loan_age_zero_loan_count,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_on_book_eom AND loan_age = 0
        ), 0) AS excluded_loan_age_zero_upb,
        count(*) FILTER (
            WHERE is_on_book_eom AND (loan_age IS NULL OR loan_age < 0)
        ) AS excluded_invalid_loan_age_loan_count,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_on_book_eom AND (loan_age IS NULL OR loan_age < 0)
        ), 0) AS excluded_invalid_loan_age_upb,
        count(*) FILTER (
            WHERE is_on_book_eom AND loan_age >= 1 AND NOT is_numeric_delinquency_status
        ) AS excluded_non_numeric_status_loan_count,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_on_book_eom AND loan_age >= 1 AND NOT is_numeric_delinquency_status
        ), 0) AS excluded_non_numeric_status_upb,
        count(*) FILTER (WHERE is_delinquency_eligible AND delinquency_months >= 1) AS dq30_loan_count,
        count(*) FILTER (WHERE is_delinquency_eligible AND delinquency_months >= 2) AS dq60_loan_count,
        count(*) FILTER (WHERE is_delinquency_eligible AND delinquency_months >= 3) AS dq90_loan_count,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_delinquency_eligible AND delinquency_months >= 1
        ), 0) AS dq30_upb,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_delinquency_eligible AND delinquency_months >= 2
        ), 0) AS dq60_upb,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_delinquency_eligible AND delinquency_months >= 3
        ), 0) AS dq90_upb
    FROM curated.fact_loan_month
    GROUP BY as_of_month
)
SELECT
    *,
    delinquency_eligible_loan_count::DOUBLE / NULLIF(on_book_loan_count, 0) AS numeric_status_coverage_count_rate,
    delinquency_eligible_upb::DOUBLE / NULLIF(on_book_upb, 0) AS numeric_status_coverage_balance_rate,
    dq30_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq30_count_rate,
    dq60_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq60_count_rate,
    dq90_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq90_count_rate,
    dq30_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq30_balance_rate,
    dq60_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq60_balance_rate,
    dq90_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq90_balance_rate
FROM aggregates
ORDER BY as_of_month;

CREATE UNIQUE INDEX idx_portfolio_monthly_pk ON mart.portfolio_monthly(as_of_month);

CREATE TABLE mart.portfolio_monthly_by_segment AS
WITH expanded AS (
    SELECT as_of_month, 'origination_year' AS segment_name,
           origination_year::VARCHAR AS segment_value,
           current_actual_upb, delinquency_months, loan_age,
           is_on_book_eom, is_numeric_delinquency_status, is_delinquency_eligible
    FROM curated.fact_loan_month
    UNION ALL
    SELECT as_of_month, 'classic_fico' AS segment_name,
           fico_band AS segment_value,
           current_actual_upb, delinquency_months, loan_age,
           is_on_book_eom, is_numeric_delinquency_status, is_delinquency_eligible
    FROM curated.fact_loan_month
    UNION ALL
    SELECT as_of_month, 'original_cltv' AS segment_name,
           cltv_band AS segment_value,
           current_actual_upb, delinquency_months, loan_age,
           is_on_book_eom, is_numeric_delinquency_status, is_delinquency_eligible
    FROM curated.fact_loan_month
    UNION ALL
    SELECT as_of_month, 'original_dti' AS segment_name,
           dti_band AS segment_value,
           current_actual_upb, delinquency_months, loan_age,
           is_on_book_eom, is_numeric_delinquency_status, is_delinquency_eligible
    FROM curated.fact_loan_month
), aggregates AS (
    SELECT
        as_of_month,
        segment_name,
        segment_value,
        count(*) FILTER (WHERE is_on_book_eom) AS on_book_loan_count,
        coalesce(sum(current_actual_upb) FILTER (WHERE is_on_book_eom), 0) AS on_book_upb,
        count(*) FILTER (WHERE is_delinquency_eligible) AS delinquency_eligible_loan_count,
        coalesce(sum(current_actual_upb) FILTER (WHERE is_delinquency_eligible), 0) AS delinquency_eligible_upb,
        count(*) FILTER (WHERE is_delinquency_eligible AND delinquency_months >= 1) AS dq30_loan_count,
        count(*) FILTER (WHERE is_delinquency_eligible AND delinquency_months >= 2) AS dq60_loan_count,
        count(*) FILTER (WHERE is_delinquency_eligible AND delinquency_months >= 3) AS dq90_loan_count,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_delinquency_eligible AND delinquency_months >= 1
        ), 0) AS dq30_upb,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_delinquency_eligible AND delinquency_months >= 2
        ), 0) AS dq60_upb,
        coalesce(sum(current_actual_upb) FILTER (
            WHERE is_delinquency_eligible AND delinquency_months >= 3
        ), 0) AS dq90_upb
    FROM expanded
    GROUP BY as_of_month, segment_name, segment_value
)
SELECT
    *,
    dq30_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq30_count_rate,
    dq60_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq60_count_rate,
    dq90_loan_count::DOUBLE / NULLIF(delinquency_eligible_loan_count, 0) AS dq90_count_rate,
    dq30_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq30_balance_rate,
    dq60_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq60_balance_rate,
    dq90_upb::DOUBLE / NULLIF(delinquency_eligible_upb, 0) AS dq90_balance_rate
FROM aggregates
ORDER BY as_of_month, segment_name, segment_value;

CREATE UNIQUE INDEX idx_portfolio_segment_pk
    ON mart.portfolio_monthly_by_segment(as_of_month, segment_name, segment_value);

CREATE TABLE mart.delinquency_transition_monthly AS
WITH expanded AS (
    SELECT from_month, to_month, 'all' AS segment_name, 'All' AS segment_value,
           from_state, to_state, from_upb
    FROM curated.v_loan_month_transition
    UNION ALL
    SELECT from_month, to_month, 'origination_year' AS segment_name,
           origination_year::VARCHAR AS segment_value, from_state, to_state, from_upb
    FROM curated.v_loan_month_transition
), flows AS (
    SELECT
        from_month,
        to_month,
        segment_name,
        segment_value,
        from_state,
        to_state,
        count(*) AS transition_loan_count,
        sum(from_upb) AS transition_from_upb
    FROM expanded
    GROUP BY ALL
), denominators AS (
    SELECT
        *,
        sum(transition_loan_count) OVER (
            PARTITION BY from_month, segment_name, segment_value, from_state
        ) AS observable_from_state_loan_count,
        sum(transition_from_upb) OVER (
            PARTITION BY from_month, segment_name, segment_value, from_state
        ) AS observable_from_state_upb
    FROM flows
)
SELECT
    *,
    transition_loan_count::DOUBLE / NULLIF(observable_from_state_loan_count, 0) AS transition_count_rate,
    transition_from_upb::DOUBLE / NULLIF(observable_from_state_upb, 0) AS transition_balance_rate
FROM denominators
ORDER BY from_month, segment_name, segment_value, from_state, to_state;

CREATE UNIQUE INDEX idx_transition_monthly_pk
    ON mart.delinquency_transition_monthly(
        from_month, segment_name, segment_value, from_state, to_state
    );
