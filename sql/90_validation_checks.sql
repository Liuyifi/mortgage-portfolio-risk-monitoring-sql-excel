-- Checks that protect the portfolio metrics used in the analysis and dashboard.

CREATE TABLE quality.validation_results AS
WITH
source_counts AS (
    SELECT
        (SELECT count(*) FROM staging.origination WHERE source_year = 2019) = 50000
        AND (SELECT count(*) FROM staging.origination WHERE source_year = 2020) = 50000
        AND (SELECT count(*) FROM staging.monthly_performance WHERE source_year = 2019) = 1934614
        AND (SELECT count(*) FROM staging.monthly_performance WHERE source_year = 2020) = 2517857
        AS passed
),
origination_duplicates AS (
    SELECT count(*) AS failures FROM (
        SELECT loan_id FROM staging.origination
        GROUP BY loan_id HAVING loan_id IS NULL OR count(*) > 1
    )
),
performance_duplicates AS (
    SELECT count(*) AS failures FROM (
        SELECT loan_id, as_of_month FROM staging.monthly_performance
        GROUP BY loan_id, as_of_month
        HAVING loan_id IS NULL OR as_of_month IS NULL OR count(*) > 1
    )
),
coverage_failures AS (
    SELECT count(*) AS failures FROM (
        SELECT o.loan_id
        FROM staging.origination o
        LEFT JOIN (SELECT DISTINCT loan_id FROM staging.monthly_performance) p USING (loan_id)
        WHERE p.loan_id IS NULL
        UNION ALL
        SELECT p.loan_id
        FROM (SELECT DISTINCT loan_id FROM staging.monthly_performance) p
        LEFT JOIN staging.origination o USING (loan_id)
        WHERE o.loan_id IS NULL
    )
),
risk_field_failures AS (
    SELECT count(*) AS failures
    FROM staging.origination
    WHERE (classic_fico_raw = '9999' AND classic_fico IS NOT NULL)
       OR (classic_fico_raw <> '9999' AND classic_fico NOT BETWEEN 300 AND 850)
       OR (original_cltv_raw = '999' AND original_cltv IS NOT NULL)
       OR (original_cltv_raw <> '999' AND original_cltv NOT BETWEEN 1 AND 998)
       OR (original_dti_raw = '999' AND original_dti IS NOT NULL)
       OR (original_dti_raw <> '999' AND original_dti NOT BETWEEN 0 AND 65)
       OR original_upb IS NULL OR original_upb <= 0
),
month_failures AS (
    SELECT count(*) AS failures
    FROM staging.monthly_performance
    WHERE as_of_month IS NULL
       OR as_of_month < make_date(source_year, 1, 1)
       OR as_of_month > DATE '2026-03-01'
),
status_failures AS (
    SELECT count(*) AS failures
    FROM staging.monthly_performance
    WHERE delinquency_status_raw IS NULL
       OR NOT (
           regexp_full_match(delinquency_status_raw, '^[0-9]{2}$')
           OR delinquency_status_raw IN ('RA', 'XX')
       )
),
zero_balance_code_failures AS (
    SELECT count(*) AS failures
    FROM staging.monthly_performance
    WHERE zero_balance_code IS NOT NULL
      AND zero_balance_code NOT IN ('01', '02', '03', '09', '15', '16', '96')
),
on_book_failures AS (
    SELECT count(*) AS failures
    FROM staging.monthly_performance
    WHERE current_actual_upb IS NULL OR current_actual_upb < 0
       OR (zero_balance_code IS NULL) <> (zero_balance_effective_month IS NULL)
       OR is_on_book_eom IS DISTINCT FROM (
           current_actual_upb > 0 AND zero_balance_code IS NULL
       )
),
eligible_failures AS (
    SELECT count(*) AS failures
    FROM curated.fact_loan_month
    WHERE is_delinquency_eligible AND (
        NOT is_on_book_eom OR loan_age < 1 OR NOT is_numeric_delinquency_status
    )
),
count_hierarchy_failures AS (
    SELECT count(*) AS failures FROM (
        SELECT delinquency_eligible_loan_count, dq30_loan_count,
               dq60_loan_count, dq90_loan_count
        FROM mart.portfolio_monthly
        UNION ALL
        SELECT delinquency_eligible_loan_count, dq30_loan_count,
               dq60_loan_count, dq90_loan_count
        FROM mart.portfolio_monthly_by_segment
    )
    WHERE dq30_loan_count > delinquency_eligible_loan_count
       OR dq60_loan_count > dq30_loan_count
       OR dq90_loan_count > dq60_loan_count
),
upb_hierarchy_failures AS (
    SELECT count(*) AS failures FROM (
        SELECT delinquency_eligible_upb, dq30_upb, dq60_upb, dq90_upb
        FROM mart.portfolio_monthly
        UNION ALL
        SELECT delinquency_eligible_upb, dq30_upb, dq60_upb, dq90_upb
        FROM mart.portfolio_monthly_by_segment
    )
    WHERE dq30_upb > delinquency_eligible_upb
       OR dq60_upb > dq30_upb
       OR dq90_upb > dq60_upb
),
required_dimensions(segment_name) AS (
    VALUES ('origination_year'), ('classic_fico'), ('original_cltv'), ('original_dti')
),
segment_rollup AS (
    SELECT
        as_of_month,
        segment_name,
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
    FROM mart.portfolio_monthly_by_segment
    GROUP BY as_of_month, segment_name
),
segment_dimension_failures AS (
    SELECT count(*) AS failures
    FROM mart.portfolio_monthly p
    CROSS JOIN required_dimensions d
    LEFT JOIN segment_rollup s
      ON s.as_of_month = p.as_of_month AND s.segment_name = d.segment_name
    WHERE s.as_of_month IS NULL
),
segment_reconciliation_failures AS (
    SELECT count(*) AS failures
    FROM segment_rollup s
    JOIN mart.portfolio_monthly p USING (as_of_month)
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
),
spot_months(as_of_month) AS (
    VALUES (DATE '2020-06-01'), (DATE '2021-06-01'), (DATE '2026-03-01')
),
spot_recalculation AS (
    SELECT
        p.as_of_month,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
        ) AS on_book_count,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
        ), 0) AS on_book_upb,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
        ) AS eligible_count,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
        ), 0) AS eligible_upb,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1 AND try_cast(p.delinquency_status_raw AS INTEGER) >= 1
        ) AS dq30_count,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1 AND try_cast(p.delinquency_status_raw AS INTEGER) >= 2
        ) AS dq60_count,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1 AND try_cast(p.delinquency_status_raw AS INTEGER) >= 3
        ) AS dq90_count,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1 AND try_cast(p.delinquency_status_raw AS INTEGER) >= 1
        ), 0) AS dq30_upb,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1 AND try_cast(p.delinquency_status_raw AS INTEGER) >= 2
        ), 0) AS dq60_upb,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1 AND try_cast(p.delinquency_status_raw AS INTEGER) >= 3
        ), 0) AS dq90_upb
    FROM staging.monthly_performance p
    JOIN spot_months m USING (as_of_month)
    GROUP BY p.as_of_month
),
spot_failures AS (
    SELECT count(*) AS failures
    FROM spot_recalculation s
    JOIN mart.portfolio_monthly p USING (as_of_month)
    WHERE s.on_book_count <> p.on_book_loan_count
       OR s.on_book_upb <> p.on_book_upb
       OR s.eligible_count <> p.delinquency_eligible_loan_count
       OR s.eligible_upb <> p.delinquency_eligible_upb
       OR s.dq30_count <> p.dq30_loan_count
       OR s.dq60_count <> p.dq60_loan_count
       OR s.dq90_count <> p.dq90_loan_count
       OR s.dq30_upb <> p.dq30_upb
       OR s.dq60_upb <> p.dq60_upb
       OR s.dq90_upb <> p.dq90_upb
),
transition_month_failures AS (
    SELECT count(*) AS failures
    FROM curated.v_loan_month_transition
    WHERE to_month <> from_month + INTERVAL '1 month'
),
transition_source_failures AS (
    SELECT count(*) AS failures
    FROM curated.v_loan_month_transition
    WHERE from_upb <= 0 OR from_state LIKE 'TERMINATED:%'
),
transition_cohort_failures AS (
    SELECT count(*) AS failures FROM (
        SELECT from_month, segment_name, segment_value, from_state,
               sum(transition_loan_count) AS transition_count,
               max(observable_from_state_loan_count) AS denominator_count,
               sum(transition_from_upb) AS transition_upb,
               max(observable_from_state_upb) AS denominator_upb
        FROM mart.delinquency_transition_monthly
        GROUP BY from_month, segment_name, segment_value, from_state
    )
    WHERE transition_count <> denominator_count OR transition_upb <> denominator_upb
),
transition_rate_failures AS (
    SELECT count(*) AS failures FROM (
        SELECT from_month, segment_name, segment_value, from_state,
               abs(sum(transition_count_rate) - 1.0) AS count_delta,
               abs(sum(transition_balance_rate) - 1.0) AS balance_delta
        FROM mart.delinquency_transition_monthly
        GROUP BY from_month, segment_name, segment_value, from_state
    )
    WHERE count_delta > 1e-10 OR balance_delta > 1e-10
),
checks AS (
    SELECT 'source_row_counts' AS check_name,
           CASE WHEN (SELECT passed FROM source_counts) THEN 0 ELSE 1 END::BIGINT AS failed_row_count,
           'Expected 100,000 originations and 4,452,471 monthly rows across the two samples.' AS details
    UNION ALL SELECT 'origination_loan_id_unique', failures,
           'loan_id must be present and unique.' FROM origination_duplicates
    UNION ALL SELECT 'performance_loan_month_unique', failures,
           '(loan_id, as_of_month) must be present and unique.' FROM performance_duplicates
    UNION ALL SELECT 'origination_performance_coverage', failures,
           'Every loan must appear in both the origination and performance data.' FROM coverage_failures
    UNION ALL SELECT 'origination_risk_fields_valid', failures,
           'FICO, CLTV, DTI, UPB and official sentinel handling must be valid.' FROM risk_field_failures
    UNION ALL SELECT 'performance_months_valid', failures,
           'Observation months must be valid and within the sample window.' FROM month_failures
    UNION ALL SELECT 'delinquency_statuses_valid', failures,
           'Allowed statuses are 00-99, RA and XX.' FROM status_failures
    UNION ALL SELECT 'zero_balance_codes_valid', failures,
           'Zero Balance Code must be one of the documented values.' FROM zero_balance_code_failures
    UNION ALL SELECT 'on_book_population_consistent', failures,
           'UPB, zero-balance dates and the on-book flag must agree.' FROM on_book_failures
    UNION ALL SELECT 'eligible_population_is_on_book', failures,
           'Eligible observations must be on book, Loan Age 1+ and numeric status.' FROM eligible_failures
    UNION ALL SELECT 'delinquency_count_hierarchy', failures,
           '30+/60+/90+ counts must be nested within the eligible denominator.' FROM count_hierarchy_failures
    UNION ALL SELECT 'delinquency_upb_hierarchy', failures,
           '30+/60+/90+ UPB must be nested within the eligible denominator.' FROM upb_hierarchy_failures
    UNION ALL SELECT 'segment_dimensions_complete', failures,
           'Every month must include vintage, FICO, CLTV and DTI summaries.' FROM segment_dimension_failures
    UNION ALL SELECT 'segment_totals_reconcile', failures,
           'Each segment dimension must sum to the overall monthly totals.' FROM segment_reconciliation_failures
    UNION ALL SELECT 'monthly_metric_spot_checks', failures,
           'Three months recalculated from staging fields must match the monthly mart.' FROM spot_failures
    UNION ALL SELECT 'transition_pairs_are_adjacent_months', failures,
           'Transition pairs must be exactly one calendar month apart.' FROM transition_month_failures
    UNION ALL SELECT 'transition_sources_are_on_book', failures,
           'A transition must start from a positive-UPB, non-terminal observation.' FROM transition_source_failures
    UNION ALL SELECT 'transition_cohort_totals_reconcile', failures,
           'Destination flows must sum to each beginning-state cohort.' FROM transition_cohort_failures
    UNION ALL SELECT 'transition_rates_sum_to_one', failures,
           'Count and beginning-UPB rates must sum to one within each cohort.' FROM transition_rate_failures
)
SELECT
    check_name,
    failed_row_count,
    failed_row_count = 0 AS passed,
    details,
    current_timestamp AS checked_at
FROM checks;

CREATE UNIQUE INDEX idx_validation_check_name
    ON quality.validation_results(check_name);
