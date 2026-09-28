-- Executable SQL quality gates. The build runner raises an exception when any
-- ERROR row has failed_row_count > 0. Documented non-blocking observations use
-- WARNING severity and remain visible in the database/build summary.

CREATE TABLE quality.loan_age_audit_summary AS
WITH classes(age_class) AS (
    VALUES ('NULL_OR_BLANK'), ('UNPARSEABLE'), ('NEGATIVE'), ('ZERO')
), classified AS (
    SELECT
        loan_id,
        source_year,
        as_of_month,
        delinquency_status_raw,
        modification_flag,
        is_on_book_eom,
        is_terminal_record,
        CASE
            WHEN loan_age_raw IS NULL OR trim(loan_age_raw) = '' THEN 'NULL_OR_BLANK'
            WHEN loan_age IS NULL THEN 'UNPARSEABLE'
            WHEN loan_age < 0 THEN 'NEGATIVE'
            WHEN loan_age = 0 THEN 'ZERO'
            ELSE 'POSITIVE'
        END AS age_class
    FROM staging.monthly_performance
), aggregates AS (
    SELECT
        age_class,
        count(*) AS record_count,
        count(DISTINCT loan_id) AS distinct_loan_count,
        count(*) FILTER (WHERE delinquency_status_raw = '00') AS status_00_count,
        count(*) FILTER (WHERE delinquency_status_raw = '01') AS status_01_count,
        count(*) FILTER (
            WHERE delinquency_status_raw NOT IN ('00', '01')
               OR delinquency_status_raw IS NULL
        ) AS other_status_count,
        count(*) FILTER (WHERE is_on_book_eom) AS on_book_count,
        count(*) FILTER (WHERE is_terminal_record) AS terminal_count,
        count(*) FILTER (WHERE modification_flag = 'Y') AS current_modification_count,
        min(as_of_month) AS min_month,
        max(as_of_month) AS max_month
    FROM classified
    WHERE age_class <> 'POSITIVE'
    GROUP BY age_class
)
SELECT
    c.age_class,
    coalesce(a.record_count, 0) AS record_count,
    coalesce(a.distinct_loan_count, 0) AS distinct_loan_count,
    coalesce(a.status_00_count, 0) AS status_00_count,
    coalesce(a.status_01_count, 0) AS status_01_count,
    coalesce(a.other_status_count, 0) AS other_status_count,
    coalesce(a.on_book_count, 0) AS on_book_count,
    coalesce(a.terminal_count, 0) AS terminal_count,
    coalesce(a.current_modification_count, 0) AS current_modification_count,
    a.min_month,
    a.max_month
FROM classes AS c
LEFT JOIN aggregates AS a USING (age_class)
ORDER BY CASE c.age_class
    WHEN 'NULL_OR_BLANK' THEN 1
    WHEN 'UNPARSEABLE' THEN 2
    WHEN 'NEGATIVE' THEN 3
    WHEN 'ZERO' THEN 4
END;

CREATE TABLE quality.loan_age_zero_monthly AS
SELECT
    source_year,
    as_of_month,
    delinquency_status_raw,
    coalesce(modification_flag, 'NOT_MODIFIED') AS modification_status,
    is_terminal_record,
    count(*) AS record_count,
    count(DISTINCT loan_id) AS distinct_loan_count
FROM staging.monthly_performance
WHERE loan_age = 0
GROUP BY ALL
ORDER BY as_of_month, source_year, delinquency_status_raw, modification_status, is_terminal_record;

-- This is deliberately recomputed from staging base fields. It does not use
-- curated eligibility flags or mart values to establish the independent side.
CREATE TABLE quality.metric_reconciliation_samples AS
WITH target_months(as_of_month) AS (
    VALUES (DATE '2020-06-01'), (DATE '2021-06-01'), (DATE '2026-03-01')
), independent AS (
    SELECT
        p.as_of_month,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
        ) AS independent_on_book_loan_count,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
        ), 0) AS independent_on_book_upb,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0
              AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
        ) AS independent_eligible_loan_count,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0
              AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
        ), 0) AS independent_eligible_upb,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
              AND try_cast(p.delinquency_status_raw AS SMALLINT) >= 1
        ) AS independent_dq30_loan_count,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
              AND try_cast(p.delinquency_status_raw AS SMALLINT) >= 2
        ) AS independent_dq60_loan_count,
        count(*) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
              AND try_cast(p.delinquency_status_raw AS SMALLINT) >= 3
        ) AS independent_dq90_loan_count,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
              AND try_cast(p.delinquency_status_raw AS SMALLINT) >= 1
        ), 0) AS independent_dq30_upb,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
              AND try_cast(p.delinquency_status_raw AS SMALLINT) >= 2
        ), 0) AS independent_dq60_upb,
        coalesce(sum(p.current_actual_upb) FILTER (
            WHERE p.current_actual_upb > 0 AND p.zero_balance_code IS NULL
              AND p.loan_age >= 1
              AND regexp_full_match(p.delinquency_status_raw, '^[0-9]{2}$')
              AND try_cast(p.delinquency_status_raw AS SMALLINT) >= 3
        ), 0) AS independent_dq90_upb
    FROM staging.monthly_performance AS p
    JOIN target_months AS t USING (as_of_month)
    GROUP BY p.as_of_month
), compared AS (
    SELECT
        t.as_of_month,
        i.* EXCLUDE (as_of_month),
        m.on_book_loan_count AS mart_on_book_loan_count,
        m.on_book_upb AS mart_on_book_upb,
        m.delinquency_eligible_loan_count AS mart_eligible_loan_count,
        m.delinquency_eligible_upb AS mart_eligible_upb,
        m.dq30_loan_count AS mart_dq30_loan_count,
        m.dq60_loan_count AS mart_dq60_loan_count,
        m.dq90_loan_count AS mart_dq90_loan_count,
        m.dq30_upb AS mart_dq30_upb,
        m.dq60_upb AS mart_dq60_upb,
        m.dq90_upb AS mart_dq90_upb
    FROM target_months AS t
    LEFT JOIN independent AS i USING (as_of_month)
    LEFT JOIN mart.portfolio_monthly AS m USING (as_of_month)
)
SELECT
    *,
    independent_on_book_loan_count = mart_on_book_loan_count
        AND independent_on_book_upb = mart_on_book_upb
        AND independent_eligible_loan_count = mart_eligible_loan_count
        AND independent_eligible_upb = mart_eligible_upb
        AND independent_dq30_loan_count = mart_dq30_loan_count
        AND independent_dq60_loan_count = mart_dq60_loan_count
        AND independent_dq90_loan_count = mart_dq90_loan_count
        AND independent_dq30_upb = mart_dq30_upb
        AND independent_dq60_upb = mart_dq60_upb
        AND independent_dq90_upb = mart_dq90_upb AS reconciled
FROM compared
ORDER BY as_of_month;

CREATE TABLE quality.dq_test_results AS
WITH
orig_counts AS (
    SELECT source_year, count(*) AS row_count
    FROM staging.origination GROUP BY source_year
),
perf_counts AS (
    SELECT source_year, count(*) AS row_count
    FROM staging.monthly_performance GROUP BY source_year
),
orig_duplicates AS (
    SELECT count(*) AS failures FROM (
        SELECT loan_id FROM staging.origination GROUP BY loan_id HAVING count(*) <> 1
    )
),
perf_duplicates AS (
    SELECT count(*) AS failures FROM (
        SELECT loan_id, as_of_month
        FROM staging.monthly_performance
        GROUP BY loan_id, as_of_month HAVING count(*) <> 1
    )
),
post_terminal AS (
    SELECT count(*) AS failures
    FROM staging.monthly_performance AS p
    JOIN (
        SELECT loan_id, min(as_of_month) AS terminal_month
        FROM staging.monthly_performance
        WHERE zero_balance_code IS NOT NULL
        GROUP BY loan_id
    ) AS t USING (loan_id)
    WHERE p.as_of_month > t.terminal_month
),
month_gaps AS (
    SELECT count(*) AS failures
    FROM (
        SELECT loan_id, as_of_month,
               lead(as_of_month) OVER (PARTITION BY loan_id ORDER BY as_of_month) AS next_month
        FROM staging.monthly_performance
    )
    WHERE next_month IS NOT NULL
      AND next_month <> as_of_month + INTERVAL '1 month'
),
metric_failures AS (
    SELECT count(*) AS failures
    FROM mart.portfolio_monthly
    WHERE dq30_loan_count > delinquency_eligible_loan_count
       OR dq60_loan_count > dq30_loan_count
       OR dq90_loan_count > dq60_loan_count
       OR dq30_upb > delinquency_eligible_upb
       OR dq60_upb > dq30_upb
       OR dq90_upb > dq60_upb
       OR on_book_loan_count < delinquency_eligible_loan_count
       OR on_book_upb < delinquency_eligible_upb
),
rate_failures AS (
    SELECT count(*) AS failures
    FROM mart.portfolio_monthly
    WHERE coalesce(dq30_count_rate NOT BETWEEN 0 AND 1, false)
       OR coalesce(dq60_count_rate NOT BETWEEN 0 AND 1, false)
       OR coalesce(dq90_count_rate NOT BETWEEN 0 AND 1, false)
       OR coalesce(dq30_balance_rate NOT BETWEEN 0 AND 1, false)
       OR coalesce(dq60_balance_rate NOT BETWEEN 0 AND 1, false)
       OR coalesce(dq90_balance_rate NOT BETWEEN 0 AND 1, false)
),
rate_null_semantics_failures AS (
    SELECT count(*) AS failures
    FROM (
        SELECT as_of_month::VARCHAR AS row_key
        FROM mart.portfolio_monthly
        WHERE (delinquency_eligible_loan_count > 0 AND (
                  dq30_count_rate IS NULL OR dq60_count_rate IS NULL OR dq90_count_rate IS NULL
              ))
           OR (delinquency_eligible_loan_count = 0 AND (
                  dq30_count_rate IS NOT NULL OR dq60_count_rate IS NOT NULL OR dq90_count_rate IS NOT NULL
              ))
           OR (delinquency_eligible_upb > 0 AND (
                  dq30_balance_rate IS NULL OR dq60_balance_rate IS NULL OR dq90_balance_rate IS NULL
              ))
           OR (delinquency_eligible_upb = 0 AND (
                  dq30_balance_rate IS NOT NULL OR dq60_balance_rate IS NOT NULL OR dq90_balance_rate IS NOT NULL
              ))
        UNION ALL
        SELECT as_of_month::VARCHAR || '|' || segment_name || '|' || segment_value
        FROM mart.portfolio_monthly_by_segment
        WHERE (delinquency_eligible_loan_count > 0 AND (
                  dq30_count_rate IS NULL OR dq60_count_rate IS NULL OR dq90_count_rate IS NULL
              ))
           OR (delinquency_eligible_loan_count = 0 AND (
                  dq30_count_rate IS NOT NULL OR dq60_count_rate IS NOT NULL OR dq90_count_rate IS NOT NULL
              ))
           OR (delinquency_eligible_upb > 0 AND (
                  dq30_balance_rate IS NULL OR dq60_balance_rate IS NULL OR dq90_balance_rate IS NULL
              ))
           OR (delinquency_eligible_upb = 0 AND (
                  dq30_balance_rate IS NOT NULL OR dq60_balance_rate IS NOT NULL OR dq90_balance_rate IS NOT NULL
              ))
    )
),
expected_months AS (
    SELECT month_value::DATE AS as_of_month
    FROM generate_series(
        (SELECT min(as_of_month) FROM curated.fact_loan_month),
        (SELECT max(as_of_month) FROM curated.fact_loan_month),
        INTERVAL '1 month'
    ) AS generated(month_value)
),
required_dimensions(segment_name) AS (
    VALUES ('origination_year'), ('classic_fico'), ('original_cltv'), ('original_dti')
),
expected_month_dimensions AS (
    SELECT m.as_of_month, d.segment_name
    FROM expected_months AS m
    CROSS JOIN required_dimensions AS d
),
segment_rollup AS (
    SELECT
        s.as_of_month,
        s.segment_name,
        sum(s.on_book_loan_count) AS on_book_loan_count,
        sum(s.on_book_upb) AS on_book_upb,
        sum(s.delinquency_eligible_loan_count) AS eligible_count,
        sum(s.delinquency_eligible_upb) AS eligible_upb,
        sum(s.dq30_loan_count) AS dq30_count,
        sum(s.dq60_loan_count) AS dq60_count,
        sum(s.dq90_loan_count) AS dq90_count,
        sum(s.dq30_upb) AS dq30_upb,
        sum(s.dq60_upb) AS dq60_upb,
        sum(s.dq90_upb) AS dq90_upb
    FROM mart.portfolio_monthly_by_segment AS s
    GROUP BY s.as_of_month, s.segment_name
),
portfolio_month_completeness AS (
    SELECT count(*) AS failures
    FROM (
        SELECT e.as_of_month
        FROM expected_months AS e
        LEFT JOIN mart.portfolio_monthly AS a USING (as_of_month)
        WHERE a.as_of_month IS NULL
        UNION ALL
        SELECT a.as_of_month
        FROM mart.portfolio_monthly AS a
        LEFT JOIN expected_months AS e USING (as_of_month)
        WHERE e.as_of_month IS NULL
    )
),
segment_dimension_completeness AS (
    SELECT count(*) AS failures
    FROM expected_month_dimensions AS e
    LEFT JOIN segment_rollup AS s
      ON s.as_of_month = e.as_of_month
     AND s.segment_name = e.segment_name
    WHERE s.as_of_month IS NULL
),
expected_segment_keys AS (
    SELECT DISTINCT as_of_month, 'origination_year' AS segment_name,
           origination_year::VARCHAR AS segment_value
    FROM curated.fact_loan_month
    UNION ALL
    SELECT DISTINCT as_of_month, 'classic_fico', fico_band
    FROM curated.fact_loan_month
    UNION ALL
    SELECT DISTINCT as_of_month, 'original_cltv', cltv_band
    FROM curated.fact_loan_month
    UNION ALL
    SELECT DISTINCT as_of_month, 'original_dti', dti_band
    FROM curated.fact_loan_month
),
segment_key_completeness AS (
    SELECT count(*) AS failures
    FROM (
        SELECT e.as_of_month, e.segment_name, e.segment_value
        FROM expected_segment_keys AS e
        LEFT JOIN mart.portfolio_monthly_by_segment AS a
          USING (as_of_month, segment_name, segment_value)
        WHERE a.as_of_month IS NULL
        UNION ALL
        SELECT a.as_of_month, a.segment_name, a.segment_value
        FROM mart.portfolio_monthly_by_segment AS a
        LEFT JOIN expected_segment_keys AS e
          USING (as_of_month, segment_name, segment_value)
        WHERE e.as_of_month IS NULL
    )
),
segment_reconciliation AS (
    SELECT count(*) AS failures
    FROM expected_month_dimensions AS e
    JOIN mart.portfolio_monthly AS p USING (as_of_month)
    LEFT JOIN segment_rollup AS s
      ON s.as_of_month = e.as_of_month
     AND s.segment_name = e.segment_name
    WHERE s.as_of_month IS NULL
       OR s.on_book_loan_count <> p.on_book_loan_count
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
transition_rate_reconciliation AS (
    SELECT count(*) AS failures
    FROM (
        SELECT from_month, segment_name, segment_value, from_state,
               abs(sum(transition_count_rate) - 1.0) AS count_delta,
               abs(sum(transition_balance_rate) - 1.0) AS balance_delta
        FROM mart.delinquency_transition_monthly
        GROUP BY from_month, segment_name, segment_value, from_state
    )
    WHERE count_delta > 1e-10 OR balance_delta > 1e-10
),
tests AS (
    SELECT 'origination_2019_row_count' AS test_name, 'ERROR' AS severity,
           CASE WHEN coalesce((SELECT row_count FROM orig_counts WHERE source_year = 2019), 0) = 50000 THEN 0 ELSE 1 END::BIGINT AS failed_row_count,
           'Expected the verified Release 47 sample baseline of 50,000 rows.' AS details
    UNION ALL
    SELECT 'origination_2020_row_count', 'ERROR',
           CASE WHEN coalesce((SELECT row_count FROM orig_counts WHERE source_year = 2020), 0) = 50000 THEN 0 ELSE 1 END,
           'Expected the verified Release 47 sample baseline of 50,000 rows.'
    UNION ALL
    SELECT 'performance_2019_row_count', 'ERROR',
           CASE WHEN coalesce((SELECT row_count FROM perf_counts WHERE source_year = 2019), 0) = 1934614 THEN 0 ELSE 1 END,
           'Expected the verified Release 47 sample baseline of 1,934,614 rows.'
    UNION ALL
    SELECT 'performance_2020_row_count', 'ERROR',
           CASE WHEN coalesce((SELECT row_count FROM perf_counts WHERE source_year = 2020), 0) = 2517857 THEN 0 ELSE 1 END,
           'Expected the verified Release 47 sample baseline of 2,517,857 rows.'
    UNION ALL
    SELECT 'origination_primary_key_unique', 'ERROR', failures,
           'loan_id must be non-null and unique across both vintages.' FROM orig_duplicates
    UNION ALL
    SELECT 'performance_primary_key_unique', 'ERROR', failures,
           '(loan_id, as_of_month) must be non-null and unique.' FROM perf_duplicates
    UNION ALL
    SELECT 'origination_primary_key_not_null', 'ERROR', count(*),
           'loan_id and origination_year are required.'
    FROM staging.origination WHERE loan_id IS NULL OR origination_year IS NULL
    UNION ALL
    SELECT 'performance_primary_key_not_null', 'ERROR', count(*),
           'loan_id and a valid YYYYMM period are required.'
    FROM staging.monthly_performance WHERE loan_id IS NULL OR as_of_month IS NULL
    UNION ALL
    SELECT 'performance_to_origination_fk', 'ERROR', count(*),
           'Every performance loan must match exactly one origination record.'
    FROM staging.monthly_performance p
    LEFT JOIN staging.origination o USING (loan_id)
    WHERE o.loan_id IS NULL
    UNION ALL
    SELECT 'origination_to_performance_coverage', 'ERROR', count(*),
           'Every sampled origination loan must have at least one performance row.'
    FROM staging.origination o
    LEFT JOIN (SELECT DISTINCT loan_id FROM staging.monthly_performance) p USING (loan_id)
    WHERE p.loan_id IS NULL
    UNION ALL
    SELECT 'source_year_matches_loan_vintage', 'ERROR', count(*),
           'Directory year, loan identifier vintage, and joined vintage must agree.'
    FROM staging.monthly_performance p
    JOIN staging.origination o USING (loan_id)
    WHERE p.source_year <> o.origination_year OR o.source_year <> o.origination_year
    UNION ALL
    SELECT 'origination_required_dates_valid', 'ERROR', count(*),
           'Nonblank First Payment Date and Maturity Date must parse as YYYYMM.'
    FROM staging.origination
    WHERE first_payment_month IS NULL OR maturity_month IS NULL
    UNION ALL
    SELECT 'origination_risk_fields_valid', 'ERROR', count(*),
           'Classic FICO, CLTV, DTI, and Original UPB must satisfy the verified Release 47 ranges or field sentinel.'
    FROM staging.origination
    WHERE (classic_fico_raw <> '9999' AND (classic_fico IS NULL OR classic_fico NOT BETWEEN 300 AND 850))
       OR (original_cltv_raw <> '999' AND (original_cltv IS NULL OR original_cltv NOT BETWEEN 1 AND 998))
       OR (original_dti_raw <> '999' AND (original_dti IS NULL OR original_dti NOT BETWEEN 0 AND 65))
       OR original_upb IS NULL OR original_upb <= 0
    UNION ALL
    SELECT 'field_sentinel_normalization_consistent', 'ERROR', count(*),
           'Confirmed FICO/CLTV/DTI/Vantage sentinel codes must normalize to NULL, while raw values remain preserved.'
    FROM staging.origination
    WHERE (classic_fico_raw = '9999' AND classic_fico IS NOT NULL)
       OR (original_cltv_raw = '999' AND original_cltv IS NOT NULL)
       OR (original_dti_raw = '999' AND original_dti IS NOT NULL)
       OR (vantage_score_4_raw = '9999' AND vantage_score_4 IS NOT NULL)
    UNION ALL
    SELECT 'performance_optional_dates_valid', 'ERROR', count(*),
           'Every populated performance date must parse as YYYYMM.'
    FROM staging.monthly_performance
    WHERE (defect_settlement_date_raw IS NOT NULL AND defect_settlement_month IS NULL)
       OR (zero_balance_effective_date_raw IS NOT NULL AND zero_balance_effective_month IS NULL)
       OR (ddlpi_raw IS NOT NULL AND ddlpi_month IS NULL)
    UNION ALL
    SELECT 'loan_age_required_nonnegative', 'ERROR', count(*),
           'Loan Age is a required Release 47 numeric field and must parse to a non-negative value.'
    FROM staging.monthly_performance
    WHERE loan_age_raw IS NULL OR trim(loan_age_raw) = ''
       OR loan_age IS NULL OR loan_age < 0
    UNION ALL
    SELECT 'loan_age_monitoring_state_consistent', 'ERROR', count(*),
           'PRE_DUE is reserved for non-terminal Loan Age 0 records with status 00; explicit delinquency codes retain their delinquency state.'
    FROM staging.monthly_performance
    WHERE monitoring_state IS DISTINCT FROM CASE
        WHEN zero_balance_code IS NOT NULL THEN 'TERMINATED:' || zero_balance_code
        WHEN loan_age IS NULL OR loan_age < 0 THEN 'UNKNOWN_LOAN_AGE'
        WHEN loan_age = 0 AND delinquency_months = 0 THEN 'PRE_DUE'
        WHEN delinquency_months = 0 THEN 'CURRENT'
        WHEN delinquency_months = 1 THEN '30'
        WHEN delinquency_months = 2 THEN '60'
        WHEN delinquency_months >= 3 THEN '90+'
        WHEN delinquency_status_raw = 'RA' THEN 'REO_ACQUISITION'
        WHEN delinquency_status_raw = 'XX' OR delinquency_status_raw IS NULL THEN 'UNKNOWN'
        ELSE 'UNKNOWN'
    END
    UNION ALL
    SELECT 'delinquency_code_valid', 'ERROR', count(*),
           'Allowed values are 00-99, RA, and XX under Release 47.'
    FROM staging.monthly_performance
    WHERE delinquency_status_raw IS NULL
       OR NOT (regexp_full_match(delinquency_status_raw, '^[0-9]{2}$') OR delinquency_status_raw IN ('RA', 'XX'))
    UNION ALL
    SELECT 'zero_balance_code_valid', 'ERROR', count(*),
           'Allowed Release 47 Zero Balance Codes: 01,02,03,09,15,16,96.'
    FROM staging.monthly_performance
    WHERE zero_balance_code IS NOT NULL
      AND zero_balance_code NOT IN ('01', '02', '03', '09', '15', '16', '96')
    UNION ALL
    SELECT 'zero_balance_code_date_pairing', 'ERROR', count(*),
           'Zero Balance Code and Zero Balance Effective Date must be populated together.'
    FROM staging.monthly_performance
    WHERE (zero_balance_code IS NULL) <> (zero_balance_effective_month IS NULL)
    UNION ALL
    SELECT 'current_actual_upb_valid', 'ERROR', count(*),
           'Current Actual UPB must be present and non-negative.'
    FROM staging.monthly_performance
    WHERE current_actual_upb IS NULL OR current_actual_upb < 0
    UNION ALL
    SELECT 'performance_period_within_verified_release', 'ERROR', count(*),
           'Performance months cannot predate the sample vintage or exceed the Release 47 cutoff of 2026-03.'
    FROM staging.monthly_performance
    WHERE as_of_month < make_date(source_year, 1, 1)
       OR as_of_month > DATE '2026-03-01'
    UNION ALL
    SELECT 'terminal_record_upb_consistency', 'ERROR', count(*),
           'In this verified sample, zero UPB and Zero Balance Code identify the same terminal rows.'
    FROM staging.monthly_performance
    WHERE (current_actual_upb = 0) <> (zero_balance_code IS NOT NULL)
    UNION ALL
    SELECT 'no_records_after_terminal', 'ERROR', failures,
           'No monthly row may occur after a loan terminal record.' FROM post_terminal
    UNION ALL
    SELECT 'loan_months_are_contiguous', 'ERROR', failures,
           'Observed consecutive rows must be exact adjacent natural months.' FROM month_gaps
    UNION ALL
    SELECT 'staging_curated_origination_reconciles', 'ERROR',
           abs((SELECT count(*) FROM staging.origination) - (SELECT count(*) FROM curated.dim_loan)),
           'Staging origination and curated loan dimension row counts must match.'
    UNION ALL
    SELECT 'staging_curated_performance_reconciles', 'ERROR',
           abs((SELECT count(*) FROM staging.monthly_performance) - (SELECT count(*) FROM curated.fact_loan_month)),
           'Staging and curated performance row counts must match.'
    UNION ALL
    SELECT 'metric_numerators_within_denominators', 'ERROR', failures,
           'Counts and UPB numerators must be nested and cannot exceed denominators.' FROM metric_failures
    UNION ALL
    SELECT 'metric_rates_between_zero_and_one', 'ERROR', failures,
           'All defined delinquency rates must fall in [0,1].' FROM rate_failures
    UNION ALL
    SELECT 'metric_rate_null_semantics', 'ERROR', failures,
           '30+/60+/90+ rates must be non-NULL when their denominator is positive and NULL only when it is zero.'
    FROM rate_null_semantics_failures
    UNION ALL
    SELECT 'portfolio_month_keys_complete', 'ERROR', failures,
           'The overall mart must contain exactly one row for every observed fact month and no extra months.'
    FROM portfolio_month_completeness
    UNION ALL
    SELECT 'segment_dimension_months_complete', 'ERROR', failures,
           'Every observed month must contain all four required segment dimensions.'
    FROM segment_dimension_completeness
    UNION ALL
    SELECT 'segment_keys_complete', 'ERROR', failures,
           'The segment mart must contain every month/dimension/value key observed in the loan-month fact and no extras.'
    FROM segment_key_completeness
    UNION ALL
    SELECT 'segment_totals_reconcile_to_portfolio', 'ERROR', failures,
           'Every segment dimension must sum exactly to the overall monthly result.' FROM segment_reconciliation
    UNION ALL
    SELECT 'independent_sample_metrics_reconcile', 'ERROR', count(*),
           'Staging-only recomputation for 2020-06, 2021-06, and 2026-03 must match every mart count and UPB component.'
    FROM quality.metric_reconciliation_samples
    WHERE NOT coalesce(reconciled, false)
    UNION ALL
    SELECT 'transition_pairs_exact_natural_month', 'ERROR', count(*),
           'Every transition pair must be exactly one natural month apart.'
    FROM curated.v_loan_month_transition
    WHERE to_month <> from_month + INTERVAL '1 month'
    UNION ALL
    SELECT 'transition_source_is_on_book', 'ERROR', count(*),
           'A transition source cannot be a terminal or zero-UPB record.'
    FROM curated.v_loan_month_transition
    WHERE from_upb <= 0 OR from_state LIKE 'TERMINATED:%'
    UNION ALL
    SELECT 'transition_rates_reconcile', 'ERROR', failures,
           'For each observable from-state, transition shares must sum to one.' FROM transition_rate_reconciliation
    UNION ALL
    SELECT 'bankruptcy_cramdown_nonterminal_nonzero', 'ERROR', count(*),
           'Non-terminal Bankruptcy Cramdown Costs must not contain nonzero values.'
    FROM staging.monthly_performance
    WHERE zero_balance_code IS NULL AND bankruptcy_cramdown_costs <> 0
    UNION ALL
    SELECT 'bankruptcy_cramdown_zero_populated_nonterminal', 'WARNING', count(*),
           'Known source observation: populated 0.00 values appear on non-terminal rows; preserved without imputation.'
    FROM staging.monthly_performance
    WHERE dq_cramdown_populated_non_terminal
    UNION ALL
    SELECT 'first_payment_more_than_one_year_after_vintage', 'WARNING', count(*),
           'Known source observation: some First Payment Dates occur more than one calendar year after the loan-ID vintage; vintage is not re-derived from this date.'
    FROM staging.origination
    WHERE year(first_payment_month) > origination_year + 1
    UNION ALL
    SELECT 'loan_age_zero_with_delinquent_status', 'WARNING', count(*),
           'Observed valid Loan Age 0 records with status 01 are excluded by the age denominator rule but retain state 30 instead of PRE_DUE.'
    FROM staging.monthly_performance
    WHERE loan_age = 0 AND delinquency_months > 0
)
SELECT
    test_name,
    severity,
    failed_row_count,
    failed_row_count = 0 AS passed,
    details,
    current_timestamp AS checked_at
FROM tests;

CREATE UNIQUE INDEX idx_dq_test_name ON quality.dq_test_results(test_name);
