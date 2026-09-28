-- Field order is locked to Freddie Mac SFLLD Release 47 (July 2026).
-- The external views read all fields as VARCHAR so DuckDB cannot infer a
-- different schema. Typed staging keeps raw status/sentinel fields needed for
-- audit and normalizes only field-specific missing-value codes.

CREATE VIEW raw_external.v_origination AS
SELECT
    regexp_extract(filename, 'sample_(\d{4})', 1)::SMALLINT AS source_year,
    filename AS source_file,
    * EXCLUDE (filename)
FROM read_csv(
    '{{RAW_DIR}}/sample_*/sample_orig_*.txt',
    delim = '|',
    header = false,
    auto_detect = false,
    strict_mode = true,
    filename = true,
    columns = {
        'classic_fico': 'VARCHAR',
        'first_payment_date': 'VARCHAR',
        'first_time_homebuyer_indicator': 'VARCHAR',
        'maturity_date': 'VARCHAR',
        'msa_or_metropolitan_division': 'VARCHAR',
        'mi_percentage': 'VARCHAR',
        'number_of_units': 'VARCHAR',
        'occupancy_status': 'VARCHAR',
        'original_cltv': 'VARCHAR',
        'original_dti': 'VARCHAR',
        'original_upb': 'VARCHAR',
        'original_ltv': 'VARCHAR',
        'original_interest_rate': 'VARCHAR',
        'channel': 'VARCHAR',
        'prepayment_penalty_indicator': 'VARCHAR',
        'amortization_type': 'VARCHAR',
        'property_state': 'VARCHAR',
        'property_type': 'VARCHAR',
        'postal_code': 'VARCHAR',
        'loan_id': 'VARCHAR',
        'loan_purpose': 'VARCHAR',
        'original_loan_term': 'VARCHAR',
        'number_of_borrowers': 'VARCHAR',
        'seller_name': 'VARCHAR',
        'super_conforming_flag': 'VARCHAR',
        'pre_harp_loan_id': 'VARCHAR',
        'special_eligibility_program': 'VARCHAR',
        'harp_indicator': 'VARCHAR',
        'property_valuation_method': 'VARCHAR',
        'interest_only_indicator': 'VARCHAR',
        'vantage_score_4': 'VARCHAR'
    }
);

CREATE VIEW raw_external.v_monthly_performance AS
SELECT
    regexp_extract(filename, 'sample_(\d{4})', 1)::SMALLINT AS source_year,
    filename AS source_file,
    * EXCLUDE (filename)
FROM read_csv(
    '{{RAW_DIR}}/sample_*/sample_perf_*.txt',
    delim = '|',
    header = false,
    auto_detect = false,
    strict_mode = true,
    filename = true,
    columns = {
        'loan_id': 'VARCHAR',
        'period': 'VARCHAR',
        'current_actual_upb': 'VARCHAR',
        'current_loan_delinquency_status': 'VARCHAR',
        'loan_age': 'VARCHAR',
        'remaining_months_to_legal_maturity': 'VARCHAR',
        'defect_settlement_date': 'VARCHAR',
        'modification_flag': 'VARCHAR',
        'zero_balance_code': 'VARCHAR',
        'zero_balance_effective_date': 'VARCHAR',
        'current_interest_rate': 'VARCHAR',
        'current_non_interest_bearing_upb': 'VARCHAR',
        'ddlpi': 'VARCHAR',
        'mi_recoveries': 'VARCHAR',
        'net_sales_proceeds': 'VARCHAR',
        'non_mi_recoveries': 'VARCHAR',
        'total_expenses': 'VARCHAR',
        'legal_costs': 'VARCHAR',
        'maintenance_and_preservation_costs': 'VARCHAR',
        'taxes_and_insurance': 'VARCHAR',
        'miscellaneous_expenses': 'VARCHAR',
        'actual_loss': 'VARCHAR',
        'cumulative_modification_costs': 'VARCHAR',
        'interest_rate_step_indicator': 'VARCHAR',
        'payment_deferral_flag': 'VARCHAR',
        'estimated_ltv': 'VARCHAR',
        'zero_balance_removal_upb': 'VARCHAR',
        'delinquent_accrued_interest': 'VARCHAR',
        'delinquency_due_to_disaster': 'VARCHAR',
        'borrower_assistance_plan': 'VARCHAR',
        'current_period_modification_costs': 'VARCHAR',
        'current_interest_bearing_upb': 'VARCHAR',
        'mi_cancellation_indicator': 'VARCHAR',
        'servicer_name': 'VARCHAR',
        'bankruptcy_cramdown_costs': 'VARCHAR'
    }
);

CREATE TABLE staging.origination AS
SELECT
    source_year,
    source_file,
    loan_id,
    CAST(('20' || substr(loan_id, 2, 2)) AS SMALLINT) AS origination_year,
    substr(loan_id, 4, 2) AS origination_quarter,
    classic_fico AS classic_fico_raw,
    CASE WHEN classic_fico IN ('', '9999') OR classic_fico IS NULL
         THEN NULL ELSE try_cast(classic_fico AS SMALLINT) END AS classic_fico,
    CASE WHEN classic_fico = '9999' THEN 'NOT_AVAILABLE'
         WHEN classic_fico IS NULL OR classic_fico = '' THEN 'BLANK'
         ELSE NULL END AS classic_fico_missing_reason,
    first_payment_date AS first_payment_date_raw,
    try_strptime(first_payment_date, '%Y%m')::DATE AS first_payment_month,
    NULLIF(first_time_homebuyer_indicator, '') AS first_time_homebuyer_indicator,
    maturity_date AS maturity_date_raw,
    try_strptime(maturity_date, '%Y%m')::DATE AS maturity_month,
    try_cast(NULLIF(msa_or_metropolitan_division, '') AS INTEGER) AS msa_or_metropolitan_division,
    mi_percentage AS mi_percentage_raw,
    CASE WHEN mi_percentage IN ('', '999') OR mi_percentage IS NULL
         THEN NULL ELSE try_cast(mi_percentage AS SMALLINT) END AS mi_percentage,
    try_cast(NULLIF(number_of_units, '') AS SMALLINT) AS number_of_units,
    NULLIF(occupancy_status, '') AS occupancy_status,
    original_cltv AS original_cltv_raw,
    CASE WHEN original_cltv IN ('', '999') OR original_cltv IS NULL
         THEN NULL ELSE try_cast(original_cltv AS SMALLINT) END AS original_cltv,
    original_dti AS original_dti_raw,
    CASE WHEN original_dti IN ('', '999') OR original_dti IS NULL
         THEN NULL ELSE try_cast(original_dti AS SMALLINT) END AS original_dti,
    try_cast(NULLIF(original_upb, '') AS DECIMAL(18,2)) AS original_upb,
    original_ltv AS original_ltv_raw,
    CASE WHEN original_ltv IN ('', '999') OR original_ltv IS NULL
         THEN NULL ELSE try_cast(original_ltv AS SMALLINT) END AS original_ltv,
    try_cast(NULLIF(original_interest_rate, '') AS DECIMAL(9,3)) AS original_interest_rate,
    NULLIF(channel, '') AS channel,
    NULLIF(prepayment_penalty_indicator, '') AS prepayment_penalty_indicator,
    NULLIF(amortization_type, '') AS amortization_type,
    NULLIF(property_state, '') AS property_state,
    NULLIF(property_type, '') AS property_type,
    NULLIF(postal_code, '') AS postal_code,
    NULLIF(loan_purpose, '') AS loan_purpose,
    try_cast(NULLIF(original_loan_term, '') AS SMALLINT) AS original_loan_term,
    try_cast(NULLIF(number_of_borrowers, '') AS SMALLINT) AS number_of_borrowers,
    NULLIF(seller_name, '') AS seller_name,
    NULLIF(super_conforming_flag, '') AS super_conforming_flag,
    NULLIF(pre_harp_loan_id, '') AS pre_harp_loan_id,
    NULLIF(special_eligibility_program, '') AS special_eligibility_program,
    NULLIF(harp_indicator, '') AS harp_indicator,
    NULLIF(property_valuation_method, '') AS property_valuation_method,
    NULLIF(interest_only_indicator, '') AS interest_only_indicator,
    vantage_score_4 AS vantage_score_4_raw,
    CASE WHEN vantage_score_4 IN ('', '9999') OR vantage_score_4 IS NULL
         THEN NULL ELSE try_cast(vantage_score_4 AS SMALLINT) END AS vantage_score_4,
    CASE WHEN vantage_score_4 = '9999' THEN 'NOT_AVAILABLE'
         WHEN vantage_score_4 IS NULL OR vantage_score_4 = '' THEN 'BLANK'
         ELSE NULL END AS vantage_score_4_missing_reason
FROM raw_external.v_origination;

CREATE TABLE staging.monthly_performance AS
WITH typed AS (
    SELECT
        source_year,
        source_file,
        loan_id,
        period AS period_raw,
        try_strptime(period, '%Y%m')::DATE AS as_of_month,
        try_cast(NULLIF(current_actual_upb, '') AS DECIMAL(18,2)) AS current_actual_upb,
        NULLIF(current_loan_delinquency_status, '') AS delinquency_status_raw,
        CASE WHEN regexp_full_match(current_loan_delinquency_status, '^[0-9]{2}$')
             THEN try_cast(current_loan_delinquency_status AS SMALLINT) ELSE NULL END AS delinquency_months,
        loan_age AS loan_age_raw,
        try_cast(NULLIF(loan_age, '') AS SMALLINT) AS loan_age,
        try_cast(NULLIF(remaining_months_to_legal_maturity, '') AS SMALLINT) AS remaining_months_to_legal_maturity,
        defect_settlement_date AS defect_settlement_date_raw,
        try_strptime(NULLIF(defect_settlement_date, ''), '%Y%m')::DATE AS defect_settlement_month,
        NULLIF(modification_flag, '') AS modification_flag,
        NULLIF(zero_balance_code, '') AS zero_balance_code,
        zero_balance_effective_date AS zero_balance_effective_date_raw,
        try_strptime(NULLIF(zero_balance_effective_date, ''), '%Y%m')::DATE AS zero_balance_effective_month,
        try_cast(NULLIF(current_interest_rate, '') AS DECIMAL(9,3)) AS current_interest_rate,
        try_cast(NULLIF(current_non_interest_bearing_upb, '') AS DECIMAL(18,2)) AS current_non_interest_bearing_upb,
        ddlpi AS ddlpi_raw,
        try_strptime(NULLIF(ddlpi, ''), '%Y%m')::DATE AS ddlpi_month,
        try_cast(NULLIF(mi_recoveries, '') AS DECIMAL(18,2)) AS mi_recoveries,
        try_cast(NULLIF(net_sales_proceeds, '') AS DECIMAL(18,2)) AS net_sales_proceeds,
        try_cast(NULLIF(non_mi_recoveries, '') AS DECIMAL(18,2)) AS non_mi_recoveries,
        try_cast(NULLIF(total_expenses, '') AS DECIMAL(18,2)) AS total_expenses,
        try_cast(NULLIF(legal_costs, '') AS DECIMAL(18,2)) AS legal_costs,
        try_cast(NULLIF(maintenance_and_preservation_costs, '') AS DECIMAL(18,2)) AS maintenance_and_preservation_costs,
        try_cast(NULLIF(taxes_and_insurance, '') AS DECIMAL(18,2)) AS taxes_and_insurance,
        try_cast(NULLIF(miscellaneous_expenses, '') AS DECIMAL(18,2)) AS miscellaneous_expenses,
        try_cast(NULLIF(actual_loss, '') AS DECIMAL(18,2)) AS actual_loss,
        try_cast(NULLIF(cumulative_modification_costs, '') AS DECIMAL(18,2)) AS cumulative_modification_costs,
        NULLIF(interest_rate_step_indicator, '') AS interest_rate_step_indicator,
        NULLIF(payment_deferral_flag, '') AS payment_deferral_flag,
        try_cast(NULLIF(estimated_ltv, '') AS DECIMAL(12,3)) AS estimated_ltv,
        try_cast(NULLIF(zero_balance_removal_upb, '') AS DECIMAL(18,2)) AS zero_balance_removal_upb,
        try_cast(NULLIF(delinquent_accrued_interest, '') AS DECIMAL(18,2)) AS delinquent_accrued_interest,
        NULLIF(delinquency_due_to_disaster, '') AS delinquency_due_to_disaster,
        NULLIF(borrower_assistance_plan, '') AS borrower_assistance_plan,
        try_cast(NULLIF(current_period_modification_costs, '') AS DECIMAL(18,2)) AS current_period_modification_costs,
        try_cast(NULLIF(current_interest_bearing_upb, '') AS DECIMAL(18,2)) AS current_interest_bearing_upb,
        NULLIF(mi_cancellation_indicator, '') AS mi_cancellation_indicator,
        NULLIF(servicer_name, '') AS servicer_name,
        try_cast(NULLIF(bankruptcy_cramdown_costs, '') AS DECIMAL(18,2)) AS bankruptcy_cramdown_costs
    FROM raw_external.v_monthly_performance
)
SELECT
    *,
    zero_balance_code IS NOT NULL AS is_terminal_record,
    current_actual_upb > 0 AND zero_balance_code IS NULL AS is_on_book_eom,
    delinquency_months IS NOT NULL AS is_numeric_delinquency_status,
    loan_age IS NOT NULL AND loan_age >= 0 AS is_valid_loan_age,
    current_actual_upb > 0
        AND zero_balance_code IS NULL
        AND loan_age >= 1
        AND delinquency_months IS NOT NULL AS is_delinquency_eligible,
    CASE
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
    END AS monitoring_state,
    bankruptcy_cramdown_costs IS NOT NULL
        AND zero_balance_code IS NULL AS dq_cramdown_populated_non_terminal
FROM typed;

CREATE INDEX idx_stg_origination_loan ON staging.origination(loan_id);
CREATE INDEX idx_stg_performance_loan_month ON staging.monthly_performance(loan_id, as_of_month);
