-- Curated loan dimension and monitoring facts. These tables keep the analysis
-- grain explicit and are intentionally narrower than the typed staging layer.

CREATE TABLE curated.dim_loan AS
SELECT
    loan_id,
    origination_year,
    origination_quarter,
    first_payment_month,
    maturity_month,
    classic_fico,
    CASE
        WHEN classic_fico IS NULL THEN 'Unknown'
        WHEN classic_fico < 620 THEN '<620'
        WHEN classic_fico <= 659 THEN '620-659'
        WHEN classic_fico <= 699 THEN '660-699'
        WHEN classic_fico <= 739 THEN '700-739'
        WHEN classic_fico <= 779 THEN '740-779'
        ELSE '780+'
    END AS fico_band,
    original_cltv,
    CASE
        WHEN original_cltv IS NULL THEN 'Unknown'
        WHEN original_cltv <= 60 THEN '<=60'
        WHEN original_cltv <= 70 THEN '61-70'
        WHEN original_cltv <= 80 THEN '71-80'
        WHEN original_cltv <= 90 THEN '81-90'
        WHEN original_cltv <= 95 THEN '91-95'
        ELSE '>95'
    END AS cltv_band,
    original_ltv,
    original_dti,
    CASE
        WHEN original_dti IS NULL THEN 'Unknown'
        WHEN original_dti <= 20 THEN '<=20'
        WHEN original_dti <= 30 THEN '21-30'
        WHEN original_dti <= 40 THEN '31-40'
        WHEN original_dti <= 45 THEN '41-45'
        WHEN original_dti <= 50 THEN '46-50'
        ELSE '>50'
    END AS dti_band,
    original_upb,
    original_interest_rate,
    original_loan_term,
    property_state,
    property_type,
    occupancy_status,
    loan_purpose,
    channel,
    source_year,
    source_file
FROM staging.origination;

CREATE UNIQUE INDEX idx_dim_loan_pk ON curated.dim_loan(loan_id);

CREATE TABLE curated.fact_loan_month AS
SELECT
    p.loan_id,
    p.as_of_month,
    d.origination_year,
    d.origination_quarter,
    d.fico_band,
    d.cltv_band,
    d.dti_band,
    p.current_actual_upb,
    p.delinquency_status_raw,
    p.delinquency_months,
    p.loan_age,
    p.zero_balance_code,
    p.zero_balance_effective_month,
    p.is_terminal_record,
    p.is_on_book_eom,
    p.is_numeric_delinquency_status,
    p.is_delinquency_eligible,
    p.monitoring_state,
    p.source_year,
    p.source_file
FROM staging.monthly_performance AS p
JOIN curated.dim_loan AS d
  ON p.loan_id = d.loan_id;

CREATE UNIQUE INDEX idx_fact_loan_month_pk
    ON curated.fact_loan_month(loan_id, as_of_month);
CREATE INDEX idx_fact_loan_month_month
    ON curated.fact_loan_month(as_of_month);

-- A view avoids storing a second multi-million-row copy. It only pairs an
-- on-book source record with the exact following calendar month. The target
-- may be a terminal record so exits are visible; no row is synthesized later.
CREATE VIEW curated.v_loan_month_transition AS
SELECT
    current_row.loan_id,
    current_row.as_of_month AS from_month,
    next_row.as_of_month AS to_month,
    current_row.origination_year,
    current_row.fico_band,
    current_row.cltv_band,
    current_row.dti_band,
    current_row.monitoring_state AS from_state,
    next_row.monitoring_state AS to_state,
    current_row.current_actual_upb AS from_upb
FROM curated.fact_loan_month AS current_row
JOIN curated.fact_loan_month AS next_row
  ON next_row.loan_id = current_row.loan_id
 AND next_row.as_of_month = current_row.as_of_month + INTERVAL '1 month'
WHERE current_row.is_on_book_eom;
