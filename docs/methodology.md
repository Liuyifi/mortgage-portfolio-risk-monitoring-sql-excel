# Risk Methodology and Data Model

## 1. Portfolio scope and data boundary

The analysis uses Freddie Mac Single-Family Loan-Level Dataset Standard Dataset samples for the 2019 and 2020 origination vintages. Together they contain 100,000 sampled originations and 4,452,471 monthly performance records. Monthly observations extend from January 2019 through March 2026.

The source files, loan-level database, detailed validation output, aggregate exports, and generated workbook remain local. The repository publishes only the independently written code and methodology required to inspect the analytical design.

## 2. Data model and table grain

| Layer | Object | Grain | Role |
|---|---|---|---|
| Raw external | `raw_external.v_origination` | Source origination row | Reads the 31 source fields as text and retains source lineage |
| Raw external | `raw_external.v_monthly_performance` | Source performance row | Reads the 35 source fields as text and retains source lineage |
| Staging | `staging.origination` | One row per loan | Applies typed parsing, sentinel handling, vintage, and source checks |
| Staging | `staging.monthly_performance` | Loan × observation month | Stores dates, balances, official status codes, terminal events, and analytical flags |
| Curated | `curated.dim_loan` | One row per loan | Stores static vintage, FICO, CLTV, and DTI attributes and bands |
| Curated | `curated.fact_loan_month` | Loan × observation month | Provides the governed month-end population, eligibility, balance, and risk state |
| Curated | `curated.v_loan_month_transition` | Loan × beginning month | Joins observations exactly one calendar month apart |
| Mart | `mart.portfolio_monthly` | Observation month | Aggregates overall exposure, coverage, and delinquency measures |
| Mart | `mart.portfolio_monthly_by_segment` | Month × dimension × segment | Aggregates vintage, FICO, CLTV, and DTI views independently |
| Mart | `mart.delinquency_transition_monthly` | Month × scope × from-state × to-state | Aggregates observable adjacent-month migrations |

`origination_year` is derived from the Loan Identifier vintage and cross-checked against the source-file year. It is not inferred from First Payment Date.

## 3. On-book and delinquency-eligible populations

A record is **on book** at month end when `current_actual_upb > 0` and `zero_balance_code IS NULL`. A zero-balance terminal record is retained for event and transition analysis but is not part of month-end exposure. No rows are generated or forward-filled after termination.

The **delinquency-eligible** population is the subset of on-book observations with `loan_age >= 1` and a numeric Current Loan Delinquency Status from `00` through `99`. Loan Age 0, `RA`, `XX`, missing or invalid status, and terminal rows are excluded from delinquency denominators. Coverage metrics make the difference between the on-book and eligible populations explicit.

## 4. Count-based and UPB-weighted delinquency metrics

Numeric status `00` is current or less than 30 days delinquent; `01` is 30–59 days, `02` is 60–89 days, and values of `03` or above enter the 90+ bucket.

| Metric family | Numerator | Denominator |
|---|---|---|
| 30+/60+/90+ count rate | Eligible loans at or above the applicable status threshold | Delinquency-eligible loan count |
| 30+/60+/90+ UPB rate | Current Actual UPB for eligible loans at or above the threshold | Current Actual UPB for all eligible loans |

Loan-count and UPB-weighted rates answer different questions and are never combined. The 90+ measure represents severe delinquency; it is not automatically equivalent to default. A default definition would require separate event criteria, such as specified zero-balance, disposition, or REO outcomes.

## 5. FICO, CLTV, DTI, and origination-vintage segmentation

Each segmentation dimension is aggregated independently:

- Origination vintage: `2019`, `2020`.
- Classic FICO: `<620`, `620-659`, `660-699`, `700-739`, `740-779`, `780+`, `Unknown`.
- Original CLTV: `<=60`, `61-70`, `71-80`, `81-90`, `91-95`, `>95`, `Unknown`.
- Original DTI: `<=20`, `21-30`, `31-40`, `41-45`, `46-50`, `>50`, `Unknown`.

Segment marts retain additive counts, balances, numerators, and denominators. Quality gates require every dimension's monthly totals to reconcile to the overall monthly mart.

## 6. Consecutive-calendar-month transition methodology

A transition exists only when the same loan has records exactly one natural calendar month apart:

```sql
next_row.loan_id = current_row.loan_id
AND next_row.as_of_month = current_row.as_of_month + INTERVAL '1 month'
```

The beginning observation must be on book. The destination may be current, delinquent, REO, unknown, or terminal. A terminal record cannot begin a later transition. A final observation without the next natural month is right-censored and excluded from the migration denominator.

Count-based transition rates use the number of loans in the same beginning-month and from-state cohort that remain observable in the next natural month. Balance-weighted transition rates use their beginning-month Current Actual UPB. Thus, the denominator includes only loans observable in the immediately following calendar month.

## 7. Missing values and official sentinel handling

Empty raw fields become SQL `NULL`; numeric fields are not globally imputed. Field-specific official sentinel values are handled during staging, including `9999` for Classic FICO and VantageScore 4.0 and `999` for Original CLTV, Original DTI, Original LTV, and MI Percentage. Raw values remain available in paired `*_raw` fields where implemented.

Missing, unparseable, or officially coded FICO, CLTV, and DTI values used for segmentation enter `Unknown`. They are neither removed nor assigned to the highest-risk band. Valid nonnumeric delinquency statuses such as `RA` and `XX` are preserved rather than coerced into numeric delinquency.

## 8. SQL quality gates and reconciliation

The DuckDB build runs inside one transaction. Any failed `ERROR`-severity check rolls back the build and returns a nonzero exit status. The accepted local build passed 39 of 39 blocking SQL checks.

Controls cover source row counts, key uniqueness, loan coverage across files, cross-vintage consistency, date and numeric ranges, Loan Age parsing, official status and zero-balance codes, terminal events, post-termination records, month continuity, staging-to-curated reconciliation, mart completeness, numerator-versus-denominator constraints, rate null semantics, segment additivity, and transition-rate reconciliation.

Selected benchmark months are recalculated directly from staging balances, termination codes, Loan Age, and raw delinquency status. This independent path does not reuse curated eligibility flags. The dashboard export layer applies 14 additional checks before workbook creation.

## 9. Excel presentation-layer design

The offline workbook contains three user-facing pages: `Portfolio Overview`, `Risk Segments`, and `Transitions`. Supporting sheets hold aggregate monthly, segment, transition, validation, and usage data. The completed workbook has eight worksheets and nine native Excel charts.

Displayed rates are formulas based on summed numerators and denominators under the active filters; stored rates are not added or arithmetically averaged. Small segment cells are grouped for display, while governed mart totals remain unchanged. Transition output is limited to aggregate cohorts. The workbook contains no Loan Identifiers, individual loan-month rows, VBA, Power Query, or external data connections.

## 10. Interpretation limitations

The 2019 and 2020 values denote origination cohorts, while monthly performance continues through March 2026. Cross-period comparisons therefore reflect different entry timing, seasoning, payoff and other exits, right-censoring, and changes in the surviving sample.

The two annual samples do not represent Freddie Mac's entire loan population or the U.S. mortgage market. Reported balances are sample balances, not a live institution's exposure. The analysis is descriptive rather than causal or predictive, and observed delinquency, transition, and segment patterns should not be generalized beyond the defined sample without additional validation.
