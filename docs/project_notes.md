# Project Notes

## 1. Data I used

I used the Freddie Mac Single-Family Loan-Level Dataset Standard Dataset samples for the 2019 and 2020 origination vintages. The four source files contain 100,000 origination rows and 4,452,471 monthly performance rows. Performance observations run from January 2019 through March 2026.

The origination files contain static attributes such as Classic FICO, Original CLTV, Original DTI and Original UPB. The performance files contain monthly balance, delinquency status, Loan Age and terminal-event fields. I read every field as text first, then parse the fields needed for this analysis in DuckDB. This avoids relying on automatic type inference for codes and sentinel values.

The main analytical grain is one row per `(loan_id, as_of_month)`. Origination attributes are joined to the monthly performance rows through `loan_id`. The 2019 and 2020 labels are origination vintages derived from the Loan Identifier and checked against the source-file year; they are not inferred from First Payment Date.

## 2. How I defined the portfolio

I treat a record as on book at month end when Current Actual UPB is greater than zero and Zero Balance Code is blank. A zero-balance row can still describe an important terminal event, so I keep it in the loan-month table and allow it to appear as a transition destination. It is not part of month-end on-book exposure.

The delinquency denominator is narrower than the on-book population. A record is delinquency eligible when it is on book, Loan Age is at least 1, and Current Loan Delinquency Status is a numeric code from `00` to `99`. This excludes Loan Age 0 records, `RA`, `XX`, missing or invalid status values, and terminal rows. I show eligible count and eligible UPB alongside on-book count and UPB so the excluded population is visible.

I do not generate rows after a loan terminates and do not forward-fill missing months. The analysis uses the observations present in the sample.

## 3. Delinquency metrics

Numeric delinquency status `00` is current or less than 30 days delinquent. Status `01` is 30–59 days, `02` is 60–89 days, and `03` or higher is included in 90+.

For each threshold I calculate two rates:

- count rate = delinquent eligible loans / all eligible loans
- UPB-weighted rate = delinquent eligible UPB / all eligible UPB

The numerators are nested: 90+ is part of 60+, and 60+ is part of 30+. I keep the underlying counts and balances in the mart and dashboard data sheets. Excel calculates filtered rates from these additive components, which avoids averaging percentages across groups.

Count and UPB rates answer different questions. The count rate describes the share of loans affected. The UPB rate describes the share of eligible balance affected. A small number of larger loans can therefore move the UPB rate more than the count rate.

I use 90+ as a severe-delinquency measure, not as an automatic default label. Default would need its own event definition, for example selected zero-balance, disposition or REO criteria.

## 4. Segmentation

I calculate monthly metrics independently for four dimensions:

- origination vintage: `2019`, `2020`
- Classic FICO: `<620`, `620-659`, `660-699`, `700-739`, `740-779`, `780+`, `Unknown`
- Original CLTV: `<=60`, `61-70`, `71-80`, `81-90`, `91-95`, `>95`, `Unknown`
- Original DTI: `<=20`, `21-30`, `31-40`, `41-45`, `46-50`, `>50`, `Unknown`

Freddie Mac sentinel values are handled field by field. For the fields used here, `9999` is treated as missing for Classic FICO and VantageScore 4.0, while `999` is treated as missing for Original CLTV, Original DTI, Original LTV and MI Percentage. I keep the raw code where useful and convert the analytical value to `NULL`. Missing or sentinel FICO, CLTV and DTI values go to `Unknown`; they are not assigned to the riskiest band.

For the local dashboard export, very small segment cells are combined into `Other / low volume`. This changes only the display grouping. The SQL mart retains the original segment rows, and the displayed additive totals still reconcile to the overall portfolio.

## 5. Monthly transitions

A transition is created only when the same loan has an observation in the exact next calendar month:

```sql
next_row.loan_id = current_row.loan_id
AND next_row.as_of_month = current_row.as_of_month + INTERVAL '1 month'
```

The beginning row must be on book. The next row may be current, 30, 60, 90+, REO, unknown or a terminal event. A terminal row cannot start another transition because it is not on book.

Transition count rates use observable loans within the same beginning month and from-state cohort. Balance rates use the same cohort's beginning-month UPB. If a loan has no observation in the next calendar month, its final observation is right-censored and does not enter the migration denominator. This is important because using the next available record after a gap would mix different time intervals.

Low-volume transition destinations are combined for the dashboard, and beginning-state cohorts below the display threshold are omitted. The DuckDB mart keeps the full aggregate transition table, including vintage cuts.

## 6. Checks I kept

The raw-data script checks that the four expected files exist, have the expected number of fields and rows, and contain usable key dates, numeric fields, delinquency statuses and zero-balance codes. It also checks duplicate origination IDs, loan coverage between the origination and performance files, important-field missingness, and the minimum and maximum performance months.

The SQL build checks duplicate keys, loan coverage, observation months, risk-field sentinels, major status codes, the on-book and eligible rules, count and UPB numerator hierarchy, segment completeness and reconciliation, and transition timing and cohort totals. Three selected months are recalculated directly from staging fields and compared with the monthly mart.

The export step is intentionally smaller. It checks nonempty row counts, output key uniqueness, segment totals, retained transition cohort totals and the absence of loan identifier columns. It does not repeat the full SQL validation.

## 7. What I learned and limitations

The main modeling decision was not the SQL syntax; it was choosing denominators that match the question. On-book exposure, delinquency eligibility and observable transition cohorts are related populations but should not be treated as interchangeable. Keeping both count and UPB views also prevents one weighting method from hiding the other.

This project is descriptive. It does not predict default, estimate lifetime loss or assign credit decisions. The two annual samples are not the full Freddie Mac population or a live lender portfolio. Comparisons across time and vintage reflect entry timing, seasoning, payoff and other exits, right-censoring, and changes in the loans remaining in the sample. The Excel workbook uses local aggregate exports, so it must be rebuilt to reflect a new database run.
