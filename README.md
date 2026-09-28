# Mortgage Portfolio Risk Monitoring with SQL and Excel

This portfolio project turns Freddie Mac Single-Family Loan-Level Dataset samples into a governed mortgage-risk monitoring layer and an offline Excel dashboard. It demonstrates how to define portfolio populations, calculate delinquency and migration metrics, reconcile results across reporting grains, and present decision-ready views without publishing loan-level data.

## Project at a glance

| Area | Verified scope |
|---|---:|
| Sampled mortgage originations | 100,000 |
| Monthly performance records | 4,452,471 |
| Origination vintages | 2019 and 2020 |
| Performance coverage | January 2019 through March 2026 |
| Blocking SQL quality checks | 39 / 39 passed |
| Dashboard export checks | 14 / 14 passed |
| Excel workbook | 8 worksheets, 9 native charts |
| Core technologies | SQL, DuckDB, Python, Excel |

The 2019 and 2020 labels identify origination vintages. Their monthly performance histories continue through March 2026.

## Business questions

- How many loans and how much unpaid principal balance remain on book each month?
- How do 30+, 60+, and 90+ delinquency rates differ when measured by loan count versus UPB?
- Where is risk concentrated across FICO, CLTV, DTI, and origination-vintage segments?
- How do observable loans migrate between current, delinquent, REO, unknown, and terminal states from one calendar month to the next?
- Are the reported populations, numerators, denominators, segment totals, and transition rates internally consistent?

## Analytical workflow

![Mortgage portfolio risk monitoring workflow](docs/assets/project_workflow.png)

The SQL layer separates raw ingestion, typed staging tables, curated analytical tables, reporting marts, and executable quality gates. Python validates the source files, orchestrates the transactional DuckDB build, exports governed aggregates, and prepares the workbook payload. The Excel builder creates the final offline presentation layer.

## Core risk metrics

- **On-book exposure:** month-end loan count and Current Actual UPB for records with positive UPB and no Zero Balance Code.
- **Delinquency rates by count:** 30+, 60+, and 90+ loan numerators divided by the delinquency-eligible loan population.
- **Delinquency rates by balance:** 30+, 60+, and 90+ UPB numerators divided by eligible UPB. These measures remain separate from count-based rates.
- **Coverage:** eligible loan count and UPB compared with the broader on-book population.
- **Segment monitoring:** independent views by origination vintage, Classic FICO, Original CLTV, and Original DTI.
- **Monthly transitions:** movements between risk states only when the same loan has an observation exactly one calendar month later.

The 90+ bucket represents severe delinquency; it is not automatically equivalent to default.

## Dashboard views

- **Portfolio Overview** tracks monthly on-book exposure, eligible populations, coverage, and count- and UPB-weighted delinquency rates.
- **Risk Segments** compares the same risk measures across vintage, FICO, CLTV, and DTI bands for a selected month.
- **Transitions** shows count and beginning-UPB migration rates for observable adjacent-month cohorts.

![Excel dashboard page structure](docs/assets/dashboard_views.png)

The workbook uses formula-driven rates based on additive numerators and denominators. It contains no loan identifiers, loan-level rows, VBA, Power Query, or external data connections.

## Data quality and validation

The build executes SQL transformations and blocking controls in one transaction. A failed `ERROR`-severity test rolls back the database build. The 39 blocking checks cover source counts, key uniqueness, cross-table coverage, valid dates and official codes, terminal-event consistency, population logic, segment reconciliation, independent metric recalculation, and transition reconciliation.

Before workbook creation, 14 additional checks validate the aggregate exports, including grain, completeness, suppression logic, rate components, and agreement with the governed marts. Three benchmark months are independently recalculated from staging fields rather than curated eligibility flags.

## Repository structure

```text
.
├── README.md
├── NOTICE.md
├── requirements.txt
├── docs/
│   └── methodology.md
├── scripts/
│   ├── validate_raw_data.py
│   ├── build_data_layer.py
│   ├── export_dashboard_data.py
│   └── build_excel_dashboard.py
└── sql/
    ├── 00_init.sql
    ├── 10_staging.sql
    ├── 20_curated.sql
    ├── 30_marts.sql
    ├── 40_dashboard_exports.sql
    └── 90_quality_checks.sql
```

Files most useful for technical review are `sql/20_curated.sql` for analytical population flags and adjacent-month joins, `sql/30_marts.sql` for metric aggregation, `sql/90_quality_checks.sql` for executable controls, `scripts/export_dashboard_data.py` for governed presentation exports, and `scripts/build_excel_dashboard.py` for workbook formulas, charts, and package validation.

## Reproduction

Obtain the applicable 2019 and 2020 sample files from the official source and place them locally at:

```text
data/raw/sample_2019/sample_orig_2019.txt
data/raw/sample_2019/sample_perf_2019.txt
data/raw/sample_2020/sample_orig_2020.txt
data/raw/sample_2020/sample_perf_2020.txt
```

From the repository root:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python scripts/validate_raw_data.py
.venv/bin/python scripts/build_data_layer.py
.venv/bin/python scripts/build_excel_dashboard.py
```

The validation step writes a local source profile. The build verifies source sizes and SHA-256 values against that profile, then runs the SQL stages in a single transaction. Source files remain read-only, and generated data stays outside Git.

The DuckDB SQL layer and Excel workbook are reproducible with the publicly listed Python dependencies. The workbook builder uses XlsxWriter and produces the offline dashboard from the validated aggregate marts.

## Methodology and data notice

- [Risk Methodology and Data Model](docs/methodology.md)
- [Data and Publication Notice](NOTICE.md)
- [Freddie Mac Single-Family Loan-Level Dataset](https://www.freddiemac.com/research/datasets/sf-loanlevel-dataset)
- [Release 47 General User Guide](https://www.freddiemac.com/fmac-resources/research/pdf/general_user_guide_july_2026.pdf)
- [Release 47 File Layout](https://www.freddiemac.com/fmac-resources/research/pdf/file_layout_july_2026.xlsx)

## Limitations

This project uses two annual Standard Dataset samples, not Freddie Mac's full loan population or the entire U.S. mortgage market. Results describe the observed sample and are not a live bank portfolio, forecast, causal estimate, or investment recommendation. Trends reflect loan entry, seasoning, payoff and other exits, right-censoring, and changes in the remaining sample composition. Raw data, rebuilt databases, aggregate exports, validation details, and the generated workbook are intentionally kept local under the applicable data terms.
