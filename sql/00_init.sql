-- Rebuildable schemas for the Freddie Mac SFLLD Release 47 sample.
-- Raw text files remain external and read-only; the DuckDB file is generated.

DROP SCHEMA IF EXISTS quality CASCADE;
DROP SCHEMA IF EXISTS mart CASCADE;
DROP SCHEMA IF EXISTS curated CASCADE;
DROP SCHEMA IF EXISTS staging CASCADE;
DROP SCHEMA IF EXISTS raw_external CASCADE;

CREATE SCHEMA raw_external;
CREATE SCHEMA staging;
CREATE SCHEMA curated;
CREATE SCHEMA mart;
CREATE SCHEMA quality;
