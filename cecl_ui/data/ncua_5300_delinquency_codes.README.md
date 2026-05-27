# NCUA 5300 Delinquency Codes — Field Code Map

This sidecar file documents `ncua_5300_delinquency_codes.csv`, the
canonical map used by `solr_5300_delq_backfill.py` to pull historical
delinquent balances from the NCUA 5300 call report via Solr.

## Format

```
field_code,loan_code,description
A210,Unsecured Credit Card,Delinquent balance (≥60 days) — ...
...
```

Each row says: "the Solr field named `<field_code>` contains the
delinquent balance for loans we group under `<loan_code>`". The
`loan_code` value must match the loan_code labels used in
`ncua_5300_chargeoff_codes.csv` exactly so that the rest of the CECL
report engine treats them as the same pool.

## Bucket convention

The CECL report renders DQ as a single percentage (≥60-day balance ÷
total balance). If the 5300 reports the same loan category split across
multiple buckets (e.g. 60-179 days and ≥180 days), add one CSV row per
bucket — they will all be summed under the same `loan_code`. Example::

```
ACH0XXX,New Vehicles,Delinquent 60-179 days — new vehicle loans
ACH0YYY,New Vehicles,Delinquent 180+ days — new vehicle loans
```

## Populating field codes

The `field_code` column is currently blank for every row — it is up to
the integrator to fill in the correct NCUA 5300 account codes. When all
rows are blank, the wizard's "Backfill from 5300" button reports
`0 of N codes mapped` and writes nothing.

The 5300 Schedule for Delinquent Loans (typically Schedule A in the
call report) carries the relevant account codes. They live in the same
namespace as the charge-off codes used by
`ncua_5300_chargeoff_codes.csv` — see that file's `field_code` column
for examples (A680, ACH0017, etc.).

Rows whose `field_code` is blank are silently skipped by the loader.
