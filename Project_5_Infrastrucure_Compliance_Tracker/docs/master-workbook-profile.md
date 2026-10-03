# Current master workbook profile

## Workbook structure

- Workbook: `NEWCopy_of_SOE_Compliance_Tracker_with_Splunk_History.xlsx`
- Inventory tab: `Main Data`
- Status history tab: `Splunk`
- Parameter tabs: `CS` (Craft), `TGIM`, `Splunk`, `Logger`, `RSA`, and `SNOW`
- `Main Data` contains 7,105 distinct hostnames and one additional row without a hostname.
- The dashboard uses the 7,105 named hosts as its initial server baseline. The previous application import used anonymized hostnames, so it is archived rather than compared as a new-server source.

## Splunk status history

| Month | Compliant | Non-Compliant | Not Applicable | `Compliant to Non-Compliant` |
| --- | ---: | ---: | ---: | ---: |
| August 2026 | 5,128 | 1,204 | 586 | 187 |
| September 2026 | 5,086 | 1,237 | 582 | 200 |
| October 2026 | 5,074 | 1,285 | 569 | 177 |

These counts exclude the row without a hostname. The transition-like source label remains Unmapped until the owner chooses a mapping. Other parameter tabs currently have no status values and therefore show “No snapshot.”

## Data quality notes

- There is one row without a hostname.
- 328 rows have no usable IP value, including blank, `nan`, and `Not Found` values.
- 87 nonblank IP strings do not parse as IP addresses.
- October includes 177 named servers with the Unmapped transition-like status.
- The server tabs have formula-linked host/IP cells. The importer reads those identities from `Main Data`; it does not import formula text as hostnames.

## New-server baseline

The first master upload establishes the baseline, so the dashboard leaves the new-server number blank. A later master upload compares its named hosts with the previous master roster. The anonymized prior dataset is retained for traceability but is not used in this comparison.
