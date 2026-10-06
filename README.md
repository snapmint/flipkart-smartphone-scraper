# flipkart-smartphone-scraper

## Jenkins output handling

Jenkins workspaces persist between builds. The pipeline deletes previous
`flipkart_mobile_*.xlsx` and `amazon_mobile_*.xlsx` workbooks before scraping,
then requires exactly one workbook per dataset before converting to CSV.
Concurrent builds are disabled to prevent overlapping cleanup and scraping.

Previously, conversion selected the first matching workbook in the workspace,
which could upload old data under a new S3 timestamp while Jenkins archived
both old and new workbooks. The fix applies to future builds; existing stale
S3 objects require separate investigation and replacement from the correct
archived workbooks, if available.

Run pipeline output regression tests with `python3 -m unittest discover -s tests -v`.
These tests execute the pipeline's cleanup and selection shell scripts in
temporary workspaces, with Excel conversion stubbed; no AWS access is needed.