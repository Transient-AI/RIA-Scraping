# RIA-Scraping

SEC Form ADV data processor for Registered Investment Advisors (RIA).

## Overview

This project automates the downloading and processing of SEC Form ADV filing data. It extracts, transforms, and organizes RIA data into structured Excel reports.

## Features

- Downloads Form ADV Part 1 filing data from SEC EDGAR
- Processes base advisor information, custodian data, and ownership structures
- Filters to latest filings per CRD number
- Exports to formatted Excel workbooks with multiple sheets
- Generates summary statistics and data quality reports

## Files

- `ADV_Filing_Processing_Part_1.py` - Main processing script
- `config.yaml` - Configuration file for paths, processing options, and settings

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure `config.yaml`:
   - Set data paths (raw_data, processed_data, reports, archive)
   - Specify month and year to process
   - Adjust processing options as needed

3. Run the processor:
```bash
python ADV_Filing_Processing_Part_1.py
```

## Configuration

Key configuration options in `config.yaml`:

- `processing.month` - Month to process (e.g., "October")
- `processing.year` - Year to process
- `processing.force_refresh` - Reprocess existing data
- `processing.use_existing_zip` - Use local ZIP file instead of downloading
- `paths.*` - Directory paths for data storage

## Output

The script generates:
- Excel workbook with Base_Data, Custodian_Data, Owners_Data sheets
- Summary statistics and CRD coverage reports
- Processing metadata and error logs
- JSON reports in the reports directory

## Data Sources

Data is sourced from the SEC's Investment Adviser Public Disclosure (IAPD) system:
- https://adviserinfo.sec.gov/
- https://reports.adviserinfo.sec.gov/

## License

MIT
