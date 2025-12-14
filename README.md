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

### Part 1 Processing
- `ADV_Filing_Processing_Part_1.py` - Main processing script for Form ADV Part 1
- `config.yaml` - Configuration file for paths, processing options, and settings

### Part 2 Brochure Processing
- `ria_brochure_two_processing.py` - Processes Form ADV Part 2 brochure PDFs, converts to markdown/JSON
- `brochure_two_vectodb_create_or_update.py` - Builds vector database in Azure Cosmos DB for brochure data with embeddings

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

## Part 2 Brochure Processing Workflow

### Step 1: Download and Convert Brochures to JSON

Use [ria_brochure_two_processing.py](ria_brochure_two_processing.py) to download Form ADV Part 2 brochures and convert them to JSON:

```bash
python ria_brochure_two_processing.py
```

Features:
- Downloads SEC Form ADV Part 2 brochures for specified month/year
- Converts PDF brochures to markdown using MarkItDown
- Filters out "wrap" brochures automatically
- Saves content as JSON files with metadata
- Generates processing summary reports

Configure the download in the script:
```python
main(
    download_month="October",
    download_year=2025,
    output_directory="/path/to/output/"
)
```

### Step 2: Build Vector Database

Use [brochure_two_vectodb_create_or_update.py](brochure_two_vectodb_create_or_update.py) to process JSON files and create vector embeddings in Azure Cosmos DB:

**Setup environment variables (required):**
```bash
export AZURE_OPENAI_KEY='your-azure-openai-api-key'
```

**Run the script:**
```bash
# Process all JSON files in directory
python brochure_two_vectodb_create_or_update.py /path/to/json/directory

# Retry failed files from previous run
python brochure_two_vectodb_create_or_update.py /path/to/json/directory --retry /path/to/previous_report.json
```

**Optional environment variables:**
- `AZURE_SUBSCRIPTION_ID` - Azure subscription ID
- `AZURE_RESOURCE_GROUP` - Resource group name
- `COSMOS_ACCOUNT_NAME` - Cosmos DB account name
- `AZURE_OPENAI_ENDPOINT` - Azure OpenAI endpoint URL
- `AZURE_OPENAI_EMBEDDING_DEPLOYMENT` - Embedding model deployment name

Features:
- Chunks documents intelligently with token-based overlap
- Generates embeddings using Azure OpenAI (text-embedding-3-large)
- Stores in Azure Cosmos DB with vector search capabilities
- Automatic retry logic for failed operations
- Replaces old versions when processing updated filings
- Detailed processing reports with error tracking

## Configuration

Key configuration options in `config.yaml`:

- `processing.month` - Month to process (e.g., "October")
- `processing.year` - Year to process
- `processing.force_refresh` - Reprocess existing data
- `processing.use_existing_zip` - Use local ZIP file instead of downloading
- `paths.*` - Directory paths for data storage

## Output

### Part 1 Output
The Part 1 processing script generates:
- Excel workbook with Base_Data, Custodian_Data, Owners_Data sheets
- Summary statistics and CRD coverage reports
- Processing metadata and error logs
- JSON reports in the reports directory

### Part 2 Brochure Output
The brochure processing generates:
- JSON files with markdown content extracted from PDFs
- Processing summary reports (Excel and CSV)
- Vector database entries in Azure Cosmos DB with embeddings
- Processing reports with success/error tracking

## Data Sources

Data is sourced from the SEC's Investment Adviser Public Disclosure (IAPD) system:
- https://adviserinfo.sec.gov/
- https://reports.adviserinfo.sec.gov/

## License

MIT
