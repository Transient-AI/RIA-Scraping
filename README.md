# RIA Web Scraping and Classification

This repository contains Python scripts for scraping, processing, and classifying data from Registered Investment Advisor (RIA) websites and SEC ADV filings.

## Files Overview

### 1. `web_scraping.py`
Advanced web scraping tool for extracting structured data from RIA websites.

**Features:**
- Adaptive web crawling with proxy rotation and session management
- Extracts company information, personnel data, services, and client details
- Uses Azure OpenAI for intelligent content extraction
- Handles multi-threading for parallel processing
- Saves output as JSON files

**Key capabilities:**
- Proxy rotation for robust scraping
- Domain-specific crawling strategies
- Error handling and retry logic
- Excel input support for batch processing

### 2. `ria_classifier.py`
Multi-category classifier for RIA data using Azure OpenAI.

**Features:**
- Classifies RIA data into three main categories:
  - Services offered
  - Target clients
  - Products and instruments
- Uses a three-prompt approach for comprehensive extraction
- Merges and deduplicates results across categories
- Generates standardized mappings for consistent categorization
- Outputs multi-sheet Excel files

**Use case:** Takes scraped RIA data and categorizes it into standardized business categories for analysis.

### 3. `ria_classifier_cosmos_brochure2_v1.py`
Extracts structured information from SEC Form ADV Part 2 brochures stored in Azure Cosmos DB.

**Features:**
- Direct integration with Azure Cosmos DB
- Processes Form ADV Part 2 brochures
- Extracts 6 key columns:
  - Services
  - Target Clients
  - Products & Instruments
  - Minimum Account Size
  - Fee Structure
  - Industry Insights
- Uses Azure OpenAI for content analysis
- Batch processing with checkpoint management
- Generates Excel output with comprehensive data

**Use case:** Automated extraction and classification of information from SEC regulatory filings.

### 4. `scraped_json_consolidation_Script_v2.py`
Data consolidation tool that processes multiple JSON files and creates organized Excel output.

**Features:**
- Consolidates scraped JSON data into multi-sheet Excel workbooks
- Creates separate sheets for:
  - Personnel information
  - Services offered
  - Target clients
  - Products and instruments
  - Industry insights
- Links all data by CRD number and original URL
- Handles data cleaning and normalization
- Deduplicates entries across categories

**Use case:** Post-processing tool to organize and consolidate scraped data into a structured format for analysis.

## Prerequisites

### Python Dependencies
```bash
pip install python-dotenv openai pandas openpyxl azure-cosmos requests tqdm httpx
```

### Environment Variables
Copy `.env.example` to `.env` and fill in your credentials:

```bash
cp .env.example .env
```

Required environment variables:
- `AZURE_OPENAI_API_KEY`: Your Azure OpenAI API key
- `AZURE_OPENAI_ENDPOINT`: Azure OpenAI endpoint URL
- `AZURE_OPENAI_API_VERSION`: API version (e.g., 2024-12-01-preview)
- `AZURE_OPENAI_DEPLOYMENT_NAME`: Deployment name (e.g., gpt-4o)
- `AZURE_COSMOSDB_KEY`: Azure Cosmos DB key (for cosmos brochure script)
- `AZURE_COSMOSDB_ENDPOINT`: Azure Cosmos DB endpoint URL

## Usage

### Web Scraping
```python
# Configure input Excel file with RIA website URLs
python web_scraping.py
```

### Data Consolidation
```python
# Process scraped JSON files into organized Excel
python scraped_json_consolidation_Script_v2.py
```

### RIA Classification
```python
# Classify consolidated data into standard categories
python ria_classifier.py
```

### SEC ADV Brochure Processing
```python
# Extract data from Cosmos DB stored Form ADV documents
python ria_classifier_cosmos_brochure2_v1.py
```

## Workflow

1. **Scrape websites** using `web_scraping.py` to collect raw data
2. **Consolidate data** using `scraped_json_consolidation_Script_v2.py` to organize into Excel
3. **Classify services** using `ria_classifier.py` to standardize categories
4. **Optional:** Process SEC filings using `ria_classifier_cosmos_brochure2_v1.py` for regulatory data

## Security Notes

- Never commit the `.env` file to version control
- The `.env.example` file is provided as a template
- All API keys and secrets should be stored in environment variables
- The repository includes GitHub secret scanning protection

## Output

All scripts generate Excel files with structured data including:
- Company information and CRD numbers
- Personnel details
- Services categorization
- Target client profiles
- Products and instruments offered
- Fee structures and minimum account sizes
- Industry insights and specializations

## License

Copyright © Transient AI
