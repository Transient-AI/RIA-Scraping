"""
SEC ADV Filing Information Extractor
Extracts structured information from Form ADV Part 2 brochures stored in Azure Cosmos DB
Uses Azure OpenAI to analyze content and generates Excel output with 6 key columns
"""

import json
import os
import logging
import time
from datetime import datetime, timezone
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
from dataclasses import dataclass
import re
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
from tqdm import tqdm
import httpx

# Azure imports
from azure.cosmos import CosmosClient, exceptions
from openai import AzureOpenAI
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import List, Dict, Any, Optional, Tuple


@dataclass
class Config:
    # Cosmos DB Configuration - Direct Connection
    cosmos_endpoint: str = os.getenv("AZURE_COSMOSDB_ENDPOINT")
    cosmos_key: str = os.getenv("AZURE_COSMOSDB_KEY")
    database_name: str = "Sec-adv-filing-part2-brochure-pdf"
    container_name: str = "part2-json-data"

    # Azure OpenAI Configuration
    azure_openai_endpoint: str = os.getenv("AZURE_OPENAI_ENDPOINT")
    azure_openai_key: str = os.getenv("AZURE_OPENAI_API_KEY")
    azure_openai_api_version: str = os.getenv("AZURE_OPENAI_API_VERSION")
    chat_deployment: str = "gpt-5.1-chat"  # Or your deployed model name
    
    # Processing Configuration
    batch_size: int = 50  # Save checkpoint every N CRDs
    call_delay: float = 0.5  # Delay between API calls in seconds
    max_retries: int = 3
    retry_delay: int = 2
    connection_timeout: int = 60
    read_timeout: int = 180  # Longer timeout for chat completions
    
    # Output Configuration
    output_dir: str = "cosmos_brochure2_extraction_output"
    checkpoint_file: str = "cosmos_brochure2_extraction_checkpoint.json"
    final_output: str = "cosmos_brochure2_sec_adv_extractions.xlsx"
    log_file: str = "cosmos_brochure2_extraction.log"


class CosmosDBClient:
    """Handles Cosmos DB operations with direct key authentication"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.cosmos_client = None
        self.container = None
        
    def initialize(self):
        """Initialize Cosmos DB client with direct key"""
        self.logger.info("Initializing Cosmos DB connection...")
        
        self.cosmos_client = CosmosClient(
            self.config.cosmos_endpoint, 
            self.config.cosmos_key
        )
        database = self.cosmos_client.get_database_client(self.config.database_name)
        self.container = database.get_container_client(self.config.container_name)
        
        self.logger.info("Cosmos DB connection established")
    
    def get_all_unique_crds(self) -> List[Dict[str, Any]]:
        """Get all unique CRD numbers with firm names and filing dates"""
        self.logger.info("Fetching all unique CRD numbers...")
        
        query = """
        SELECT c.crd_number, c.firm_name, c.filing_date, c.source_file
        FROM c 
        WHERE c.chunk_index = 0
        """
        
        items = list(self.container.query_items(
            query=query,
            enable_cross_partition_query=True
        ))
        
        # Deduplicate by CRD, keeping only the latest filing_date
        crd_map = {}
        for item in items:
            crd = item['crd_number']
            filing_date = item.get('filing_date', '')
            
            if crd not in crd_map:
                crd_map[crd] = item
            else:
                # Compare filing dates, keep the latest
                existing_date = crd_map[crd].get('filing_date', '')
                if filing_date and (not existing_date or filing_date > existing_date):
                    crd_map[crd] = item
        
        unique_items = list(crd_map.values())
        self.logger.info(f"Found {len(unique_items)} unique CRDs (from {len(items)} total records)")
        return unique_items
    
    def get_document_by_crd(self, crd_number: int, filing_date: str = None) -> tuple:
        """Reconstruct full document from chunks for a given CRD and filing date"""
        
        # If filing_date provided, filter by it to get specific filing
        if filing_date:
            query = f"""
            SELECT c.content, c.chunk_index, c.total_chunks, c.firm_name, c.filing_date
            FROM c 
            WHERE c.crd_number = {crd_number} AND c.filing_date = '{filing_date}'
            ORDER BY c.chunk_index
            """
        else:
            query = f"""
            SELECT c.content, c.chunk_index, c.total_chunks, c.firm_name, c.filing_date
            FROM c 
            WHERE c.crd_number = {crd_number}
            ORDER BY c.chunk_index
            """
        
        chunks = list(self.container.query_items(
            query=query,
            enable_cross_partition_query=True
        ))
        
        if not chunks:
            return None, None, None
        
        # Sort by chunk_index and concatenate
        chunks.sort(key=lambda x: x['chunk_index'])
        full_content = " ".join([c['content'] for c in chunks])
        firm_name = chunks[0].get('firm_name', '')
        filing_date = chunks[0].get('filing_date', '')
        
        return full_content, firm_name, filing_date


class ADVExtractor:
    """Extracts structured information from ADV filings using Azure OpenAI"""
    
    EXTRACTION_PROMPT = """I am providing a Form ADV Part 2 brochure content.

Extract ONLY the information that is explicitly stated in the brochure — 
do NOT use assumptions, external sources, or interpretations.

Search the entire document including:
- core sections
- appendix
- footnotes
- fee examples
- risk disclosures
- supplementary pages

Extract information for EXACTLY these 6 categories:

1. Services Provided - What advisory services does the firm offer?
2. Investment Strategies - What investment strategies do they use?
3. Methods of Analysis - What methods do they use to analyze investments?
4. Asset Classes / Products Used - List EVERY asset or investment product referenced
5. Clients Served - What types of clients do they serve?
6. Investment Discretion - Do they have discretionary authority? What are the terms?

CRITICAL OUTPUT FORMAT RULES:
- Each field must contain SHORT, DISTINCT items separated by " | " (pipe with spaces)
- Each item should be a concise phrase (2-6 words), NOT a full sentence
- Extract individual categories/types, NOT paragraphs of text
- If a field has no data, return "NOT FOUND"

EXAMPLES OF CORRECT FORMAT:
- services_provided: "Portfolio Management | Financial Planning | Retirement Planning | Tax Advisory | Estate Planning"
- investment_strategies: "Long-term Growth | Value Investing | Index Tracking | Tactical Allocation | Income Generation"
- methods_of_analysis: "Fundamental Analysis | Technical Analysis | Quantitative Models | Economic Research | Company Financials Review"
- asset_classes_products: "Mutual Funds | ETFs | Stocks | Bonds | Options | REITs | Private Equity | Commodities"
- clients_served: "High Net Worth Individuals | Pension Plans | Trusts | Corporations | Charitable Organizations"
- investment_discretion: "Full Discretionary | Non-Discretionary | Limited Discretionary for Rebalancing"

EXAMPLES OF INCORRECT FORMAT (DO NOT DO THIS):
- "The firm provides portfolio management services including..." (too verbose)
- "Mutual funds, ETFs, stocks" (use pipes, not commas)

Return your response as a JSON object with these exact keys:
{
    "services_provided": "Item1 | Item2 | Item3",
    "investment_strategies": "Item1 | Item2 | Item3",
    "methods_of_analysis": "Item1 | Item2 | Item3",
    "asset_classes_products": "Item1 | Item2 | Item3",
    "clients_served": "Item1 | Item2 | Item3",
    "investment_discretion": "Item1 | Item2 | Item3"
}

Here is the Form ADV Part 2 brochure content:

"""

    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.openai_client = None
        
    def initialize(self):
        """Initialize Azure OpenAI client"""
        self.logger.info("Initializing Azure OpenAI client...")
        
        self.openai_client = AzureOpenAI(
            azure_endpoint=self.config.azure_openai_endpoint,
            api_key=self.config.azure_openai_key,
            api_version=self.config.azure_openai_api_version,
            http_client=httpx.Client(
                timeout=httpx.Timeout(
                    connect=self.config.connection_timeout,
                    read=self.config.read_timeout,
                    write=self.config.connection_timeout,
                    pool=self.config.connection_timeout
                )
            )
        )
        
        self.logger.info("Azure OpenAI client initialized")
    
    def extract_information(self, content: str, crd_number: int) -> Optional[Dict[str, str]]:
        """Extract structured information from document content with chunking for long docs"""
        
        # Token estimation: ~4 chars per token, leave room for prompt (~2000 tokens)
        max_content_chars = 120000  # ~30k tokens for content
        
        if len(content) <= max_content_chars:
            # Single extraction for normal-sized documents
            return self._extract_single(content, crd_number)
        else:
            # Chunked extraction for long documents
            return self._extract_chunked(content, crd_number, max_content_chars)
        
        for attempt in range(self.config.max_retries):
            try:
                response = self.openai_client.chat.completions.create(
                    model=self.config.chat_deployment,
                    messages=[
                        {
                            "role": "system",
                            "content": "You are a financial document analyst specializing in SEC Form ADV filings. Extract information precisely as requested, using only the document content provided. Return valid JSON only."
                        },
                        {
                            "role": "user",
                            "content": self.EXTRACTION_PROMPT + content
                        }
                    ],
                    temperature=0.1,
                    max_tokens=4000,
                    response_format={"type": "json_object"}
                )
                
                result_text = response.choices[0].message.content
                result = json.loads(result_text)
                
                # Validate required fields
                required_fields = [
                    "services_provided", "investment_strategies", 
                    "methods_of_analysis", "asset_classes_products",
                    "clients_served", "investment_discretion"
                ]
                
                for field in required_fields:
                    if field not in result:
                        result[field] = "NOT FOUND"
                
                return result
                
            except json.JSONDecodeError as e:
                self.logger.warning(f"JSON decode error for CRD {crd_number}, attempt {attempt + 1}: {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (2 ** attempt))
                    
            except Exception as e:
                self.logger.warning(f"Extraction error for CRD {crd_number}, attempt {attempt + 1}: {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (2 ** attempt))
        
        self.logger.error(f"Failed to extract information for CRD {crd_number} after {self.config.max_retries} attempts")
        return None
    
    
    
    def _extract_single(self, content: str, crd_number: int) -> Optional[Dict[str, str]]:
        """Extract from a single content block"""
        for attempt in range(self.config.max_retries):
            try:
                response = self.openai_client.chat.completions.create(
                    model=self.config.chat_deployment,
                    messages=[
                        {
                            "role": "system",
                            "content": "You are a financial document analyst specializing in SEC Form ADV filings. Extract information precisely as requested, using only the document content provided. Return valid JSON only. Use pipe separators between items."
                        },
                        {
                            "role": "user",
                            "content": self.EXTRACTION_PROMPT + content
                        }
                    ],
                    temperature=0.1,
                    max_tokens=4000,
                    response_format={"type": "json_object"}
                )
                
                result_text = response.choices[0].message.content
                result = json.loads(result_text)
                
                # Validate and normalize required fields
                required_fields = [
                    "services_provided", "investment_strategies", 
                    "methods_of_analysis", "asset_classes_products",
                    "clients_served", "investment_discretion"
                ]
                
                for field in required_fields:
                    if field not in result or not result[field]:
                        result[field] = "NOT FOUND"
                    else:
                        # Normalize: ensure pipe separation, clean up
                        result[field] = self._normalize_field(result[field])
                
                return result
                
            except json.JSONDecodeError as e:
                self.logger.warning(f"JSON decode error for CRD {crd_number}, attempt {attempt + 1}: {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (2 ** attempt))
                    
            except Exception as e:
                self.logger.warning(f"Extraction error for CRD {crd_number}, attempt {attempt + 1}: {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (2 ** attempt))
        
        return None
    
    def _extract_chunked(self, content: str, crd_number: int, max_chunk_size: int) -> Optional[Dict[str, str]]:
        """Extract from multiple chunks and merge results"""
        self.logger.info(f"CRD {crd_number}: Document too long ({len(content)} chars), using chunked extraction")
        
        # Split content into overlapping chunks
        chunks = []
        overlap = 5000  # Overlap to avoid missing info at boundaries
        start = 0
        
        while start < len(content):
            end = min(start + max_chunk_size, len(content))
            chunks.append(content[start:end])
            start = end - overlap
            if start >= len(content) - overlap:
                break
        
        self.logger.info(f"CRD {crd_number}: Split into {len(chunks)} chunks")
        
        # Extract from each chunk
        all_extractions = []
        for i, chunk in enumerate(chunks):
            self.logger.info(f"CRD {crd_number}: Processing chunk {i+1}/{len(chunks)}")
            extraction = self._extract_single(chunk, crd_number)
            if extraction:
                all_extractions.append(extraction)
            time.sleep(self.config.call_delay)
        
        if not all_extractions:
            return None
        
        # Merge results from all chunks
        return self._merge_extractions(all_extractions)
    
    def _merge_extractions(self, extractions: List[Dict[str, str]]) -> Dict[str, str]:
        """Merge multiple extractions, deduplicating items"""
        required_fields = [
            "services_provided", "investment_strategies", 
            "methods_of_analysis", "asset_classes_products",
            "clients_served", "investment_discretion"
        ]
        
        merged = {}
        
        for field in required_fields:
            all_items = set()
            
            for extraction in extractions:
                value = extraction.get(field, "")
                if value and value != "NOT FOUND":
                    # Split by pipe and add unique items
                    items = [item.strip() for item in value.split("|") if item.strip()]
                    all_items.update(items)
            
            if all_items:
                # Sort for consistency and join with pipes
                merged[field] = " | ".join(sorted(all_items))
            else:
                merged[field] = "NOT FOUND"
        
        return merged
    
    def _normalize_field(self, value: str) -> str:
        """Normalize field value to ensure proper pipe separation"""
        if not value or value == "NOT FOUND":
            return "NOT FOUND"
        
        # If already has pipes, clean up
        if "|" in value:
            items = [item.strip() for item in value.split("|") if item.strip()]
            return " | ".join(items)
        
        # If comma-separated, convert to pipes
        if "," in value and len(value) < 500:  # Short lists
            items = [item.strip() for item in value.split(",") if item.strip()]
            return " | ".join(items)
        
        # Single item or sentence - return as is
        return value.strip()


class CheckpointManager:
    """Manages checkpointing for resumable processing"""
    
    def __init__(self, config: Config):
        self.config = config
        self.output_dir = Path(config.output_dir)
        self.checkpoint_path = self.output_dir / config.checkpoint_file
        self.logger = logging.getLogger(__name__)
        
        # Ensure output directory exists
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def load_checkpoint(self) -> Dict[str, Any]:
        """Load checkpoint data if exists"""
        if self.checkpoint_path.exists():
            with open(self.checkpoint_path, 'r') as f:
                checkpoint = json.load(f)
            self.logger.info(f"Loaded checkpoint: {checkpoint.get('processed_count', 0)} CRDs already processed")
            return checkpoint
        return {
            "processed_crds": [],
            "failed_crds": [],
            "processed_count": 0,
            "batch_files": [],
            "last_updated": None
        }
    
    def save_checkpoint(self, checkpoint: Dict[str, Any]):
        """Save checkpoint data"""
        checkpoint["last_updated"] = datetime.now(timezone.utc).isoformat()
        with open(self.checkpoint_path, 'w') as f:
            json.dump(checkpoint, f, indent=2)
    
    def save_batch(self, batch_data: List[Dict], batch_number: int) -> str:
        """Save batch results to JSON file"""
        batch_file = self.output_dir / f"batch_{batch_number:04d}.json"
        with open(batch_file, 'w') as f:
            json.dump(batch_data, f, indent=2)
        self.logger.info(f"Saved batch {batch_number} with {len(batch_data)} records to {batch_file}")
        return str(batch_file)
    
    def load_all_batches(self) -> List[Dict]:
        """Load all batch files and merge"""
        all_data = []
        batch_files = sorted(self.output_dir.glob("batch_*.json"))
        
        for batch_file in batch_files:
            with open(batch_file, 'r') as f:
                batch_data = json.load(f)
            all_data.extend(batch_data)
        
        self.logger.info(f"Loaded {len(all_data)} records from {len(batch_files)} batch files")
        return all_data


class ExcelGenerator:
    """Generates final Excel output"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
    
    def generate_excel(self, data: List[Dict], output_path: str):
        """Generate formatted Excel file from extraction results"""
        self.logger.info(f"Generating Excel file with {len(data)} records...")
        
        # Prepare DataFrame
        rows = []
        for item in data:
            row = {
                "CRD Number": item.get("crd_number", ""),
                "Firm Name": item.get("firm_name", ""),
                "Filing Date": item.get("filing_date", ""),  # Add filing date
                "Services Provided": item.get("services_provided", "NOT FOUND"),
                "Investment Strategies": item.get("investment_strategies", "NOT FOUND"),
                "Methods of Analysis": item.get("methods_of_analysis", "NOT FOUND"),
                "Asset Classes / Products Used": item.get("asset_classes_products", "NOT FOUND"),
                "Clients Served": item.get("clients_served", "NOT FOUND"),
                "Investment Discretion": item.get("investment_discretion", "NOT FOUND"),
                "Extraction Status": item.get("status", "SUCCESS"),
                "Extracted At": item.get("extracted_at", "")
            }
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        # Create workbook with formatting
        wb = Workbook()
        ws = wb.active
        ws.title = "ADV Extractions"
        
        # Header styling
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="4472C4")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Border styling
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Write headers
        headers = list(df.columns)
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Write data
        for row_idx, row_data in enumerate(df.values, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = thin_border
        
        # Set column widths (updated for new column)
        column_widths = {
            "A": 12,   # CRD Number
            "B": 30,   # Firm Name
            "C": 12,   # Filing Date
            "D": 50,   # Services Provided
            "E": 50,   # Investment Strategies
            "F": 40,   # Methods of Analysis
            "G": 60,   # Asset Classes
            "H": 40,   # Clients Served
            "I": 40,   # Investment Discretion
            "J": 15,   # Status
            "K": 20,   # Extracted At
        }
        
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # Freeze header row
        ws.freeze_panes = "A2"
        
        # Add autofilter
        ws.auto_filter.ref = ws.dimensions
        
        # Save
        wb.save(output_path)
        self.logger.info(f"Excel file saved to {output_path}")


class ADVExtractionPipeline:
    """Main orchestrator for the extraction pipeline"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = self._setup_logging()
        
        self.cosmos_client = CosmosDBClient(config)
        self.extractor = ADVExtractor(config)
        self.checkpoint_mgr = CheckpointManager(config)
        self.excel_gen = ExcelGenerator(config)
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging configuration"""
        log_dir = Path(self.config.output_dir)
        log_dir.mkdir(parents=True, exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_dir / self.config.log_file),
                logging.StreamHandler()
            ]
        )
        return logging.getLogger(__name__)
    
    def run(self, skip_existing: bool = True, retry_failed: bool = True):
        """Run the extraction pipeline"""
        self.logger.info("=" * 60)
        self.logger.info("Starting SEC ADV Extraction Pipeline")
        self.logger.info("=" * 60)
        
        # Initialize clients
        self.cosmos_client.initialize()
        self.extractor.initialize()
        
        # Load checkpoint
        checkpoint = self.checkpoint_mgr.load_checkpoint()
        processed_crds = set(checkpoint.get("processed_crds", []))
        failed_crds = set(checkpoint.get("failed_crds", []))
        batch_files = checkpoint.get("batch_files", [])
        
        # Get all CRDs
        all_crds = self.cosmos_client.get_all_unique_crds()
        self.logger.info(f"Total CRDs in database: {len(all_crds)}")
        
        # Determine which CRDs to process
        crds_to_process = []
        for crd_info in all_crds:
            crd = crd_info['crd_number']
            if skip_existing and crd in processed_crds:
                continue
            if retry_failed or crd not in failed_crds:
                crds_to_process.append(crd_info)
        
        self.logger.info(f"CRDs to process: {len(crds_to_process)}")
        
        if not crds_to_process:
            self.logger.info("No CRDs to process. Generating final output...")
            self._generate_final_output()
            return
        
        # Process CRDs in batches
        current_batch = []
        batch_number = len(batch_files) + 1
        
        for idx, crd_info in enumerate(tqdm(crds_to_process, desc="Processing CRDs")):
            crd_number = crd_info['crd_number']
            firm_name = crd_info.get('firm_name', '')
            
            try:
                # Get document content for specific filing date
                filing_date = crd_info.get('filing_date', '')
                content, db_firm_name, db_filing_date = self.cosmos_client.get_document_by_crd(
                    crd_number, 
                    filing_date=filing_date
                )
                
                if not content:
                    self.logger.warning(f"No content found for CRD {crd_number}")
                    failed_crds.add(crd_number)
                    continue
                
                firm_name = firm_name or db_firm_name or f"Unknown_CRD_{crd_number}"
                filing_date = db_filing_date or filing_date or ""
                
                # Extract information
                extraction = self.extractor.extract_information(content, crd_number)
                
                if extraction:
                    record = {
                        "crd_number": crd_number,
                        "firm_name": firm_name,
                        "filing_date": filing_date,  # Add filing date
                        **extraction,
                        "status": "SUCCESS",
                        "extracted_at": datetime.now(timezone.utc).isoformat()
                    }
                    current_batch.append(record)
                    processed_crds.add(crd_number)
                    
                    # Remove from failed if it was there
                    failed_crds.discard(crd_number)
                else:
                    # Retry once more for failed extractions
                    self.logger.info(f"Retrying CRD {crd_number}...")
                    time.sleep(self.config.call_delay)
                    
                    extraction = self.extractor.extract_information(content, crd_number)
                    if extraction:
                        record = {
                            "crd_number": crd_number,
                            "firm_name": firm_name,
                            "filing_date": filing_date,  # Add filing date
                            **extraction,
                            "status": "SUCCESS_RETRY",
                            "extracted_at": datetime.now(timezone.utc).isoformat()
                        }
                        current_batch.append(record)
                        processed_crds.add(crd_number)
                        failed_crds.discard(crd_number)
                    else:
                        failed_crds.add(crd_number)
                        self.logger.error(f"Failed to extract CRD {crd_number} even after retry")
                
                # Delay between calls
                time.sleep(self.config.call_delay)
                
                # Save batch checkpoint
                if len(current_batch) >= self.config.batch_size:
                    batch_file = self.checkpoint_mgr.save_batch(current_batch, batch_number)
                    batch_files.append(batch_file)
                    
                    # Update checkpoint
                    checkpoint = {
                        "processed_crds": list(processed_crds),
                        "failed_crds": list(failed_crds),
                        "processed_count": len(processed_crds),
                        "batch_files": batch_files
                    }
                    self.checkpoint_mgr.save_checkpoint(checkpoint)
                    
                    current_batch = []
                    batch_number += 1
                    
                    self.logger.info(f"Progress: {len(processed_crds)}/{len(all_crds)} CRDs processed")
                    
            except Exception as e:
                self.logger.error(f"Error processing CRD {crd_number}: {e}")
                failed_crds.add(crd_number)
        
        # Save final batch
        if current_batch:
            batch_file = self.checkpoint_mgr.save_batch(current_batch, batch_number)
            batch_files.append(batch_file)
        
        # Final checkpoint save
        checkpoint = {
            "processed_crds": list(processed_crds),
            "failed_crds": list(failed_crds),
            "processed_count": len(processed_crds),
            "batch_files": batch_files
        }
        self.checkpoint_mgr.save_checkpoint(checkpoint)
        
        # Generate final output
        self._generate_final_output()
        
        # Summary
        self.logger.info("=" * 60)
        self.logger.info("EXTRACTION COMPLETE")
        self.logger.info("=" * 60)
        self.logger.info(f"Total processed: {len(processed_crds)}")
        self.logger.info(f"Total failed: {len(failed_crds)}")
        if failed_crds:
            self.logger.info(f"Failed CRDs: {list(failed_crds)[:20]}...")
    
    def _generate_final_output(self):
        """Generate final Excel file from all batches"""
        self.logger.info("Generating final Excel output...")
        
        # Load all batch data
        all_data = self.checkpoint_mgr.load_all_batches()
        
        if not all_data:
            self.logger.warning("No data to generate Excel file")
            return
        
        # Generate Excel
        output_path = Path(self.config.output_dir) / self.config.final_output
        self.excel_gen.generate_excel(all_data, str(output_path))
        
        self.logger.info(f"Final Excel file: {output_path}")
    
    def retry_failed_only(self):
        """Retry only the failed CRDs from previous run"""
        self.logger.info("Retrying failed CRDs only...")
        
        checkpoint = self.checkpoint_mgr.load_checkpoint()
        failed_crds = checkpoint.get("failed_crds", [])
        
        if not failed_crds:
            self.logger.info("No failed CRDs to retry")
            return
        
        self.logger.info(f"Retrying {len(failed_crds)} failed CRDs")
        
        # Run with retry_failed=True, which will process failed CRDs
        self.run(skip_existing=True, retry_failed=True)
    
    def run_test(self, num_crds: int = 3):
        """Run test mode with limited CRDs"""
        self.logger.info("=" * 60)
        self.logger.info(f"TEST MODE: Processing only {num_crds} CRDs")
        self.logger.info("=" * 60)
        
        # Initialize clients
        self.cosmos_client.initialize()
        self.extractor.initialize()
        
        # Get all CRDs but only process first N
        all_crds = self.cosmos_client.get_all_unique_crds()
        test_crds = all_crds[:num_crds]
        
        self.logger.info(f"Total CRDs in database: {len(all_crds)}")
        self.logger.info(f"Testing with CRDs: {[c['crd_number'] for c in test_crds]}")
        
        results = []
        
        for crd_info in tqdm(test_crds, desc="Processing test CRDs"):
            crd_number = crd_info['crd_number']
            firm_name = crd_info.get('firm_name', '')
            
            try:
                # Get document content
                filing_date = crd_info.get('filing_date', '')
                content, db_firm_name, db_filing_date = self.cosmos_client.get_document_by_crd(
                    crd_number,
                    filing_date=filing_date
                )
                
                if not content:
                    self.logger.warning(f"No content found for CRD {crd_number}")
                    continue
                
                firm_name = firm_name or db_firm_name or f"Unknown_CRD_{crd_number}"
                filing_date = db_filing_date or filing_date or ""
                
                self.logger.info(f"Processing CRD {crd_number} ({firm_name}) - Filing: {filing_date}")
                self.logger.info(f"Content length: {len(content)} characters")
                
                # Extract information
                extraction = self.extractor.extract_information(content, crd_number)
                
                if extraction:
                    record = {
                        "crd_number": crd_number,
                        "firm_name": firm_name,
                        "filing_date": filing_date,  # Add filing date
                        **extraction,
                        "status": "SUCCESS",
                        "extracted_at": datetime.now(timezone.utc).isoformat()
                    }
                    results.append(record)
                    
                    # Print extraction result
                    self.logger.info(f"Extraction result for CRD {crd_number}:")
                    for key, value in extraction.items():
                        preview = value[:100] + "..." if len(str(value)) > 100 else value
                        self.logger.info(f"  {key}: {preview}")
                else:
                    self.logger.error(f"Failed to extract CRD {crd_number}")
                
                time.sleep(self.config.call_delay)
                
            except Exception as e:
                self.logger.error(f"Error processing CRD {crd_number}: {e}")
                import traceback
                traceback.print_exc()
        
        # Save test results
        if results:
            test_output = Path(self.config.output_dir) / "test_results.json"
            test_output.parent.mkdir(parents=True, exist_ok=True)
            with open(test_output, 'w') as f:
                json.dump(results, f, indent=2)
            self.logger.info(f"Test results saved to {test_output}")
            
            # Also generate test Excel
            test_excel = Path(self.config.output_dir) / "test_results.xlsx"
            self.excel_gen.generate_excel(results, str(test_excel))
        
        self.logger.info("=" * 60)
        self.logger.info(f"TEST COMPLETE: {len(results)}/{num_crds} successful")
        self.logger.info("=" * 60)


def main():
    """Main entry point"""
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(description='SEC ADV Filing Information Extractor')
    parser.add_argument('--retry-failed', action='store_true', help='Retry failed CRDs only')
    parser.add_argument('--generate-excel', action='store_true', help='Generate Excel from existing batches')
    parser.add_argument('--test', type=int, metavar='N', help='Test mode: process only N CRDs')
    parser.add_argument('--chat-model', type=str, default='gpt-4o', help='Azure OpenAI chat model deployment name')
    parser.add_argument('--batch-size', type=int, default=50, help='Batch size for checkpointing')
    parser.add_argument('--output-dir', type=str, default='extraction_output', help='Output directory')
    
    args = parser.parse_args()
    
    config = Config()
    config.chat_deployment = args.chat_model
    config.batch_size = args.batch_size
    config.output_dir = args.output_dir
    
    pipeline = ADVExtractionPipeline(config)
    
    if args.retry_failed:
        pipeline.retry_failed_only()
    elif args.generate_excel:
        pipeline._generate_final_output()
    elif args.test:
        pipeline.run_test(args.test)
    else:
        pipeline.run()


if __name__ == "__main__":
    main()