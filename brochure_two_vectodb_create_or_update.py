"""
SEC ADV Filing Vector Database Builder for Azure Cosmos DB - With Retry Logic
Processes JSON documents, generates embeddings, and stores in Cosmos DB with vector search capabilities
"""

import json
import os
import logging
import hashlib
import tiktoken
from datetime import datetime, timezone
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
from pathlib import Path
import re
from dataclasses import dataclass, field
from tqdm import tqdm
import time
import httpx
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Azure imports
from azure.cosmos import CosmosClient, PartitionKey, exceptions
from azure.identity import DefaultAzureCredential
from azure.mgmt.cosmosdb import CosmosDBManagementClient
from azure.mgmt.cosmosdb.models import (
    DatabaseAccountCreateUpdateParameters,
    Location,
    DatabaseAccountKind,
    ConsistencyPolicy,
    DefaultConsistencyLevel
)
from openai import AzureOpenAI
import numpy as np

# Configuration
@dataclass
class Config:
    # Azure Configuration
    subscription_id: str = os.getenv("AZURE_SUBSCRIPTION_ID", "d59e2404-f3c7-4580-a8a0-32651ba1c1b1")
    resource_group: str = os.getenv("AZURE_RESOURCE_GROUP", "sec-form-adv-downloader-rg")
    location: str = os.getenv("AZURE_LOCATION", "eastus")

    # Cosmos DB Configuration
    cosmos_account_name: str = os.getenv("COSMOS_ACCOUNT_NAME", "sec-adv-cosmos")
    database_name: str = os.getenv("COSMOS_DATABASE_NAME", "Sec-adv-filing-part2-brochure-pdf")
    container_name: str = os.getenv("COSMOS_CONTAINER_NAME", "part2-json-data")

    # Azure OpenAI Configuration
    azure_openai_endpoint: str = os.getenv("AZURE_OPENAI_ENDPOINT", "https://model-deployments-hurricanecap-eastus2.openai.azure.com/")
    azure_openai_key: str = os.getenv("AZURE_OPENAI_KEY")  # Must be set via environment variable
    azure_openai_api_version: str = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01")
    embedding_deployment: str = os.getenv("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", "text-embedding-3-large")
    embedding_dimensions: int = int(os.getenv("AZURE_OPENAI_EMBEDDING_DIMENSIONS", "3072"))
    
    # Document Processing Configuration
    chunk_size: int = 1500
    chunk_overlap: int = 200
    batch_size: int = 20
    
    # Retry Configuration
    max_retries: int = 3
    retry_delay: int = 2
    connection_timeout: int = 60
    read_timeout: int = 120
    
    # Logging
    log_file: str = "vector_db_processing_report.log"

class DocumentProcessor:
    """Handles document chunking and metadata extraction"""
    
    def __init__(self, config: Config):
        self.config = config
        self.encoder = tiktoken.get_encoding("cl100k_base")
        self.logger = logging.getLogger(__name__)
    
    def extract_crd_from_filename(self, filename: str) -> Optional[int]:
        """Extract CRD number from filename"""
        match = re.match(r'^(\d+)_', filename)
        if match:
            return int(match.group(1))
        return None
    
    def count_tokens(self, text: str) -> int:
        """Count tokens in text"""
        return len(self.encoder.encode(text))
    
    def chunk_text(self, text: str, metadata: Dict) -> List[Dict]:
        """Chunk text with overlap while preserving document structure"""
        chunks = []
        text = re.sub(r'\s+', ' ', text).strip()
        sentences = re.split(r'(?<=[.!?])\s+', text)
        
        current_chunk = ""
        current_tokens = 0
        chunk_index = 0
        
        for sentence in sentences:
            sentence_tokens = self.count_tokens(sentence)
            
            if sentence_tokens > self.config.chunk_size:
                if current_chunk:
                    chunks.append(self._create_chunk_dict(
                        current_chunk, chunk_index, metadata
                    ))
                    chunk_index += 1
                
                words = sentence.split()
                temp_chunk = ""
                for word in words:
                    if self.count_tokens(temp_chunk + " " + word) <= self.config.chunk_size:
                        temp_chunk = (temp_chunk + " " + word).strip()
                    else:
                        if temp_chunk:
                            chunks.append(self._create_chunk_dict(
                                temp_chunk, chunk_index, metadata
                            ))
                            chunk_index += 1
                        temp_chunk = word
                
                if temp_chunk:
                    current_chunk = temp_chunk
                    current_tokens = self.count_tokens(temp_chunk)
            
            elif current_tokens + sentence_tokens <= self.config.chunk_size:
                current_chunk = (current_chunk + " " + sentence).strip()
                current_tokens += sentence_tokens
            else:
                chunks.append(self._create_chunk_dict(
                    current_chunk, chunk_index, metadata
                ))
                chunk_index += 1
                
                overlap_text = self._get_overlap_text(current_chunk)
                current_chunk = overlap_text + " " + sentence if overlap_text else sentence
                current_tokens = self.count_tokens(current_chunk)
        
        if current_chunk:
            chunks.append(self._create_chunk_dict(
                current_chunk, chunk_index, metadata
            ))
        
        return chunks
    
    def _get_overlap_text(self, text: str) -> str:
        """Get overlap text from the end of current chunk"""
        words = text.split()
        overlap_text = ""
        
        for word in reversed(words):
            temp = word + " " + overlap_text if overlap_text else word
            temp_tokens = self.count_tokens(temp)
            if temp_tokens <= self.config.chunk_overlap:
                overlap_text = temp
            else:
                break
        
        return overlap_text.strip()
    
    def _create_chunk_dict(self, text: str, index: int, metadata: Dict) -> Dict:
        """Create chunk dictionary with metadata"""
        return {
            "content": text,
            "chunk_index": index,
            "crd_number": metadata["crd_number"],
            "firm_name": metadata.get("firm_name", ""),
            "source_file": metadata["source_file"],
            "filing_date": metadata.get("filing_date", ""),
            "total_chunks": -1,
            "is_current": True,
            "processed_date": datetime.now(timezone.utc).isoformat()
        }

class VectorDBManager:
    """Manages Azure Cosmos DB operations with retry logic"""
    
    def __init__(self, config: Config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.cosmos_client = None
        self.database = None
        self.container = None
        self.openai_client = None
        
    def initialize(self):
        """Initialize Cosmos DB and Azure OpenAI clients"""
        self.logger.info("Initializing Azure services...")
        
        self._ensure_cosmos_account()
        
        endpoint = f"https://{self.config.cosmos_account_name}.documents.azure.com:443/"
        key = self._get_cosmos_key()
        self.cosmos_client = CosmosClient(endpoint, key)
        
        self._ensure_database_and_container()
        
        # Initialize OpenAI client with retry and timeout configuration
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
        
    def _ensure_cosmos_account(self):
        """Create Cosmos DB account if it doesn't exist"""
        credential = DefaultAzureCredential()
        cosmos_mgmt_client = CosmosDBManagementClient(
            credential, self.config.subscription_id
        )
        
        try:
            cosmos_mgmt_client.database_accounts.get(
                self.config.resource_group,
                self.config.cosmos_account_name
            )
            self.logger.info(f"Cosmos DB account {self.config.cosmos_account_name} exists")
        except Exception:
            self.logger.info(f"Creating Cosmos DB account {self.config.cosmos_account_name}...")
            
            params = DatabaseAccountCreateUpdateParameters(
                location=self.config.location,
                kind=DatabaseAccountKind.GLOBAL_DOCUMENT_DB,
                database_account_offer_type="Standard",
                locations=[Location(location_name=self.config.location)],
                consistency_policy=ConsistencyPolicy(
                    default_consistency_level=DefaultConsistencyLevel.SESSION
                ),
                capabilities=[{"name": "EnableNoSQLVectorSearch"}]
            )
            
            operation = cosmos_mgmt_client.database_accounts.begin_create_or_update(
                self.config.resource_group,
                self.config.cosmos_account_name,
                params
            )
            operation.wait()
            self.logger.info("Cosmos DB account created successfully")
    
    def _get_cosmos_key(self) -> str:
        """Get Cosmos DB account key"""
        credential = DefaultAzureCredential()
        cosmos_mgmt_client = CosmosDBManagementClient(
            credential, self.config.subscription_id
        )
        
        keys = cosmos_mgmt_client.database_accounts.list_keys(
            self.config.resource_group,
            self.config.cosmos_account_name
        )
        return keys.primary_master_key
    
    def _ensure_database_and_container(self):
        """Create database and container if they don't exist"""
        try:
            self.database = self.cosmos_client.create_database(
                id=self.config.database_name,
                offer_throughput=400
            )
            self.logger.info(f"Created database {self.config.database_name}")
        except exceptions.CosmosResourceExistsError:
            self.database = self.cosmos_client.get_database_client(self.config.database_name)
            self.logger.info(f"Database {self.config.database_name} already exists")
        
        try:
            self.container = self.database.create_container(
                id=self.config.container_name,
                partition_key=PartitionKey(path="/crd_number"),
                offer_throughput=1000
            )
            self.logger.info(f"Created container {self.config.container_name}")
        except exceptions.CosmosResourceExistsError:
            self.container = self.database.get_container_client(self.config.container_name)
            self.logger.info(f"Container {self.config.container_name} already exists")
    
    def generate_embeddings(self, texts: List[str]) -> List[List[float]]:
        """Generate embeddings with retry logic"""
        for attempt in range(self.config.max_retries):
            try:
                response = self.openai_client.embeddings.create(
                    model=self.config.embedding_deployment,
                    input=texts,
                    dimensions=self.config.embedding_dimensions
                )
                return [e.embedding for e in response.data]
            except Exception as e:
                self.logger.warning(f"Embedding attempt {attempt + 1} failed: {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (2 ** attempt))
                else:
                    self.logger.error(f"Failed to generate embeddings after {self.config.max_retries} attempts")
                    raise
    
    def delete_old_versions(self, crd_number: int):
        """Delete all existing documents for a CRD with retry logic"""
        for attempt in range(self.config.max_retries):
            try:
                query = f"SELECT * FROM c WHERE c.crd_number = {crd_number}"
                items = list(self.container.query_items(
                    query=query,
                    enable_cross_partition_query=True
                ))

                for item in items:
                    try:
                        self.container.delete_item(item=item['id'], partition_key=crd_number)
                    except exceptions.CosmosResourceNotFoundError:
                        # Item already deleted, ignore
                        pass

                self.logger.info(f"Deleted {len(items)} old documents for CRD {crd_number}")
                return
            except Exception as e:
                self.logger.warning(f"Delete attempt {attempt + 1} failed for CRD {crd_number}: {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (2 ** attempt))
                else:
                    self.logger.error(f"Failed to delete old versions for CRD {crd_number}")
                    # Continue anyway - we'll upsert the new documents
    
    def insert_documents(self, documents: List[Dict]):
        """Insert documents with embeddings into Cosmos DB with retry logic"""
        for batch_start in range(0, len(documents), self.config.batch_size):
            batch_end = min(batch_start + self.config.batch_size, len(documents))
            batch = documents[batch_start:batch_end]
            
            # Generate embeddings for batch with retry
            texts = [doc['content'] for doc in batch]
            embeddings = self.generate_embeddings(texts)
            
            # Add embeddings and insert with retry
            for doc, embedding in zip(batch, embeddings):
                doc['content_vector'] = embedding
                doc['id'] = f"{doc['crd_number']}_{doc['source_file']}_{doc['chunk_index']}"
                
                # Insert with retry logic
                for attempt in range(self.config.max_retries):
                    try:
                        self.container.upsert_item(body=doc)
                        break
                    except Exception as e:
                        if attempt < self.config.max_retries - 1:
                            self.logger.warning(f"Insert attempt {attempt + 1} failed for {doc['id']}: {e}")
                            time.sleep(self.config.retry_delay)
                        else:
                            self.logger.error(f"Failed to insert {doc['id']} after {self.config.max_retries} attempts")
                            raise
            
            self.logger.info(f"Inserted batch {batch_start//self.config.batch_size + 1}")
            time.sleep(0.5)

class ADVVectorDBBuilder:
    """Main orchestrator with retry capability for failed files"""
    
    def __init__(self, config: Config):
        self.config = config
        self.processor = DocumentProcessor(config)
        self.db_manager = VectorDBManager(config)
        self.logger = self._setup_logging()
        self.report = {
            "processed": [],
            "errors": [],
            "warnings": [],
            "summary": {}
        }
    
    def _setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.config.log_file),
                logging.StreamHandler()
            ]
        )
        return logging.getLogger(__name__)
    
    def load_csv_mapping(self, csv_path: str) -> pd.DataFrame:
        """Load CSV mapping file"""
        try:
            df = pd.read_csv(csv_path, encoding='cp1252')
            self.logger.info(f"Loaded CSV with {len(df)} entries")
            return df
        except Exception as e:
            self.logger.error(f"Error loading CSV: {e}")
            raise
    
    def process_directory(self, directory_path: str, retry_failed: bool = False, failed_files_list: List[str] = None):
        """Process all JSON files in directory or retry failed files"""

        self.db_manager.initialize()

        csv_files = list(Path(directory_path).glob("ADV_Brochure_Mapping_*.csv"))
        if not csv_files:
            raise FileNotFoundError(f"No CSV mapping file found")

        csv_path = str(csv_files[0])
        self.logger.info(f"Found CSV mapping file: {csv_path}")
        csv_df = self.load_csv_mapping(csv_path)

        # Determine which files to process
        if retry_failed and failed_files_list:
            # Only process the specific failed files
            json_files = [Path(f) for f in failed_files_list if Path(f).suffix == '.json']
            self.logger.info(f"RETRY MODE: Processing {len(json_files)} failed files ONLY")
            print(f"\n🔄 RETRY MODE: Processing {len(json_files)} failed files")
        else:
            # Process all JSON files in directory
            json_files = list(Path(directory_path).glob("*.json"))
            # Filter out CSV files
            json_files = [f for f in json_files if f.suffix == '.json']
            self.logger.info(f"FULL MODE: Processing {len(json_files)} files")
            print(f"\n📂 FULL MODE: Processing {len(json_files)} files")

        processed_crds = set()

        for json_file in tqdm(json_files, desc="Processing files"):
            try:
                self.process_single_file(json_file, csv_df, processed_crds)
            except Exception as e:
                self.logger.error(f"Error processing {json_file}: {e}")
                self.report["errors"].append({
                    "file": str(json_file),
                    "error": str(e)
                })

        # Only generate report if there were errors or if not in retry mode
        if not retry_failed or self.report["errors"]:
            self.generate_report()
        else:
            # In retry mode with no new errors - all retries succeeded
            self.logger.info("All retry attempts completed successfully - no new report generated")
            print("\n✅ All failed files processed successfully on retry!")
            print(f"   Processed: {len(self.report['processed'])} files")
            print(f"   No new errors to report")
    
    def process_single_file(self, json_path: Path, csv_df: pd.DataFrame, processed_crds: set):
        """Process a single JSON file with retry logic"""
        filename = json_path.name
        
        # Load JSON with retry
        for attempt in range(self.config.max_retries):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                break
            except (json.JSONDecodeError, IOError) as e:
                if attempt < self.config.max_retries - 1:
                    self.logger.warning(f"File read attempt {attempt + 1} failed for {filename}: {e}")
                    time.sleep(self.config.retry_delay)
                else:
                    self.logger.error(f"Failed to read {filename} after {self.config.max_retries} attempts")
                    self.report["errors"].append({"file": filename, "error": f"File read error: {e}"})
                    return
        
        content = data.get('content', '')
        metadata_info = data.get('metadata', {})
        source_file = metadata_info.get('source_file', filename)
        
        crd_number = self.processor.extract_crd_from_filename(filename)
        if not crd_number:
            self.logger.warning(f"Could not extract CRD from {filename}")
            self.report["warnings"].append({"file": filename, "warning": "Could not extract CRD number"})
            return
        
        pdf_filename = filename.replace('.json', '.pdf')
        csv_match = csv_df[csv_df['PDFFileName'] == pdf_filename]
        
        if csv_match.empty:
            self.logger.warning(f"No CSV match for {filename}")
            firm_name = f"Unknown_CRD_{crd_number}"
        else:
            firm_name = csv_match.iloc[0]['FirmName']
        
        if crd_number not in processed_crds:
            self.db_manager.delete_old_versions(crd_number)
            processed_crds.add(crd_number)
        
        metadata = {
            "crd_number": crd_number,
            "firm_name": firm_name,
            "source_file": source_file,
            "filing_date": csv_match.iloc[0]['DateFiled'] if not csv_match.empty else "",
            "filing_id": int(csv_match.iloc[0]['FilingID']) if not csv_match.empty else None
        }
        
        chunks = self.processor.chunk_text(content, metadata)
        
        for chunk in chunks:
            chunk['total_chunks'] = len(chunks)
        
        self.db_manager.insert_documents(chunks)
        
        self.report["processed"].append({
            "file": filename,
            "crd": crd_number,
            "firm_name": firm_name,
            "chunks_created": len(chunks)
        })
        
        self.logger.info(f"Processed {filename}: {len(chunks)} chunks")
    
    def generate_report(self):
        """Generate processing report"""
        self.report["summary"] = {
            "total_files_processed": len(self.report["processed"]),
            "total_errors": len(self.report["errors"]),
            "total_warnings": len(self.report["warnings"]),
            "timestamp": datetime.now(timezone.utc).isoformat()
        }
        
        report_path = f"processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(report_path, 'w') as f:
            json.dump(self.report, f, indent=2)
        
        self.logger.info(f"Report saved to {report_path}")
        
        print("\n" + "="*50)
        print("PROCESSING SUMMARY")
        print("="*50)
        print(f"Total files processed: {self.report['summary']['total_files_processed']}")
        print(f"Total errors: {self.report['summary']['total_errors']}")
        print(f"Total warnings: {self.report['summary']['total_warnings']}")
        
        if self.report["errors"]:
            print("\nERRORS:")
            for error in self.report["errors"]:
                print(f"  - {error['file']}: {error['error']}")
    
    def retry_failed_files(self, previous_report_path: str, directory_path: str):
        """Retry processing of failed files from a previous run"""
        self.logger.info(f"Loading previous report: {previous_report_path}")
        
        with open(previous_report_path, 'r') as f:
            previous_report = json.load(f)
        
        # Get full paths for failed files
        failed_files = []
        for error in previous_report.get('errors', []):
            file_path = error['file']
            # If it's just a filename, prepend the directory path
            if not file_path.startswith('/'):
                file_path = os.path.join(directory_path, os.path.basename(file_path))
            
            if os.path.exists(file_path):
                failed_files.append(file_path)
            else:
                self.logger.warning(f"Failed file not found: {file_path}")
        
        if not failed_files:
            self.logger.info("No failed files to retry")
            print("\n✅ No failed files found to retry!")
            return
        
        self.logger.info(f"Retrying {len(failed_files)} failed files")
        print(f"\n{'='*60}")
        print(f"RETRY MODE - Processing {len(failed_files)} failed files ONLY")
        print(f"{'='*60}")
        print(f"\n📋 Failed files to retry:")
        for f in failed_files:
            print(f"  - {os.path.basename(f)}")
        print()
        
        # Process only failed files
        self.process_directory(directory_path, retry_failed=True, failed_files_list=failed_files)


def main():
    """Main entry point"""
    import sys

    if len(sys.argv) < 2:
        print("Usage:")
        print("  Process all files: python script.py /path/to/data/directory")
        print("  Retry failed: python script.py /path/to/data/directory --retry /path/to/previous_report.json")
        print("\nRequired environment variable:")
        print("  AZURE_OPENAI_KEY - Your Azure OpenAI API key")
        sys.exit(1)

    directory_path = sys.argv[1]
    config = Config()

    # Validate required configuration
    if not config.azure_openai_key:
        print("ERROR: AZURE_OPENAI_KEY environment variable is required")
        print("Please set it before running this script:")
        print("  export AZURE_OPENAI_KEY='your-api-key'")
        sys.exit(1)

    builder = ADVVectorDBBuilder(config)
    
    # Check if retry mode
    if len(sys.argv) == 4 and sys.argv[2] == '--retry':
        previous_report = sys.argv[3]
        builder.retry_failed_files(previous_report, directory_path)
    else:
        builder.process_directory(directory_path)


if __name__ == "__main__":
    main()