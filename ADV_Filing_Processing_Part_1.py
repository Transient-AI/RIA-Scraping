#!/usr/bin/env python3
"""
RIA Data Processor
Automates downloading and processing of SEC Form ADV data using functional programming
"""
import os
import sys
import yaml
import logging
import requests
import zipfile
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from urllib.parse import urljoin
import traceback
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import hashlib

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def load_config(config_path: str = "config.yaml") -> Tuple[Dict, List[Dict]]:
    """Load configuration from YAML file"""
    try:
        with open(config_path, 'r') as f:
            config = yaml.safe_load(f)
        logger.info(f"Configuration loaded from {config_path}")
        return config, []
    except Exception as e:
        logger.error(f"Failed to load configuration: {e}")
        sys.exit(1)

def setup_directories(config: Dict, error_log: List[Dict]) -> List[Dict]:
    """Create necessary directories if they don't exist"""
    directories = [
        config['paths']['raw_data'],
        config['paths']['processed_data'],
        config['paths']['reports'],
        config['paths']['archive']
    ]
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
    logger.info("Directory structure initialized")
    return error_log

def get_file_hash(filepath: str) -> str:
    """Calculate MD5 hash of a file for change detection"""
    hash_md5 = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def check_if_processed(config: Dict, month: str, year: int, error_log: List[Dict]) -> Tuple[bool, List[Dict]]:
    """Check if data for given month/year has already been processed"""
    processed_file = os.path.join(
        config['paths']['processed_data'],
        f"RIA_Data_{year}_{month}.xlsx"
    )
    metadata_file = os.path.join(
        config['paths']['processed_data'],
        f"metadata_{year}_{month}.json"
    )

    if os.path.exists(processed_file) and os.path.exists(metadata_file):
        with open(metadata_file, 'r') as f:
            metadata = json.load(f)

        if config['processing']['force_refresh']:
            logger.info(f"Force refresh enabled, reprocessing {month} {year}")
            return False, error_log

        logger.info(f"Data for {month} {year} already processed")
        return not config['processing']['incremental_update'], error_log

    return False, error_log

def download_file(url: str, filepath: str, error_log: List[Dict]) -> List[Dict]:
    """Download a file from URL to filepath"""
    try:
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()

        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        logger.info(f"Downloaded: {filepath}")
        return error_log
    except requests.exceptions.RequestException as e:
        error_msg = f"Download failed for {url}: {str(e)}"
        logger.error(error_msg)
        return error_log + [{
            'timestamp': datetime.now(),
            'function': 'download_file',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]

def download_adv_files(config: Dict, month: str, year: int, error_log: List[Dict]) -> Tuple[Optional[Tuple[str, str]], List[Dict]]:
    """Download Form ADV Part 1 and Part 2 files for specified month"""
    try:
        # Calculate date range for Part 1 (needs date format)
        month_num = datetime.strptime(month, "%B").month
        
        first_day = datetime(year, month_num, 1)
        if month_num == 12:
            last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            last_day = datetime(year, month_num + 1, 1) - timedelta(days=1)
        
        date_format = f"{first_day.strftime('%Y%m%d')}_{last_day.strftime('%Y%m%d')}"
        
        # Headers to avoid 403 errors
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Referer': 'https://adviserinfo.sec.gov/'
        }
        
        # Download Part 1 - ADV Filing Data
        part1_base_url = "https://reports.adviserinfo.sec.gov/reports/foia/advFilingData"
        part1_filename = f"ADV_Filing_Data_{date_format}.zip"
        part1_url = f"{part1_base_url}/{year}/{part1_filename}"
        part1_path = os.path.join(config['paths']['raw_data'], part1_filename)
        
        logger.info(f"Downloading Part 1 from: {part1_url}")
        
        response = requests.get(part1_url, headers=headers, stream=True, timeout=60)
        response.raise_for_status()
        
        with open(part1_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        logger.info(f"Downloaded Part 1: {part1_path} ({os.path.getsize(part1_path) / 1024 / 1024:.2f} MB)")
        
     
        
        return part1_path, error_log

        
    except Exception as e:
        error_msg = f"Failed to download ADV files: {str(e)}"
        logger.error(error_msg)
        return None, error_log + [{
            'timestamp': datetime.now(),
            'function': 'download_adv_files',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]
    
    

def extract_part1_files(config: Dict, zip_path: str, error_log: List[Dict]) -> Tuple[Optional[Dict[str, pd.DataFrame]], List[Dict]]:
    """Extract and load the three required CSV files from Part 1 zip"""
    try:
        extracted_data = {}

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            file_list = zip_ref.namelist()

            files_to_extract = {
                'base': 'IA_ADV_Base_A_',
                'custodian': 'IA_Schedule_D_5K3_',
                'owners': 'IA_Schedule_A_B_'
            }

            for key, pattern in files_to_extract.items():
                matching_file = [f for f in file_list if pattern in f and f.endswith('.csv')]

                if matching_file:
                    file_name = matching_file[0]
                    logger.info(f"Extracting {key}: {file_name}")

                    zip_ref.extract(file_name, config['paths']['raw_data'])
                    file_path = os.path.join(config['paths']['raw_data'], file_name)

                    df = pd.read_csv(
                        file_path,
                        encoding='latin-1',
                        low_memory=False,
                        on_bad_lines='warn'
                    )

                    extracted_data[key] = df
                    logger.info(f"Loaded {key}: {len(df)} records")

                    if config['processing'].get('cleanup_raw_files', False):
                        os.remove(file_path)
                else:
                    error_msg = f"Required file pattern '{pattern}' not found in zip"
                    logger.error(error_msg)
                    error_log = error_log + [{
                        'timestamp': datetime.now(),
                        'function': 'extract_part1_files',
                        'error': error_msg
                    }]

        return extracted_data, error_log

    except Exception as e:
        error_msg = f"Failed to extract Part 1 files: {str(e)}"
        logger.error(error_msg)
        return None, error_log + [{
            'timestamp': datetime.now(),
            'function': 'extract_part1_files',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]

def process_base_data(df: pd.DataFrame, error_log: List[Dict]) -> Tuple[pd.DataFrame, List[Dict]]:
    """Process base ADV data with column mappings"""
    try:
        # Debug: Check for the CRD column with various possible names
        logger.info(f"Checking columns around position K (index 10): {df.columns[8:15].tolist()}")
        
        # The CRD column might be read as '10', 10, 10.0, or '1.00E+01'
        crd_column = None
        possible_crd_names = ['10', 10, 10.0, '10.0', '1.00E+01', '1.0E+01', '1E+01']
        
        for col in possible_crd_names:
            if col in df.columns:
                crd_column = col
                logger.info(f"Found CRD column as: {col}")
                break
        
        # If not found by name, check by position (column K is index 10)
        if crd_column is None and len(df.columns) > 10:
            potential_crd = df.columns[10]  # Column K (0-indexed as 10)
            logger.info(f"Column at position K (index 10) is: {potential_crd}")
            # Check if this column contains CRD-like data
            sample_values = df[potential_crd].dropna().head(5)
            if sample_values.apply(lambda x: str(x).replace('.0', '').isdigit() and len(str(x).replace('.0', '')) in range(5, 9)).all():
                crd_column = potential_crd
                logger.info(f"Using column K as CRD column: {potential_crd}")
        
        column_mapping = {
            'FilingID': 'Filing_ID',
            'FormVersion': 'Form_Version',
            'DateSubmitted': 'Latest_ADV_Filing_Date',
            '1A': 'Legal_Name',
            '1B1': 'Primary_Business_Name',
            '1B2': 'More_Than_One_IA',
            '1C-Legal': 'New_Legal_Name',
            '1C-Business': 'New_Primary_Business_Name',
            '1C-New Name': 'New_Legal_Name_Alt',
            '1D': 'SEC_ID',
            '1F1-Street 1': 'RIA_Street1',
            '1F1-Street 2': 'RIA_Street2',
            '1F1-City': 'RIA_City',
            '1F1-State': 'RIA_State',
            '1F1-Country': 'RIA_Country',
            '1F1-Postal': 'RIA_Postal'
        }
        
        # Add the CRD column mapping dynamically
        if crd_column is not None:
            column_mapping[crd_column] = 'CRD'
            logger.info(f"Added CRD mapping: {crd_column} -> CRD")
        else:
            logger.warning("CRD column not found!")

        available_columns = [col for col in column_mapping.keys() if col in df.columns]
        processed_df = df[available_columns].rename(columns=column_mapping)

        # Convert CRD to integer if it exists
        if 'CRD' in processed_df.columns:
            processed_df['CRD'] = pd.to_numeric(processed_df['CRD'], errors='coerce').fillna(0).astype('Int64')
            logger.info(f"CRD column processed with {processed_df['CRD'].notna().sum()} non-null values")

        if 'Latest_ADV_Filing_Date' in processed_df.columns:
            processed_df = processed_df.assign(
                Latest_ADV_Filing_Date=pd.to_datetime(
                    processed_df['Latest_ADV_Filing_Date'],
                    errors='coerce'
                )
            )

        text_columns = ['Legal_Name', 'Primary_Business_Name', 'RIA_City', 'RIA_State']
        for col in text_columns:
            if col in processed_df.columns:
                processed_df = processed_df.assign(**{col: processed_df[col].str.strip().str.upper()})

        logger.info(f"Processed base data: {len(processed_df)} records")
        logger.info(f"Final columns: {processed_df.columns.tolist()}")
        
        return processed_df, error_log

    except Exception as e:
        error_msg = f"Failed to process base data: {str(e)}"
        logger.error(error_msg)
        return pd.DataFrame(), error_log + [{
            'timestamp': datetime.now(),
            'function': 'process_base_data',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]
    
    

    
    
def process_custodian_data(df: pd.DataFrame, error_log: List[Dict]) -> Tuple[pd.DataFrame, List[Dict]]:
    """Process custodian data with column mappings"""
    try:
        # Handle both "FilingID" and "Filing ID" column names
        if 'Filing ID' in df.columns:
            df = df.rename(columns={'Filing ID': 'FilingID'})
            logger.info("Renamed 'Filing ID' to 'FilingID' for consistency")
        
        column_mapping = {
            'FilingID': 'Filing_ID',
            '5K(3)(a)': 'Legal_Name_Custodian',
            '5K(3)(b)': 'Business_Name_Custodian',
            '5K(3)(c) City': 'Address_City',
            '5K(3)(c) State': 'Address_State',
            '5K(3)(c) Country': 'Address_Country',
            '5K(3)(d)': 'Is_Custodian_Related_To_Firm',
            '5K(3)(e)': 'SEC_Registration_Number',
            '5K(3)(f)': 'Custodian_LEI',
            '5K(3)(g)': 'AUM'
        }

        available_columns = [col for col in column_mapping.keys() if col in df.columns]
        processed_df = df[available_columns].rename(columns=column_mapping)

        if 'AUM' in processed_df.columns:
            processed_df = processed_df.assign(AUM=pd.to_numeric(processed_df['AUM'], errors='coerce'))

        if 'Is_Custodian_Related_To_Firm' in processed_df.columns:
            processed_df = processed_df.assign(
                Is_Custodian_Related_To_Firm=processed_df['Is_Custodian_Related_To_Firm'].map(
                    {'Y': True, 'N': False}
                )
            )
            
            
        # Add broker dealer flag based on SEC Registration Number
        if 'SEC_Registration_Number' in processed_df.columns:
            processed_df = processed_df.assign(
                is_broker_dealer=processed_df['SEC_Registration_Number'].notna() & 
                                (processed_df['SEC_Registration_Number'] != '') &
                                (processed_df['SEC_Registration_Number'].astype(str).str.strip() != '')
            )
            logger.info(f"Added broker dealer flag: {processed_df['is_broker_dealer'].sum()} records marked as broker dealers")
        else:
            logger.warning("SEC_Registration_Number column not found, broker dealer flag not added")
            
        
        

        logger.info(f"Processed custodian data: {len(processed_df)} records")
        logger.info(f"Custodian columns: {processed_df.columns.tolist()}")
        return processed_df, error_log

    except Exception as e:
        error_msg = f"Failed to process custodian data: {str(e)}"
        logger.error(error_msg)
        return pd.DataFrame(), error_log + [{
            'timestamp': datetime.now(),
            'function': 'process_custodian_data',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]

def process_owners_data(df: pd.DataFrame, error_log: List[Dict]) -> Tuple[pd.DataFrame, List[Dict]]:
    """Process owners data with column mappings and code translations"""
    try:
        # Handle both "FilingID" and "Filing ID" column names
        if 'Filing ID' in df.columns:
            df = df.rename(columns={'Filing ID': 'FilingID'})
            logger.info("Renamed 'Filing ID' to 'FilingID' for consistency")
            
        column_mapping = {
            'FilingID': 'Filing_ID',
            'SchA-3': 'Indirect_Owners_In_SchB',
            'Schedule': 'Direct_Indirect_Owners',
            'Full Legal Name': 'Full_Legal_Name',
            'DE/FE/I': 'Owner_Type',
            'Entity in Which': 'Entity',
            'Title or Status': 'Title_Or_Status',
            'Status Acquired': 'Acquired_Date',
            'Ownership Code': 'Ownership_Code',
            'Control Person': 'Control_Person',
            'PR': 'Public_Reporting',
            'OwnerID': 'Owner_ID'
        }

        available_columns = [col for col in column_mapping.keys() if col in df.columns]
        processed_df = df[available_columns].rename(columns=column_mapping)

        ownership_mapping = {
            'NA': 'Less than 5%',
            'A': '5-10%',
            'B': '10-25%',
            'C': '25-50%',
            'D': '50-75%',
            'E': '75% or more'
        }

        if 'Ownership_Code' in processed_df.columns:
            processed_df = processed_df.assign(
                Ownership_Percentage=processed_df['Ownership_Code'].map(ownership_mapping)
            )

        owner_type_mapping = {
            'DE': 'Domestic Entity',
            'FE': 'Foreign Entity',
            'I': 'Individual'
        }

        if 'Owner_Type' in processed_df.columns:
            processed_df = processed_df.assign(
                Owner_Type_Description=processed_df['Owner_Type'].map(owner_type_mapping)
            )

        schedule_mapping = {
            'A': 'Direct Owner',
            'B': 'Indirect Owner'
        }

        if 'Direct_Indirect_Owners' in processed_df.columns:
            processed_df = processed_df.assign(
                Owner_Classification=processed_df['Direct_Indirect_Owners'].map(schedule_mapping)
            )

        if 'Acquired_Date' in processed_df.columns:
            processed_df = processed_df.assign(
                Acquired_Date=pd.to_datetime(processed_df['Acquired_Date'], errors='coerce')
            )

        logger.info(f"Processed owners data: {len(processed_df)} records")
        logger.info(f"Owners columns: {processed_df.columns.tolist()}")
        return processed_df, error_log

    except Exception as e:
        error_msg = f"Failed to process owners data: {str(e)}"
        logger.error(error_msg)
        return pd.DataFrame(), error_log + [{
            'timestamp': datetime.now(),
            'function': 'process_owners_data',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]



def create_summary_statistics(data_dict: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Create summary statistics for the processed data"""
    summary_data = [
        {
            'Sheet': sheet_name,
            'Total_Records': len(df),
            'Columns': len(df.columns),
            'Memory_Usage_MB': df.memory_usage(deep=True).sum() / 1024 / 1024,
            'Null_Values': df.isnull().sum().sum(),
            'Duplicate_Rows': df.duplicated().sum()
        }
        for sheet_name, df in data_dict.items() if not df.empty
    ]
    return pd.DataFrame(summary_data)

def save_to_excel(config: Dict, data_dict: Dict[str, pd.DataFrame], month: str, year: int, error_log: List[Dict]) -> List[Dict]:
    """Save processed data to Excel workbook with formatting"""
    try:
        output_file = os.path.join(
            config['paths']['processed_data'],
            f"RIA_Data_{year}_{month}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in data_dict.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    worksheet = writer.sheets[sheet_name]

                    for cell in worksheet[1]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

        logger.info(f"Data saved to: {output_file}")

        metadata = {
            'processing_date': datetime.now().isoformat(),
            'month': month,
            'year': year,
            'sheets': list(data_dict.keys()),
            'record_counts': {k: len(v) for k, v in data_dict.items()},
            'file_hash': get_file_hash(output_file)
        }

        metadata_file = os.path.join(
            config['paths']['processed_data'],
            f"metadata_{year}_{month}.json"
        )

        with open(metadata_file, 'w') as f:
            json.dump(metadata, f, indent=2)

        return error_log

    except Exception as e:
        error_msg = f"Failed to save Excel file: {str(e)}"
        logger.error(error_msg)
        return error_log + [{
            'timestamp': datetime.now(),
            'function': 'save_to_excel',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]

def generate_error_report(config: Dict, month: str, year: int, processing_stats: Dict, error_log: List[Dict]) -> List[Dict]:
    """Generate comprehensive error report"""
    try:
        report_file = os.path.join(
            config['paths']['reports'],
            f"error_report_{year}_{month}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )

        report = {
            'processing_info': {
                'month': month,
                'year': year,
                'start_time': processing_stats['start_time'].isoformat(),
                'end_time': datetime.now().isoformat(),
                'duration_seconds': (datetime.now() - processing_stats['start_time']).total_seconds(),
                'files_processed': processing_stats['files_processed'],
                'records_processed': processing_stats['records_processed']
            },
            'errors': error_log,
            'error_count': len(error_log),
            'status': 'SUCCESS' if len(error_log) == 0 else 'COMPLETED_WITH_ERRORS'
        }

        with open(report_file, 'w') as f:
            json.dump(report, f, indent=2, default=str)

        logger.info(f"Error report generated: {report_file}")

        if error_log:
            error_df = pd.DataFrame(error_log)
            error_excel = os.path.join(
                config['paths']['reports'],
                f"error_summary_{year}_{month}.xlsx"
            )
            error_df.to_excel(error_excel, index=False)

        return error_log

    except Exception as e:
        logger.error(f"Failed to generate error report: {str(e)}")
        return error_log + [{
            'timestamp': datetime.now(),
            'function': 'generate_error_report',
            'error': f"Failed to generate error report: {str(e)}",
            'traceback': traceback.format_exc()
        }]


def get_latest_filing_ids(base_df: pd.DataFrame, error_log: List[Dict]) -> Tuple[pd.DataFrame, List[Dict]]:
    """
    Filter base data to only keep records with the latest filing date per CRD.
    When multiple Filing_IDs exist for the same date, pick the largest Filing_ID.
    Returns DataFrame with only the latest filings
    """
    try:
        # Ensure we have the necessary columns
        if 'CRD' not in base_df.columns or 'Latest_ADV_Filing_Date' not in base_df.columns:
            logger.warning("CRD or Latest_ADV_Filing_Date column missing, returning all records")
            return base_df, error_log
        
        # Remove rows where CRD is null or 0
        filtered_df = base_df[base_df['CRD'].notna() & (base_df['CRD'] != 0)].copy()
        
        # Convert Filing_ID to numeric for proper comparison (in case it's string)
        if 'Filing_ID' in filtered_df.columns:
            filtered_df['Filing_ID_numeric'] = pd.to_numeric(filtered_df['Filing_ID'], errors='coerce')
        else:
            logger.warning("Filing_ID column not found")
            return base_df, error_log
        
        # First, get the latest filing date for each CRD
        latest_dates = filtered_df.groupby('CRD')['Latest_ADV_Filing_Date'].max().reset_index()
        latest_dates.columns = ['CRD', 'Max_Filing_Date']
        
        # Merge to get all records with the latest filing date per CRD
        latest_date_filings = filtered_df.merge(
            latest_dates, 
            left_on=['CRD', 'Latest_ADV_Filing_Date'],
            right_on=['CRD', 'Max_Filing_Date'],
            how='inner'
        )
        
        # Now, for each CRD with the latest date, get the maximum Filing_ID
        # This handles cases where multiple filings exist on the same date
        latest_filing_ids = latest_date_filings.groupby('CRD')['Filing_ID_numeric'].max().reset_index()
        latest_filing_ids.columns = ['CRD', 'Max_Filing_ID_numeric']
        
        # Final merge to get only the records with both latest date AND largest Filing_ID
        latest_filings = latest_date_filings.merge(
            latest_filing_ids,
            left_on=['CRD', 'Filing_ID_numeric'],
            right_on=['CRD', 'Max_Filing_ID_numeric'],
            how='inner'
        )
        
        # Drop the temporary columns
        columns_to_drop = ['Max_Filing_Date', 'Filing_ID_numeric', 'Max_Filing_ID_numeric']
        latest_filings = latest_filings.drop(columns=[col for col in columns_to_drop if col in latest_filings.columns], axis=1)
        
        # Remove any duplicate rows (in case of exact duplicates)
        latest_filings = latest_filings.drop_duplicates()
        
        # Log statistics
        original_count = len(base_df)
        filtered_count = len(latest_filings)
        unique_crds = latest_filings['CRD'].nunique()
        
        logger.info(f"Filtered to latest filings: {original_count} -> {filtered_count} records")
        logger.info(f"Unique CRDs with latest filings: {unique_crds}")
        
        # Log some examples of the filtering
        if len(latest_filings) > 0:
            sample_crd = latest_filings['CRD'].iloc[0]
            sample_records = base_df[base_df['CRD'] == sample_crd][['CRD', 'Filing_ID', 'Latest_ADV_Filing_Date']].head()
            logger.debug(f"Example - Original records for CRD {sample_crd}:\n{sample_records}")
            selected_record = latest_filings[latest_filings['CRD'] == sample_crd][['CRD', 'Filing_ID', 'Latest_ADV_Filing_Date']].head()
            logger.debug(f"Selected record for CRD {sample_crd}:\n{selected_record}")
        
        return latest_filings, error_log
        
    except Exception as e:
        error_msg = f"Failed to filter latest filings: {str(e)}"
        logger.error(error_msg)
        return base_df, error_log + [{
            'timestamp': datetime.now(),
            'function': 'get_latest_filing_ids',
            'error': error_msg,
            'traceback': traceback.format_exc()
        }]

def run(config: Dict, error_log: List[Dict], processing_stats: Dict) -> Tuple[Dict, List[Dict]]:
    """Main execution function"""
    try:
        month = config['processing']['month']
        year = config['processing']['year']

        logger.info(f"Starting RIA data processing for {month} {year}")

        # Check if already processed
        is_processed, error_log = check_if_processed(config, month, year, error_log)
        if is_processed and not config['processing']['force_refresh']:
            logger.info("Data already processed and incremental update not required")
            return processing_stats, error_log

        # Check if using existing ZIP or downloading new one
        if config['processing'].get('use_existing_zip', False):
            # Use existing ZIP file
            part1_zip = config['processing'].get('existing_zip_path', '')
            if not part1_zip:
                raise Exception("use_existing_zip is True but existing_zip_path is not set in config.yaml")
            if not os.path.exists(part1_zip):
                raise Exception(f"Existing ZIP file not found: {part1_zip}")
            logger.info(f"Using existing ZIP file: {part1_zip}")
        else:
            # Download files
            logger.info("Downloading ADV files...")
            file_paths, error_log = download_adv_files(config, month, year, error_log)
            if file_paths is None:
                raise Exception("Download failed")
            part1_zip = file_paths

        # Extract and process Part 1 files
        logger.info("Extracting Part 1 files...")
        part1_data, error_log = extract_part1_files(config, part1_zip, error_log)
        if part1_data is None:
            raise Exception("Extraction failed")

        # Process each dataset
        processed_data = {}
        processing_stats = processing_stats.copy()  # Ensure immutability

        # Process base data first to get CRD mapping and latest filing IDs
        base_data_df = pd.DataFrame()
        crd_mapping = pd.DataFrame()
        latest_filing_ids = set()  # Track Filing_IDs that represent latest filings
        
        if 'base' in part1_data:
            logger.info("Processing base data...")
            base_data_df, error_log = process_base_data(part1_data['base'], error_log)
            
            # Filter to only latest filings
            logger.info("Filtering to latest filings only...")
            base_data_df, error_log = get_latest_filing_ids(base_data_df, error_log)
            
            # Store the Filing_IDs that represent latest filings
            if 'Filing_ID' in base_data_df.columns:
                latest_filing_ids = set(base_data_df['Filing_ID'].unique())
                logger.info(f"Identified {len(latest_filing_ids)} latest Filing_IDs to keep")
            
            processed_data['Base_Data'] = base_data_df
            
            # Extract CRD mapping for use in other sheets
            # Extract CRD mapping with filing date for use in other sheets
            if 'Filing_ID' in base_data_df.columns and 'CRD' in base_data_df.columns:
                crd_mapping_cols = ['Filing_ID', 'CRD']
                if 'Latest_ADV_Filing_Date' in base_data_df.columns:
                    crd_mapping_cols.append('Latest_ADV_Filing_Date')
                crd_mapping = base_data_df[crd_mapping_cols].drop_duplicates()
                logger.info(f"Extracted CRD mapping for {len(crd_mapping)} unique Filing IDs")
            
            
            
            processing_stats = {
                **processing_stats,
                'files_processed': processing_stats['files_processed'] + 1,
                'records_processed': processing_stats['records_processed'] + len(base_data_df)
            }

        if 'custodian' in part1_data:
            logger.info("Processing custodian data...")
            custodian_df, error_log = process_custodian_data(part1_data['custodian'], error_log)
            
            # Filter to only latest Filing_IDs
            if latest_filing_ids and 'Filing_ID' in custodian_df.columns:
                original_count = len(custodian_df)
                custodian_df = custodian_df[custodian_df['Filing_ID'].isin(latest_filing_ids)]
                logger.info(f"Filtered custodian data: {original_count} -> {len(custodian_df)} records")
            
            # Merge CRD numbers if available
            # Merge CRD numbers and filing date if available
            if not crd_mapping.empty and 'Filing_ID' in custodian_df.columns:
                logger.info("Adding CRD numbers and filing date to custodian data...")
                original_cols = custodian_df.columns.tolist()
                custodian_df = custodian_df.merge(crd_mapping, on='Filing_ID', how='left')
                # Reorder columns to put CRD and filing date after Filing_ID
                base_cols = ['Filing_ID', 'CRD']
                if 'Latest_ADV_Filing_Date' in crd_mapping.columns:
                    base_cols.append('Latest_ADV_Filing_Date')
                cols = base_cols + [col for col in original_cols if col != 'Filing_ID']
                custodian_df = custodian_df[[col for col in cols if col in custodian_df.columns]]
                logger.info(f"Added CRD to {custodian_df['CRD'].notna().sum()} custodian records")
            
            processed_data['Custodian_Data'] = custodian_df
            processing_stats = {
                **processing_stats,
                'files_processed': processing_stats['files_processed'] + 1,
                'records_processed': processing_stats['records_processed'] + len(custodian_df)
            }

        if 'owners' in part1_data:
            logger.info("Processing owners data...")
            owners_df, error_log = process_owners_data(part1_data['owners'], error_log)
            
            # Filter to only latest Filing_IDs
            if latest_filing_ids and 'Filing_ID' in owners_df.columns:
                original_count = len(owners_df)
                owners_df = owners_df[owners_df['Filing_ID'].isin(latest_filing_ids)]
                logger.info(f"Filtered owners data: {original_count} -> {len(owners_df)} records")
            
            # Merge CRD numbers if available
            # Merge CRD numbers and filing date if available
            if not crd_mapping.empty and 'Filing_ID' in owners_df.columns:
                logger.info("Adding CRD numbers and filing date to owners data...")
                original_cols = owners_df.columns.tolist()
                owners_df = owners_df.merge(crd_mapping, on='Filing_ID', how='left')
                # Reorder columns to put CRD and filing date after Filing_ID
                base_cols = ['Filing_ID', 'CRD']
                if 'Latest_ADV_Filing_Date' in crd_mapping.columns:
                    base_cols.append('Latest_ADV_Filing_Date')
                cols = base_cols + [col for col in original_cols if col != 'Filing_ID']
                owners_df = owners_df[[col for col in cols if col in owners_df.columns]]
                logger.info(f"Added CRD to {owners_df['CRD'].notna().sum()} owner records")
            
            processed_data['Owners_Data'] = owners_df
            processing_stats = {
                **processing_stats,
                'files_processed': processing_stats['files_processed'] + 1,
                'records_processed': processing_stats['records_processed'] + len(owners_df)
            }

 

        # Create summary statistics
        processed_data['Summary_Statistics'] = create_summary_statistics(processed_data)
        
        # Add CRD coverage statistics
        crd_coverage_data = []
        for sheet_name, df in processed_data.items():
            if not df.empty and 'CRD' in df.columns and sheet_name != 'Summary_Statistics':
                crd_coverage_data.append({
                    'Sheet': sheet_name,
                    'Total_Records': len(df),
                    'Records_with_CRD': df['CRD'].notna().sum(),
                    'CRD_Coverage_%': round((df['CRD'].notna().sum() / len(df) * 100), 2) if len(df) > 0 else 0
                })
        
        if crd_coverage_data:
            processed_data['CRD_Coverage_Stats'] = pd.DataFrame(crd_coverage_data)
        
        # Add filtering statistics
        filtering_stats_data = []
        if latest_filing_ids:
            filtering_stats_data.append({
                'Metric': 'Total Latest Filing IDs',
                'Value': len(latest_filing_ids)
            })
            filtering_stats_data.append({
                'Metric': 'Unique CRDs with Latest Filings',
                'Value': base_data_df['CRD'].nunique() if 'CRD' in base_data_df.columns else 0
            })
            filtering_stats_data.append({
                'Metric': 'Latest Filing Date Range',
                'Value': f"{base_data_df['Latest_ADV_Filing_Date'].min().date()} to {base_data_df['Latest_ADV_Filing_Date'].max().date()}" 
                         if 'Latest_ADV_Filing_Date' in base_data_df.columns and not base_data_df.empty else 'N/A'
            })
        
        if filtering_stats_data:
            processed_data['Filtering_Statistics'] = pd.DataFrame(filtering_stats_data)

        # Save to Excel
        logger.info("Saving processed data to Excel...")
        error_log = save_to_excel(config, processed_data, month, year, error_log)

        # Generate reports
        error_log = generate_error_report(config, month, year, processing_stats, error_log)

        # Archive raw files if configured
        if config['processing'].get('archive_raw_files', False):
            logger.info("Archiving raw files...")
            archive_dir = os.path.join(config['paths']['archive'], f"{year}_{month}")
            Path(archive_dir).mkdir(parents=True, exist_ok=True)
            
            import shutil
            if os.path.exists(part1_zip):
                shutil.move(part1_zip, os.path.join(archive_dir, os.path.basename(part1_zip)))
            logger.info(f"Raw files archived to {archive_dir}")

        logger.info("RIA data processing completed successfully")
        logger.info(f"Total latest filings processed: {len(latest_filing_ids)}")
        
        return processing_stats, error_log

    except Exception as e:
        logger.error(f"Fatal error in processing: {str(e)}")
        error_log = error_log + [{
            'timestamp': datetime.now(),
            'function': 'run',
            'error': str(e),
            'traceback': traceback.format_exc()
        }]
        error_log = generate_error_report(config, config['processing']['month'], config['processing']['year'], processing_stats, error_log)
        sys.exit(1)



def main():
    """Main entry point"""
    config, error_log = load_config()
    error_log = setup_directories(config, error_log)
    processing_stats = {
        'start_time': datetime.now(),
        'files_processed': 0,
        'records_processed': 0,
        'errors': []
    }
    run(config, error_log, processing_stats)

if __name__ == "__main__":
    main()