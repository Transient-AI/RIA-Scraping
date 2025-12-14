#!/usr/bin/env python3
"""
Enhanced RIA Data Processor with Multi-Sheet Excel Output
Creates separate sheets for personnel, services, target_clients, products_instruments, and industry_insights
All linked by CRD and original_url
"""
import json
import pandas as pd
from pathlib import Path
from typing import List, Dict, Any, Optional, Set, Tuple
import argparse
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from urllib.parse import urlparse
import re
from datetime import datetime

# ============================================================================
# URL NORMALIZATION AND MATCHING
# ============================================================================

def normalize_url(url: str) -> str:
    """Normalize URL for matching."""
    if pd.isna(url) or not url:
        return ""
    
    url = str(url).lower().strip()
    
    # Remove protocol
    url = re.sub(r'^https?://', '', url)
    url = re.sub(r'^www\.', '', url)
    
    # Remove trailing slashes
    url = url.rstrip('/')
    
    return url

def create_crd_mapping(crd_file_path: str) -> Dict[str, int]:
    """
    Create a mapping dictionary from normalized URL to Organization CRD.
    Returns: dict of {normalized_url: crd_number}
    """
    print(f"\n{'='*80}")
    print("LOADING CRD MAPPING FROM EXCEL")
    print(f"{'='*80}")
    
    df_crd = pd.read_excel(crd_file_path)
    
    print(f"Loaded {len(df_crd):,} URL-CRD mappings")
    print(f"Unique CRDs: {df_crd['Organization CRD#'].nunique():,}")
    
    # Create mapping dictionary
    crd_mapping = {}
    
    for idx, row in df_crd.iterrows():
        url = row['Website Address']
        crd = row['Organization CRD#']
        
        if pd.notna(url) and pd.notna(crd):
            normalized = normalize_url(url)
            if normalized:
                # Store mapping (if duplicate, last one wins)
                crd_mapping[normalized] = int(crd)
    
    print(f"Created {len(crd_mapping):,} normalized URL mappings")
    
    return crd_mapping

def find_crd_for_url(url: str, crd_mapping: Dict[str, int]) -> Optional[int]:
    """Find CRD for a given URL using the mapping."""
    if not url:
        return None
    
    normalized = normalize_url(url)
    
    # Try exact match first
    if normalized in crd_mapping:
        return crd_mapping[normalized]
    
    # Try domain match (extract domain and check)
    try:
        parsed = urlparse(f"http://{normalized}")
        domain = parsed.netloc or parsed.path.split('/')[0]
        if domain in crd_mapping:
            return crd_mapping[domain]
    except:
        pass
    
    return None

# ============================================================================
# JSON Loading Functions
# ============================================================================

def load_json_file(file_path: str) -> Any:
    """Load a single JSON file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"  ⚠️  Error loading {Path(file_path).name}: {e}")
        return {}

def safe_get(data: Any, key: str, default: Any = None) -> Any:
    """Safely get a value from data."""
    if isinstance(data, dict):
        return data.get(key, default)
    return default

def safe_nested_get(data: Any, *keys, default: Any = None) -> Any:
    """Safely get nested values."""
    result = data
    for key in keys:
        if isinstance(result, dict):
            result = result.get(key)
            if result is None:
                return default
        else:
            return default
    return result if result is not None else default

# ============================================================================
# Data Extraction Functions for Multi-Sheet Structure
# ============================================================================

def extract_personnel_rows(json_data: dict, original_url: str, crd: Optional[int]) -> List[Dict]:
    """
    Extract personnel information as individual rows.
    Each person becomes one row with flattened contact information.
    """
    rows = []
    
    personnel_list = safe_get(json_data, 'key_personnel', [])
    if not personnel_list:
        return rows
    
    company_name = safe_get(json_data, 'company_name', '')
    
    for idx, person in enumerate(personnel_list, 1):
        if not isinstance(person, dict):
            continue
        
        # Extract contact information (flatten the dictionary)
        contact = person.get('contact', {})
        contact_dict = {}
        
        if isinstance(contact, dict):
            # Flatten contact dictionary
            for key, value in contact.items():
                if isinstance(value, list):
                    contact_dict[f'contact_{key}'] = ' | '.join(str(v) for v in value if v)
                elif value:
                    contact_dict[f'contact_{key}'] = str(value)
        
        row = {
            'original_url': original_url,
            'organization_crd': crd,
            'company_name': company_name,
            'person_index': idx,
            'name': person.get('name', ''),
            'title': person.get('title', ''),
            'bio': person.get('bio', ''),
            **contact_dict  # Add all flattened contact fields
        }
        
        rows.append(row)
    
    return rows

def extract_list_field_rows(json_data: dict, field_name: str, original_url: str, crd: Optional[int]) -> List[Dict]:
    """
    Extract list field information as individual rows.
    Each item in the list becomes one row.
    """
    rows = []
    
    field_data = safe_get(json_data, field_name, [])
    if not field_data:
        return rows
    
    company_name = safe_get(json_data, 'company_name', '')
    
    # Handle both list and string types
    if isinstance(field_data, str):
        field_data = [field_data]
    elif not isinstance(field_data, list):
        return rows
    
    for idx, item in enumerate(field_data, 1):
        if not item:
            continue
        
        row = {
            'original_url': original_url,
            'organization_crd': crd,
            'company_name': company_name,
            'item_index': idx,
            field_name: str(item).strip()
        }
        
        rows.append(row)
    
    return rows

def extract_text_field_rows(json_data: dict, field_name: str, original_url: str, crd: Optional[int]) -> List[Dict]:
    """
    Extract text field information as a single row.
    For fields that contain text/string data.
    """
    rows = []
    
    field_data = safe_get(json_data, field_name)
    if not field_data:
        return rows
    
    company_name = safe_get(json_data, 'company_name', '')
    
    row = {
        'original_url': original_url,
        'organization_crd': crd,
        'company_name': company_name,
        field_name: str(field_data).strip()
    }
    
    rows.append(row)
    
    return rows

def extract_company_overview(json_data: dict, original_url: str, crd: Optional[int]) -> Dict:
    """
    Extract company overview information for the main sheet.
    """
    # Extract contact info
    contact_info = safe_nested_get(json_data, 'contact_info', default={})
    
    # Handle offices (could be a list)
    offices = safe_nested_get(json_data, 'contact_info', 'offices', default=[])
    if isinstance(offices, list):
        offices_str = ' | '.join(str(office) for office in offices if office)
    else:
        offices_str = str(offices) if offices else None
    
    row = {
        'original_url': original_url,
        'organization_crd': crd,
        'success': safe_get(json_data, 'success'),
        'error': safe_get(json_data, 'error'),
        'company_name': safe_get(json_data, 'company_name'),
        'website_url': safe_get(json_data, 'url'),
        'scraped_at': safe_get(json_data, 'scraped_at'),
        'processing_index': safe_get(json_data, 'processing_index'),
        'total_pages_crawled': safe_get(json_data, 'total_pages_crawled'),
        'pages_analyzed': safe_get(json_data, 'pages_analyzed'),
        'address': safe_nested_get(json_data, 'contact_info', 'address'),
        'phone': safe_nested_get(json_data, 'contact_info', 'phone'),
        'email': safe_nested_get(json_data, 'contact_info', 'email'),
        'offices': offices_str,
        # Count fields for summary
        'services_count': len(safe_get(json_data, 'services', [])),
        'target_clients_count': len(safe_get(json_data, 'target_clients', [])),
        'products_count': len(safe_get(json_data, 'products_instruments', [])),
        'personnel_count': len(safe_get(json_data, 'key_personnel', [])),
        'has_investment_strategy': bool(safe_get(json_data, 'investment_strategy')),
        'has_industry_insights': bool(safe_get(json_data, 'industry_insights'))
    }
    
    return row

# ============================================================================
# Batch Processing for Multi-Sheet Structure
# ============================================================================

def process_all_json_files_multisheet(base_folder: str, crd_mapping: Dict[str, int]) -> Dict[str, pd.DataFrame]:
    """
    Process all JSON files and return multiple DataFrames for different sheets.
    """
    base_path = Path(base_folder)
    
    # Initialize collectors for each sheet
    company_overview_rows = []
    personnel_rows = []
    services_rows = []
    target_clients_rows = []
    products_rows = []
    investment_strategy_rows = []
    industry_insights_rows = []
    
    batch_folders = sorted([f for f in base_path.iterdir() 
                           if f.is_dir() and f.name.startswith('batch_')])
    
    print(f"\nFound {len(batch_folders)} batch folders")
    
    total_files = 0
    total_errors = 0
    total_successful = 0
    crd_matched = 0
    
    for batch_folder in batch_folders:
        print(f"\nProcessing: {batch_folder.name}")
        
        json_files = [f for f in batch_folder.glob('*.json') 
                     if f.name != 'batch_summary.json']
        
        print(f"  Found {len(json_files)} JSON files")
        total_files += len(json_files)
        
        batch_successful = 0
        batch_crd_matched = 0
        
        for json_file in json_files:
            try:
                json_data = load_json_file(json_file)
                
                if not isinstance(json_data, dict):
                    continue
                
                # Get original URL and CRD
                original_url = safe_get(json_data, 'original_url', '')
                if not original_url:
                    original_url = json_file.name.replace('_result.json', '').replace('__result.json', '')
                
                website_url = safe_get(json_data, 'url')
                
                # Find CRD
                crd = find_crd_for_url(original_url, crd_mapping)
                if crd is None:
                    crd = find_crd_for_url(website_url, crd_mapping)
                
                if crd is not None:
                    batch_crd_matched += 1
                
                if safe_get(json_data, 'success'):
                    batch_successful += 1
                
                # Extract data for each sheet
                # 1. Company Overview
                overview_row = extract_company_overview(json_data, original_url, crd)
                company_overview_rows.append(overview_row)
                
                # Only process details if scraping was successful
                if safe_get(json_data, 'success'):
                    # 2. Personnel
                    personnel = extract_personnel_rows(json_data, original_url, crd)
                    personnel_rows.extend(personnel)
                    
                    # 3. Services
                    services = extract_list_field_rows(json_data, 'services', original_url, crd)
                    services_rows.extend(services)
                    
                    # 4. Target Clients
                    clients = extract_list_field_rows(json_data, 'target_clients', original_url, crd)
                    target_clients_rows.extend(clients)
                    
                    # 5. Products/Instruments
                    products = extract_list_field_rows(json_data, 'products_instruments', original_url, crd)
                    products_rows.extend(products)
                    
                    # 6. Investment Strategy
                    strategy = extract_text_field_rows(json_data, 'investment_strategy', original_url, crd)
                    investment_strategy_rows.extend(strategy)
                    
                    # 7. Industry Insights
                    insights = extract_text_field_rows(json_data, 'industry_insights', original_url, crd)
                    industry_insights_rows.extend(insights)
                
            except Exception as e:
                print(f"    Error: {json_file.name} - {e}")
                total_errors += 1
                company_overview_rows.append({
                    'original_url': json_file.name.replace('_result.json', '').replace('__result.json', ''),
                    'organization_crd': None,
                    'success': False,
                    'error': f'FILE_PROCESSING_ERROR: {str(e)}'
                })
        
        total_successful += batch_successful
        crd_matched += batch_crd_matched
        
        print(f"  ✓ Processed: {len(json_files)}, Successful: {batch_successful}, CRD Matched: {batch_crd_matched}")
    
    print(f"\n{'='*80}")
    print(f"PROCESSING SUMMARY")
    print(f"{'='*80}")
    print(f"Total files processed: {total_files}")
    print(f"Total successful: {total_successful}")
    print(f"Total errors: {total_errors}")
    print(f"CRD Matched: {crd_matched} ({crd_matched/total_files*100:.1f}%)")
    print(f"{'='*80}")
    
    # Create DataFrames
    dataframes = {
        'Company_Overview': pd.DataFrame(company_overview_rows),
        'Key_Personnel': pd.DataFrame(personnel_rows),
        'Services': pd.DataFrame(services_rows),
        'Target_Clients': pd.DataFrame(target_clients_rows),
        'Products_Instruments': pd.DataFrame(products_rows),
        'Investment_Strategy': pd.DataFrame(investment_strategy_rows),
        'Industry_Insights': pd.DataFrame(industry_insights_rows)
    }
    
    # Set index for Company Overview
    if not dataframes['Company_Overview'].empty and 'original_url' in dataframes['Company_Overview'].columns:
        dataframes['Company_Overview'] = dataframes['Company_Overview'].set_index('original_url')
    
    return dataframes

# ============================================================================
# Excel Formatting Functions
# ============================================================================

def format_excel_sheet(worksheet, df: pd.DataFrame, sheet_type: str):
    """
    Apply formatting to Excel worksheet based on sheet type.
    """
    # Define styles
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Header colors based on sheet type
    header_colors = {
        'Company_Overview': 'C5504D',
        'Key_Personnel': '4472C4',
        'Services': '70AD47',
        'Target_Clients': 'FFC000',
        'Products_Instruments': '7030A0',
        'Investment_Strategy': '00B0F0',
        'Industry_Insights': 'FF6600'
    }
    
    header_color = header_colors.get(sheet_type, '4472C4')
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    
    # Format headers
    for col_num, col_name in enumerate(df.columns, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
        # Auto-adjust column width
        max_width = len(str(col_name)) + 2
        
        # Check data for width (sample first 100 rows)
        for row_idx in range(2, min(len(df) + 2, 102)):
            cell_value = worksheet.cell(row=row_idx, column=col_num).value
            if cell_value:
                # Limit width calculation to prevent excessive widths
                value_lines = str(cell_value).split('\n')
                max_line_width = max(len(line) for line in value_lines) if value_lines else 0
                max_width = max(max_width, min(max_line_width, 50))
        
        # Set column width (minimum 10, maximum 60)
        worksheet.column_dimensions[get_column_letter(col_num)].width = max(10, min(max_width + 2, 60))
    
    # Format data rows
    for row in worksheet.iter_rows(min_row=2, max_row=len(df) + 1):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = thin_border
            
            # Special formatting for CRD column
            if cell.column == 2 and sheet_type != 'Company_Overview':  # CRD is usually column 2
                cell.alignment = Alignment(horizontal="center", vertical="top")
    
    # Freeze top row
    worksheet.freeze_panes = 'A2'
    
    # Add autofilter
    worksheet.auto_filter.ref = worksheet.dimensions

# ============================================================================
# Summary Statistics Functions
# ============================================================================

def create_summary_sheet(dataframes: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """
    Create a summary statistics sheet.
    """
    summary_data = []
    
    for sheet_name, df in dataframes.items():
        if df.empty:
            record_count = 0
            unique_urls = 0
            unique_crds = 0
        else:
            record_count = len(df)
            unique_urls = df['original_url'].nunique() if 'original_url' in df.columns else 0
            unique_crds = df['organization_crd'].nunique() if 'organization_crd' in df.columns else 0
        
        summary_data.append({
            'Sheet Name': sheet_name,
            'Total Records': record_count,
            'Unique URLs': unique_urls,
            'Unique CRDs': unique_crds,
            'Columns': len(df.columns) if not df.empty else 0
        })
    
    # Add overview statistics
    overview_df = dataframes.get('Company_Overview', pd.DataFrame())
    if not overview_df.empty:
        summary_data.append({'Sheet Name': '---', 'Total Records': '---', 'Unique URLs': '---', 'Unique CRDs': '---', 'Columns': '---'})
        
        if 'success' in overview_df.columns:
            summary_data.append({
                'Sheet Name': 'Success Rate',
                'Total Records': f"{overview_df['success'].sum()} / {len(overview_df)}",
                'Unique URLs': f"{(overview_df['success'].sum() / len(overview_df) * 100):.1f}%",
                'Unique CRDs': '',
                'Columns': ''
            })
        
        if 'organization_crd' in overview_df.columns:
            summary_data.append({
                'Sheet Name': 'CRD Coverage',
                'Total Records': f"{overview_df['organization_crd'].notna().sum()} / {len(overview_df)}",
                'Unique URLs': f"{(overview_df['organization_crd'].notna().sum() / len(overview_df) * 100):.1f}%",
                'Unique CRDs': '',
                'Columns': ''
            })
    
    return pd.DataFrame(summary_data)

# ============================================================================
# Main Excel Creation Function
# ============================================================================

def create_multisheet_excel(dataframes: Dict[str, pd.DataFrame], output_path: str):
    """
    Create a multi-sheet Excel file with all extracted data.
    """
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    print(f"\n{'='*80}")
    print("CREATING MULTI-SHEET EXCEL FILE")
    print(f"{'='*80}")
    
    # Create summary sheet
    summary_df = create_summary_sheet(dataframes)
    
    # Write to Excel
    with pd.ExcelWriter(str(output_file), engine='openpyxl') as writer:
        # Write Summary first
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        worksheet = writer.sheets['Summary']
        format_excel_sheet(worksheet, summary_df, 'Summary')
        
        # Write each dataframe to its sheet
        for sheet_name, df in dataframes.items():
            if df.empty:
                # Create empty sheet with message
                empty_df = pd.DataFrame({'Message': ['No data available for this category']})
                empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  ⚠️  {sheet_name}: No data (empty sheet created)")
            else:
                # Write data
                if sheet_name == 'Company_Overview':
                    df.to_excel(writer, sheet_name=sheet_name, index=True)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                worksheet = writer.sheets[sheet_name]
                format_excel_sheet(worksheet, df, sheet_name)
                
                print(f"  ✓ {sheet_name}: {len(df):,} rows, {len(df.columns)} columns")
    
    print(f"\n{'='*80}")
    print(f"✓ Multi-sheet Excel file created successfully!")
    print(f"Location: {output_file}")
    print(f"{'='*80}")
    
    return output_file

# ============================================================================
# Additional Export Functions
# ============================================================================

def export_individual_csvs(dataframes: Dict[str, pd.DataFrame], output_folder: str):
    """
    Export each dataframe as a separate CSV file for additional processing.
    """
    output_path = Path(output_folder) / 'csv_exports'
    output_path.mkdir(parents=True, exist_ok=True)
    
    print(f"\n{'='*80}")
    print("EXPORTING INDIVIDUAL CSV FILES")
    print(f"{'='*80}")
    
    for sheet_name, df in dataframes.items():
        if not df.empty:
            csv_file = output_path / f"{sheet_name.lower()}.csv"
            if sheet_name == 'Company_Overview':
                df.to_csv(csv_file, index=True)
            else:
                df.to_csv(csv_file, index=False)
            print(f"  ✓ {sheet_name}: {csv_file.name}")
    
    print(f"\nAll CSV files saved to: {output_path}")

# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Enhanced RIA Data Processor with Multi-Sheet Excel Output')
    parser.add_argument('--input', '-i', type=str,
                       default='/Users/nikitakumar/Documents/RIA_Web_Scraping/WebScraping-Output/ria_outputs',
                       help='Input folder with batch_* folders')
    parser.add_argument('--output', '-o', type=str,
                       default='/Users/nikitakumar/Documents/RIA_Web_Scraping/WebScraping-Output/final_output',
                       help='Output folder')
    parser.add_argument('--crd-file', '-c', type=str,
                       default='/Users/nikitakumar/Documents/RIA_Web_Scraping/RIA_websites_company_websites_only.xlsx',
                       help='Excel file with Organization CRD and Website Address mapping')
    parser.add_argument('--excel-name', type=str,
                       default='ria_data_multisheet.xlsx',
                       help='Name for the multi-sheet Excel file')
    parser.add_argument('--export-csv', action='store_true',
                       help='Also export individual CSV files')
    
    args = parser.parse_args()
    
    print("="*80)
    print("ENHANCED RIA DATA PROCESSOR - MULTI-SHEET EXCEL")
    print("="*80)
    print(f"Input: {args.input}")
    print(f"Output: {args.output}")
    print(f"CRD Mapping File: {args.crd_file}")
    print(f"Excel Output Name: {args.excel_name}")
    print("="*80)
    
    # Load CRD mapping
    crd_mapping = create_crd_mapping(args.crd_file)
    
    # Process all JSON files
    print("\nProcessing JSON files and extracting multi-sheet data...")
    dataframes = process_all_json_files_multisheet(args.input, crd_mapping)
    
    # Display extraction summary
    print(f"\n{'='*80}")
    print("DATA EXTRACTION SUMMARY")
    print(f"{'='*80}")
    for sheet_name, df in dataframes.items():
        if not df.empty:
            print(f"{sheet_name}:")
            print(f"  - Records: {len(df):,}")
            print(f"  - Columns: {len(df.columns)}")
            if 'organization_crd' in df.columns:
                crd_coverage = df['organization_crd'].notna().sum()
                print(f"  - CRD Coverage: {crd_coverage:,} ({crd_coverage/len(df)*100:.1f}%)")
    
    # Create output folder
    output_path = Path(args.output)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Create multi-sheet Excel file
    excel_path = output_path / args.excel_name
    excel_file = create_multisheet_excel(dataframes, str(excel_path))
    
    # Export individual CSVs if requested
    if args.export_csv:
        export_individual_csvs(dataframes, args.output)
    
    # Display sample data
    print(f"\n{'='*80}")
    print("SAMPLE DATA")
    print(f"{'='*80}")
    
    # Show sample from Key Personnel
    if not dataframes['Key_Personnel'].empty:
        print("\nKey Personnel (first 3 records):")
        sample_cols = ['organization_crd', 'company_name', 'name', 'title']
        available_cols = [col for col in sample_cols if col in dataframes['Key_Personnel'].columns]
        print(dataframes['Key_Personnel'][available_cols].head(3).to_string(index=False))
    
    # Show sample from Services
    if not dataframes['Services'].empty:
        print("\nServices (first 3 unique services):")
        sample_cols = ['organization_crd', 'company_name', 'services']
        available_cols = [col for col in sample_cols if col in dataframes['Services'].columns]
        print(dataframes['Services'][available_cols].head(3).to_string(index=False))
    
    # Final summary
    print(f"\n{'='*80}")
    print("✅ PROCESSING COMPLETE!")
    print(f"{'='*80}")
    print(f"\nMain Output:")
    print(f"  📊 Multi-Sheet Excel: {excel_path.name}")
    print(f"\nSheets Created:")
    print(f"  1. Summary - Overview statistics")
    print(f"  2. Company_Overview - Main company data")
    print(f"  3. Key_Personnel - Individual personnel records")
    print(f"  4. Services - Extracted services")
    print(f"  5. Target_Clients - Target client categories")
    print(f"  6. Products_Instruments - Financial products")
    print(f"  7. Investment_Strategy - Investment strategies")
    print(f"  8. Industry_Insights - Industry insights")
    
    if args.export_csv:
        print(f"\nAdditional CSV exports in: {output_path / 'csv_exports'}")
    
    print(f"\n{'='*80}")
    print(f"All data is linked by 'organization_crd' and 'original_url' columns")
    print(f"{'='*80}")