import pandas as pd
import os
import json
from pathlib import Path
from markitdown import MarkItDown
import logging
from typing import List, Dict, Optional, Tuple
import requests
import zipfile
from datetime import datetime, timedelta



# Set up logging and suppress pdfminer warnings at module level
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Suppress pdfminer warnings globally
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='pdfminer')
logging.getLogger('pdfminer').setLevel(logging.ERROR)
logging.getLogger('pdfminer.pdfinterp').setLevel(logging.ERROR)

def detect_file_encoding(file_path: str) -> str:
    """
    Detect the encoding of a file by trying common encodings
    """
    import chardet
    
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(100000)  # Read first 100KB for detection
        
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        confidence = result['confidence']
        
        logger.info(f"Detected encoding: {encoding} (confidence: {confidence:.2f})")
        return encoding
    except Exception as e:
        logger.warning(f"Could not detect encoding: {str(e)}")
        return 'utf-8'
    
def download_adv_brochures(month: str, year: int, output_directory: str) -> Optional[str]:
    """
    Download Form ADV Part 2 brochure files for specified month
    Returns path to downloaded zip file or None if failed
    """
    try:
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
        
        # Create output directory
        Path(output_directory).mkdir(parents=True, exist_ok=True)
        
        # Download Part 2 - ADV Brochures
        part2_base_url = "https://reports.adviserinfo.sec.gov/reports/foia/advBrochures"
        part2_filename = f"ADV_Brochures_{year}_{month}.zip"
        part2_url = f"{part2_base_url}/{year}/{part2_filename}"
        part2_path = os.path.join(output_directory, part2_filename)
        
        logger.info(f"Downloading ADV Brochures from: {part2_url}")
        
        response = requests.get(part2_url, headers=headers, stream=True, timeout=60)
        response.raise_for_status()
        
        with open(part2_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        file_size_mb = os.path.getsize(part2_path) / 1024 / 1024
        logger.info(f"Downloaded ADV Brochures: {part2_path} ({file_size_mb:.2f} MB)")
        
        return part2_path
        
    except Exception as e:
        logger.error(f"Failed to download ADV brochures: {str(e)}")
        return None

def extract_brochure_zip(zip_path: str, extract_directory: str) -> bool:
    """
    Extract the brochure zip file to specified directory
    Returns True if successful, False otherwise
    """
    try:
        extract_dir = Path(extract_directory)
        extract_dir.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Extracting brochure zip: {zip_path}")
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        # List extracted files
        extracted_files = list(extract_dir.rglob("*"))
        pdf_files = [f for f in extracted_files if f.suffix.lower() == '.pdf']
        csv_files = [f for f in extracted_files if f.suffix.lower() == '.csv']
        
        logger.info(f"Extraction complete:")
        logger.info(f"  Total files: {len(extracted_files)}")
        logger.info(f"  PDF files: {len(pdf_files)}")
        logger.info(f"  CSV files: {len(csv_files)}")
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to extract zip file: {str(e)}")
        return False

    

def read_csv_with_encoding(file_path: str) -> pd.DataFrame:
    """
    Read CSV file trying different encodings
    """
    # List of common encodings to try
    encodings_to_try = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16']
    
    # First try to detect encoding
    try:
        detected_encoding = detect_file_encoding(file_path)
        if detected_encoding:
            encodings_to_try.insert(0, detected_encoding)
    except ImportError:
        logger.warning("chardet not available, trying common encodings")
        pass
    
    # Remove duplicates while preserving order
    encodings_to_try = list(dict.fromkeys(encodings_to_try))
    
    for encoding in encodings_to_try:
        try:
            logger.info(f"Trying to read CSV with encoding: {encoding}")
            df = pd.read_csv(file_path, encoding=encoding)
            logger.info(f"Successfully read CSV with encoding: {encoding}")
            return df
        except UnicodeDecodeError:
            logger.warning(f"Failed to read with encoding: {encoding}")
            continue
        except Exception as e:
            logger.error(f"Error reading CSV with {encoding}: {str(e)}")
            continue
    
    # If all encodings fail, try with error handling
    try:
        logger.info("Trying to read CSV with utf-8 and error handling")
        df = pd.read_csv(file_path, encoding='utf-8', encoding_errors='ignore')
        logger.warning("Read CSV with some characters ignored due to encoding issues")
        return df
    except Exception as e:
        raise Exception(f"Could not read CSV file with any encoding: {str(e)}")

def read_and_filter_data(file_path: str) -> pd.DataFrame:
    """
    Read CSV/Excel file and filter out PDFs with 'wrap' in filename or brochure name
    """
    try:
        # Determine file type and read accordingly
        file_extension = Path(file_path).suffix.lower()
        
        if file_extension == '.csv':
            df = read_csv_with_encoding(file_path)
        elif file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")
        
        # Print basic info about the file
        logger.info(f"File shape: {df.shape}")
        logger.info(f"Columns in file: {df.columns.tolist()}")
        
        # Check if required columns exist
        required_columns = ['BrochureName', 'PDFFileName']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {missing_columns}")
        
        # Filter conditions - exclude records with 'wrap' in either column:
        # 1. PDFFileName does not contain 'wrap' (case insensitive)
        # 2. BrochureName does not contain 'wrap' (case insensitive)
        filtered_df = df[
            (~df['PDFFileName'].str.contains('wrap', case=False, na=False)) &
            (~df['BrochureName'].str.contains('wrap', case=False, na=False))
        ]
        
        # Select only required columns
        result_df = filtered_df[['BrochureName', 'PDFFileName']].copy()
        
        logger.info(f"Found {len(result_df)} matching records (excluded {len(df) - len(result_df)} records with 'wrap')")
        
        # Show sample of filtered data
        if len(result_df) > 0:
            logger.info("Sample filtered data:")
            logger.info(f"\n{result_df.head()}")
        
        return result_df
        
    except Exception as e:
        logger.error(f"Error reading file: {str(e)}")
        raise

def find_pdf_files(pdf_directory: str, pdf_filenames: List[str]) -> Dict[str, str]:
    """
    Find PDF files in the given directory that match the filenames from Excel
    Returns dict with filename as key and full path as value
    """
    pdf_dir = Path(pdf_directory)
    found_files = {}
    
    if not pdf_dir.exists():
        logger.error(f"PDF directory does not exist: {pdf_directory}")
        return found_files
    
    # Get all PDF files in directory (recursive search)
    all_pdf_files = list(pdf_dir.rglob("*.pdf"))
    
    for filename in pdf_filenames:
        # Try to find exact match first
        matching_files = [f for f in all_pdf_files if f.name == filename]
        
        if matching_files:
            found_files[filename] = str(matching_files[0])
            
        else:
            # Try partial match (in case of slight differences)
            base_name = filename.replace('.pdf', '').lower()
            partial_matches = [f for f in all_pdf_files if base_name in f.name.lower()]
            
            if partial_matches:
                found_files[filename] = str(partial_matches[0])
                logger.info(f"Found (partial match): {filename} -> {partial_matches[0].name}")
            else:
                logger.warning(f"PDF not found: {filename}")
    
    logger.info(f"Found {len(found_files)} out of {len(pdf_filenames)} PDF files")
    return found_files

def convert_pdf_to_markdown(pdf_path: str) -> Optional[str]:
    """
    Convert a single PDF file to markdown using MarkItDown
    
    Args:
        pdf_path: Path to the PDF file
    """
    try:
        # Suppress pdfminer warnings temporarily
        import warnings
        import logging as logging_module
        
        # Store original warning settings
        original_warnings = warnings.filters[:]
        original_pdfminer_level = logging_module.getLogger('pdfminer').level
        
        # Suppress pdfminer warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='pdfminer')
        logging_module.getLogger('pdfminer').setLevel(logging_module.ERROR)
        logging_module.getLogger('pdfminer.pdfinterp').setLevel(logging_module.ERROR)
        
        try:
            # Initialize MarkItDown converter
            markitdown = MarkItDown()
            
            # Convert PDF to markdown
            result = markitdown.convert(pdf_path)
            
            if result and hasattr(result, 'text_content') and result.text_content:
                markdown_content = result.text_content.strip()
                if markdown_content:
                    logger.info(f"Successfully converted PDF to markdown: {pdf_path}")
                    return markdown_content
                else:
                    logger.warning(f"Empty markdown content from: {pdf_path}")
                    return None
            else:
                logger.warning(f"No markdown content extracted from: {pdf_path}")
                return None
                
        finally:
            # Restore original warning settings
            warnings.filters[:] = original_warnings
            logging_module.getLogger('pdfminer').setLevel(original_pdfminer_level)
            
    except Exception as e:
        logger.error(f"Error converting PDF to markdown: {pdf_path}. Error: {str(e)}")
        return None

def convert_pdf_to_markdown_with_options(pdf_path: str, 
                                       extract_images: bool = False) -> Optional[str]:
    """
    Convert PDF to markdown with additional options and error handling
    
    Args:
        pdf_path: Path to the PDF file
        extract_images: Whether to attempt image extraction (MarkItDown may support this)
    """
    try:
        # Suppress pdfminer warnings
        import warnings
        import logging as logging_module
        
        # Store original warning settings
        original_warnings = warnings.filters[:]
        original_pdfminer_level = logging_module.getLogger('pdfminer').level
        
        # Suppress pdfminer warnings
        warnings.filterwarnings('ignore', category=UserWarning, module='pdfminer')
        logging_module.getLogger('pdfminer').setLevel(logging_module.ERROR)
        logging_module.getLogger('pdfminer.pdfinterp').setLevel(logging_module.ERROR)
        
        try:
            # Initialize MarkItDown converter
            markitdown = MarkItDown()
            
            # Convert PDF to markdown
            # Note: MarkItDown may not support all advanced options like PyMuPDF4LLM
            result = markitdown.convert(pdf_path)
            
            if result and hasattr(result, 'text_content') and result.text_content:
                markdown_content = result.text_content.strip()
                if markdown_content:
                    logger.info(f"Successfully converted PDF to markdown: {pdf_path}")
                    return markdown_content
                else:
                    logger.warning(f"Empty markdown content from: {pdf_path}")
                    return None
            else:
                logger.warning(f"No markdown content extracted from: {pdf_path}")
                return None
                
        finally:
            # Restore original warning settings
            warnings.filters[:] = original_warnings
            logging_module.getLogger('pdfminer').setLevel(original_pdfminer_level)
            
    except Exception as e:
        logger.error(f"Error converting PDF to markdown: {pdf_path}. Error: {str(e)}")
        return None

def save_as_json(content: str, filename: str, output_directory: str, 
                include_metadata: bool = True) -> bool:
    """
    Save markdown content as JSON file with optional metadata
    """
    try:
        output_dir = Path(output_directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create JSON structure
        json_data = {
            "content": content
        }
        
        # Add metadata if requested
        if include_metadata:
            json_data["metadata"] = {
                "source_file": filename,
                "content_length": len(content),
                "conversion_method": "MarkItDown"
            }
        
        # Create JSON filename (replace .pdf with .json)
        json_filename = filename.replace('.pdf', '.json')
        json_path = output_dir / json_filename
        
        # Save JSON file
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"Saved JSON: {json_path}")
        return True
        
    except Exception as e:
        logger.error(f"Error saving JSON for {filename}: {str(e)}")
        return False

def process_single_pdf(pdf_filename: str, pdf_path: str, output_directory: str,
                      extract_images: bool = False) -> bool:
    """
    Process a single PDF: convert to markdown and save as JSON
    """
    logger.info(f"Processing: {pdf_filename}")
    
    # Convert PDF to markdown using MarkItDown
    markdown_content = convert_pdf_to_markdown_with_options(
        pdf_path, 
        extract_images=extract_images
    )
    
    if markdown_content is None:
        logger.error(f"Failed to convert PDF to markdown: {pdf_filename}")
        return False
    
    # Save as JSON
    success = save_as_json(markdown_content, pdf_filename, output_directory)
    return success

def process_all_pdfs(filtered_data: pd.DataFrame, pdf_directory: str, 
                    output_directory: str = "adv-filing-brochure",
                    extract_images: bool = False,
                    source_csv_path: str = None) -> Dict[str, bool]:
    """
    Process all filtered PDFs: find files, convert to markdown, and save as JSON
    Also copies the source CSV to the output directory
    """
    # Get list of PDF filenames
    pdf_filenames = filtered_data['PDFFileName'].tolist()
    
    # Find PDF files
    found_files = find_pdf_files(pdf_directory, pdf_filenames)
    
    # Process each found PDF
    results = {}
    
    for pdf_filename, pdf_path in found_files.items():
        success = process_single_pdf(
            pdf_filename, 
            pdf_path, 
            output_directory,
            extract_images=extract_images
        )
        results[pdf_filename] = success
    
    # Copy source CSV to output directory if provided
    if source_csv_path and os.path.exists(source_csv_path):
        try:
            import shutil
            output_dir = Path(output_directory)
            output_dir.mkdir(parents=True, exist_ok=True)
            
            csv_filename = Path(source_csv_path).name
            destination_path = output_dir / csv_filename
            
            shutil.copy2(source_csv_path, destination_path)
            logger.info(f"Copied source CSV to output directory: {destination_path}")
            
        except Exception as e:
            logger.error(f"Failed to copy CSV file: {str(e)}")
    
    # Report results
    successful = sum(results.values())
    total = len(results)
    logger.info(f"Processing complete: {successful}/{total} files processed successfully")
    
    return results

def create_summary_report(filtered_data: pd.DataFrame, results: Dict[str, bool], 
                         output_directory: str) -> None:
    """
    Create a summary report of the processing results
    """
    try:
        output_dir = Path(output_directory)
        
        # Create summary data
        summary_data = []
        for _, row in filtered_data.iterrows():
            pdf_filename = row['PDFFileName']
            brochure_name = row['BrochureName']
            
            summary_data.append({
                'PDFFileName': pdf_filename,
                'BrochureName': brochure_name,
                'ProcessingStatus': 'Success' if results.get(pdf_filename, False) else 'Failed/Not Found',
                'JSONCreated': results.get(pdf_filename, False)
            })
        
        # Create summary DataFrame
        summary_df = pd.DataFrame(summary_data)
        
        # Save as Excel
        summary_path = output_dir / "processing_summary.xlsx"
        summary_df.to_excel(summary_path, index=False)
        
        # Save as CSV
        summary_csv_path = output_dir / "processing_summary.csv"
        summary_df.to_csv(summary_csv_path, index=False)
        
        logger.info(f"Summary report saved: {summary_path}")
        
    except Exception as e:
        logger.error(f"Error creating summary report: {str(e)}")

def main(file_path: str = None, pdf_directory: str = None, 
         output_directory: str = "adv-filing-brochure",
         extract_images: bool = False,
         create_summary: bool = True,
         download_month: str = None,
         download_year: int = None,
         download_only: bool = False):
    """
    Main function to orchestrate the entire process
    
    Args:
        file_path: Path to the CSV/Excel file (optional if downloading)
        pdf_directory: Directory containing PDF files (optional if downloading)
        output_directory: Directory to save JSON files
        extract_images: Whether to extract images from PDFs
        create_summary: Whether to create processing summary report
        download_month: Month to download (e.g., 'August')
        download_year: Year to download (e.g., 2025)
        download_only: If True, only download and extract, don't process PDFs
    """
    try:
        logger.info("Starting ADV Brochure processing...")
        
        # Handle download if requested
        if download_month and download_year:
            logger.info(f"Downloading brochures for {download_month} {download_year}...")
            
            # Create download directory
            download_dir = os.path.join(output_directory, f"ADV_Brochures_{download_year}_{download_month}")
            
            # Download the zip file
            zip_path = download_adv_brochures(download_month, download_year, download_dir)
            
            if zip_path is None:
                logger.error("Download failed, aborting process")
                return
            
            # Extract the zip file
            extract_success = extract_brochure_zip(zip_path, download_dir)
            
            if not extract_success:
                logger.error("Extraction failed, aborting process")
                return
            
            # Update paths for processing
            if file_path is None:
                # Look for CSV file in extracted directory
                csv_files = list(Path(download_dir).rglob("*.csv"))
                if csv_files:
                    file_path = str(csv_files[0])
                    logger.info(f"Found CSV file: {file_path}")
                else:
                    logger.error("No CSV file found in extracted directory")
                    if download_only:
                        logger.info("Download and extraction completed successfully")
                        return
                    else:
                        return
            
            if pdf_directory is None:
                pdf_directory = download_dir
                logger.info(f"Using download directory for PDFs: {pdf_directory}")
            
            if download_only:
                logger.info("Download and extraction completed successfully")
                return
        
        # Validate required parameters for processing
        if file_path is None or pdf_directory is None:
            raise ValueError("file_path and pdf_directory are required for processing")
        
        logger.info("Starting PDF to Markdown conversion process with MarkItDown...")
        
        # Step 1: Read and filter data
        logger.info("Step 1: Reading and filtering data...")
        filtered_data = read_and_filter_data(file_path)
        
        if filtered_data.empty:
            logger.warning("No records found matching the filter criteria")
            return
        
        # Print sample of filtered data
        print("\nFiltered Data Sample:")
        print(filtered_data.head(10))
        
        # Step 2: Process all PDFs
        logger.info("Step 2: Processing PDF files...")
        results = process_all_pdfs(
            filtered_data, 
            pdf_directory, 
            output_directory,
            extract_images=extract_images,
            source_csv_path=file_path  # Add this line
        )
        
        # Step 3: Create summary report
        if create_summary:
            logger.info("Step 3: Creating summary report...")
            create_summary_report(filtered_data, results, output_directory)
        
        # Final Summary
        print(f"\n=== PROCESSING SUMMARY ===")
        print(f"Total records from Excel: {len(filtered_data)}")
        print(f"PDFs found and processed: {len(results)}")
        print(f"Successful conversions: {sum(results.values())}")
        print(f"Failed conversions: {sum(1 for v in results.values() if not v)}")
        print(f"Output directory: {output_directory}")
        
        if extract_images:
            print("Image extraction: Enabled")
        
    except Exception as e:
        logger.error(f"Error in main process: {str(e)}")
        raise

# Example usage
# Example usage
if __name__ == "__main__":
    # Option 1: Download and process in one go
    main(
        download_month="October",
        download_year=2025,
        output_directory="/Users/nikitakumar/Documents/RIA_Excel_Processing/data/raw/"
    )
    
    # Option 2: Download only (no processing)
    # main(
    #     download_month="August",
    #     download_year=2025,
    #     output_directory="/Users/nikitakumar/Documents/RIA_Excel_Processing/data/raw/",
    #     download_only=True
    # )
    
    # Option 3: Process existing files (original functionality)
    # file_path = "/path/to/your/csv/file.csv"
    # pdf_directory = "/path/to/your/pdf/directory"
    # output_directory = "/path/to/output/directory"
    # main(file_path, pdf_directory, output_directory)