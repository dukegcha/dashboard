import pandas as pd
import os
import logging
from datetime import datetime
from typing import Dict, List, Optional, Union
from pathlib import Path
import chardet

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class CSVCleaner:
    """
    A comprehensive CSV cleaning utility for GI (Goods Issue) data.
    
    This class handles the transformation of raw CSV data into a clean format
    suitable for MySQL database insertion, with proper data type conversion,
    date formatting, and error handling.
    """
    
    # Configuration constants
    COLUMN_MAPPING = {
        'Material': 'material_id',
        'Delivery #': 'delivery_number',
        'Ship-to': 'ship_to',
        'Carrier': 'carrier_name',
        'ShpPoint': 'shipping_point',
        'SO Created Date': 'so_created_date',
        'Ac.GI date': 'ac_gi_date',
        'Delivery Date': 'delivery_date',
        'IOD from 3PL': 'iod_from_3pl',
        'PlanShipSt': 'planned_ship_start',
        'S.Org(G)': 'sales_org',
        'P/O #': 'purchase_order',
        'Type': 'record_type',
        'Shipment': 'shipment_number',
        'Sold-to': 'sold_to',
        '[WE]Name1': 'customer_name',
        '[WE]State': 'customer_state',
        'Pro #': 'pro_number',
        'DOCrtDate': 'document_created_date',
        'Serial no. profile': 'serial_no_profile',
        'Ship Crt': 'ship_crt',
        'G/I Date': 'g_i_date',
        'PlanLoadSt': 'planloadst',
        '[WE]Street': 'we_street',
        '[WE]City': 'we_city',
        '[WE]Country': 'we_country',
        '[WE]Zipcode': 'we_zipcode',
        'Division': 'division',
        'Quantity': 'quantity',
        'Plan G/I (DO)': 'plan_g_i_do_',
        'Qty.Unit': 'qty_unit',
        'Delivery type': 'delivery_type',
        'DOCrtTime': 'docrttime',
        'Material Group': 'material_group',
        'Volume': 'volume',
        'Vol.Unit': 'vol_unit',
        'Weight': 'weight',
        'Wgt.Unit': 'wgt_unit',
        'ShTy': 'shty',
        'S/O #': 's_o_',
        'S/O item#': 's_o_item_',
        'P/O item#': 'p_o_item_',
        'Cust.Grp': 'cust_grp',
        'Escort/Txt3': 'escort_txt3',
        'ActLT': 'actlt',
    }
    
    NUMERIC_COLUMNS = [
        "delivery_number", "sales_org", "shipment_number", "quantity", 
        "volume", "weight", "s_o_", "s_o_item_", "actlt"
    ]
    
    DATE_COLUMNS = [
        "so_created_date", "ac_gi_date", "delivery_date", "iod_from_3pl",
        "planned_ship_start", "document_created_date", "ship_crt", "g_i_date",
        "planloadst", "plan_g_i_do_"
    ]
    
    COLUMNS_TO_DROP = ['Status']
    
    def __init__(self, chunk_size: Optional[int] = None):
        """
        Initialize the CSV cleaner.
        
        Args:
            chunk_size: Optional chunk size for processing large files
        """
        self.chunk_size = chunk_size
        self.stats = {
            'files_processed': 0,
            'rows_processed': 0,
            'errors': []
        }
    
    def detect_encoding(self, file_path: Union[str, Path]) -> str:
        """
        Detect the encoding of a CSV file.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            Detected encoding string
        """
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)  # Read first 10KB
                result = chardet.detect(raw_data)
                encoding = result['encoding']
                confidence = result['confidence']
                
                logger.info(f"Detected encoding: {encoding} (confidence: {confidence:.2f})")
                return encoding or 'utf-8'
        except Exception as e:
            logger.warning(f"Failed to detect encoding for {file_path}: {e}")
            return 'utf-8'
    
    def read_csv_with_fallback(self, file_path: Union[str, Path]) -> pd.DataFrame:
        """
        Read CSV file with encoding fallback and error handling.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            DataFrame with the CSV data
            
        Raises:
            Exception: If all encoding attempts fail
        """
        encodings = [self.detect_encoding(file_path), 'utf-8', 'latin1', 'cp1252']
        
        for encoding in encodings:
            try:
                logger.info(f"Attempting to read {file_path} with encoding: {encoding}")
                
                read_kwargs = {
                    'dtype': str, 
                    'encoding': encoding,
                    'na_filter': False  # Prevent pandas from converting empty strings to NaN
                }
                
                if self.chunk_size:
                    read_kwargs['chunksize'] = self.chunk_size
                
                df = pd.read_csv(file_path, **read_kwargs)
                logger.info(f"Successfully read {file_path} with {encoding}")
                return df
                
            except UnicodeDecodeError:
                logger.warning(f"Failed to read with {encoding}, trying next encoding")
                continue
            except Exception as e:
                logger.error(f"Unexpected error reading {file_path} with {encoding}: {e}")
                continue
        
        raise Exception(f"Failed to read {file_path} with any encoding")
    
    def clean_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and standardize column names.
        
        Args:
            df: Input DataFrame
            
        Returns:
            DataFrame with cleaned column names
        """
        # Remove columns with null/empty names
        df = df.loc[:, df.columns.notnull() & (df.columns != '')]
        
        # Strip whitespace from column names
        df.columns = df.columns.str.strip()
        
        # Log unknown columns
        unknown_cols = [col for col in df.columns if col not in self.COLUMN_MAPPING and col not in self.COLUMNS_TO_DROP]
        if unknown_cols:
            logger.warning(f"Unknown columns found: {unknown_cols}")
        
        return df
    
    def filter_and_rename_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filter to keep only mapped columns and rename them.
        
        Args:
            df: Input DataFrame
            
        Returns:
            DataFrame with filtered and renamed columns
        """
        # Keep only columns that are in our mapping
        valid_cols = [col for col in df.columns if col in self.COLUMN_MAPPING]
        df = df[valid_cols]
        
        # Drop specified columns
        for col in self.COLUMNS_TO_DROP:
            if col in df.columns:
                df = df.drop(columns=[col])
                logger.info(f"Dropped column: {col}")
        
        # Rename columns
        df = df.rename(columns=self.COLUMN_MAPPING)
        
        return df
    
    def clean_string_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean string data by stripping whitespace.
        
        Args:
            df: Input DataFrame
            
        Returns:
            DataFrame with cleaned string data
        """
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
                # Convert 'nan' strings back to empty strings
                df[col] = df[col].replace('nan', '')
        
        return df
    
    def parse_date_flexible(self, date_str: str) -> str:
        """
        Parse date string with multiple format attempts.
        
        Args:
            date_str: Date string to parse
            
        Returns:
            Formatted date string (YYYY-MM-DD) or empty string if parsing fails
        """
        if pd.isnull(date_str) or date_str == '' or date_str == 'nan':
            return ''
        
        # Try multiple date formats
        date_formats = [
            '%Y-%m-%d',     # YYYY-MM-DD
            '%m/%d/%Y',     # MM/DD/YYYY
            '%d/%m/%Y',     # DD/MM/YYYY
            '%Y/%m/%d',     # YYYY/MM/DD
            '%m-%d-%Y',     # MM-DD-YYYY
            '%d-%m-%Y',     # DD-MM-YYYY
        ]
        
        for fmt in date_formats:
            try:
                dt = datetime.strptime(str(date_str).strip(), fmt)
                return dt.strftime('%Y-%m-%d')
            except ValueError:
                continue
        
        # If manual parsing fails, try pandas
        try:
            dt = pd.to_datetime(date_str, errors='coerce', dayfirst=False)
            if pd.isnull(dt):
                return ''
            return dt.strftime('%Y-%m-%d')
        except:
            logger.warning(f"Failed to parse date: {date_str}")
            return ''
    
    def process_date_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Process and format date columns.
        
        Args:
            df: Input DataFrame
            
        Returns:
            DataFrame with processed date columns
        """
        for col in self.DATE_COLUMNS:
            if col in df.columns:
                logger.info(f"Processing date column: {col}")
                df[col] = df[col].apply(self.parse_date_flexible)
        
        return df
    
    def process_numeric_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Process and clean numeric columns.
        
        Args:
            df: Input DataFrame
            
        Returns:
            DataFrame with processed numeric columns
        """
        for col in self.NUMERIC_COLUMNS:
            if col in df.columns:
                logger.info(f"Processing numeric column: {col}")
                # Clean the data: remove commas, strip spaces, handle empty strings
                df[col] = df[col].astype(str).str.replace(",", "", regex=False).str.strip()
                df[col] = df[col].replace('', None)  # Convert empty strings to None
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    
    def validate_data(self, df: pd.DataFrame) -> Dict[str, Union[int, List[str]]]:
        """
        Validate the cleaned data and return quality metrics.
        
        Args:
            df: DataFrame to validate
            
        Returns:
            Dictionary with validation results
        """
        validation_results = {
            'total_rows': len(df),
            'empty_rows': df.isnull().all(axis=1).sum(),
            'columns_with_nulls': [],
            'date_parsing_issues': [],
            'numeric_conversion_issues': []
        }
        
        # Check for columns with high null percentages
        for col in df.columns:
            null_pct = (df[col].isnull().sum() / len(df)) * 100
            if null_pct > 50:  # More than 50% nulls
                validation_results['columns_with_nulls'].append(f"{col}: {null_pct:.1f}%")
        
        return validation_results
    
    def clean_csv(self, input_path: Union[str, Path], output_path: Union[str, Path]) -> bool:
        """
        Clean a single CSV file.
        
        Args:
            input_path: Path to input CSV file
            output_path: Path to output cleaned CSV file
            
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Starting to clean: {input_path}")
            
            # Read CSV with fallback encoding
            df = self.read_csv_with_fallback(input_path)
            
            if df.empty:
                logger.warning(f"Empty DataFrame for {input_path}")
                return False
            
            original_shape = df.shape
            logger.info(f"Original shape: {original_shape}")
            
            # Processing pipeline
            df = self.clean_column_names(df)
            df = self.filter_and_rename_columns(df)
            df = self.clean_string_data(df)
            df = self.process_date_columns(df)
            df = self.process_numeric_columns(df)
            
            # Validate results
            validation = self.validate_data(df)
            logger.info(f"Validation results: {validation}")
            
            # Ensure output directory exists
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            
            # Write cleaned CSV
            df.to_csv(output_path, index=False, na_rep='')
            
            final_shape = df.shape
            logger.info(f"Cleaned CSV written to {output_path}")
            logger.info(f"Shape changed from {original_shape} to {final_shape}")
            
            # Update stats
            self.stats['files_processed'] += 1
            self.stats['rows_processed'] += len(df)
            
            return True
            
        except Exception as e:
            error_msg = f"Failed to clean {input_path}: {str(e)}"
            logger.error(error_msg)
            self.stats['errors'].append(error_msg)
            return False
    
    def clean_directory(self, input_dir: Union[str, Path], output_dir: Union[str, Path], 
                       file_pattern: str = "*.csv") -> Dict[str, int]:
        """
        Clean all CSV files in a directory.
        
        Args:
            input_dir: Directory containing input CSV files
            output_dir: Directory for output cleaned CSV files
            file_pattern: File pattern to match (default: "*.csv")
            
        Returns:
            Dictionary with processing statistics
        """
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        
        if not input_path.exists():
            logger.error(f"Input directory does not exist: {input_dir}")
            return self.stats
        
        # Create output directory
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Find all CSV files
        csv_files = list(input_path.glob(file_pattern))
        
        if not csv_files:
            logger.warning(f"No CSV files found in {input_dir}")
            return self.stats
        
        logger.info(f"Found {len(csv_files)} CSV files to process")
        
        # Process each file
        for csv_file in csv_files:
            output_file = output_path / f"cleaned_{csv_file.name}"
            self.clean_csv(csv_file, output_file)
        
        # Log final stats
        logger.info(f"Processing complete. Stats: {self.stats}")
        return self.stats


def main():
    """
    Main function for command-line usage.
    """
    # Configuration
    script_dir = Path(__file__).parent
    input_folder = script_dir / 'csv'
    output_folder = script_dir / 'cleaned_csv'
    
    # Initialize cleaner
    cleaner = CSVCleaner(chunk_size=None)  # Set chunk_size for large files
    
    # Process all CSV files
    stats = cleaner.clean_directory(input_folder, output_folder)
    
    # Print summary
    print("\n" + "="*50)
    print("PROCESSING SUMMARY")
    print("="*50)
    print(f"Files processed: {stats['files_processed']}")
    print(f"Total rows processed: {stats['rows_processed']}")
    
    if stats['errors']:
        print(f"Errors encountered: {len(stats['errors'])}")
        for error in stats['errors']:
            print(f"  - {error}")
    else:
        print("No errors encountered!")


if __name__ == "__main__":
    main()