"""
MRP Critical Items Analyzer - Core Analysis Module
This module handles the analysis of Material Requirements Planning (MRP) data,
identifying critical items and generating reports.

Author: Humberto Rodrigues
Date: August 2025
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
import logging
from pathlib import Path
from typing import Tuple, Optional, List
from dataclasses import dataclass, field

# Configure detailed logging
log_dir = Path.home() / '.mrp_analyzer'
log_dir.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(log_dir / 'mrp_analyzer.log')
    ]
)

logger = logging.getLogger(__name__)

@dataclass
class MRPConfig:
    """Configuration settings for MRP analysis."""
    REQUIRED_COLUMNS: List[str] = field(default_factory=lambda: [
        "CÓD", "DESCRIÇÃOPROMOB", "ESTQ10", "ESTQ20", "DEMANDAMRP",
        "ESTOQSEG", "FORNECEDORPRINCIPAL", "PEDIDOS", "OBS"
    ])
    NUMERIC_COLUMNS: List[str] = field(default_factory=lambda: [
        "ESTQ10", "ESTQ20", "DEMANDAMRP", "ESTOQSEG", "PEDIDOS"
    ])
    OUTPUT_COLUMNS: List[str] = field(default_factory=lambda: [
        "CÓD", "FORNECEDOR PRINCIPAL", "DESCRIÇÃOPROMOB", "ESTQ10", "ESTQ20",
        "DEMANDAMRP", "ESTOQSEG", "PEDIDOS", "ESTOQUE DISPONÍVEL",
        "QUANTIDADE A SOLICITAR", "OBS"
    ])
    HISTORY_DIR: str = "historico_mrp"
    
class ValidationError(Exception):
    """Custom exception for data validation errors."""
    pass

class DataValidator:
    """Handles validation of input data for MRP analysis."""
    
    @staticmethod
    def validate_numeric_columns(df: pd.DataFrame, columns: list) -> None:
        """
        Validates that numeric columns contain only valid numbers.
        
        Args:
            df: DataFrame to validate
            columns: List of columns that should be numeric
            
        Raises:
            ValidationError: If non-numeric values are found
        """
        for col in columns:
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            if not numeric_series.notna().all():
                invalid_rows = df[numeric_series.isna()]
                raise ValidationError(
                    f"Non-numeric values found in column {col}. "
                    f"Rows: {invalid_rows.index.tolist()}"
                )

    @staticmethod
    def validate_positive_values(df: pd.DataFrame, columns: list) -> None:
        """
        Validates that columns contain only positive values.
        
        Args:
            df: DataFrame to validate
            columns: List of columns to check
            
        Raises:
            ValidationError: If negative values are found
        """
        for col in columns:
            if (df[col] < 0).any():
                negative_rows = df[df[col] < 0]
                raise ValidationError(
                    f"Negative values found in column {col}. "
                    f"Rows: {negative_rows.index.tolist()}"
                )
    
    @staticmethod
    def validate_required_columns(df: pd.DataFrame, required_cols: list) -> None:
        """
        Validates that all required columns are present.
        
        Args:
            df: DataFrame to validate
            required_cols: List of required column names
            
        Raises:
            ValidationError: If any required columns are missing
        """
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValidationError(f"Required columns missing: {', '.join(missing_cols)}")

class MRPAnalyzer:
    """Handles MRP analysis operations."""
    
    def __init__(self, config: MRPConfig = MRPConfig()):
        self.config = config
        self.validator = DataValidator()
    
    @staticmethod
    def _calculate_available_stock(df: pd.DataFrame) -> pd.Series:
        """Calculates available stock considering ESTQ10 and ESTQ20."""
        return np.add(df["ESTQ10"], np.divide(df["ESTQ20"], 3))
    
    @staticmethod
    def _calculate_required_quantity(df: pd.DataFrame) -> pd.Series:
        """Calculates the quantity to be ordered."""
        return (
            df["DEMANDAMRP"] - df["ESTOQUE DISPONÍVEL"] + 
            df["ESTOQSEG"] - df["PEDIDOS"]
        ).clip(lower=0).round().astype(int)

    def _load_and_validate_data(self, input_file: str, sheet_name: str) -> pd.DataFrame:
        """Loads Excel data and executes all validation checks."""
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"File not found: {input_file}")

        df = pd.read_excel(
            input_file,
            sheet_name=sheet_name,
            dtype={col: 'float64' for col in self.config.NUMERIC_COLUMNS}
        )

        df.columns = [col.strip().upper() for col in df.columns]
        self.validator.validate_required_columns(df, self.config.REQUIRED_COLUMNS)
        self.validator.validate_numeric_columns(df, self.config.NUMERIC_COLUMNS)
        self.validator.validate_positive_values(df, self.config.NUMERIC_COLUMNS)
        return df

    def _prepare_critical_items(self, df: pd.DataFrame) -> pd.DataFrame:
        """Computes stock-related metrics and filters critical items."""
        working_df = df.copy()
        working_df["ESTOQUE DISPONÍVEL"] = self._calculate_available_stock(working_df)

        critical_mask = (working_df["ESTOQUE DISPONÍVEL"] - working_df["DEMANDAMRP"]) < working_df["ESTOQSEG"]
        critical_items = working_df[critical_mask].copy()
        critical_items["QUANTIDADE A SOLICITAR"] = self._calculate_required_quantity(critical_items)
        critical_items["FORNECEDOR PRINCIPAL"] = critical_items["FORNECEDORPRINCIPAL"]
        critical_items["ESTOQUE DISPONÍVEL"] = critical_items["ESTOQUE DISPONÍVEL"].round().astype(int)
        return critical_items

    def _build_output_dataframe(self, critical_items: pd.DataFrame) -> pd.DataFrame:
        """Builds the final output shape expected by exporters/UI."""
        return critical_items[self.config.OUTPUT_COLUMNS].fillna("")
    
    def analyze(self, input_file: str, sheet_name: str, 
               output_file: str = 'itens_criticos.xlsx') -> Tuple[Optional[int], Optional[str], Optional[pd.DataFrame]]:
        """
        Performs MRP analysis from Excel file, saves results and history, returns critical items count.
        
        Args:
            input_file: Path to input Excel file
            sheet_name: Name of the worksheet to analyze
            output_file: Path to save results
            
        Returns:
            Tuple containing:
            - Number of critical items (or None if error)
            - Error message (or None if successful)
            - DataFrame with results (or None if error)
        """
        try:
            logger.info(f"Starting analysis of file: {input_file}")

            df = self._load_and_validate_data(input_file, sheet_name)
            critical_items = self._prepare_critical_items(df)
            output_df = self._build_output_dataframe(critical_items)
            self._save_results(output_df, output_file)
            
            return len(output_df), None, output_df
            
        except ValidationError as e:
            logger.error(f"Validation error: {str(e)}")
            return None, f"Validation error: {str(e)}", None
        except Exception as e:
            logger.error(f"Error during analysis: {str(e)}", exc_info=True)
            return None, f"Error during analysis: {str(e)}", None
            
    def _save_results(self, df: pd.DataFrame, output_file: str) -> None:
        """Saves results and creates a historical copy."""
        self._save_formatted_excel(df, output_file)
        self._save_history(df, output_file)
    
    def _save_history(self, df: pd.DataFrame, output_file: str) -> None:
        """Saves a historical copy of the file with timestamp."""
        hist_dir = Path(output_file).parent / self.config.HISTORY_DIR
        hist_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        hist_path = hist_dir / f"itens_criticos_{timestamp}.xlsx"
        self._save_formatted_excel(df, str(hist_path))
        logger.info(f"History saved to: {hist_path}")
    
    def _save_formatted_excel(self, df: pd.DataFrame, output_file: str) -> None:
        """Saves DataFrame to Excel with formatting."""
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        df.to_excel(writer, sheet_name="Critical Items", index=False)
        self._format_excel(writer, df)
        writer.close()
        logger.info(f"Excel file saved to: {output_file}")

    def _format_excel(self, writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
        """
        Formats Excel spreadsheet with styles and highlights.
        
        Args:
            writer: Excel writer object
            df: DataFrame to format
        """
        workbook = writer.book
        worksheet = writer.sheets['Critical Items']

        # Define formats
        formats = {
            'header': workbook.add_format({
                'bold': True, 
                'text_wrap': True, 
                'valign': 'top',
                'fg_color': '#D7E4BC', 
                'border': 1
            }),
            'integer': workbook.add_format({
                'num_format': '0', 
                'border': 1
            }),
            'text': workbook.add_format({
                'border': 1
            }),
            'highlight': workbook.add_format({
                'bg_color': '#F4CCCC', 
                'border': 1
            }),
            'alternate_row': workbook.add_format({
                'bg_color': '#F9F9F9'
            })
        }

        # Write headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, formats['header'])

        # Set column formats and widths
        for i, col in enumerate(df.columns):
            fmt = formats['integer'] if pd.api.types.is_numeric_dtype(df[col]) else formats['text']
            worksheet.set_column(i, i, 20, fmt)

        # Add worksheet features
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

        # Write data with conditional formatting
        for row_idx, row in enumerate(df.itertuples(index=False), start=1):
            for col_idx, value in enumerate(row):
                fmt = (formats['highlight'] 
                      if df.columns[col_idx] == "QUANTIDADE A SOLICITAR" 
                         and isinstance(value, (int, float)) and value > 0 
                      else formats['alternate_row'] if row_idx % 2 == 0 
                      else None)
                worksheet.write(row_idx, col_idx, value, fmt)


def analyze_mrp(input_file: str, sheet_name: str, output_file: str = 'itens_criticos.xlsx') -> Tuple[Optional[int], Optional[str], Optional[pd.DataFrame]]:
    """
    Convenience function for backward compatibility.
    Performs MRP analysis using the MRPAnalyzer class.
    """
    analyzer = MRPAnalyzer()
    return analyzer.analyze(input_file, sheet_name, output_file)


if __name__ == "__main__":
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(description="MRP Critical Items Analyzer")
    parser.add_argument("input_file", help="Path to the input Excel file")
    parser.add_argument("sheet_name", help="Name of the worksheet to analyze")
    parser.add_argument("-o", "--output", help="Output file path", default="itens_criticos.xlsx")
    
    args = parser.parse_args()
    
    analyzer = MRPAnalyzer()
    count, error, df = analyzer.analyze(args.input_file, args.sheet_name, args.output)
    
    if error:
        print(f"Error: {error}", file=sys.stderr)
        sys.exit(1)
    else:
        print(f"{count} critical items identified.")
        print(f"Results saved to: {args.output}")
        sys.exit(0)
