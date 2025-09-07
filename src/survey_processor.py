#!/usr/bin/env python3
"""
Survey Data Processor

Converts raw survey data files into formatted Excel workbooks.
"""

import pandas as pd
import numpy as np
from pathlib import Path
import chardet
import logging
from typing import Dict, List, Optional
import json

class SurveyProcessor:
    def __init__(self, config_file: Optional[str] = None):
        self.config = self._load_config(config_file)
        self._setup_logging()
        
    def _load_config(self, config_file: Optional[str] = None) -> Dict:
        if config_file and Path(config_file).exists():
            with open(config_file, 'r') as f:
                return json.load(f)
        return {
            "input_encoding": "auto",
            "output_format": "xlsx",
            "clean_column_names": True,
            "remove_empty_rows": True,
            "date_formats": ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]
        }
    
    def _setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('survey_processing.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def detect_encoding(self, file_path: str) -> str:
        if self.config["input_encoding"] == "auto":
            with open(file_path, 'rb') as f:
                result = chardet.detect(f.read(10000))
                return result['encoding']
        return self.config["input_encoding"]
    
    def load_raw_data(self, file_path: str) -> pd.DataFrame:
        file_path = Path(file_path)
        encoding = self.detect_encoding(file_path)
        
        self.logger.info(f"Loading {file_path} with encoding {encoding}")
        
        if file_path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path, encoding=encoding)
        elif file_path.suffix.lower() in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_path.suffix}")
        
        self.logger.info(f"Loaded {len(df)} rows and {len(df.columns)} columns")
        return df
    
    def clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        if self.config["clean_column_names"]:
            df.columns = df.columns.str.strip().str.replace(' ', '_').str.lower()
        
        if self.config["remove_empty_rows"]:
            df = df.dropna(how='all')
        
        return df
    
    def process_survey_data(self, input_file: str, output_file: str):
        try:
            df = self.load_raw_data(input_file)
            df_cleaned = self.clean_data(df)
            
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'bg_color': '#D7E4BC',
                    'border': 1
                })
                
                df_cleaned.to_excel(writer, sheet_name='Survey_Data', index=False)
                worksheet = writer.sheets['Survey_Data']
                
                for col_num, value in enumerate(df_cleaned.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                worksheet.autofit()
            
            self.logger.info(f"Successfully processed {input_file} -> {output_file}")
            
        except Exception as e:
            self.logger.error(f"Error processing {input_file}: {str(e)}")
            raise

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Process survey data files')
    parser.add_argument('input_file', help='Input survey data file')
    parser.add_argument('-o', '--output', help='Output Excel file', default=None)
    parser.add_argument('-c', '--config', help='Configuration file', default=None)
    
    args = parser.parse_args()
    
    if not args.output:
        input_path = Path(args.input_file)
        args.output = f"output/{input_path.stem}_processed.xlsx"
    
    processor = SurveyProcessor(args.config)
    processor.process_survey_data(args.input_file, args.output)

if __name__ == "__main__":
    main()