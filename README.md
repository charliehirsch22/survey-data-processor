# Survey Data Processor

A Python tool for converting raw survey data into formatted Excel workbooks.

## Setup

1. Activate the virtual environment:
   ```bash
   source survey_processor_env/bin/activate
   ```

2. Install dependencies (already done):
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage
```bash
python src/survey_processor.py data/raw/your_survey_data.csv
```

### With custom output file
```bash
python src/survey_processor.py data/raw/survey.csv -o output/formatted_survey.xlsx
```

### With configuration file
```bash
python src/survey_processor.py data/raw/survey.csv -c config/survey_config.json
```

## Project Structure

```
├── data/
│   ├── raw/           # Place your raw survey files here
│   └── processed/     # Intermediate processed files
├── src/
│   └── survey_processor.py  # Main processing script
├── config/
│   └── survey_config.json   # Configuration settings
├── output/            # Final Excel workbooks
├── tests/             # Unit tests
└── requirements.txt   # Python dependencies
```

## Features

- Auto-detects file encoding
- Supports CSV and Excel input formats
- Cleans column names and removes empty rows
- Creates formatted Excel output with headers
- Configurable processing options
- Logging for debugging

## Configuration

Edit `config/survey_config.json` to customize:
- Input encoding detection
- Column name cleaning
- Excel formatting options
- Data cleaning rules