# Survey Data Processor - Developer Guide

## Project Overview

This is a **Survey Data Processor** that transforms raw survey data from platforms like Forsta into a structured, analyzed Excel workbook with multiple specialized worksheets. The processor automatically creates question-specific tabs (Q1-Q10) with appropriate formatting, formulas, and cross-tabulation analysis based on question types.

### What This Tool Does

1. **Ingests** raw survey data Excel files with participant responses
2. **Structures** the data into multiple worksheets: raw data, data map, column question map, loop variables
3. **Analyzes** each question based on its type (single select, matrix, rank/loop, etc.)
4. **Generates** formatted question tabs (Q1-Q10) with:
   - Response options and counts
   - Percentage calculations
   - Cross-cut analysis sections with demographic filters
   - Human-readable filters
   - Properly formatted tables with borders, colors, and alignment

### Key Features

- **Automatic Question Type Detection**: Analyzes question structure to determine type (single select, single select with "Other", matrix, rank/loop)
- **Cross-Tabulation Analysis**: Creates "cross cuts" section with demographic filters (gender, age, employment, location, etc.)
- **Formula-Based Calculations**: Uses Excel formulas for dynamic calculations (counts, percentages, filters)
- **Version Management**: Auto-increments output file versions (v1, v2, v3, etc.)
- **Windows Excel Integration**: Optional win32com integration for formula calculation on Windows

---

## Architecture Overview

### Folder Structure

```
src/
├── main.py                           # Entry point with CLI
├── constants.py                      # All constants (sheet names, widths, colors)
├── pipeline.py                       # Main processing pipeline orchestration
│
├── setup/                            # Worksheet setup scripts
│   ├── initial_setup.py             # Rename tabs, create blank sheets
│   ├── raw_data.py                  # Format raw data tab
│   ├── data_map.py                  # Create data map with 27 formulas
│   └── column_question_map.py       # Create column-to-question mapping
│
├── question_types/                   # Question processing modules
│   ├── single_select.py             # Single select questions
│   └── single_select_with_other.py  # Single select with "Other Specify"
│
├── formatters/                       # Styling and formatting utilities
│   ├── styles.py                    # Border, fill, font creators
│   └── worksheet.py                 # Worksheet formatting helpers
│
├── data_extractors/                  # Data extraction from worksheets
│   └── data_map_extractor.py        # Extract from data map tab
│
└── utils/                            # General utilities
    ├── file_utils.py                # File versioning
    └── excel_calculator.py          # Excel formula calculation (win32com)
```

### Dependency Hierarchy

The codebase follows a strict import hierarchy to avoid circular dependencies:

```
Level 0: constants.py
         ↓
Level 1: formatters/styles.py, utils/*
         ↓
Level 2: data_extractors/data_map_extractor.py
         ↓
Level 3: formatters/worksheet.py
         ↓
Level 4: setup/*, question_types/*
         ↓
Level 5: pipeline.py
         ↓
Level 6: main.py
```

**Rule**: Lower-level modules never import from higher-level modules.

---

## Data Structures

### Worksheet Overview

The processor creates and manages 5 core worksheets plus 10 question tabs:

#### Core Worksheets

1. **raw data** (`SHEET_RAW_DATA`)
   - Contains original survey responses from participants
   - Each row = 1 participant
   - Each column = 1 survey field/question response
   - Headers are formatted and frozen
   - Column widths: 13 for data columns, 3 for narrow spacers

2. **data map** (`SHEET_DATA_MAP`)
   - Maps survey questions to their properties
   - Key columns:
     - **Column A**: Question marker (e.g., "Q1", "Q2_1", "Q2_2_other")
     - **Column B**: Question number (e.g., "1", "2", "3")
     - **Column C**: Question text
     - **Column D-G**: Various question attributes
     - **Column H**: Question signature (single select, matrix, etc.)
     - **Column I**: Child sequence number
     - **Column J**: Response text
     - **Column K**: Response code/value
     - **Column L**: Label text (for matrix questions, this is row label)
   - Contains 27 different formulas for analyzing question structure
   - Each response option gets its own row
   - Used by question processors to extract question metadata

3. **column question map** (`SHEET_COLUMN_QUESTION_MAP`)
   - Maps raw data columns to question numbers
   - Transpose structure: columns become rows
   - Formula-based mapping using VLOOKUP and MATCH
   - Used to identify which columns belong to which questions

4. **loop variables** (`SHEET_LOOP_VARIABLES`)
   - Stores variables for rank/loop questions
   - Currently a placeholder for future expansion

5. **Q1 through Q10** (Question Tabs)
   - One tab per question (Q1, Q2, Q3, etc.)
   - Each tab has two main sections:
     - **Top Section**: Response options with counts/percentages (rows 1-N)
     - **Cross Cuts Section**: Demographic filters with filtered counts/percentages (rows N+5 onwards)

### Question Tab Structure

Every question tab follows this structure:

```
Row 1:  [Question Text]                    [Section Number]
Row 2:  [Question Type Info]               [Bracketed Text]
Row 3:  [Blue header row]
Row 4:  Q  |  Response Options  |  n  |  %
Row 5+: Data rows with response options
...
Row X:  [Blank spacer rows]
Row X+1: "Cross Cuts" title row (blue background)
Row X+2: Q  |  Filter  |  HumRead Filter  |  n  |  %  |  record
Row X+3+: Filter rows (gender, age, employment, location, etc.)
```

#### Column Structure for Question Tabs

- **Column A-H**: Narrow spacer columns (width 3)
- **Column I**: "Q" marker column (width 3) - marks where question data starts
- **Column J**: Response options or filter names (width 50)
- **Column K**: "n" - count of responses (width 13)
- **Column L**: "%" - percentage (width 13)
- **Column M**: "HumRead Filter" - human-readable filter description (width 50, only in cross cuts)
- **Column N**: "record" - helper column for filtering (width 13, only in cross cuts)

### Question Types

#### 1. Single Select
- One response per participant
- Example: "What is your gender?" → Male, Female, Other
- **Data Map Pattern**:
  - Column A: "Q1", "Q1", "Q1" (same marker repeated)
  - Column H: "single select"
  - Column J: Response text for each option
  - Column K: Response code (1, 2, 3, etc.)

#### 2. Single Select with "Other Specify"
- Single select + free-text "Other" option
- Example: "What is your favorite color?" → Red, Blue, Green, Other (please specify)
- **Data Map Pattern**:
  - Column A: "Q5", "Q5", "Q5", "Q5_other"
  - Column H: "single select"
  - Last row has "_other" suffix in Column A
  - Uses `find_other_specify_child_text()` to detect "Other" child

#### 3. Matrix (Grid)
- Multiple questions sharing same response scale
- Example: "Rate the following on a scale of 1-5: Price, Quality, Service"
- **Data Map Pattern**:
  - Column A: "Q7_1", "Q7_2", "Q7_3" (numbered children)
  - Column H: "matrix"
  - Column I: Child sequence (1, 2, 3)
  - Column L: Row label text ("Price", "Quality", "Service")
  - Response options are shared across all children

#### 4. Rank/Loop with "Other"
- Participants rank or select multiple items
- Can include "Other Specify" option
- **Data Map Pattern**:
  - Similar to matrix but with ranking logic
  - May have "_other" child for free-text responses

### Cross Cuts (Demographic Filters)

The "Cross Cuts" section appears in every question tab starting ~10 rows below the main data. It contains filtered analysis based on demographics.

#### Filter Structure

Each filter row uses Excel formulas to count responses matching specific criteria:

```
Filter Name                 | HumRead Filter              | n              | %
Gender - Male              | Q1=1                        | =COUNTIFS(...) | =IFERROR(K24/SUM(K$24:K$33),0)
Gender - Female            | Q1=2                        | =COUNTIFS(...) | =IFERROR(K25/SUM(K$24:K$33),0)
Age - 18-24                | Q2=1                        | =COUNTIFS(...) | =IFERROR(K26/SUM(K$24:K$33),0)
```

#### Standard Cross Cut Filters

1. **Filter Q #1** - First filter question
2. **Filter Q #2** - Second filter question
3. **Gender** - Male (Q1=1), Female (Q1=2)
4. **Age groups** - 18-24, 25-34, 35-44, 45-54, 55-64, 65+
5. **Employment** - Full-time, Part-time, Self-employed, Unemployed, Student, Retired
6. **Location/Region** - Various geographic filters

#### How Filters Work

- **Column J**: Filter name (e.g., "Gender - Male")
- **Column K**: Formula using COUNTIFS to count matching responses
  ```excel
  =COUNTIFS('raw data'!$A:$A,"<>",IF(ISBLANK('raw data'!B$1),'raw data'!A$1,'raw data'!B$1),J24,'raw data'!$D:$D,I$5)
  ```
  This counts rows where:
  - Column A is not blank (participant exists)
  - Filter column matches filter value
  - Response column matches current response option

- **Column L**: Percentage formula
  ```excel
  =IFERROR(K24/SUM(K$24:K$33),0)
  ```

- **Column M**: Human-readable filter description (e.g., "Q1=1" for Gender - Male)

- **Column N**: "record" check - validates if filter has responses
  ```excel
  =IF(K24>0,"record","")
  ```

---

## Processing Pipeline

### Main Flow (`pipeline.py`)

```python
def process_excel_file(input_path, output_path):
    workbook = load_raw_excel_file(input_path)
    initial_set_up(workbook)
    raw_data_initial_setup(workbook)
    data_map_initial_setup(workbook)
    column_question_map_initial_setup(workbook)
    create_question_tabs(workbook)
    question_cutting_processor(workbook)
    save_processed_excel(workbook, output_path)
    calculate_excel_formulas(output_path)  # Optional, Windows only
```

### Step-by-Step Processing

1. **Load Raw Excel File** (`load_raw_excel_file`)
   - Opens Excel file with openpyxl
   - Returns workbook object

2. **Initial Setup** (`initial_set_up`)
   - Renames first sheet to "raw data"
   - Creates blank sheets: "data map", "column question map", "loop variables"

3. **Raw Data Setup** (`raw_data_initial_setup`)
   - Freezes top row
   - Sets column widths (13 for data, 3 for spacers)
   - Formats 937 column headers
   - Applies thin borders to all cells

4. **Data Map Setup** (`data_map_initial_setup`)
   - Creates 27 different formulas in columns A-AA
   - Key formulas:
     - **Column A**: Question marker extraction
     - **Column B**: Question number extraction
     - **Column C**: Question text extraction
     - **Column H**: Question type signature (single select, matrix, etc.)
     - **Column I**: Child sequence number
     - **Column J**: Response text
     - **Column K**: Response code
     - **Column L**: Label text (for matrix questions)
   - Copies formulas down to cover all possible question/response combinations
   - Formats headers with blue background
   - Sets column widths

5. **Column Question Map Setup** (`column_question_map_initial_setup`)
   - Transposes column headers to rows
   - Creates VLOOKUP formulas to map columns to questions
   - Identifies unique question numbers
   - Formats with borders and alignment

6. **Create Question Tabs** (`create_question_tabs`)
   - Creates 10 blank worksheets: Q1, Q2, ..., Q10
   - Positions them after core worksheets

7. **Question Cutting Processor** (`question_cutting_processor`)
   - Iterates through questions 1-10
   - For each question:
     - Detects question type using `find_question_column_h_text()`
     - Checks for "Other Specify" using `find_other_specify_child_text()`
     - Calls appropriate processor:
       - `cut_single_select_with_other()` if "Other" exists
       - `cut_single_select()` if single select
       - (Future: matrix, rank/loop processors)

8. **Save Processed Excel** (`save_processed_excel`)
   - Saves workbook to output path
   - Uses versioned filename (v1, v2, v3, etc.)

9. **Calculate Formulas** (`calculate_excel_formulas`) - Optional
   - Only runs on Windows with win32com available
   - Opens Excel via COM
   - Forces formula recalculation
   - Saves and closes
   - Ensures all formulas display calculated values

---

## Adding New Question Types

To add a new question type (e.g., multiple select, ranking, text entry):

### Step 1: Create New Question Type File

Create `src/question_types/new_question_type.py`:

```python
"""Processor for [question type] questions."""

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from ..constants import *
from ..formatters.styles import create_pale_blue_fill, create_thin_border
from ..formatters.worksheet import (
    setup_question_basic_formatting,
    add_question_text_and_section_header,
    add_row4_headers,
    apply_center_alignment_to_columns,
    add_cross_cut_section,
)
from ..data_extractors.data_map_extractor import (
    find_question_text_from_data_map,
    extract_response_options,
)

def cut_new_question_type(question_ws: Worksheet, question_number: int, workbook: Workbook):
    """
    Process a [question type] question.

    Args:
        question_ws: The worksheet for this question (e.g., Q1, Q2)
        question_number: The question number (1-10)
        workbook: The main workbook containing all sheets
    """
    # 1. Basic setup and formatting
    setup_question_basic_formatting(question_ws, include_other=False)
    add_question_text_and_section_header(question_ws, question_number, workbook)
    add_row4_headers(question_ws, include_q_header=True)

    # 2. Extract and populate response options
    data_map_ws = workbook[SHEET_DATA_MAP]
    extract_response_options(data_map_ws, question_ws, question_number)

    # 3. Add your custom formulas and logic here
    # ...

    # 4. Add cross cuts section
    apply_center_alignment_to_columns(question_ws, include_q_column=True)
    add_cross_cut_section(question_ws)
```

### Step 2: Update `question_types/__init__.py`

```python
from .single_select import cut_single_select
from .single_select_with_other import cut_single_select_with_other
from .new_question_type import cut_new_question_type  # Add this

__all__ = [
    'cut_single_select',
    'cut_single_select_with_other',
    'cut_new_question_type',  # Add this
]
```

### Step 3: Update `pipeline.py`

Modify `question_cutting_processor()` to detect and call your new processor:

```python
def question_cutting_processor(workbook: Workbook):
    for i in range(1, 11):
        question_ws = workbook[f'Q{i}']
        data_map_ws = workbook[SHEET_DATA_MAP]

        question_type = find_question_column_h_text(data_map_ws, i)
        has_other = find_other_specify_child_text(data_map_ws, i)

        # Add your detection logic
        if question_type == "your_new_type_signature":
            cut_new_question_type(question_ws, i, workbook)
        elif has_other:
            cut_single_select_with_other(question_ws, i, workbook)
        elif question_type and 'single select' in question_type.lower():
            cut_single_select(question_ws, i, workbook)
        else:
            logging.info(f"Question {i}: Unknown type '{question_type}', skipping")
```

---

## Adding New Setup Scripts

To add a new worksheet setup (e.g., for a new data analysis tab):

### Step 1: Create New Setup File

Create `src/setup/new_worksheet_setup.py`:

```python
"""Setup script for [worksheet name] worksheet."""

from openpyxl.workbook.workbook import Workbook
from ..constants import *
from ..formatters.styles import create_pale_blue_fill, create_thin_border
from ..formatters.worksheet import apply_column_widths

def new_worksheet_initial_setup(workbook: Workbook):
    """
    Set up the [worksheet name] worksheet.

    Args:
        workbook: The main workbook
    """
    ws = workbook['worksheet_name']

    # Your setup logic here
    # - Add headers
    # - Add formulas
    # - Apply formatting
    # - Set column widths
```

### Step 2: Update `setup/__init__.py`

```python
from .initial_setup import initial_set_up
from .raw_data import raw_data_initial_setup
from .data_map import data_map_initial_setup
from .column_question_map import column_question_map_initial_setup
from .new_worksheet_setup import new_worksheet_initial_setup  # Add this

__all__ = [
    'initial_set_up',
    'raw_data_initial_setup',
    'data_map_initial_setup',
    'column_question_map_initial_setup',
    'new_worksheet_initial_setup',  # Add this
]
```

### Step 3: Update `pipeline.py`

Add your setup call to `process_excel_file()`:

```python
def process_excel_file(input_path: str, output_path: str):
    workbook = load_raw_excel_file(input_path)
    initial_set_up(workbook)
    raw_data_initial_setup(workbook)
    data_map_initial_setup(workbook)
    column_question_map_initial_setup(workbook)
    new_worksheet_initial_setup(workbook)  # Add this
    create_question_tabs(workbook)
    question_cutting_processor(workbook)
    save_processed_excel(workbook, output_path)
    calculate_excel_formulas(output_path)
```

---

## Development Guidelines

### Code Style

- Use type hints for function parameters and return values
- Use descriptive variable names (e.g., `question_ws` not `ws`, `data_map_ws` not `dm`)
- Add docstrings to all functions with Args and Returns sections
- Keep functions focused on a single responsibility
- Use constants from `constants.py` instead of hardcoding values

### Import Rules

1. **Never import from higher-level modules**
   - ❌ `constants.py` importing from `pipeline.py`
   - ✅ `pipeline.py` importing from `constants.py`

2. **Use relative imports within src/**
   - ✅ `from ..constants import *`
   - ✅ `from ..formatters.styles import create_pale_blue_fill`
   - ❌ `from src.constants import *`

3. **Import order**:
   - Standard library imports
   - Third-party imports (openpyxl, pandas)
   - Local imports (from ..constants, from ..formatters, etc.)

### Testing

Before committing changes:

1. **Test with sample data**:
   ```bash
   python -m src.main data/raw/raw_pilot_forsta.xlsx
   ```

2. **Verify output file**:
   - Check all worksheets are created
   - Check formulas are correct
   - Check formatting (borders, colors, widths)
   - Spot-check calculations

3. **Test edge cases**:
   - Questions with no responses
   - Questions with "Other Specify"
   - Matrix questions with many children
   - Missing data in raw file

### Common Pitfalls

1. **Don't hardcode sheet names** - Use constants:
   ```python
   # ❌ Bad
   ws = workbook['raw data']

   # ✅ Good
   ws = workbook[SHEET_RAW_DATA]
   ```

2. **Don't forget to apply formatting**:
   - Borders
   - Column widths
   - Blue headers
   - Alignment

3. **Use proper Excel formula syntax**:
   - Single quotes around sheet names with spaces: `'raw data'!A:A`
   - Absolute references for column headers: `$A$1`
   - Relative references for data rows: `A2`

4. **Handle missing data gracefully**:
   ```python
   # ✅ Good - check before accessing
   if question_text:
       question_ws['C1'] = question_text
   ```

---

## Dependencies

### Required
- **Python 3.7+**
- **openpyxl** - Excel file manipulation
- **pandas** - Data manipulation (minimal use)

### Optional
- **pywin32** (win32com.client) - Windows-only, for Excel formula calculation
  - Only needed if you want formulas to show calculated values immediately
  - Falls back gracefully if not available
  - Mac/Linux users: formulas will calculate when file is opened in Excel

### Installation

```bash
# Install required dependencies
pip install openpyxl pandas

# Optional: Install Windows COM support
pip install pywin32  # Windows only
```

---

## Common Workflows

### Processing a Survey File

```bash
# Basic usage - auto-versions output
python -m src.main data/raw/raw_pilot_forsta.xlsx

# Specify output file
python -m src.main data/raw/raw_pilot_forsta.xlsx --output_file output/custom_name.xlsx
```

### Adding a New Question

1. Ensure your raw data file has the question data
2. Update data map formulas if needed (usually automatic)
3. If it's a new question type:
   - Create processor in `src/question_types/`
   - Update `pipeline.py` to detect and call it
4. Run processor and verify output

### Debugging Formulas

1. **Check data map first**:
   - Open output file
   - Go to "data map" tab
   - Verify Column A (question marker), B (question number), H (question type)
   - Check if response options are correctly extracted in Column J

2. **Check question tab**:
   - Verify response options are populated
   - Check formula syntax in count columns
   - Verify filter formulas in cross cuts section

3. **Use Excel formula auditing**:
   - Open in Excel
   - Use "Trace Precedents" to see what cells formulas reference
   - Use "Evaluate Formula" to step through complex formulas

---

## Glossary

- **Question Marker**: Unique identifier for a question in data map (e.g., "Q1", "Q2_1", "Q5_other")
- **Question Signature**: Question type indicator in Column H of data map (e.g., "single select", "matrix")
- **Child**: Sub-question in a matrix or loop (e.g., Q7_1, Q7_2, Q7_3)
- **Other Specify**: Free-text field for "Other" option in questions
- **Cross Cuts**: Demographic filter analysis section in question tabs
- **HumRead Filter**: Human-readable filter description (e.g., "Q1=1" for Gender - Male)
- **Record Check**: Formula that verifies if a filter row has any responses

---

## Future Enhancements

Potential areas for expansion:

1. **More Question Types**:
   - Multiple select (select all that apply)
   - Ranking questions
   - Open-ended text questions
   - Numeric entry questions

2. **Advanced Analysis**:
   - Statistical significance testing
   - Correlation analysis
   - Trend analysis across waves

3. **Performance**:
   - Parallel processing for large files
   - Chunked reading for very large datasets
   - Caching of intermediate results

4. **Export Formats**:
   - CSV export of cross tabs
   - PowerPoint chart generation
   - PDF report generation

5. **Configuration**:
   - YAML/JSON config files for filter definitions
   - Customizable cross cut filters per project
   - Template support for different survey platforms

---

## Contact & Support

For questions about this codebase, refer to:
- This claude.md file
- Inline code comments and docstrings
- Git commit history for implementation details
