# Data Processing for LMS

This project provides Python scripts to help LMS administrators optimize and automate data processing tasks related to student records, results, and reporting. The workflow streamlines merging data from multiple systems (Stars, Canvas/GOALS), filtering and cleaning records, and generating formatted Excel reports with outcome highlights.

## Overview

The repository includes three main scripts:

- **dataModelling.py:** Filters and organizes student info into categorized Excel sheets.
- **mergeFile2.py:** Merges and cleans data from Stars and Canvas/GOALS CSVs to produce unified results.
- **convertTable2.py:** Transforms merged results into a long format and applies outcome-based cell color formatting in Excel.

## Features

- Data ingestion from CSV files exported via LMS subsystems (Stars, Canvas/GOALS)
- Data cleaning and standardization (removing special characters, formatting columns)
- Robust merging of different sources via Student ID
- Filtering by status and outstanding fees
- Unpivoting scores to facilitate analytics/reporting
- Automated creation of multi-sheet Excel files with outcome-based highlights

## Workflow

Below is the recommended step-by-step workflow using your scripts:

1. **Prepare CSVs:** Export required CSVs from Stars and Canvas/GOALS.
2. **Merge Data:** Use `mergeFile2.py` to create a unified dataset (`merged_final_scores_2.csv`).
3. **Model Data:** Use `dataModelling.py` for categorized student reports (`student_report.xlsx`).
4. **Transform & Format:** Use `convertTable2.py` to unpivot scores and create color-coded Excel reports (`output.xlsx`).

## Requirements

- Python 3.x
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage

### 1. Merge the CSV files

Edit the file paths in `mergeFile2.py` to point to your exported Stars and Canvas/GOALS CSVs.

```python
CSV_1 = r"path/to/canvas.csv"
CSV_2 = r"path/to/stars.csv"
```

Run the script:

```bash
python mergeFile2.py
```

This creates `merged_final_scores_2.csv`.

### 2. Model and Filter Data

Edit the file paths in `dataModelling.py`:

```python
CSV_1 = r"path/to/Book1.csv"
```

Run:

```bash
python dataModelling.py
```

This creates `student_report.xlsx` with three sheets:
- Basic_Info: active students
- Status_Filtered: students on leave or finished
- Outstanding_500_Plus: students with outstanding fees ≥ 500

### 3. Convert and Format Scores

Update the file path in `convertTable2.py`:

```python
FILE_PATH = r"path/to/merged_final_scores_2.csv"
```

Run:

```bash
python convertTable2.py
```

This generates `output.xlsx` with:
- Sheet "Original": cleaned, merged data
- Sheet "Unpivoted": long format, each outcome colored
  - Green: Pass
  - Red: Fail
  - Yellow: Credit Transfer
  - Black: Invalid/Not found

## Project Structure

```
data-processing-for-LMS/
├── convertTable2.py
├── dataModelling.py
├── mergeFile2.py
└── README.md
```

## Notes

- Ensure your CSV files have correct headers as expected by the scripts.
- All file paths are absolute by default. Change them as needed for your environment.
- The scripts do not modify source files.


## Contact

aphiwut.jan@gmail.com
