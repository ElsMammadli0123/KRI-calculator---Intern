# KRI Calculator

## Overview
This script calculates Key Risk Indicators (KRIs) for IT systems based on incident data provided in an Excel file. It outputs an Excel report with KRI values and highlights any threshold breaches.

---

## 1. Setup Instructions

### Prerequisites
- Python 3.8+
- Required Python packages:
  - openpyxl

### Installation
1. Clone or download this repository to your local machine.
2. Install dependencies:
   ```bash
   pip install openpyxl
   ```
3. Place your input Excel file (named `internship_task_data.xlsx`) in the same directory as the script, or update the `INPUT_FILE` variable in `KRI_calculator.py` to match your file name.

---

## 2. Calculation Methodology

### KRI1: MTBF (Mean Time Between Failures)
- For each system, incidents are sorted by start time.
- MTBF is calculated as the average number of days between consecutive incidents.
- If a system has fewer than 2 incidents, MTBF is marked as "N/A".
- If MTBF is below the threshold (default: 90 days), status is "Breach"; otherwise, "OK".

### KRI2: Monthly Incident Count
- For each system, the total number of incidents is counted (across all months).
- If the count exceeds the threshold (default: 3), status is "Breach"; otherwise, "OK".

### KRI3 & KRI4: RTO (Recovery Time Objective) Analysis
- For each incident, if duration > RTO threshold (default: 120 minutes), it is counted as "Exceeded RTO"; otherwise, "Within RTO".
- For each system:
  - If any incident exceeded RTO, "Exceed Status" is "Breach".
  - If more than 3 incidents are within RTO, "Within Status" is "Breach".

---

## 3. User Guide

### Running the Script
1. Ensure your input Excel file is named `internship_task_data.xlsx` and is formatted as described below.
2. Run the script:
   ```bash
   python KRI_calculator.py
   ```
3. The script will:
   - Print loading and calculation logs to the terminal.
   - Generate an output Excel file named `KRI_FINAL_RESULTS.xlsx` in the same directory.


### Input File Format
- Sheet name: `incident_data`
- Columns:
  - **A:** System (system name row, then blank for incidents)
  - **B:** Incident start date (e.g., 01/24/25 02:30:00 PM)
  - **C:** Incident end date (same format)
  - **D:** Incident duration (mins)
- Each system starts with a row with the system name in column A, followed by incident rows with data in columns B, C, D.

### Output
- The output Excel file will highlight any threshold breaches in red.
- Console output will show data loading and calculation logs for transparency.

---

## Troubleshooting
- Permission Denied Error: Close the output Excel file before running the script again.
- No Data Loaded: Check that your input file matches the expected format and worksheet name.
- Ensure date formats in your Excel file match the expected format: `MM/DD/YY HH:MM:SS AM/PM`.
- Unexpected Results: Review the printed logs for errors or warnings.
