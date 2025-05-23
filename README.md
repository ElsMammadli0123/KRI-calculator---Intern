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

### KRI1: Mean Time Between Critical System Failures (MTBF)
- **Definition:** Average number of days between two consecutive incidents for each system.
- **Threshold:** 90 days (breach if MTBF < 90 days).
- **Calculation:** For each system, calculate the days between each pair of consecutive incidents, then average these values.

### KRI2: Number of Incidents Affecting Critical Applications (Monthly)
- **Definition:** Number of incidents per system per month.
- **Threshold:** 3 incidents per month (breach if count > 3).
- **Calculation:** Count the number of incidents for each system in each month.

### KRI3: Number of Incidents Resulting in Downtime Exceeding RTO
- **Definition:** Number of incidents where downtime exceeds the Recovery Time Objective (RTO).
- **Threshold:** 0 (breach if any incident exceeds RTO).
- **Calculation:** Count incidents per system where duration > 120 minutes (2 hours).

### KRI4: Number of Incidents Resulting in Downtime Within RTO
- **Definition:** Number of incidents where downtime is within the RTO.
- **Threshold:** 3 (breach if count > 3).
- **Calculation:** Count incidents per system where duration <= 120 minutes.

---

## 3. User Guide

### Running the Script
1. Ensure your input Excel file is named `internship_task_data.xlsx` and is formatted as described below.
2. Run the script:
   ```bash
   python KRI_calculator.py
   ```
3. The script will output `KRI_FINAL_RESULTS.xlsx` with three sheets:
   - `KRI1 - MTBF`
   - `KRI2 - Monthly`
   - `KRI3-KRI4 - RTO`

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
- If the output file is empty, check the console for errors or data loading issues.
- Ensure date formats in your Excel file match the expected format: `MM/DD/YY HH:MM:SS AM/PM`. 