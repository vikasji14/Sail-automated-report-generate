# Cobble Detection Analysis Report Generator

This tool automatically generates comprehensive reports from cobble detection data stored in Excel files. The report includes detailed visualizations and insights across multiple dimensions of analysis.

## Features

- **Time-Based Analysis**: Visualizes cobble event trends over time, including daily frequencies, hourly distributions, and shift patterns.
- **Block-Specific Analysis**: Identifies blocks with highest/lowest cobble occurrences and compares performance across blocks.
- **Profile-Based Analysis**: Analyzes impact of profile variations on cobble detection.
- **Predictive Performance**: Evaluates short-term cobble prediction effectiveness (10-min and 20-min detection).
- **Anomaly Detection**: Identifies unusual patterns or spikes in cobble occurrences.
- **Shift Analysis**: Compares cobble occurrence rates across different shifts.
- **Sequential Analysis**: Examines patterns between consecutive cobble events.
- **Machine Learning Insights**: Provides performance metrics of ML models in forecasting cobble events.

## Installation

1. Ensure Python 3.7+ is installed on your system
2. Install required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Report Generation

1. Run the report generation script:
   ```
   python report_automation.py
   ```
2. When prompted, select your Excel file containing cobble detection data
3. The tool will generate a comprehensive report in DOCX format and save it in a "Reports" folder next to your input file
4. The report folder will automatically open when processing is complete

### Data Cleaning Utility

If you encounter issues with your data files, use the data cleaning utility:

1. Run the data cleaner:
   ```
   python data_cleaner.py
   ```
2. Select the Excel file that needs cleaning
3. Review the identified issues
4. Save the cleaned file to use with the report generator

## Input Data Format

Your Excel file should have the following columns:
- `Date`: Date of the event (MM/DD/YYYY format)
- `Time`: Time of the event (standard time format like "1:30 PM" or "13:30")
- `Block`: Block identifier
- `Profile`: Profile value
- `cobble_detected_10min`: Whether cobble was detected in 10 min window (YES/NO)
- `cobble_detected_20min`: Whether cobble was detected in 20 min window (YES/NO)
- Data quality assessment

## System Requirements

- Windows, macOS, or Linux
- Python 3.7+
- 4GB RAM minimum (8GB recommended for large datasets)
- Microsoft Word or compatible document viewer to open reports
