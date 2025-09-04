# VN01 Future Ship Mode & OH Merge Automation

Automates merging of VN01 On-Hand (OH) Parts data with Future Ship Mode (FSM) files and outputs a consolidated Excel file filtered by specific criteria.

This script is designed to streamline weekly reporting for the Europe & Other Asia Team at Jabil, reducing manual Excel handling and improving data accuracy.

---

## Features

- Reads VN01 OH Parts Excel file and Future Ship Mode Excel file for the current ISO week.
- Standardizes column names for consistent merging.
- Merges OH and FSM data on Customer Part. column.
- Filters merged data for PGr codes starting with 'U'.
- Selects only the relevant columns for reporting.
- Outputs a weekly Excel file named 'JB <file_name>.xlsx' in the corresponding week folder.
- Prints a list of unique Buyer Names for quick reference.

---

## How It Works

1. Determine Current Week & File Paths
   - Automatically generates paths based on the current ISO week for FSM files.
   - Generates the current week's Monday date to locate OH part files.

2. Data Processing
   - Reads Excel files using pandas.
   - Renames columns for consistency.
   - Merges OH and FSM data.
   - Filters merged data for specific PGr codes.
   - Selects desired columns for the final report.

3. Output
   - Saves the consolidated Excel report in the structured folder path:
     //sgsind0nsifsv01a/.../FUTURE SHIP MODE/WWxx'yy/JB <file_name>.xlsx
   - Prints unique Buyer Names in the console.

---
