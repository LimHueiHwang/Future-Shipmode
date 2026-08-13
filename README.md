# VN01 Future Ship Mode & OH Merge Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![Python](https://img.shields.io/badge/Python-Automation-blue)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-blue)
![Excel](https://img.shields.io/badge/Excel-Automation-green)

> Python automation component for consolidating VN01 On-Hand data with Future Supplier Ship Mode data, applying Purchasing Group filtering, and generating a standardized Excel report.

---

## Overview

This project automates a weekly Purchasing reporting process involving two Excel reports:

1. **VN01 On-Hand (OH) report**
2. **Future Supplier Ship Mode (FSM) report**

The original process required the Purchasing user to manually combine the two reports, match parts, identify the relevant Future Purchasing Group, filter the required records, and prepare the final report.

The Python automation was developed to handle these repetitive data-processing steps automatically.

The automation matches the two reports using `Customer Part.`, uses the Future supplier's `PGr`, filters records where the `PGr` starts with `U`, and generates the consolidated Excel report.

A Buyer list is also generated for Purchasing reference.

**Project Status:** Production

---

## Business Problem

The weekly VN01 Future Ship Mode process requires information from multiple Excel reports to be combined and filtered.

The manual process involved:

* Opening the VN01 On-Hand report.
* Opening the Future Ship Mode report.
* Matching Customer Parts.
* Identifying the relevant Purchasing Group.
* Filtering the required Purchasing Groups.
* Preparing the final consolidated report.
* Identifying the relevant Buyers.
* Saving the final report to the correct server location.

This created repetitive manual work and introduced the possibility of:

* Incorrect part matching.
* Incorrect Purchasing Group filtering.
* Manual data-processing errors.
* Incorrect output file locations.
* Additional processing time.

The automation was developed to standardize these steps.

---

## Solution

The Python automation processes the two source Excel reports and produces the required Purchasing output.

The automation:

1. Loads the VN01 On-Hand report.
2. Loads the Future Ship Mode report.
3. Validates the required input columns.
4. Standardizes the VN01 `Material` field to `Customer Part.`.
5. Matches the two datasets using `Customer Part.`.
6. Separates the VN01 and Future `PGr` fields.
7. Uses the Future supplier's `PGr` as the primary Purchasing Group.
8. Filters records where `PGr` starts with `U`.
9. Generates the consolidated Excel report.
10. Generates a unique Buyer list.
11. Saves the output to the designated server location.

The Python automation focuses on the data-processing portion of the workflow while the VN01 report generation from SAP remains outside the current automation.

---

## Key Features

### Automated Excel Processing

The automation loads and processes both source Excel reports without requiring manual data copying between workbooks.

---

### Customer Part Matching

The two reports are matched using:

```text
Customer Part.
```

The VN01 On-Hand report originally uses:

```text
Material
```

The automation standardizes this field before performing the merge.

---

### Future Purchasing Group Filtering

The Future supplier's `PGr` is used as the authoritative Purchasing Group.

Only records where:

```text
PGr starts with "U"
```

are included in the final output.

---

### Duplicate PGr Handling

Both source reports contain a `PGr` column.

When the datasets are merged, pandas can generate duplicate column names such as:

```text
PGr_x
PGr_y
```

The automation handles this explicitly:

```text
VN01 PGr
    ↓
PGr_oh

Future PGr
    ↓
PGr
```

This ensures that the Future supplier's `PGr` remains the primary field used for filtering.

---

### Buyer List Generation

After filtering the data, the automation generates a unique Buyer list for Purchasing reference.

Example:

```text
Buyer Names:
------------------------------
Buyer A
Buyer B
Buyer C
```

---

### Server Output

The final report is automatically saved to the designated VN01 server location using the required output naming convention.

---

# Technologies Used

| Category             | Technology            |
| -------------------- | --------------------- |
| Programming Language | Python                |
| Data Processing      | pandas                |
| Excel Processing     | openpyxl              |
| Input                | Excel `.xlsx`         |
| Output               | Excel `.xlsx`         |
| File Storage         | Windows Network Share |

---

# Workflow

The complete automation workflow is shown below.

![VN01 Future Ship Mode Workflow](docs/images/workflow.png)

### Workflow Steps

**1. VN01 On-Hand Report**

The existing VN01 On-Hand Excel report is used as the first input.

**2. Future Ship Mode Report**

The Future Supplier Ship Mode Excel report is used as the second input.

**3. Load and Validate**

Python loads both Excel files and validates the required columns.

**4. Standardize Customer Part**

The VN01 `Material` field is standardized to `Customer Part.`.

**5. Match Data**

The two datasets are merged using `Customer Part.`.

**6. Apply Future PGr**

The Future supplier's `PGr` is retained as the primary Purchasing Group.

**7. Filter Purchasing Group**

Only records where `PGr` starts with `U` are retained.

**8. Generate Report**

The filtered data is written to the consolidated Excel report.

**9. Generate Buyer List**

The automation generates a unique Buyer list for reference.

**10. Save Output**

The final report is saved to the designated server folder.

---

# Architecture

The solution uses Python as the data-processing layer between the source Excel reports and the final Purchasing output.

![VN01 Future Ship Mode Architecture](docs/images/architecture.png)

## Architecture Components

| Component              | Responsibility                                  |
| ---------------------- | ----------------------------------------------- |
| VN01 On-Hand Excel     | Provides VN01 On-Hand information               |
| Future Ship Mode Excel | Provides Future supplier shipment information   |
| Python                 | Controls the automation workflow                |
| pandas                 | Performs data processing, merging and filtering |
| openpyxl               | Reads and writes Excel files                    |
| Server                 | Stores input and output reports                 |
| Purchasing User        | Provides/initiates the required input process   |

---

## SAP Boundary

The current Python automation does **not** directly extract the VN01 report from SAP.

The current process starts after the VN01 On-Hand report has already been generated.

```text
SAP
 │
 ▼
VN01 On-Hand Excel
 │
 ├──────────────────┐
 │                  │
 ▼                  ▼
Python          Future Ship Mode
Automation          Excel
 │
 ▼
Consolidated Report
```

This keeps the current automation focused on the Excel data-processing requirement.

---

# Data Processing

## VN01 On-Hand

The VN01 On-Hand report uses:

```text
Material
```

The automation converts this into:

```text
Customer Part.
```

This allows it to be matched against the Future report.

---

## Future Ship Mode

The Future Ship Mode report already contains:

```text
Customer Part.
```

Therefore, the matching key is:

```text
Customer Part.
```

---

## Purchasing Group

Both source reports contain:

```text
PGr
```

The automation separates them during the merge:

```text
VN01 PGr     → PGr_oh
Future PGr   → PGr
```

The Future supplier's `PGr` is then used for the final filtering.

---

## Filtering Rule

The final report includes only records where:

```text
PGr starts with "U"
```

This is the primary business rule used to determine which Future records are included in the output.

---

# Input Files

## VN01 On-Hand

Current server location:

```text
\\sgsind0nsifsv01a\IMAC Data\IMAC Senior or Teams\Europe & Other Asia Team\VN01\VN01 on hand part\
```

Example:

```text
VN01 OH 0810.xlsx
```

---

## Future Ship Mode

Current server location:

```text
\\sgsind0nsifsv01a\IMAC Data\IMAC Senior or Teams\Europe & Other Asia Team\VN01\FUTURE SHIP MODE\WWxx'yy\
```

Example weekly folder:

```text
WW33'26
```

Example input:

```text
VN01-Backlog Report Shipment.xlsx
```

The current implementation requires the user to provide the Future Ship Mode filename.

---

# Future Ship Mode Input Structure

The current Future Ship Mode report contains:

```text
Ship To
Order Number
Line Number
Customer Part.
Mfr
Mfr Part.
Qty Required
Qty Allocated
Value Alloc USD
Value USD
Date Ordered
Date Required
Current Customer Commit
Current Commit Date
Resale Price
MPQ
Ship (Transport) Date
Invoice Number
Weight(grams)
Date Code
Via Code
Ship mode
PGr
```

The automation validates the input structure before processing.

---

# Output

The production output follows the naming convention:

```text
JB <file_name>.xlsx
```

Example:

```text
JB VN01-Backlog Report Shipment.xlsx
```

During testing, `_NEW` was used to prevent overwriting the existing production output:

```text
JB VN01-Backlog Report Shipment_NEW.xlsx
```

The final output contains the consolidated records after applying the required matching and `PGr` filtering rules.

---

# Error Handling

The automation performs input validation before processing.

### Missing Input Files

The automation detects when the expected source file cannot be loaded.

### Missing Columns

Required columns are validated before the data-processing stage.

### Invalid Report Structure

Unexpected changes to the source report structure can stop the process before an incorrect report is generated.

### Excel Processing Errors

Excel read/write errors are handled as processing failures rather than silently producing incomplete output.

### Server Access

The automation depends on access to the designated network share.

---

# Project Structure

```text
Future-Shipmode/
│
├── README.md
├── Future_Shipmode.py
├── .gitignore
│
└── docs/
    └── images/
        ├── workflow.png
        └── architecture.png
```

---

# Installation

Install the required Python packages:

```powershell
python -m pip install pandas openpyxl
```

The automation requires access to:

* Python environment.
* VN01 server folders.
* Source Excel reports.
* Output server location.

---

# Testing

The automation was tested using the actual working server files.

| Test                            | Status   |
| ------------------------------- | -------- |
| VN01 On-Hand file detection     | ✅ Passed |
| Future Ship Mode file detection | ✅ Passed |
| Excel file loading              | ✅ Passed |
| Required column validation      | ✅ Passed |
| Customer Part. matching         | ✅ Passed |
| Duplicate PGr handling          | ✅ Passed |
| Future PGr filtering            | ✅ Passed |
| Consolidated report generation  | ✅ Passed |
| Buyer list generation           | ✅ Passed |
| Server output                   | ✅ Passed |

**Current Status:** Tested and working.

---

# Business Impact

The automation improves the weekly Purchasing reporting process by:

* Reducing repetitive Excel manipulation.
* Standardizing the part-matching process.
* Applying the same Purchasing Group filtering rule every time.
* Reducing the risk of manual filtering errors.
* Producing a consistent report format.
* Reducing the time required to prepare the report.
* Providing a Buyer reference list.

No percentage-based improvement is claimed because formal time measurements have not yet been recorded.

---

# Current Limitations

### Server Dependency

The automation depends on the existing network folder structure.

Changes to the server location may require configuration or code changes.

### Input Report Dependency

The source reports must maintain the expected column structure.

Changes to the report format may require updates to the automation.

### Manual Future Filename

The Future Ship Mode filename currently needs to be provided by the user.

### SAP Extraction

The Python automation does not currently extract the VN01 report directly from SAP.

The VN01 On-Hand report must already exist as an Excel file.

---

# Future Improvements

Potential future improvements include:

### Input Automation

* Automatically detect the latest Future Ship Mode report.
* Automatically identify the current weekly folder.
* Add a file-selection interface.
* Improve source-file validation.

### Reporting

* Add a summary worksheet.
* Display total input records.
* Display total matched records.
* Display total filtered records.
* Display total Buyers.
* Add processing timestamps.

### Logging

* Add structured processing logs.
* Add detailed error logs.
* Record processing duration.

### Notification

* Add email notification.
* Add Microsoft Teams notification.

### Deployment

* Add `requirements.txt`.
* Package the automation as a standalone executable.
* Improve deployment documentation.

### SAP Integration

If required, the VN01 report extraction could be integrated into the automation in a future version.

---

# Lessons Learned

This project demonstrated several important automation principles.

### Validate Real-World Input

Testing with the actual server reports was important because the real report structure and file paths needed to be handled correctly.

### Understand the Business Rule

The Future supplier's `PGr` is the authoritative Purchasing Group for this process.

Therefore, the automation must filter using the Future `PGr`, rather than simply using whichever `PGr` column appears first.

### Handle Duplicate Columns

When merging datasets with identical column names, pandas can automatically generate `_x` and `_y` suffixes.

Explicitly handling these columns prevents unexpected processing errors.

### Keep Automation Focused

The Python automation focuses on the repetitive Excel data-processing task rather than attempting to replace the entire SAP reporting process.

---

# My Role

I was responsible for developing and improving the automation for the VN01 Future Ship Mode reporting process.

My responsibilities included:

* Understanding the existing Purchasing workflow.
* Identifying the manual data-processing steps.
* Designing the Python automation.
* Developing the Excel data-processing logic.
* Implementing Customer Part matching.
* Implementing Future PGr filtering.
* Handling duplicate `PGr` columns.
* Implementing input validation.
* Generating the consolidated report.
* Generating the Buyer reference list.
* Testing the automation using actual server files.
* Troubleshooting file-path and Excel dependency issues.
* Preparing the project for GitHub documentation.

---

# Engineering Skills Demonstrated

* Python
* pandas
* openpyxl
* Excel Automation
* Data Processing
* Data Merging
* Data Validation
* File Handling
* Network Share Integration
* Error Handling
* Business Process Automation
* SAP Purchasing Process Understanding
* Troubleshooting
* Automation Documentation

---

# Project Information

| Item                 | Details                                  |
| -------------------- | ---------------------------------------- |
| Project Status       | Production                               |
| Project Type         | Purchasing Process Automation            |
| Primary Function     | Future Ship Mode & VN01 OH Consolidation |
| Programming Language | Python                                   |
| Data Processing      | pandas                                   |
| Excel Processing     | openpyxl                                 |
| Input                | VN01 On-Hand + Future Ship Mode          |
| Output               | Consolidated Excel Report                |
| Business Area        | Purchasing                               |
| Site                 | VN01                                     |

---

# Key Takeaway

This project demonstrates how Python can be used to remove repetitive data-processing work from an existing Purchasing workflow.

The original manual process:

```text
Multiple Excel Reports
        ↓
Manual Matching
        ↓
Manual Filtering
        ↓
Manual Report Preparation
```

was replaced with:

```text
VN01 On-Hand
      +
Future Ship Mode
      ↓
Python Automation
      ↓
Customer Part Matching
      ↓
Future PGr Filtering
      ↓
Consolidated Report
      ↓
Buyer Reference
```

The result is a more consistent and repeatable weekly reporting process with less manual Excel manipulation.
