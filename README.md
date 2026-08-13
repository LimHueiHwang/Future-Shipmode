# VN01 Future Ship Mode & OH Merge Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![Python](https://img.shields.io/badge/Python-Automation-blue)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-blue)
![Excel](https://img.shields.io/badge/Excel-Automation-green)

> Python automation for consolidating VN01 On-Hand data with Future Supplier Ship Mode data, matching Customer Parts, applying VN01 Purchasing Group filtering, and generating a consolidated Excel report.

---

## Overview

This project automates a weekly Purchasing reporting process involving two Excel reports:

1. **VN01 On-Hand (OH) report**
2. **Future Supplier Ship Mode (FSM) report**

The original process required the Purchasing user to manually combine the two reports, match parts, identify the required Purchasing Groups, filter the relevant records, and prepare the final report.

The Python automation was developed to handle these repetitive data-processing steps automatically.

The automation:

* Loads the VN01 On-Hand report.
* Loads the Future Ship Mode report.
* Standardizes the VN01 `Material` field to `Customer Part.`.
* Matches the two reports using `Customer Part.`.
* Retains the `PGr` from the VN01 On-Hand report.
* Filters records where `PGr` starts with `U`.
* Generates the consolidated Excel report.
* Generates a Buyer reference list.

**Project Status:** Production

---

## Business Problem

The weekly VN01 Future Ship Mode process requires information from multiple Excel reports to be combined and filtered.

The original manual process involved:

* Opening the VN01 On-Hand report.
* Opening the Future Ship Mode report.
* Matching Customer Parts.
* Applying the required Purchasing Group filtering.
* Preparing the final consolidated report.
* Identifying the relevant Buyers.
* Saving the final report to the correct server location.

This created repetitive manual work and introduced the possibility of:

* Incorrect part matching.
* Incorrect Purchasing Group filtering.
* Manual data-processing errors.
* Incorrect output files.
* Additional processing time.

The automation was developed to standardize these steps and make the process repeatable.

---

## Solution

The Python automation processes the two source Excel reports and produces the required Purchasing output.

The automation:

1. Loads the VN01 On-Hand report.
2. Loads the Future Ship Mode report.
3. Validates the required input columns.
4. Converts VN01 `Material` to `Customer Part.`.
5. Matches both datasets using `Customer Part.`.
6. Retains the VN01 `PGr` in the merged dataset.
7. Filters records where `PGr` starts with `U`.
8. Selects the required output columns.
9. Generates the consolidated Excel report.
10. Generates a unique Buyer list.
11. Saves the output to the designated server location.

The Python automation focuses on the Excel data-processing portion of the workflow.

---

## Key Features

### Automated Excel Processing

The automation loads and processes the VN01 On-Hand and Future Ship Mode Excel reports without requiring manual data copying between files.

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

The automation renames this field to:

```text
Customer Part.
```

before the merge.

---

### VN01 Purchasing Group Filtering

The `PGr` field comes from the **VN01 On-Hand report**.

The Future Ship Mode report does not provide `PGr`.

After the two datasets are merged, the VN01 `PGr` is retained and used for filtering.

The business rule is:

```text
PGr starts with "U"
```

Only records satisfying this rule are included in the final output.

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

The final consolidated report is automatically saved to the designated VN01 server location.

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

![VN01 Future Ship Mode Workflow](images/workflow.png)

### Workflow Steps

**1. VN01 On-Hand Report**

The existing VN01 On-Hand Excel report provides the VN01 part information and Purchasing Group.

**2. Future Ship Mode Report**

The Future Supplier Ship Mode Excel report provides the Future shipment information and `Customer Part.`.

**3. Load and Validate**

Python loads both Excel files and validates the required columns.

**4. Standardize Customer Part**

The VN01 `Material` field is renamed to `Customer Part.`.

**5. Match Data**

The two datasets are merged using `Customer Part.`.

**6. Retain VN01 PGr**

The `PGr` from the VN01 On-Hand report remains in the merged dataset.

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

![VN01 Future Ship Mode Architecture](images/architecture.png)

## Architecture Components

| Component              | Responsibility                                      |
| ---------------------- | --------------------------------------------------- |
| VN01 On-Hand Excel     | Provides VN01 part and Purchasing Group information |
| Future Ship Mode Excel | Provides Future shipment information                |
| Python                 | Controls the automation workflow                    |
| pandas                 | Performs data processing, merging and filtering     |
| openpyxl               | Reads and writes Excel files                        |
| Server                 | Stores input and output reports                     |
| Purchasing User        | Initiates the required input process                |

---

## SAP Boundary

The current Python automation does **not** directly extract the VN01 report from SAP.

The process starts after the VN01 On-Hand report has already been generated.

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

# Before and After

## Before Automation

The original process required manual handling of the VN01 On-Hand and Future Ship Mode reports before producing the final Purchasing output.

![Before Process](images/before-process.png)

---

## After Automation

The Python automation performs the matching, Purchasing Group filtering, and report generation automatically.

![After Process](images/after-process.png)

---

# Data Processing

## VN01 On-Hand

The VN01 On-Hand report provides:

```text
Material
PGr
```

The `Material` field is renamed to:

```text
Customer Part.
```

The `PGr` field remains from the VN01 On-Hand data and is carried into the merged dataset.

---

## Future Ship Mode

The Future Ship Mode report provides:

```text
Customer Part.
```

The Future Ship Mode report **does not provide `PGr`**.

Its `Customer Part.` field is used to match the VN01 On-Hand data.

---

## Merge

The two datasets are merged using:

```text
Customer Part.
```

The resulting data contains:

```text
Customer Part.
PGr
Future Ship Mode information
```

The `PGr` in the merged dataset originates from the VN01 On-Hand report.

---

## Filtering Rule

After the merge, the automation applies:

```text
PGr starts with "U"
```

Only matching records are retained for the final report.

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

The VN01 On-Hand filename is generated based on the Monday date of the current week.

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

The current implementation requires the user to enter the Future Ship Mode filename.

---

# Future Ship Mode Input Structure

The current Future Ship Mode report contains fields including:

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
```

`PGr` is **not required from the Future Ship Mode report**.

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

During testing, `_NEW` was used to prevent overwriting an existing output:

```text
JB VN01-Backlog Report Shipment_NEW.xlsx
```

The final output contains the records that:

1. Match between the VN01 On-Hand and Future Ship Mode reports.
2. Have a VN01 `PGr` beginning with `U`.

---

# Output Columns

The final report contains:

```text
Ship To
Order Number
Line Number
Customer Part.
Mfr
Mfr Part.
Qty Required
Qty Allocated
Date Ordered
Current Commit Date
Resale Price
MPQ
Invoice Number
Weight(grams)
Date Code
Ship mode
PGr
Buyer Name
```

---

# Error Handling

The automation performs validation before processing.

### Missing Input Files

The automation checks whether the expected source files exist.

### Missing Required Columns

Required columns are validated before processing.

VN01 On-Hand requires:

```text
Material
PGr
```

Future Ship Mode requires:

```text
Customer Part.
```

### No Matching Parts

If no matching `Customer Part.` records are found, the automation stops and reports the issue.

### No Matching Purchasing Groups

If matching parts are found but none have a `PGr` beginning with `U`, the automation stops and reports the issue.

### Invalid Output Structure

The automation validates that all required output columns exist before creating the final report.

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
        ├── architecture.png
        ├── before-process.png
        └── after-process.png
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
* VN01 On-Hand report.
* Future Ship Mode report.
* Output server location.

---

# Testing

The automation has been tested using the actual working server files.

| Test                            | Status                              |
| ------------------------------- | ----------------------------------- |
| VN01 On-Hand file detection     | ✅ Passed                            |
| Future Ship Mode file detection | ✅ Passed                            |
| Excel file loading              | ✅ Passed                            |
| Required column validation      | ✅ Passed                            |
| Customer Part. matching         | ✅ Passed                            |
| VN01 PGr retention              | 🔄 Re-testing after code correction |
| PGr filtering                   | 🔄 Re-testing after code correction |
| Consolidated report generation  | 🔄 Re-testing after code correction |
| Buyer list generation           | 🔄 Re-testing after code correction |
| Server output                   | 🔄 Re-testing after code correction |

> **Note:** The original version was tested successfully. The code was subsequently corrected so that `PGr` is sourced from the VN01 On-Hand report rather than being expected from the Future Ship Mode report. The corrected version should be re-tested before marking the final tests as passed.

---

# Business Impact

The automation improves the weekly Purchasing reporting process by:

* Reducing repetitive Excel manipulation.
* Standardizing the part-matching process.
* Applying the same Purchasing Group filtering rule consistently.
* Reducing the risk of manual filtering errors.
* Producing a consistent report format.
* Reducing the time required to prepare the report.
* Providing a Buyer reference list.

No percentage-based improvement is claimed because formal time measurements have not yet been recorded.

---

# Current Limitations

### Server Dependency

The automation depends on the existing network folder structure.

Changes to the server location may require code changes.

### Input Report Dependency

The source reports must maintain the expected column structure.

Changes to the report format may require updates to the automation.

### Manual Future Filename

The Future Ship Mode filename currently needs to be entered by the user.

### SAP Extraction

The Python automation does not currently extract the VN01 report directly from SAP.

The VN01 On-Hand report must already exist as an Excel file before the automation starts.

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

If required, VN01 report extraction could be integrated into the automation in a future version.

---

# Lessons Learned

### Validate the Actual Input Structure

The actual report structure should be treated as the source of truth when designing the automation.

In this process, `PGr` belongs to the VN01 On-Hand report and is not required from the Future Ship Mode report.

---

### Keep Business Logic Simple

The required business flow is straightforward:

```text
VN01 On-Hand
      +
Future Ship Mode
      ↓
Match Customer Part.
      ↓
Retain VN01 PGr
      ↓
Filter PGr starting with U
      ↓
Generate Report
```

Avoid introducing additional logic when the existing business process does not require it.

---

### Validate Before Processing

Checking the required files and columns before performing the merge prevents incomplete or incorrectly structured reports from being processed.

---

### Test Against Real Files

Testing against the actual server files was important because the real file paths and report structures needed to be handled correctly.

---

# My Role

I was responsible for developing and improving the automation for the VN01 Future Ship Mode reporting process.

My responsibilities included:

* Understanding the existing Purchasing workflow.
* Identifying the manual data-processing steps.
* Designing the Python automation.
* Developing the Excel data-processing logic.
* Implementing Customer Part matching.
* Implementing VN01 Purchasing Group filtering.
* Implementing input validation.
* Generating the consolidated report.
* Generating the Buyer reference list.
* Testing the automation using actual server files.
* Troubleshooting file-path and Excel dependency issues.
* Preparing the project documentation for GitHub.

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

This project demonstrates how Python can be used to remove repetitive Excel processing from an existing Purchasing workflow.

### Before

```text
VN01 On-Hand
      +
Future Ship Mode
      ↓
Manual Matching
      ↓
Manual Filtering
      ↓
Manual Report Preparation
```

### After

```text
VN01 On-Hand
      +
Future Ship Mode
      ↓
Python Automation
      ↓
Customer Part. Matching
      ↓
VN01 PGr Filtering
      ↓
Consolidated Report
      ↓
Buyer Reference
```

The result is a more consistent and repeatable weekly reporting process with less manual Excel manipulation.
