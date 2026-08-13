# VN01 Future Ship Mode & OH Merge Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![Python](https://img.shields.io/badge/Python-Automation-blue)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-blue)
![Excel](https://img.shields.io/badge/Excel-Automation-green)

> Python automation developed to consolidate VN01 On-Hand and Future Ship Mode reports, match Customer Parts, apply Purchasing Group rules, and generate a standardized Purchasing report.

---

## Overview

This project automates a weekly Purchasing reporting process that previously required manual Excel data processing.

The automation combines:

* VN01 On-Hand data
* Future Ship Mode data
* Customer Part matching
* Purchasing Group filtering
* Buyer reference generation

The automation standardizes the process and generates a consolidated Excel report.

**Project Status:** Production

---

# Business Problem

The original process required Purchasing users to manually:

* Open multiple Excel reports.
* Match VN01 Material with Future Ship Mode Customer Parts.
* Identify the applicable Purchasing Group.
* Filter records based on Purchasing Group.
* Prepare the final report.
* Identify the relevant Buyers.

This repetitive process increased processing time and created opportunities for manual matching and filtering errors.

---

# Solution

The solution uses Python and Pandas to automate the data-processing workflow.

The automation:

1. Loads the VN01 On-Hand report.
2. Loads the Future Ship Mode report.
3. Converts VN01 `Material` to `Customer Part.`.
4. Standardizes matching values.
5. Merges both datasets using `Customer Part.`.
6. Retains the required Future Ship Mode `PGr`.
7. Filters records where `PGr` starts with `U`.
8. Generates the consolidated Excel report.
9. Generates a Buyer reference list.

---

# Key Features

* Excel file validation
* Customer Part matching
* Data standardization
* Purchasing Group filtering
* Automated report generation
* Buyer reference generation
* Consistent output format
* Repeatable weekly workflow

---

# Technologies Used

| Category             | Technology      |
| -------------------- | --------------- |
| Programming Language | Python          |
| Data Processing      | Pandas          |
| Excel Processing     | OpenPyXL        |
| Input / Output       | Microsoft Excel |

---

# Workflow

![VN01 Future Ship Mode Workflow](images/workflow.png)

```text
VN01 On-Hand
      +
Future Ship Mode
      │
      ▼
Validate Input
      │
      ▼
Standardize Customer Part.
      │
      ▼
Merge Data
      │
      ▼
Resolve PGr
      │
      ▼
Filter PGr = "U*"
      │
      ▼
Generate Report
      │
      ▼
Generate Buyer List
```

---

# Data Processing

### Customer Part Matching

VN01 `Material` is converted to `Customer Part.` and standardized before matching with the Future Ship Mode data.

### Purchasing Group

The Future Ship Mode `PGr` is retained as the primary Purchasing Group for the final report.

### Filtering Rule

Only records where:

```text
PGr starts with "U"
```

are included in the final output.

---

# Architecture

![VN01 Future Ship Mode Architecture](images/architecture.png)

| Component        | Responsibility                        |
| ---------------- | ------------------------------------- |
| VN01 On-Hand     | Provides material information         |
| Future Ship Mode | Provides shipment and PGr information |
| Python           | Controls the automation workflow      |
| Pandas           | Processes and merges datasets         |
| OpenPyXL         | Handles Excel files                   |
| Output Report    | Provides the final Purchasing dataset |
| Buyer List       | Provides Buyer reference information  |

---

# Business Impact

The automation helps the Purchasing team by:

* Reducing repetitive Excel processing.
* Reducing manual data manipulation.
* Standardizing Customer Part matching.
* Applying Purchasing Group rules consistently.
* Producing a repeatable reporting output.
* Automatically identifying relevant Buyers.

No percentage-based efficiency improvement is claimed because formally measured performance data is not available.

---

# My Role

I was responsible for developing and maintaining this automation.

My responsibilities included:

* Understanding the Purchasing reporting process.
* Identifying manual activities suitable for automation.
* Developing the Python processing workflow.
* Implementing Customer Part matching.
* Implementing Purchasing Group filtering.
* Developing Excel output generation.
* Testing the automation against the reporting workflow.
* Troubleshooting and maintaining the solution.

---

# Current Limitations

* Depends on the expected Excel report structure.
* Required column names must remain consistent.
* Input files must be available in the expected location.
* Changes to source report formats may require code updates.
* Business rules are currently designed specifically for the VN01 workflow.

---

# Future Improvements

Potential improvements include:

* Automated weekly file detection.
* Configuration-based business rules.
* Enhanced data validation.
* Structured execution logging.
* User interface for file selection and execution.
* Scheduled report generation.

---

# Summary

The **VN01 Future Ship Mode & OH Merge Automation** converts a repetitive Purchasing reporting process into a standardized Python workflow.

The project demonstrates practical experience in:

* Python automation
* Pandas data processing
* Excel automation
* Data matching
* Business-rule implementation
* Reporting automation
* Process standardization

The result is a repeatable and maintainable workflow that reduces manual Excel processing and produces a consistent Purchasing report.
