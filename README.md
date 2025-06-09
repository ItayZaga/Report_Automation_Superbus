# Automated Report Generation System

> ⚠️ **Macros may be blocked.**
> For each macro file, right-click > **Properties** > **General** tab > check **Unblock** > click **Apply**.

A VBA-based Excel system for generating complex reports at the click of a button, designed for branch managers and senior management.
The system significantly reduces the time spent on manual report creation and enables non-technical users to easily access critical business data.

---

## System Goals

* Automatically generate a variety of operational reports, accessible even to non-technical users
* Streamline repetitive tasks and free time for important data analysis

---

## Main Files

### Main Controller File:

* `macro BI.xlsm`

### Report Macro Files:

* `macro performence.xlsm`
* `macro fines.xlsm`
* `macro lates analysis.xlsm`
* `macro last station.xlsm`
* `macro report stations.xlsm`

### Auxiliary File:

* `macro yoman long.xlsm` – optional. Used for data conversion when applying "Event Log Exclusion"

> These files are integral to the system. Do not add or remove any files.
> **Do not rename any of the files or worksheets. Do not hide any worksheets. All macro files must be kept in the same folder.**

---

## Input Files

Report files are empty by default. The input data files contain fictitious data (for information security) covering only a three-day period to ensure smooth performance. To generate reports, the following input files must be loaded via the interface in `macro BI.xlsm`:

* `VM.xlsx`
* `Work Schedule.xlsx`
* `Event Log.xlsx` (optional)

---

## System Requirements

* Microsoft Excel 2021 or newer
* Macro-enabled (`.xlsm`)
* Macro authorization required (see notice above)

---

## Getting Started

1. Open the file `macro BI.xlsm`
   - The main menu will open automatically. To reopen it later, reopen the worksheet named **"פתיחת תפריט"** ("Open Menu").
2. Load the required input files
3. Select the desired report from the interface
4. If event log exclusion is selected, load the Event Log file

---

### What is "Event Log Exclusion"?

Some reports allow the user to apply an "Event Log Exclusion".
This feature filters out trips that were marked as invalid  due to unforeseen events.

---

## About the Project

This tool was independently developed by the data analyst at Superbus, as an internal tool for report creation and improving work efficiency through automation.
