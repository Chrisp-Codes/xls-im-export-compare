# Excel Import-Export Compare Tool (POC)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Status](https://img.shields.io/badge/status-POC-orange)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
[![English](https://img.shields.io/badge/README-English-informational?style=flat-square)](README_en.md)
[![Deutsch](https://img.shields.io/badge/README-Deutsch-informational?style=flat-square)](README.md)

## Overview
This repository contains a **proof of concept** Python tool to compare an import file against the corresponding export result in Excel.  
The goal is to verify whether an import was successful and to highlight mismatches automatically, reducing manual checks.

## Features
- GUI built with **Tkinter** for easy file selection  
- Compares two Excel files (import vs. export)  
- Matches employees based on **Personalnummer, Vorname, Nachname** (at least 2 out of 3 must match)  
- Highlights issues directly in the import file:
  - **Red**: no match or mismatched fields  
  - **Orange**: multiple potential matches found  
- Field-specific comparison logic:
  - Names (case-insensitive, ignores spaces and hyphens)  
  - Emails (case-insensitive)  
  - Dates (normalized to `DD.MM.YYYY`)  
  - Certain HR fields (e.g. Familienstand, Elterneigenschaft)  
- Saves results as a new Excel file with all mismatches clearly marked  

## Motivation
During one project, import success was uncertain because the system did not generate clean logs.  
Manual verification of import vs. export was error-prone and time-consuming.  
This tool automates the process and provides a clear, auditable result.

## Usage
1. Run the tool (`imexportchk.py` or via compiled `.exe`).  
2. Select:
   - Import file  
   - Export file  
   - Output file name  
3. Start the comparison.  
4. The tool will save a new Excel file where mismatches are highlighted directly in the import data.

## Example
- Input:  
  - Import file with employee master data  
  - Export file after system import  
- Output:  
  - Import file copy with highlighted differences  

## Status
- Proof of Concept (POC)  
- Not intended for production use  
- Tested with HR/Payroll-like Excel structures (up to ~500 rows)  

## Technologies
- Python  
- OpenPyXL  
- Tkinter  

## Disclaimer
This is a **POC project** and not officially released. Use at your own risk.

## License
This project is licensed under the MIT License.  
See the [LICENSE](LICENSE) file for details.
