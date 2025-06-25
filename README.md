# Salary Document Analyzer

## Overview

This project is a Python-based tool designed to analyze **SD Worx** salary sheets in PDF format.  
It automatically processes all documents stored in a local folder, detects and combines entries that relate to the **same work month** (even if paid later), and generates a summary.

Optionally, it can also incorporate data from an Excel file containing **per diems** and **work mission expenses**.  
All results are displayed via a local web interface with interactive graphs and tables.

## Features

- Parses and analyzes SD Worx salary PDFs.
- Merges multiple entries corresponding to the same work month.
- Optional integration of mission-related expenses from Excel.
- Visualizes salary breakdowns and expenses in a web browser.


## Folder Structure

To use the project, organize your files as follows:

```
project_folder_name/
├── salary_analysis.py  # Main Python script (not needed if you use the .exe)
├── salary_analysis.exe # Executable
├── sdworks_JohnDoe/    # Folder containing SD Worx PDF salary sheets
│ ├── January2024.pdf
│ ├── February2024_part2.pdf
│ ├── any_name.pdf
│ └── ...
├── perdiems.xlsx # (Optional) Excel file with per diem and mission expenses
```

- The folder with salary PDFs **must begin with `sdworks_`** followed by the name of the person. If you have multiple folders starting with `sdworks_`, the program will let you choose the one to analyse. 
- The name of the salary PDFs has no importance. 
- All PDF files inside that folder will be scanned but only the relevant ones will be used. The rest will not block the execution of the code. 
- The optional `perdiems.xlsx` file (if provided) will be included in the analysis.

## How to Use

1. Place your SD Worx PDF salary sheets in a folder named `sdworks_<yourname>`.
2. (Optional) Add your per diem Excel file as `perdiems.xlsx` in the main directory.
3. Run the script:
 - The Windows executable: **`salary_analysis.exe`**
 - Or the Python source code: **`salary_analysis.py`** (you need python installed and the libraries)
4. A local web page will open automatically with interactive graphs and summaries. (http://localhost:127.0.0.1:56146/)

## Feedback

If you encounter incorrect results or have suggestions for improvement, feel free to contact me.