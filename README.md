# Multithreading-excel-sheet-program
C# automation tool developed during CIB internship for MSISDN validation and Excel data processing.


Company: CIB (Commercial International Bank)

Role: Internship Project

Project Overview
Developed a robust C# console application designed to automate the validation of mobile subscriber numbers (MSISDN) from Excel spreadsheets. This tool was built to ensure high data quality and format compliance for banking operations.

Key Features
Excel Processing: Utilized C# libraries (like ClosedXML or EPPlus) to read and parse data rows from .xlsx files.

Format Validation: Implemented strict logic to verify MSISDN standards, ensuring numbers met length, country code, and network prefix requirements.

Granular Logging: Developed a custom logging system that tracks the status of every row:

Success: Logs the row index and the validated number.

Failure: Flags the specific row and provides a detailed error message (e.g., "Invalid Prefix at Row 45", "Input Not Numeric").

Batch Summary: Automatically generates a final report indicating whether the file is "Clean" or listing the specific rows that require manual correction.

Tech Stack
Language: C# (.NET)

Concepts: File I/O, Exception Handling, Object-Oriented Programming (OOP), String Manipulation (Regex).
