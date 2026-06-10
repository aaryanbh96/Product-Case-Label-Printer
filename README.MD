**The Tao of Tea:2×2 Label Maker Software**
===========================================

**ProcessDocumentation (For Reference)**
========================================

_A Step-by-Step Guide fromEnvironment Setup to Software Distribution_

**1\. Introduction**

This software automates theprinting of 2×2 inch product labels for The Tao of Tea products. The systemreads product data from a centralized Excel file on a network drive (preferred)or on a local device and generates formatted labels with barcodes, TAO IDs,SKUs, and product information. Labels are printed 20 per sheet (4 columns × 5rows) on standard 8.5×11 inch label sheets (OL330).

**Key Features:**

*   Network-compatible: Excel data stored on shared drive accessible to all users
    

*   Barcode generation using UPC-A format
    

*   User-friendly search interface
    

*   Remembers folder locations between sessions
    

**2\. Environment Setup**

**Current Installation**

**Primary System:**Lenovo 3**Location:** C:\\LABELPRINTER2x2\\NewProgram\\**Virtual Environment:** C:\\Users\\veeri\\quickbooks\_env\\

**Dependencies Required (For Any System)**

Python 3.8+

pandas - Excel file reading

reportlab - PDF generation

python-barcode - Barcode generation

Pillow (PIL) - Image processing

pywin32 - Windows printer integration

openpyxl - Excel file handling

tkinter - GUI (included with Python)

**Installation Commands**

pip install pandas reportlabpython-barcode\[images\] pillow pywin32 openpyxl

**Excel File Requirements:**

*   Must contain columns: Tea, Tao ID, Tin SKU, Barcode\_UPC
    

*   Barcode\_UPC must be valid 12-digit UPC codes
    

*   Tin SKU formatted as text to preserve leading zeros
    

**3\. Code Development**

**Source Code**

**Filename:** label\_printer.py**Language:** Python 3**Storage:**

*   Physical: Printed and filed in office documentation
    

*   Digital: GitHub repository (private)
    

*   Access: Available with Aryan's permission
    

**Key Components of label\_printer.py**

2.  **Config Management:** Saves/loads data folder location using JSON
    

4.  **Folder Discovery:** Prompts user on first run, remembers location
    

6.  **Label Generation:** Creates PDFs with precise 2×2 inch label positioning
    

8.  **Barcode Integration:** Generates UPC-A barcodes dynamically
    

10.  **Print Integration:** Supports SumatraPDF, Adobe Reader, Windows default
    

**Label Specifications**

*   **Size:** 2×2 inches per label
    

*   **Layout:** 20 labels per sheet (4×5 grid)
    

*   **Margins:** Top: 0.321", Left: 0.161"
    

*   **Content:** TAO ID (right-aligned), Tea name, SKU, QTY field, UPC barcode
    

**4\. Compilation Process**

**Prerequisites**

2.  Navigate to project directory: cd C:\\LABELPRINTER2x2\\NewProgram
    

4.  Activate virtual environment: C:\\Users\\veeri\\quickbooks\_env\\Scripts\\activate
    

**Compilation Command**

pyinstaller --onefile --windowed ^

 --hidden-import=win32print ^

 --hidden-import=win32api ^

 --hidden-import=win32timezone ^

 --hidden-import=PIL.\_tkinter\_finder ^

 --collect-all barcode ^

 --collect-all reportlab ^

 --collect-all pandas ^

  --name"LabelPrinter" ^

 label\_printer.py

**Output Location**

C:\\LABELPRINTER2x2\\NewProgram\\dist\\LabelPrinter.exe

**Compilation Notes**

*   \--onefile: Creates single executable (no DLL dependencies)
    

*   \--windowed: No console window appears (GUI only)
    

*   \--hidden-import: Includes modules PyInstaller might miss
    

*   \--collect-all: Bundles entire package with data files
    

*   Compilation time: ~2-3 minutes depending on system
    

**Quick Compile Script (Optional)**

Create compile.bat in project folder forone-click compilation:

@echo off

callC:\\Users\\veeri\\quickbooks\_env\\Scripts\\activate

if exist build rmdir /s /q build

if exist dist rmdir /s /q dist

pyinstaller --onefile --windowed--hidden-import=win32print ^

 --hidden-import=win32api --collect-all barcode --collect-all reportlab ^

 --collect-all pandas --name "LabelPrinter" label\_printer.py

pause

**5\. Distribution to Computers**

**Deployment Package**

**Single File Required:** LabelPrinter.exe(from dist folder)

**Installation Steps (Per Computer)**

2.  **Copy Executable**
    

2.  **First Run Setup**
    

5.  **Printer Setup (One-Time)**
    

**Network Requirements**

*   All computers must have network access to shared Excel folder (if excel hosted on network)
    

*   Read/write permissions needed for Labels and Barcodes subfolders
    

*   Excel file must remain at: \[Network Path\]\\UPC\_With\_Barcodes\_FOR\_PRINT.xlsx
    

**Troubleshooting Common Issues**

**Issue**

**Solution**

"No application associated" error

Install SumatraPDF or Adobe Reader

Can't find Excel file

Use "Change Data Folder" button to reselect location

Printer not found

Verify printer installed in Windows devices

PDF created but won't print

Manually print PDF from Labels folder

**6\. User Operation**

**Workflow**

2.  Launch LabelPrinter.exe
    

4.  Type tea name in search box
    

6.  Select product from list
    

8.  Enter quantity (number of labels, not sheets)
    

10.  Select printer from dropdown
    

12.  Click green "PRINT" button
    

**Important Notes**

*   **Quantity = Labels:** Enter total labels needed (e.g., "40" = 2 sheets)
    

*   **Each sheet contains 20 labels** (4 across × 5 down)
    

*   PDFs saved for records
    

*   Search is real-time: results appear as you type
    

**7\. Maintenance & Updates**

**Updating Product Data**

*   Edit UPC\_With\_Barcodes\_FOR\_PRINT.xlsx on network drive(preferred) or on local system.
    

*   Changes immediately available to all users if excel file stays on a network drive
    

*   Software restart not needed but recommended.
    

**Software Updates**

1.  Modify label\_printer.py on development machine
    

2.  Recompile using compilation command
    

3.  Distribute new LabelPrinter.exe to users
    

4.  User data/settings preserved (stored separately in config file)
    

**Changing Network Location**

*   Click "Change Data Folder" button in application
    

*   Browse to new location
    

*   Restart program
    

**8\. Technical Specifications**

**Software Version:** 1.0**Build Date:** October 28 2025**Platform:** Windows 7/10/11 (64-bit)**Executable Size:** ~85-120 MB (includes all dependencies)**Config Storage:** User's home directory (%USERPROFILE%)**Label Printer Compatibility:** Any Windows-compatible printer (tested withBrother label printers)

## Screenshots

![Screenshot](Screenshots/Screenshot%202026-03-25%20184617.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20184731.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20184752.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20184835.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185218.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185310.png)

**Contact & Support**

**Developed by:** Aryan Bhardwaj for The Taoof Tea**Code Repository:** GitHub (private, access by permission)**Documentation:** This file + printed code in office files

_Document Version 1.0 | Last Updated: October 28,2025_



