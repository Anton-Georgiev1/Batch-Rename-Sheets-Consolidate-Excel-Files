## Batch_Rename_Sheets_Consolidate_Excel_Files

**The Problem**
Managing data across dozens of Excel workbooks can feel like herding cats at times. “Sales_2023” vs. “2023 Sales”) and the trailing blank rows make it a nightmare for automated processing. You find yourself spending hours renaming tabs and copy-pasting data to get a clean, consolidated report.

**The Answer**
**Batch_Rename_Sheets_Consolidate_Excel_Files** is a professional desktop utility which turns messy file collections into structured, consolidated data. It’s the glue that holds disparate source files to your master reporting structure, doing the tedious “prep work” that usually eats up your morning.

## Important Features
**Smart Sheet Renaming:** The tool implements advanced fuzzy logic (`rapidfuzz`) to search and rename sheets using a master template. It knows that "Mktg_Data" probably means "Marketing Data" , and prevents you from correcting it yourself .
**Formatting-Aware Consolidation:** This tool does "smart copies" using Windows COM automation, unlike common data scripts that rip away all your formatting. It preserves the bold headers, the cell colors, and the border styles as it merges.
**Proactive Data Hygiene:** Automatically detects and removes trailing blank rows before processing, resulting in a lean consolidated file free of “phantom” data.
**Master-Centric Architecture:** Point the tool at a “Master” file and it will ensure that every processed workbook matches exactly the sheet names and column structures you want.
**Live Transparency:** A built-in live logging console gives you instant feedback on every rename and merge operation, so you never have to wonder what's happening with your data.

## Technology Base
The app is written in **Python** with a **Tkinter** GUI and designed to be reliable and responsive. It uses **Pandas** and **Openpyxl** for core data manipulation and **Win32COM** for deep integration with Microsoft Excel's native formatting engine. All heavy lifting is done in threads to keep the interface responsive and interactive.

## Why Use It?
This is not just a script but a workflow accelerator for data analysts and project managers. It turns a multi-hour manual process into a three-click automated process that makes sure that your final consolidated workbook is not only accurate, but presentation-ready.

---

### Requirements
*   Windows OS (for consolidation that is aware of formatting)
*   Microsoft Excel installed*
*   Python 3.12+

### Installing
1. Clone the repo.
2. Install required dependencies:
   ```bash
   pip install pandas openpyxl rapidfuzz pywin32
   ```

### How to use
1. Start the app:
   ```bash
   python Batch_Rename_Sheets_Consolidate_Excel_Files.py
   ```
2. **Choose Master Excel:** Choose the template file that contains the correct sheet names and headers.
3. **Choose Folder**: Select the folder that contains the files you wish to process.
4. **Rename Sheets:** If sheet names do not match the master template, rename them to synchronize.
5. **Consolidate:** Combine all documents together into one, formatted master workbook. Retain all visual aesthetics.

---
*Note: To keep the formatting, the “Consolidate Sheets” function currently needs a Windows environment with Excel installed to use COM automation.