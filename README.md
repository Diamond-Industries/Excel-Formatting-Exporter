# Excel Formatting Exporter (VBA to Web)

## Overview
This repository contains a comprehensive VBA macro designed to "reverse engineer" Excel templates. It loops through every sheet in a workbook and extracts **60 distinct data points** per cell, including exact styles, borders, alignments, and validation rules. 

It compiles this data into a single master CSV report. This output is specifically formatted to be consumed by web developers using TypeScript/JavaScript libraries (like `exceljs` or `SheetJS`) to dynamically reconstruct identical Excel files from a web frontend.

## Key Features
* **Massive Extraction:** Captures 60 properties including Fonts, Backgrounds, Margins, Orientations, and Merged Areas.
* **Web-Ready Colors:** Automatically converts Excel's decimal color codes and Theme Colors into web-ready **HEX codes** (e.g., `#C0A0C0`).
* **Human-Readable Alignments:** Translates internal VBA enumerations (like `-4108`) into CSS-friendly terminology (like `Center`).
* **Dropdown Extraction:** Pulls the exact string arrays out of Data Validation dropdowns.
* **Safe Execution:** Bypasses standard VBA crashes (Error 5, Error 91) caused by blank comments or missing theme colors.

## How to Run in Excel
1. Open your target Excel workbook.
2. Press `Alt + F11` to open the VBA Editor.
3. Click `Insert > Module`.
4. Paste the entire script (including the helper functions at the bottom) into the text window.
5. Press `F5` to run.
6. A new sheet named `Web_Audit_[Timestamp]` will be generated at the front of your workbook containing the data.

## Parsing the Output in TypeScript
To keep the CSV compact, this script combines the original VBA enumerations with web-ready translations, separated by a pipe (` | `). 

When parsing the CSV in your frontend code, simply split the string to grab the value you need:

```typescript
// Example: Parsing a row from the CSV
const rawColor = row['FillColor'];             // "12611584 | #C0A0C0"
const hexColor = rawColor.split(' | ')[1];     // "#C0A0C0"

const rawHAlign = row['H-Align'];              // "-4108 | Center"
const alignString = rawHAlign.split(' | ')[1]; // "Center"
