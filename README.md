# Excel Extractor Tool (`xlse-unpacker`)

A zero-external-dependency toolset in **Python** (`xlse_unpacker.py`) and **PowerShell** (`Unpack-Xlse.ps1`) to unpack, deconstruct, and extract human-readable code and structure from **XLSX**, **XLSM**, and **XLSB** workbooks.

Following the architecture of `pbix-unpacker`, it operates recursively across directories, extracts containers into sibling folders (`./WorkbookName/`), calculates SHA-256 delta hashes to avoid redundant writes, and extracts opaque binary streams into version-controllable, grep-searchable text files.

---

## Key Features

- **Sibling Directory Unpacking**: Automatically extracts `./Report.xlsx` into `./Report/`.
- **Pure Stdlib VBA Decompiler (Zero Dependencies)**:
  - Custom pure Python implementation of Microsoft Compound File Binary Format (CFBF/OLE) and [MS-OVBA] decompression algorithms.
  - Decompiles `vbaProject.bin` into clean `.bas` (standard modules), `.cls` (classes and document/sheet code), and `.frm` (userforms) files under `VBA_Extracted/`.
  - Exports project properties and library references to `Project_Info.json` and `References.txt`.
- **Power Query (M) Code Decompiler**:
  - Locates embedded `DataMashup` binary streams (MS-QDEFF) in `customXml` or root parts.
  - Deconstructs `Formulas/Section1.m` into individual query `.m` files under `PowerQuery_Extracted/Queries/`.
  - Generates `Queries_Summary.json` and pretty-prints XML configuration packages.
- **OpenXML Structure & Formula Deconstruction**:
  - Memory-efficient streaming parser extracts cell formulas across sheets into `Workbook_Structure/Formulas_Summary.json`.
  - Extracts sheet names, visibility states, and tab IDs into `Sheets.json`.
  - Extracts named ranges and scoped formulas into `DefinedNames.json`.
  - Extracts structured table definitions and columns into `Tables.json`.
  - Summarizes pivot table cache fields in `PivotTables_Summary.json`.
- **Excel Binary Workbook (.xlsb) Support**:
  - Pure stdlib reader for Biff12 variable-length record streams (`RECORD_BrtBundleSh`, etc.) to parse workbook and sheet definitions.
  - Automatically extracts embedded `vbaProject.bin` and `DataMashup` packages from XLSB containers.
- **Content-Aware Delta Synchronization**:
  - Checks file size and SHA-256 hashes to update only new or modified entries.
- **Granular Traversal & Filtering**:
  - Wildcard and substring filtering on file names (`--filter-in`, `--filter-out`).
  - Directory filtering during recursive scanning (`--dir-filter-in`, `--dir-filter-out`).
- **Dry-Run Simulation**:
  - Preview creations, updates, and unchanged counts without modifying disk.

---

## Output Layout

```
Directory/
├── FinancialModel.xlsm
└── FinancialModel/                           <-- Sibling extraction directory
    ├── [OpenXML Internal Files]              <-- Formatted XMLs / internal streams
    ├── VBA_Extracted/                        <-- Decompiled VBA code
    │   ├── Modules/
    │   │   ├── Module1.bas
    │   │   └── Calculations.bas
    │   ├── Classes/
    │   │   └── CustomModel.cls
    │   ├── Forms/
    │   │   └── UserForm1.frm
    │   ├── Document/
    │   │   ├── ThisWorkbook.cls
    │   │   └── Sheet1.cls
    │   ├── References.txt                   <-- Project references & GUIDs
    │   └── Project_Info.json                <-- Project metadata
    ├── PowerQuery_Extracted/                 <-- Power Query M code
    │   ├── Queries/
    │   │   ├── RawTransactions.m
    │   │   └── DimCustomers.m
    │   └── Queries_Summary.json
    └── Workbook_Structure/                  <-- Human-readable workbook metadata
        ├── Sheets.json                      <-- Sheet names, states, IDs
        ├── DefinedNames.json                <-- Named ranges and formulas
        ├── Formulas_Summary.json            <-- Grouped cell formulas per sheet
        ├── Tables.json                      <-- Excel ListObject table definitions
        └── PivotTables_Summary.json         <-- Pivot cache sources and fields
```

---

## 1. Python Tool (`xlse_unpacker.py`)

Requires Python 3.8+ (uses strictly the standard library, zero pip packages).

### Usage

```bash
# Basic unpack of a folder
python3 xlse_unpacker.py -p "/path/to/workbooks"

# Recursive traversal across all subfolders
python3 xlse_unpacker.py -p "/path/to/workbooks" -r

# Filter specific workbooks
python3 xlse_unpacker.py -p "/path/to/workbooks" -r --filter-in "*Budget*" "*2026*"

# Exclude temporary or draft files
python3 xlse_unpacker.py -p "/path/to/workbooks" -r --filter-out "*temp*" "~*"

# Dry-run preview
python3 xlse_unpacker.py -p "/path/to/workbooks" --dry-run
```

---

## 2. PowerShell Utility (`Unpack-Xlse.ps1`)

Compatible with Windows PowerShell 5.1 and PowerShell 7+ on Windows, macOS, and Linux.

### Usage

```powershell
# Basic unpack
.\Unpack-Xlse.ps1 -Path "C:\Financials"

# Recursive traversal
.\Unpack-Xlse.ps1 -Path "C:\Financials" -Recurse

# Granular filtering
.\Unpack-Xlse.ps1 -Path "C:\Financials" -Recurse -FileInclude "*Report*", "*Model*" -FileExclude "*Draft*"

# Dry-run mode
.\Unpack-Xlse.ps1 -Path "C:\Financials" -DryRun
```

---

## Testing

Run unit tests via Python's standard unittest runner:

```bash
python3 test_xlse_unpacker.py
```
