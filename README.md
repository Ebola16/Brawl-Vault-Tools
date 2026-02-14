# Brawl Vault Tools

Tools for retrieving data from https://forums.kc-mm.com/

## Overview

This repository contains three scripts that work together to process Brawl Vault data:

1. **BVexport.py** - Exports entries from saved web pages to Excel
2. **BVexportToJSON.py** - Converts the Excel export to JSON format
3. **BVextractImages.py** - Downloads all images referenced in the data

## Requirements

### Python Version
- Python 3.6 or higher

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd <repository-name>
```

2. Install dependencies:
```bash
pip install pandas openpyxl beautifulsoup4 requests
```

## Usage

### Step 1: Export Data from Saved Web Pages

**Script:** `BVexport.py`

1. Save your Brawl Vault web pages to a folder
2. Place `BVexport.py` in the same folder as the saved web pages
3. Run the script

**Output:** `brawl_vault_full_export.xlsx` - An Excel file containing all exported entries

### Step 2: Convert to JSON (Optional)

**Script:** `BVexportToJSON.py`

1. Place `BVexportToJSON.py` in the same folder as `brawl_vault_full_export.xlsx`
2. Run the script

**Output:** `brawl_data.json` - A JSON file with the exported data

### Step 3: Download Images (Optional)

**Script:** `BVextractImages.py`

1. Place `BVextractImages.py` in the same folder as `brawl_vault_full_export.xlsx`
2. Run the script

**Output:** 
- `images/` - Folder containing all downloaded images
- `Missing Images.txt` - List of image URLs that failed to download

## File Structure

```
your-folder/
├── BVexport.py
├── BVexportToJSON.py
├── BVextractImages.py
├── [saved web pages]
├── brawl_vault_full_export.xlsx    (generated)
├── brawl_data.json                  (generated)
├── Missing Images.txt               (generated)
└── images/                          (generated)
    └── [downloaded images]
```

## Workflow

The typical workflow is:

1. Save Brawl Vault web pages to a folder
2. Run `BVexport.py` to create the Excel export
3. (Optional) Run `BVexportToJSON.py` if you need JSON format
4. (Optional) Run `BVextractImages.py` to download all images