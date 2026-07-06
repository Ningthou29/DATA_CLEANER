# Quick Start Guide - Meitei Dataset Cleaner

## 5-Minute Setup & Usage Guide

### Step 1: Install Python Requirements

```bash
pip install pandas openpyxl
```

### Step 2: Run the Program

Open your terminal/command prompt and run:

```bash
python meitei_dataset_cleaner.py
```

### Step 3: Provide Your Excel File

When the program asks for the file path, you can:

**Option A - Type the path:**
```
> C:\Users\YourName\Downloads\dataset.xlsx
```

**Option B - Drag & Drop (Easiest!):**
- Drag your Excel file into the terminal window
- The path will auto-fill
- Press Enter

### Step 4: Wait for Processing

The program will:
- Load your file
- Remove rows with English letters in column B
- Remove incomplete rows in column A
- Convert numbers to Meitei script in column B
- Save the cleaned file

### Step 5: Find Your Output

The cleaned file is saved in:
```
same_folder_as_program/cleaned_data/cleaned_meitei_TIMESTAMP.xlsx
```

Example:
```
C:\Program Files\Meitei Cleaner\cleaned_data\cleaned_meitei_20250609_143022.xlsx
```

---

## Example Workflow

### Before Running:
```
C:\Program Files\Meitei Cleaner\
├── meitei_dataset_cleaner.py
└── README.md
```

### After First Run:
```
C:\Program Files\Meitei Cleaner\
├── meitei_dataset_cleaner.py
├── README.md
└── cleaned_data\              ← Created automatically
    └── cleaned_meitei_20250609_143022.xlsx
```

### After Multiple Runs:
```
C:\Program Files\Meitei Cleaner\
├── meitei_dataset_cleaner.py
├── README.md
└── cleaned_data\
    ├── cleaned_meitei_20250609_143022.xlsx
    ├── cleaned_meitei_20250609_150015.xlsx
    └── cleaned_meitei_20250609_152341.xlsx
```

---

## Command Line Tips

### On Windows:
1. Right-click in the folder where the program is located
2. Select "Open in Terminal" or "Open Command Window Here"
3. Type: `python meitei_dataset_cleaner.py`
4. Press Enter

### On Mac/Linux:
1. Open Terminal
2. Navigate to program folder: `cd /path/to/program`
3. Run: `python3 meitei_dataset_cleaner.py`
4. Press Enter

---

## What Gets Cleaned

### Rows Removed - Column B (Meitei Script)
- Any row where column B contains English letters (a-z, A-Z)
- Example removed: `ꯂꯥꯟꯅ hello ꯂꯥꯟꯅ` ❌

### Rows Removed - Column A (English)
- Empty cells
- Text shorter than 20 characters
- Single words or fragments
- Mostly special characters
- Incomplete sentences

Examples removed:
- `terday` ❌ (fragment)
- `police` ❌ (too short/single word)
- `The arrested have` ❌ (incomplete)

### Numbers Converted - Column B Only
| Original | Converted |
|----------|-----------|
| 0-9 | ꯰-꯹ |
| 23 | ꯲꯳ |
| 840 | ꯸꯴꯰ |
| 1.10 | ꯱.꯱꯰ |

Columns A and C keep original numbers!

---

## Output File Information

### File Name Format
```
cleaned_meitei_[YYYYMMDD]_[HHMMSS].xlsx
```

Example: `cleaned_meitei_20250609_143022.xlsx`
- Date: June 9, 2025
- Time: 14:30:22

### Columns in Output
- **Column A (english)**: Complete, meaningful English text
- **Column B (meitei_script)**: Pure Meitei script with numbers as Meitei numerals
- **Column C (roman_standard)**: Roman transliteration with original numbers

---

## Common Issues & Solutions

| Issue | Solution |
|-------|----------|
| "File not found" | Check the path is correct, use drag & drop |
| Can't open terminal | Right-click folder > "Open in Terminal" |
| Python not found | Install Python 3.6+ from python.org |
| Missing pandas | Run: `pip install pandas openpyxl` |
| Permission denied | Run terminal as Administrator (Windows) |

---

## Typical Session

```
$ python meitei_dataset_cleaner.py

============================================================
MEITEI DATASET CLEANER
============================================================

Supported formats: .md (Markdown), .xlsx or .xls (Excel)

Please enter the input file path:
(You can drag and drop the file here or type the path)

> /Users/yourname/Documents/dataset.xlsx

============================================================
Setting up output folder...
============================================================

✓ Output folder: /Users/yourname/Documents/cleaned_data

Loading file: dataset.xlsx
✓ Loaded 4177 rows from dataset.xlsx

============================================================
MEITEI DATASET CLEANER - PROCESSING
============================================================

Step 1: Removing rows with English letters in column B...
✓ Removed 155 rows with English letters in column B
  Remaining: 4022 rows

Step 2: Removing incomplete/meaningless rows in column A...
✓ Removed 265 rows with incomplete/meaningless information in column A
  Remaining: 3757 rows

Step 3: Converting numbers to Meitei script in column B...
✓ Converted numbers to Meitei script in column B

Saving cleaned data...
✓ Saved to /Users/yourname/Documents/cleaned_data/cleaned_meitei_20250609_143022.xlsx

============================================================
CLEANING SUMMARY
============================================================

Initial rows:              4177
Rows removed (English):    155
Rows removed (Incomplete): 265
Total rows removed:        420
Final rows:                3757
Data retention:            89.9%

============================================================
OUTPUT FOLDER LOCATION
============================================================

Output folder: /Users/yourname/Documents/cleaned_data
Output file: cleaned_meitei_20250609_143022.xlsx

Full path: /Users/yourname/Documents/cleaned_data/cleaned_meitei_20250609_143022.xlsx

============================================================
✓ CLEANING COMPLETED SUCCESSFULLY!
============================================================
```

---

## Next Steps

1. Open the cleaned Excel file from the `cleaned_data` folder
2. Review the data
3. Run the program again with another file if needed
4. All cleaned files are saved in the `cleaned_data` folder

---

## Support

For detailed information, see **README.md**

For technical details, see the program documentation in the code.

---

**Version**: 1.0  
**Last Updated**: June 2025
