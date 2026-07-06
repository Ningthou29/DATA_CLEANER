# Meitei Dataset Cleaner

A Python utility for cleaning multilingual **Meitei language datasets** stored in Excel (`.xlsx`, `.xls`) or Markdown (`.md`) files.

The program automatically removes invalid records, converts Arabic numerals into Meitei numerals, and exports both the cleaned dataset and the rejected records into separate Excel files.

---

# Features

* Supports **Excel (.xlsx/.xls)** and **Markdown (.md)** input files.
* Interactive file selection through the terminal.
* Automatically creates an output folder.
* Removes rows where the **Meitei Script** column contains English letters.
* Removes rows with incomplete or meaningless English text.
* Converts digits (0–9) in the Meitei Script column into Meitei numerals.
* Keeps track of every deleted row.
* Adds a reason for each deleted record.
* Exports:

  * Cleaned dataset
  * Deleted/rejected dataset
* Displays a detailed cleaning summary after processing.

---

# Folder Structure

```
Project/
│
├── meitei_dataset_cleaner.py
├── cleaned_data/
│   ├── cleaned_meitei_YYYYMMDD_HHMMSS.xlsx
│   └── deleted_meitei_YYYYMMDD_HHMMSS.xlsx
└── README.md
```

---

# Requirements

* Python 3.9 or later

Required Python package:

```
pandas
openpyxl
```

Install dependencies:

```bash
pip install pandas openpyxl
```

---

# Supported Input Formats

### Excel

```
.xlsx
.xls
```

### Markdown

Markdown tables with the following columns:

| english | meitei_script | roman_standard |
| ------- | ------------- | -------------- |
| Example | Example       | Example        |

---

# Expected Columns

The input file should contain three columns in this order:

| Column         | Description                |
| -------------- | -------------------------- |
| english        | English sentence or phrase |
| meitei_script  | Meitei Mayek text          |
| roman_standard | Romanized Meitei text      |

If an Excel file has different column names, the program automatically renames the first three columns.

---

# Cleaning Rules

## 1. Remove English Letters from Meitei Script

Any row where the **Meitei Script** column contains English alphabet characters (A–Z or a–z) is removed.

Example:

```
ꯍꯥꯏ ABC
```

Result:

```
Deleted
```

---

## 2. Remove Incomplete English Entries

Rows are removed if the English column is:

* Empty
* Too short
* Meaningless
* Repeated characters
* Mostly punctuation
* Contains insufficient meaningful words

Example:

```
Hello
```

or

```
.....
```

Result:

```
Deleted
```

---

## 3. Convert Digits to Meitei Numerals

Only the **Meitei Script** column is modified.

Example:

```
123
```

becomes

```
꯱꯲꯳
```

Digit mapping:

| Arabic | Meitei |
| ------ | ------ |
| 0      | ꯰      |
| 1      | ꯱      |
| 2      | ꯲      |
| 3      | ꯳      |
| 4      | ꯴      |
| 5      | ꯵      |
| 6      | ꯶      |
| 7      | ꯷      |
| 8      | ꯸      |
| 9      | ꯹      |

---

# Output Files

After processing, two Excel files are generated inside the **cleaned_data** folder.

### Cleaned Dataset

```
cleaned_meitei_YYYYMMDD_HHMMSS.xlsx
```

Contains only valid records.

### Deleted Dataset

```
deleted_meitei_YYYYMMDD_HHMMSS.xlsx
```

Contains every removed row along with a **Reason for Deletion** column.

Example:

| english | meitei_script | roman_standard | Reason for Deletion              |
| ------- | ------------- | -------------- | -------------------------------- |
| Example | ABC           | example        | English letters in Meitei Script |

---

# Running the Program

Run:

```bash
python meitei_dataset_cleaner.py
```

The program will ask for the dataset path.

Example:

```
Please enter the input file path:

> D:\Datasets\meitei.xlsx
```

The program then:

1. Creates the output folder.
2. Loads the dataset.
3. Cleans the data.
4. Saves the cleaned dataset.
5. Saves deleted records.
6. Displays a processing summary.

---

# Example Summary

```
==================================================
CLEANING SUMMARY
==================================================

Initial rows:              12500
Rows removed (English):      340
Rows removed (Incomplete):   212
Total rows removed:          552
Final rows:                11948
Data retention:            95.6%
```

---

# Error Handling

The program checks for:

* Missing files
* Unsupported file formats
* Invalid input paths
* File loading errors
* Output folder creation errors
* Excel export errors

Clear error messages are displayed to help diagnose problems.

---

# Technologies Used

* Python
* Pandas
* OpenPyXL
* pathlib
* datetime
* regular expressions (re)

---

# License

This project is provided for educational and research purposes.

You are free to modify and extend the code to suit your own Meitei language dataset cleaning workflow.
