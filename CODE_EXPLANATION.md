# Meitei Dataset Cleaner - Code Explanation (Line by Line)

## Overview

The program is an **Object-Oriented** Python application with:
- A main `MeiteiDatasetCleaner` class that handles all cleaning operations
- Helper functions for user interaction
- A main entry point that orchestrates everything

---

## PART 1: IMPORTS & SETUP

```python
#!/usr/bin/env python3
```
- **`#!`** = Shebang line (tells system to run with Python 3)
- Allows you to run the file directly: `./meitei_dataset_cleaner.py`

```python
import pandas as pd
import re
import sys
import os
from pathlib import Path
from datetime import datetime
from typing import Tuple, List
```

**What each import does:**
- **`pandas`** - Library for working with Excel and data tables (creates DataFrames)
- **`re`** - Regular expressions for pattern matching (finds English letters, etc.)
- **`sys`** - System operations (exit program, command-line arguments)
- **`os`** - Operating system functions (file/folder operations)
- **`Path`** - Modern way to work with file paths (from `pathlib`)
- **`datetime`** - For creating timestamps in filenames
- **`Tuple, List`** - Type hints for function documentation

---

## PART 2: THE MeiteiDatasetCleaner CLASS

### 2.1 Class Definition & Meitei Numerals

```python
class MeiteiDatasetCleaner:
    """Main class for cleaning Meitei language datasets"""
```
- Creates a class (blueprint) for the cleaning tool
- Docstring explains what the class does

```python
    MEITEI_NUMERALS = {
        '0': '꯰',
        '1': '꯱',
        '2': '꯲',
        '3': '꯳',
        '4': '꯴',
        '5': '꯵',
        '6': '꯶',
        '7': '꯷',
        '8': '꯸',
        '9': '꯹'
    }
```
- **Dictionary** that maps regular digits (0-9) to Meitei numerals
- **Capital letters** = class variable (shared by all instances)
- **Example**: When we want to convert "23" to "꯲꯳", we look up '2' and '3' in this dictionary
- **Why a dictionary?** Fast lookup - Python finds the value instantly

---

### 2.2 The __init__ Method (Initialization)

```python
    def __init__(self, verbose=True):
        """
        Initialize the cleaner.
        
        Args:
            verbose (bool): Print detailed information during processing
        """
```
- **`__init__`** = Constructor method (runs when you create a new instance)
- **`self`** = Refers to the current object being created
- **`verbose=True`** = Parameter with default value. If True, print progress messages
- **Docstring** explains what the method does

```python
        self.verbose = verbose
        self.df = None
        self.stats = {
            'initial_rows': 0,
            'removed_english_in_b': 0,
            'removed_incomplete_a': 0,
            'final_rows': 0
        }
```
- **`self.verbose`** = Store the verbose setting as an instance variable
- **`self.df = None`** = DataFrame starts empty (will be filled when we load a file)
- **`self.stats`** = Dictionary to track cleaning statistics:
  - `initial_rows` - How many rows we started with
  - `removed_english_in_b` - Rows removed due to English in Column B
  - `removed_incomplete_a` - Rows removed due to incomplete data in Column A
  - `final_rows` - How many rows we end up with

```python
        self.program_dir = Path(__file__).parent
        self.output_dir = self.program_dir / 'cleaned_data'
```
- **`Path(__file__)`** = Get the path to this program file
- **`.parent`** = Get the folder containing this program
- **`self.program_dir`** = Store the program's folder location
- **`self.output_dir`** = Create path to 'cleaned_data' folder (using `/` operator for paths)
  - Example: If program is at `/home/user/meitei_cleaner.py`
  - Then output_dir will be `/home/user/cleaned_data`

---

### 2.3 create_output_folder Method

```python
    def create_output_folder(self) -> bool:
        """
        Create output folder in the program directory.
        
        Returns:
            bool: True if successful or already exists, False otherwise
        """
```
- **`-> bool`** = Type hint saying this method returns a boolean (True or False)

```python
        try:
            self.output_dir.mkdir(exist_ok=True)
```
- **`try:`** = Start error handling block (if something goes wrong, we'll catch it)
- **`.mkdir()`** = Create the directory/folder
- **`exist_ok=True`** = Don't fail if folder already exists (if it's already there, that's fine)

```python
            if self.verbose:
                print(f"✓ Output folder: {self.output_dir}")
            return True
        except Exception as e:
            print(f"❌ Error creating output folder: {e}")
            return False
```
- **`if self.verbose:`** = Only print if verbose mode is ON
- **`f"...{variable}..."`** = F-string (formatted string) - inserts variable values
- **`return True`** = Success - folder was created
- **`except Exception as e:`** = If ANY error occurs, catch it
- **`e`** = The error object
- **`return False`** = Failure - couldn't create folder

---

### 2.4 load_file Method

```python
    def load_file(self, filepath: str) -> bool:
        """
        Load data from markdown or Excel file.
        
        Args:
            filepath (str): Path to input file
            
        Returns:
            bool: True if successful, False otherwise
        """
```
- **`filepath: str`** = Parameter must be a string
- **`-> bool`** = Returns a boolean

```python
        filepath = Path(filepath)
        
        if not filepath.exists():
            print(f"❌ Error: File not found: {filepath}")
            return False
```
- **`Path(filepath)`** = Convert string to Path object (better for file operations)
- **`if not filepath.exists():`** = If file does NOT exist
- **`not`** = Logical NOT (inverts True/False)
- Print error message and return False

```python
        try:
            if filepath.suffix.lower() == '.md':
                self._load_markdown(filepath)
            elif filepath.suffix.lower() in ['.xlsx', '.xls']:
                self._load_excel(filepath)
            else:
                print(f"❌ Error: Unsupported file format: {filepath.suffix}")
                return False
```
- **`filepath.suffix`** = File extension (e.g., '.xlsx')
- **`.lower()`** = Convert to lowercase (so '.XLSX' becomes '.xlsx')
- **`== '.md'`** = Check if it's a markdown file
- **`elif ... in [...]`** = Check if it's Excel format (could be .xlsx OR .xls)
- **`in`** = Check if value is in the list
- **`else:`** = If it's neither markdown nor Excel

```python
            self.stats['initial_rows'] = len(self.df)
            if self.verbose:
                print(f"✓ Loaded {len(self.df)} rows from {filepath.name}")
            return True
```
- **`len(self.df)`** = Count rows in the DataFrame
- **`filepath.name`** = Just the filename (without the path)
- Return True for success

```python
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            return False
```
- Catch any errors during loading

---

### 2.5 _load_markdown Method

```python
    def _load_markdown(self, filepath: Path):
        """Load data from markdown table"""
```
- **`_` prefix** = Convention for "private" methods (used internally only)

```python
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
```
- **`with open(...) as f:`** = Open file and automatically close it when done
- **`'r'`** = Read mode (not write)
- **`encoding='utf-8'`** = Use UTF-8 encoding (supports Meitei script)
- **`f.read()`** = Read entire file as a string
- **`content`** = The entire markdown file as a string

```python
        lines = content.strip().split('\n')
```
- **`.strip()`** = Remove whitespace from beginning/end
- **`.split('\n')`** = Split by newline character, creating a list of lines
- **`lines`** = List of all lines in the file

```python
        data = []
        
        for line in lines:
            if line.startswith('##') or line.startswith('| ---'):
                continue
```
- **`data = []`** = Empty list to store rows
- **`for line in lines:`** = Loop through each line
- **`.startswith(...)`** = Check if line starts with something
- **`or`** = Logical OR (skip if line starts with either ## or | ---)
- **`continue`** = Skip to next iteration, don't process this line

```python
            if line.startswith('|') and '|' in line:
                parts = [p.strip() for p in line.split('|')]
                parts = [p for p in parts if p]
```
- **`and`** = Both conditions must be true
- **`.split('|')`** = Split the line by pipe character (|)
- **`[p.strip() for p in ...]`** = List comprehension (creates new list)
  - For each part `p`, strip whitespace: `p.strip()`
- **`[p for p in parts if p]`** = Keep only non-empty parts
  - `if p` = if p is truthy (non-empty string is truthy)

```python
                if len(parts) == 3:
                    data.append(parts)
```
- **`len(parts) == 3`** = Check if we have exactly 3 columns
- **`.append()`** = Add this row to the data list

```python
        if data and data[0][0].lower() == 'english':
            data = data[1:]
```
- **`if data and ...`** = If data list is not empty AND...
- **`data[0][0]`** = First row, first column
- Remove the header row (it contains column names)
- **`data[1:]`** = All rows except the first

```python
        self.df = pd.DataFrame(data, columns=['english', 'meitei_script', 'roman_standard'])
```
- **`pd.DataFrame()`** = Create a DataFrame from the data
- **`columns=...`** = Name the columns
- **`self.df`** = Store the DataFrame in the instance

---

### 2.6 _load_excel Method

```python
    def _load_excel(self, filepath: Path):
        """Load data from Excel file"""
        self.df = pd.read_excel(filepath)
```
- **`pd.read_excel()`** = Pandas built-in function to read Excel files
- Automatically creates a DataFrame

```python
        if len(self.df.columns) >= 3:
            self.df.columns = ['english', 'meitei_script', 'roman_standard']
```
- **`len(self.df.columns)`** = Number of columns in the DataFrame
- **`>= 3`** = If we have at least 3 columns
- Rename the columns (pandas auto-names them 0, 1, 2, etc.)

---

### 2.7 _contains_english_letters Method

```python
    def _contains_english_letters(self, text: str) -> bool:
        """Check if text contains English alphabet letters (a-z, A-Z)"""
        if pd.isna(text):
            return False
```
- **`pd.isna(text)`** = Check if value is NaN (missing/null)
- If it's NaN (missing), return False (missing value doesn't contain English)

```python
        return bool(re.search(r'[a-zA-Z]', str(text)))
```
- **`str(text)`** = Convert to string (just in case it's not)
- **`re.search()`** = Search for a pattern in the text
- **`r'[a-zA-Z]'`** = Regular expression pattern
  - `[...]` = Character class (match any character inside)
  - `a-z` = Any lowercase letter
  - `A-Z` = Any uppercase letter
- **`.search()`** = Returns the match object if found, None if not
- **`bool(...)`** = Convert to True/False
  - Match object = True
  - None = False

---

### 2.8 _is_meaningful_text Method

```python
    def _is_meaningful_text(self, text: str) -> bool:
        """
        Check if text is meaningful and complete.
        
        Returns True if text should be KEPT, False if it should be REMOVED.
        """
```

```python
        if pd.isna(text) or text is None:
            return False
        
        text = str(text).strip()
```
- **`or`** = Either condition removes the row
- **`.strip()`** = Remove spaces from beginning/end

```python
        if len(text) == 0:
            return False
```
- If empty after stripping, remove it

```python
        if len(text) < 20:
            return False
```
- If less than 20 characters, it's too short (incomplete)

```python
        if len(set(text)) == 1:
            return False
```
- **`set(text)`** = Get unique characters
- **`len(set(...)) == 1`** = Only 1 unique character (like "aaaa" or "----")
- These are meaningless repeated characters

```python
        word_pattern = re.compile(r'[a-zA-Z]{2,}')
        words = word_pattern.findall(text)
        if len(words) < 2:
            return False
```
- **`re.compile()`** = Pre-compile a regex pattern (more efficient)
- **`r'[a-zA-Z]{2,}'`** = Find sequences of 2+ letters (words)
  - `{2,}` = 2 or more occurrences
- **`.findall()`** = Find all matches, return as list
- **`len(words) < 2`** = Need at least 2 meaningful words

```python
        if re.match(r'^[0-9\s\.,;:\-\—\—\/\(\)\[\]\{\}\.]+$', text):
            return False
```
- **`re.match()`** = Check if pattern matches from the beginning
- **`^`** = Start of string
- **`$`** = End of string
- **`[...]+`** = One or more of these characters
- **Inside `[]`**: 0-9 (digits), \s (whitespace), \. (periods), etc.
- **`\/`** = Forward slash (escaped because / is special)
- **`\(`** = Parenthesis (escaped)
- If text is ONLY numbers/punctuation, remove it

```python
        special_char_ratio = sum(1 for c in text if not c.isalnum() and c != ' ') / len(text)
        if special_char_ratio > 0.5:
            return False
```
- **`sum(1 for c in text if ...)`** = Count characters matching condition
- **`c.isalnum()`** = Is character alphanumeric? (letter or digit)
- **`not c.isalnum() and c != ' '`** = Special characters (not letters, digits, or spaces)
- **`/ len(text)`** = Divide by total length to get ratio
- **`> 0.5`** = If more than 50% are special characters, it's garbage

```python
        if re.search(r'[—\-_\…]{2,}$', text):
            return False
```
- **`{2,}$`** = 2+ dashes/underscores at the END of the string
- Indicates incomplete text that got cut off

```python
        return True
```
- If it passed all checks, keep it!

---

### 2.9 _convert_numbers_to_meitei Method

```python
    def _convert_numbers_to_meitei(self, text: str) -> str:
        """Convert all digits to Meitei numerals"""
        for digit, meitei_digit in self.MEITEI_NUMERALS.items():
            text = str(text).replace(digit, meitei_digit)
        return text
```
- **`for digit, meitei_digit in dict.items():`** = Loop through dictionary key-value pairs
  - Example: digit='0', meitei_digit='꯰'
  - Example: digit='1', meitei_digit='꯱'
- **`.replace(old, new)`** = Replace all occurrences of `old` with `new`
- **`text = ...`** = Update text with the replacement
- **`return text`** = Return the converted text

**Example:**
- Input: "23 apples"
- After digit '2': "꯲3 apples"
- After digit '3': "꯲꯳ apples"
- Return: "꯲꯳ apples"

---

### 2.10 remove_english_in_column_b Method

```python
    def remove_english_in_column_b(self) -> int:
        """
        Remove rows where column B (meitei_script) contains English letters.
        
        Returns:
            int: Number of rows removed
        """
        if self.df is None:
            print("❌ Error: No data loaded")
            return 0
```
- Check if DataFrame is empty

```python
        initial_count = len(self.df)
        self.df = self.df[~self.df['meitei_script'].apply(self._contains_english_letters)]
        removed = initial_count - len(self.df)
```
- **`initial_count`** = Count rows before filtering
- **`self.df['meitei_script']`** = Select just column B
- **`.apply(function)`** = Apply function to each value in the column
- **`~`** = NOT (invert True/False)
  - `~True` = False (remove this row)
  - `~False` = True (keep this row)
- **`self.df[boolean_mask]`** = Keep only rows where mask is True
- **`removed = ...`** = Calculate how many were removed

```python
        self.stats['removed_english_in_b'] = removed
        
        if self.verbose:
            print(f"✓ Removed {removed} rows with English letters in column B")
            print(f"  Remaining: {len(self.df)} rows")
        
        return removed
```
- Update statistics
- Print if verbose mode is on
- Return the count

---

### 2.11 remove_incomplete_in_column_a Method

```python
    def remove_incomplete_in_column_a(self) -> int:
        """
        Remove rows with incomplete or meaningless information in column A (english).
        
        Returns:
            int: Number of rows removed
        """
        if self.df is None:
            print("❌ Error: No data loaded")
            return 0
        
        initial_count = len(self.df)
        self.df = self.df[self.df['english'].apply(self._is_meaningful_text)]
        removed = initial_count - len(self.df)
```
- Same pattern as previous method
- Uses `_is_meaningful_text()` instead of `_contains_english_letters()`
- No `~` (NOT) because `_is_meaningful_text()` returns True for KEEP, False for REMOVE

```python
        self.stats['removed_incomplete_a'] = removed
        
        if self.verbose:
            print(f"✓ Removed {removed} rows with incomplete/meaningless information in column A")
            print(f"  Remaining: {len(self.df)} rows")
        
        return removed
```

---

### 2.12 convert_column_b_to_meitei Method

```python
    def convert_column_b_to_meitei(self) -> None:
        """Convert numbers to Meitei script ONLY in column B"""
        if self.df is None:
            print("❌ Error: No data loaded")
            return
```
- **`-> None`** = This method doesn't return anything

```python
        self.df['meitei_script'] = self.df['meitei_script'].apply(self._convert_numbers_to_meitei)
```
- Select column B
- Apply the conversion function to each value
- Replace column with converted values

```python
        if self.verbose:
            print(f"✓ Converted numbers to Meitei script in column B")
```

---

### 2.13 clean Method

```python
    def clean(self) -> None:
        """Execute all cleaning steps in sequence"""
        if self.df is None:
            print("❌ Error: No data loaded")
            return
        
        if self.verbose:
            print("\n" + "="*60)
            print("MEITEI DATASET CLEANER - PROCESSING")
            print("="*60 + "\n")
```
- **`"="*60`** = Repeat "=" 60 times to make a line
- Print header

```python
        print("Step 1: Removing rows with English letters in column B...")
        self.remove_english_in_column_b()
        
        print("\nStep 2: Removing incomplete/meaningless rows in column A...")
        self.remove_incomplete_in_column_a()
        
        print("\nStep 3: Converting numbers to Meitei script in column B...")
        self.convert_column_b_to_meitei()
        
        self.stats['final_rows'] = len(self.df)
```
- Execute all three cleaning steps in order
- Update final row count

---

### 2.14 save_excel Method

```python
    def save_excel(self, filename: str = None) -> str:
        """
        Save cleaned data to Excel file in output folder.
        
        Args:
            filename (str): Optional filename (default: auto-generated with timestamp)
            
        Returns:
            str: Full path to saved file, or empty string if failed
        """
        if self.df is None:
            print("❌ Error: No data to save")
            return ""
```
- **`filename: str = None`** = Optional parameter (can be None)
- **`-> str`** = Returns a string (the file path)

```python
        try:
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"cleaned_meitei_{timestamp}.xlsx"
```
- **`datetime.now()`** = Get current date and time
- **`.strftime(...)`** = Format as string
  - `%Y` = 4-digit year
  - `%m` = Month (01-12)
  - `%d` = Day (01-31)
  - `%H` = Hour (00-23)
  - `%M` = Minute (00-59)
  - `%S` = Second (00-59)
- Example: "20250609_143022"
- **`filename = f"...{timestamp}..."`** = Create filename with timestamp

```python
            output_path = self.output_dir / filename
            self.df.to_excel(output_path, index=False, sheet_name='Sheet1')
```
- **`self.output_dir / filename`** = Combine folder and filename (Path operation)
- **`.to_excel()`** = Pandas method to write DataFrame to Excel
- **`index=False`** = Don't write row numbers to the file
- **`sheet_name='Sheet1'`** = Name of the sheet

```python
            if self.verbose:
                print(f"✓ Saved to {output_path}")
            
            return str(output_path)
```
- Return the full path as a string

```python
        except Exception as e:
            print(f"❌ Error saving file: {e}")
            return ""
```
- If anything goes wrong, return empty string

---

### 2.15 print_summary Method

```python
    def print_summary(self) -> None:
        """Print cleaning summary statistics"""
        print("\n" + "="*60)
        print("CLEANING SUMMARY")
        print("="*60)
```

```python
        print(f"\nInitial rows:              {self.stats['initial_rows']}")
        print(f"Rows removed (English):    {self.stats['removed_english_in_b']}")
        print(f"Rows removed (Incomplete): {self.stats['removed_incomplete_a']}")
        print(f"Total rows removed:        {self.stats['removed_english_in_b'] + self.stats['removed_incomplete_a']}")
        print(f"Final rows:                {self.stats['final_rows']}")
```
- Print each statistic from the stats dictionary

```python
        if self.stats['initial_rows'] > 0:
            percentage = (self.stats['final_rows'] / self.stats['initial_rows']) * 100
            print(f"Data retention:            {percentage:.1f}%")
```
- **`/ self.stats['initial_rows']`** = Divide to get ratio
- **`* 100`** = Convert to percentage
- **`{percentage:.1f}%`** = Format to 1 decimal place

```python
        print("\n" + "="*60 + "\n")
```

---

## PART 3: HELPER FUNCTIONS

### 3.1 get_user_input Function

```python
def get_user_input() -> str:
    """
    Get file path from user with validation.
    
    Returns:
        str: Valid file path
    """
    while True:
```
- **`while True:`** = Infinite loop (keep asking until user provides valid input)

```python
        print("\n" + "="*60)
        print("MEITEI DATASET CLEANER")
        print("="*60)
        print("\nSupported formats: .md (Markdown), .xlsx or .xls (Excel)")
        print("\nPlease enter the input file path:")
        print("(You can drag and drop the file here or type the path)")
        
        filepath = input("\n> ").strip()
```
- **`input()`** = Get text from user
- **`.strip()`** = Remove whitespace from beginning/end
- **`"> "`** = Prompt symbol

```python
        filepath = filepath.strip('"').strip("'")
```
- **`.strip('"')`** = Remove double quotes (if user dragged a file with quotes)
- **`.strip("'")`** = Remove single quotes

```python
        if not Path(filepath).exists():
            print(f"\n❌ File not found: {filepath}")
            print("Please check the path and try again.\n")
            continue
```
- **`.exists()`** = Check if file exists
- **`continue`** = Skip to next iteration of while loop (ask again)

```python
        ext = Path(filepath).suffix.lower()
        if ext not in ['.md', '.xlsx', '.xls']:
            print(f"\n❌ Unsupported file format: {ext}")
            print("Please use .md, .xlsx, or .xls files.\n")
            continue
```
- **`.suffix`** = Get file extension
- **`not in [...]`** = Check if it's NOT in the list
- Ask again if format is wrong

```python
        return filepath
```
- Once validation passes, return the valid path and exit the loop

---

## PART 4: MAIN FUNCTION

```python
def main():
    """Main entry point with interactive user input"""
    try:
        input_file = get_user_input()
```
- Get the input file path from user
- **`try:`** = Start error handling

```python
        cleaner = MeiteiDatasetCleaner(verbose=True)
```
- Create a new cleaner instance with verbose mode ON

```python
        print("\n" + "="*60)
        print("Setting up output folder...")
        print("="*60)
        
        if not cleaner.create_output_folder():
            print("\n❌ Failed to create output folder. Exiting.")
            sys.exit(1)
```
- Create the output folder
- **`sys.exit(1)`** = Exit program with error code 1 (means error)

```python
        print(f"\nLoading file: {input_file}")
        if not cleaner.load_file(input_file):
            print("\n❌ Failed to load file. Exiting.")
            sys.exit(1)
```
- Load the input file
- Exit if it fails

```python
        cleaner.clean()
```
- Execute all cleaning steps

```python
        print("\nSaving cleaned data...")
        output_path = cleaner.save_excel()
        
        if not output_path:
            print("\n❌ Failed to save file. Exiting.")
            sys.exit(1)
```
- Save the cleaned data
- Exit if saving fails

```python
        cleaner.print_summary()
```
- Print the statistics

```python
        print("="*60)
        print("OUTPUT FOLDER LOCATION")
        print("="*60)
        print(f"\nOutput folder: {cleaner.output_dir}")
        print(f"Output file: {Path(output_path).name}")
        print(f"\nFull path: {output_path}")
        print("\n" + "="*60)
        print("✓ CLEANING COMPLETED SUCCESSFULLY!")
        print("="*60 + "\n")
```
- Display the output location to the user

```python
    except KeyboardInterrupt:
        print("\n\n❌ Operation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)
```
- **`except KeyboardInterrupt:`** = User pressed Ctrl+C
- **`except Exception as e:`** = Any other error
- Print error message and exit

---

## PART 5: ENTRY POINT

```python
if __name__ == '__main__':
    main()
```
- **`if __name__ == '__main__':`** = "Only run this if the file was run directly (not imported)"
- **`main()`** = Call the main function

**Why this matters:**
- If someone imports this file: `from meitei_dataset_cleaner import MeiteiDatasetCleaner`
- The `main()` function will NOT run (good - we only want the class)
- If someone runs the file: `python meitei_dataset_cleaner.py`
- The `main()` function WILL run (good - we want the interactive program)

---

## WORKFLOW SUMMARY

Here's how everything works together:

```
1. User runs: python meitei_dataset_cleaner.py
   ↓
2. __name__ == '__main__' is True
   ↓
3. main() is called
   ↓
4. get_user_input() asks for file path
   ↓
5. MeiteiDatasetCleaner instance is created
   ↓
6. create_output_folder() creates "cleaned_data" folder
   ↓
7. load_file() reads the Excel/Markdown file
   ↓
8. clean() executes three cleaning steps:
   - remove_english_in_column_b()
   - remove_incomplete_in_column_a()
   - convert_column_b_to_meitei()
   ↓
9. save_excel() saves the cleaned file with timestamp
   ↓
10. print_summary() shows statistics
   ↓
11. Program displays output folder location
   ↓
12. Program ends successfully
```

---

## KEY CONCEPTS USED

### Object-Oriented Programming (OOP)
- **Classes** - Blueprint for objects
- **Methods** - Functions inside classes
- **Instance variables** - Data stored in objects (`self.df`, `self.stats`)
- **`self`** - Reference to current object

### Data Processing
- **Pandas DataFrame** - Table-like data structure
- **Filtering** - Removing rows based on conditions
- **Applying functions** - Running a function on each row

### Regular Expressions
- **Pattern matching** - Finding text that matches a pattern
- **`[a-zA-Z]`** - Any letter
- **`{2,}`** - 2 or more occurrences
- **`^` and `$`** - Start and end of string

### File Operations
- **Path objects** - Modern way to work with files
- **Reading/Writing** - Loading data and saving results
- **Encoding** - UTF-8 for supporting Meitei script

### Error Handling
- **try/except** - Catching errors gracefully
- **Returning values** - Functions return True/False or results
- **Exit codes** - 0 = success, 1 = error

---

This complete explanation covers every significant line of code! Let me know if you want more detail on any specific part.
