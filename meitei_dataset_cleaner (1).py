
"""
Meitei Language Dataset Cleaner - Interactive Version (With Deleted Rows Export)
==============================================================================
A comprehensive tool for cleaning multilingual Meitei language datasets.

Features:
- Interactive user input for file selection
- Automatic output folder creation
- Removes rows with English letters in column B (meitei_script)
- Removes rows with incomplete/meaningless information in column A (english)
- Converts numbers to Meitei script numerals in column B only
- Supports markdown and Excel input formats
- Outputs cleaned data as Excel file in dedicated folder
- Outputs DELETED/REJECTED rows as an Excel file with reason for deletion

Author: Dataset Cleaning Tool
Date: 2026
"""

import pandas as pd
import re
import sys
import os
from pathlib import Path
from datetime import datetime
from typing import Tuple, List


class MeiteiDatasetCleaner:
    """Main class for cleaning Meitei language datasets"""
    
    # Meitei numeral mapping
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
    
    def __init__(self, verbose=True):
        """
        Initialize the cleaner.
        
        Args:
            verbose (bool): Print detailed information during processing
        """
        self.verbose = verbose
        self.df = None
        self.deleted_df = pd.DataFrame()  # To track deleted/rejected rows
        self.stats = {
            'initial_rows': 0,
            'removed_english_in_b': 0,
            'removed_incomplete_a': 0,
            'final_rows': 0
        }
        self.program_dir = Path(__file__).parent
        self.output_dir = self.program_dir / 'cleaned_data'
    
    def create_output_folder(self) -> bool:
        """
        Create output folder in the program directory.
        
        Returns:
            bool: True if successful or already exists, False otherwise
        """
        try:
            self.output_dir.mkdir(exist_ok=True)
            if self.verbose:
                print(f"✓ Output folder: {self.output_dir}")
            return True
        except Exception as e:
            print(f"❌ Error creating output folder: {e}")
            return False
    
    def load_file(self, filepath: str) -> bool:
        """
        Load data from markdown or Excel file.
        
        Args:
            filepath (str): Path to input file
            
        Returns:
            bool: True if successful, False otherwise
        """
        filepath = Path(filepath)
        
        if not filepath.exists():
            print(f"❌ Error: File not found: {filepath}")
            return False
        
        try:
            if filepath.suffix.lower() == '.md':
                self._load_markdown(filepath)
            elif filepath.suffix.lower() in ['.xlsx', '.xls']:
                self._load_excel(filepath)
            else:
                print(f"❌ Error: Unsupported file format: {filepath.suffix}")
                return False
            
            self.stats['initial_rows'] = len(self.df)
            self.deleted_df = pd.DataFrame(columns=list(self.df.columns) + ['Reason for Deletion'])
            
            if self.verbose:
                print(f"✓ Loaded {len(self.df)} rows from {filepath.name}")
            return True
            
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            return False
    
    def _load_markdown(self, filepath: Path):
        """Load data from markdown table"""
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        lines = content.strip().split('\n')
        data = []
        
        for line in lines:
            if line.startswith('##') or line.startswith('| ---'):
                continue
            
            if line.startswith('|') and '|' in line:
                parts = [p.strip() for p in line.split('|')]
                parts = [p for p in parts if p]
                
                if len(parts) == 3:
                    data.append(parts)
        
        # Remove header row if present
        if data and data[0][0].lower() == 'english':
            data = data[1:]
        
        self.df = pd.DataFrame(data, columns=['english', 'meitei_script', 'roman_standard'])
    
    def _load_excel(self, filepath: Path):
        """Load data from Excel file"""
        self.df = pd.read_excel(filepath)
        
        # Ensure correct column names
        if len(self.df.columns) >= 3:
            self.df.columns = ['english', 'meitei_script', 'roman_standard']
    
    def _contains_english_letters(self, text: str) -> bool:
        """Check if text contains English alphabet letters (a-z, A-Z)"""
        if pd.isna(text):
            return False
        return bool(re.search(r'[a-zA-Z]', str(text)))
    
    def _is_meaningful_text(self, text: str) -> bool:
        """
        Check if text is meaningful and complete.
        
        Returns True if text should be KEPT, False if it should be REMOVED.
        """
        if pd.isna(text) or text is None:
            return False
        
        text = str(text).strip()
        
        # Check: is empty
        if len(text) == 0:
            return False
        
        # Check: too short (less than 20 characters - incomplete info)
        if len(text) < 20:
            return False
        
        # Check: just repeated characters
        if len(set(text)) == 1:
            return False
        
        # Check: has at least 2 meaningful words
        word_pattern = re.compile(r'[a-zA-Z]{2,}')
        words = word_pattern.findall(text)
        if len(words) < 2:
            return False
        
        # Check: not just punctuation and numbers
        if re.match(r'^[0-9\s\.,;:\-\—\—\/\(\)\[\]\{\}\.]+$', text):
            return False
        
        # Check: not too many special characters (more than 50%)
        special_char_ratio = sum(1 for c in text if not c.isalnum() and c != ' ') / len(text)
        if special_char_ratio > 0.5:
            return False
        
        # Check: doesn't end with incomplete marks
        if re.search(r'[—\-_\…]{2,}$', text):
            return False
        
        return True
    
    def _convert_numbers_to_meitei(self, text: str) -> str:
        """Convert all digits to Meitei numerals"""
        for digit, meitei_digit in self.MEITEI_NUMERALS.items():
            text = str(text).replace(digit, meitei_digit)
        return text
    
    def remove_english_in_column_b(self) -> int:
        """
        Remove rows where column B (meitei_script) contains English letters.
        
        Returns:
            int: Number of rows removed
        """
        if self.df is None:
            print("❌ Error: No data loaded")
            return 0
        
        # Identify rows to remove
        mask_to_remove = self.df['meitei_script'].apply(self._contains_english_letters)
        rejected_rows = self.df[mask_to_remove].copy()
        
        if not rejected_rows.empty:
            rejected_rows['Reason for Deletion'] = 'English letters in Meitei Script (Column B)'
            self.deleted_df = pd.concat([self.deleted_df, rejected_rows], ignore_index=True)
            
        initial_count = len(self.df)
        self.df = self.df[~mask_to_remove]
        removed = initial_count - len(self.df)
        self.stats['removed_english_in_b'] = removed
        
        if self.verbose:
            print(f"✓ Removed {removed} rows with English letters in column B")
            print(f"  Remaining: {len(self.df)} rows")
        
        return removed
    
    def remove_incomplete_in_column_a(self) -> int:
        """
        Remove rows with incomplete or meaningless information in column A (english).
        
        Returns:
            int: Number of rows removed
        """
        if self.df is None:
            print("❌ Error: No data loaded")
            return 0
        
        # Identify rows to remove (where _is_meaningful_text is False)
        mask_to_keep = self.df['english'].apply(self._is_meaningful_text)
        rejected_rows = self.df[~mask_to_keep].copy()
        
        if not rejected_rows.empty:
            rejected_rows['Reason for Deletion'] = 'Incomplete/Meaningless English text (Column A)'
            self.deleted_df = pd.concat([self.deleted_df, rejected_rows], ignore_index=True)
            
        initial_count = len(self.df)
        self.df = self.df[mask_to_keep]
        removed = initial_count - len(self.df)
        self.stats['removed_incomplete_a'] = removed
        
        if self.verbose:
            print(f"✓ Removed {removed} rows with incomplete/meaningless information in column A")
            print(f"  Remaining: {len(self.df)} rows")
        
        return removed
    
    def convert_column_b_to_meitei(self) -> None:
        """Convert numbers to Meitei script ONLY in column B"""
        if self.df is None:
            print("❌ Error: No data loaded")
            return
        
        self.df['meitei_script'] = self.df['meitei_script'].apply(self._convert_numbers_to_meitei)
        
        if self.verbose:
            print(f"✓ Converted numbers to Meitei script in column B")
    
    def clean(self) -> None:
        """Execute all cleaning steps in sequence"""
        if self.df is None:
            print("❌ Error: No data loaded")
            return
        
        if self.verbose:
            print("\n" + "="*60)
            print("MEITEI DATASET CLEANER - PROCESSING")
            print("="*60 + "\n")
        
        print("Step 1: Removing rows with English letters in column B...")
        self.remove_english_in_column_b()
        
        print("\nStep 2: Removing incomplete/meaningless rows in column A...")
        self.remove_incomplete_in_column_a()
        
        print("\nStep 3: Converting numbers to Meitei script in column B...")
        self.convert_column_b_to_meitei()
        
        self.stats['final_rows'] = len(self.df)
    
    def save_excel(self, filename: str = None) -> Tuple[str, str]:
        """
        Save cleaned data and deleted data to Excel files in output folder.
        
        Args:
            filename (str): Optional filename base
            
        Returns:
            Tuple[str, str]: (Cleaned file path, Deleted file path)
        """
        if self.df is None:
            print("❌ Error: No data to save")
            return "", ""
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if filename is None:
                cleaned_filename = f"cleaned_meitei_{timestamp}.xlsx"
                deleted_filename = f"deleted_meitei_{timestamp}.xlsx"
            else:
                cleaned_filename = f"cleaned_{filename}"
                deleted_filename = f"deleted_{filename}"
            
            output_path_cleaned = self.output_dir / cleaned_filename
            output_path_deleted = self.output_dir / deleted_filename
            
            # Save cleaned data
            self.df.to_excel(output_path_cleaned, index=False, sheet_name='Cleaned Data')
            if self.verbose:
                print(f"✓ Cleaned data saved to {output_path_cleaned}")
                
            # Save deleted data
            self.deleted_df.to_excel(output_path_deleted, index=False, sheet_name='Deleted Rows')
            if self.verbose:
                print(f"✓ Deleted rows saved to {output_path_deleted}")
            
            return str(output_path_cleaned), str(output_path_deleted)
        except Exception as e:
            print(f"❌ Error saving files: {e}")
            return "", ""
    
    def print_summary(self) -> None:
        """Print cleaning summary statistics"""
        print("\n" + "="*60)
        print("CLEANING SUMMARY")
        print("="*60)
        print(f"\nInitial rows:              {self.stats['initial_rows']}")
        print(f"Rows removed (English):    {self.stats['removed_english_in_b']}")
        print(f"Rows removed (Incomplete): {self.stats['removed_incomplete_a']}")
        print(f"Total rows removed:        {self.stats['removed_english_in_b'] + self.stats['removed_incomplete_a']}")
        print(f"Final rows:                {self.stats['final_rows']}")
        
        if self.stats['initial_rows'] > 0:
            percentage = (self.stats['final_rows'] / self.stats['initial_rows']) * 100
            print(f"Data retention:            {percentage:.1f}%")
        print("\n" + "="*60 + "\n")


def get_user_input() -> str:
    """
    Get file path from user with validation.
    
    Returns:
        str: Valid file path
    """
    while True:
        print("\n" + "="*60)
        print("MEITEI DATASET CLEANER")
        print("="*60)
        print("\nSupported formats: .md (Markdown), .xlsx or .xls (Excel)")
        print("\nPlease enter the input file path:")
        print("(You can drag and drop the file here or type the path)")
        
        filepath = input("\n> ").strip()
        
        # Remove quotes if user dragged and dropped
        filepath = filepath.strip('"').strip("'")
        
        # Check if file exists
        if not Path(filepath).exists():
            print(f"\n❌ File not found: {filepath}")
            print("Please check the path and try again.\n")
            continue
        
        # Check file extension
        ext = Path(filepath).suffix.lower()
        if ext not in ['.md', '.xlsx', '.xls']:
            print(f"\n❌ Unsupported file format: {ext}")
            print("Please use .md, .xlsx, or .xls files.\n")
            continue
        
        return filepath


def main():
    """Main entry point with interactive user input"""
    try:
        # Get input file from user
        input_file = get_user_input()
        
        # Create cleaner instance
        cleaner = MeiteiDatasetCleaner(verbose=True)
        
        # Create output folder
        print("\n" + "="*60)
        print("Setting up output folder...")
        print("="*60)
        
        if not cleaner.create_output_folder():
            print("\n❌ Failed to create output folder. Exiting.")
            sys.exit(1)
        
        # Load file
        print(f"\nLoading file: {input_file}")
        if not cleaner.load_file(input_file):
            print("\n❌ Failed to load file. Exiting.")
            sys.exit(1)
        
        # Execute cleaning
        cleaner.clean()
        
        # Save output
        print("\nSaving files...")
        output_path, deleted_path = cleaner.save_excel()
        
        if not output_path or not deleted_path:
            print("\n❌ Failed to save files. Exiting.")
            sys.exit(1)
        
        # Print summary
        cleaner.print_summary()
        
        # Print output folder information
        print("="*60)
        print("OUTPUT FOLDER LOCATION")
        print("="*60)
        print(f"\nOutput folder: {cleaner.output_dir}")
        print(f"Cleaned file:  {Path(output_path).name}")
        print(f"Deleted file:  {Path(deleted_path).name}")
        print(f"\nFull Cleaned Path: {output_path}")
        print(f"Full Deleted Path: {deleted_path}")
        print("\n" + "="*60)
        print("✓ CLEANING COMPLETED SUCCESSFULLY!")
        print("="*60 + "\n")
        
    except KeyboardInterrupt:
        print("\n\n❌ Operation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
