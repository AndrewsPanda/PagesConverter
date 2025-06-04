#!/usr/bin/env python3
"""
Filename Checker for Pages Documents
Finds filenames with problematic characters before conversion
"""

import sys
from pathlib import Path
import re

def check_filenames(directory):
    """Check for problematic characters in Pages filenames"""
    
    target_dir = Path(directory)
    if not target_dir.exists():
        print(f"Error: Directory not found: {directory}")
        return
    
    # Find all Pages files
    pages_files = list(target_dir.rglob("*.pages"))
    print(f"\nChecking {len(pages_files)} Pages files for problematic characters...\n")
    
    # Problematic characters
    problematic_chars = {
        '\\': 'backslash',
        '"': 'double quote',
        '/': 'forward slash in filename',
        ':': 'colon (appears as slash on Mac)',
        '*': 'asterisk',
        '?': 'question mark',
        '<': 'less than',
        '>': 'greater than',
        '|': 'pipe',
        '\n': 'newline',
        '\r': 'carriage return',
        '\t': 'tab'
    }
    
    issues_found = 0
    
    for file in pages_files:
        filename = file.name
        problems = []
        
        # Check for each problematic character
        for char, description in problematic_chars.items():
            if char in filename:
                problems.append(f"{description} ({char})")
        
        # Check for other issues
        if filename.endswith(' .pages'):
            problems.append("trailing space before extension")
        if '  ' in filename:
            problems.append("double spaces")
        if filename.startswith('.'):
            problems.append("hidden file (starts with dot)")
        
        # Report problems
        if problems:
            issues_found += 1
            print(f"❌ {filename}")
            for problem in problems:
                print(f"   Issue: {problem}")
            
            # Suggest fixed name
            fixed_name = filename
            for char in ['\\', '/', ':', '*', '?', '<', '>', '|', '"']:
                fixed_name = fixed_name.replace(char, '-')
            fixed_name = fixed_name.replace('\n', '').replace('\r', '').replace('\t', ' ')
            fixed_name = re.sub(r'\s+', ' ', fixed_name)  # Replace multiple spaces with single
            fixed_name = fixed_name.rstrip()  # Remove trailing spaces
            
            if fixed_name != filename:
                print(f"   Suggested: {fixed_name}")
            print()
    
    # Summary
    if issues_found == 0:
        print("✅ All filenames are clean! No problematic characters found.")
    else:
        print(f"\n⚠️  Found {issues_found} files with problematic names")
        print("\nTo fix these issues:")
        print("1. Rename the files to remove problematic characters")
        print("2. Replace special characters with dashes or underscores")
        print("3. Remove trailing spaces and fix double spaces")

def main():
    if len(sys.argv) > 1:
        directory = sys.argv[1]
    else:
        print("Filename Checker for Pages Documents")
        print("Enter the path to check:")
        directory = input().strip().strip("'\"")
    
    check_filenames(directory)

if __name__ == "__main__":
    main()