#!/usr/bin/env python3
"""
Batch Filename Fixer for Pages Documents
Automatically fixes problematic characters in filenames
"""

import sys
import os
from pathlib import Path
import re

def fix_filename(filename):
    """Fix problematic characters in a filename"""
    # Keep the stem and extension separate
    if filename.endswith('.pages'):
        stem = filename[:-6]
        ext = '.pages'
    else:
        stem = filename
        ext = ''
    
    # Replace problematic characters
    fixed = stem
    replacements = {
        '\\': '-',
        '/': '-',
        ':': '-',
        '*': '-',
        '?': '',
        '<': '-',
        '>': '-',
        '|': '-',
        '"': "'",
        '\n': ' ',
        '\r': ' ',
        '\t': ' '
    }
    
    for old, new in replacements.items():
        fixed = fixed.replace(old, new)
    
    # Fix multiple spaces and trailing spaces
    fixed = re.sub(r'\s+', ' ', fixed)
    fixed = fixed.strip()
    
    # Remove any remaining non-printable characters
    fixed = ''.join(char for char in fixed if char.isprintable())
    
    return fixed + ext

def batch_fix_filenames(directory, dry_run=True):
    """Fix all problematic filenames in directory"""
    
    target_dir = Path(directory)
    if not target_dir.exists():
        print(f"Error: Directory not found: {directory}")
        return
    
    # Find all Pages files
    pages_files = list(target_dir.rglob("*.pages"))
    print(f"\n{'DRY RUN - ' if dry_run else ''}Checking {len(pages_files)} Pages files...\n")
    
    files_to_fix = []
    
    # Check each file
    for file in pages_files:
        original_name = file.name
        fixed_name = fix_filename(original_name)
        
        if original_name != fixed_name:
            files_to_fix.append((file, fixed_name))
            print(f"📝 {original_name}")
            print(f"   → {fixed_name}")
    
    if not files_to_fix:
        print("✅ All filenames are already clean!")
        return
    
    print(f"\n{'Would fix' if dry_run else 'Fixing'} {len(files_to_fix)} files...")
    
    if dry_run:
        print("\nThis is a DRY RUN. No files were changed.")
        print("To actually rename files, run with --fix flag")
        return
    
    # Actually rename files
    renamed_count = 0
    error_count = 0
    
    for file, new_name in files_to_fix:
        new_path = file.parent / new_name
        
        # Handle duplicates
        if new_path.exists():
            base = new_name[:-6]  # Remove .pages
            counter = 1
            while (file.parent / f"{base}_{counter}.pages").exists():
                counter += 1
            new_path = file.parent / f"{base}_{counter}.pages"
            new_name = new_path.name
        
        try:
            file.rename(new_path)
            print(f"✅ Renamed to: {new_name}")
            renamed_count += 1
        except Exception as e:
            print(f"❌ Error renaming {file.name}: {e}")
            error_count += 1
    
    print(f"\n{'='*60}")
    print(f"Renamed: {renamed_count} files")
    if error_count > 0:
        print(f"Errors: {error_count} files")
    print(f"{'='*60}")

def main():
    import argparse
    parser = argparse.ArgumentParser(description='Fix problematic filenames in Pages documents')
    parser.add_argument('directory', nargs='?', help='Directory to process')
    parser.add_argument('--fix', action='store_true', help='Actually rename files (default is dry run)')
    
    args = parser.parse_args()
    
    if args.directory:
        directory = args.directory
    else:
        print("Batch Filename Fixer for Pages Documents")
        print("Enter the directory path:")
        directory = input().strip().strip("'\"")
    
    batch_fix_filenames(directory, dry_run=not args.fix)
    
    if not args.fix:
        print("\n💡 Tip: To actually rename files, run:")
        print(f"   python3 {sys.argv[0]} '{directory}' --fix")

if __name__ == "__main__":
    main()