#!/usr/bin/env python3
"""
Pages to Word Converter - Python Version
More stable alternative to AppleScript for macOS Sequoia
Requires: Python 3.6+ (included with macOS)

UPDATED: Now skips files already in 'pages' folders and checks for existing conversions
"""

import os
import sys
import subprocess
import time
import logging
from pathlib import Path
from datetime import datetime
import shutil

class PagesToWordConverter:
    def __init__(self, target_dir):
        self.target_dir = Path(target_dir)
        self.batch_size = 5
        self.files_processed = 0
        self.converted_count = 0
        self.error_count = 0
        self.skipped_count = 0
        
        # Setup logging
        log_filename = f"pages_conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        self.log_path = Path.home() / "Desktop" / log_filename
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_path),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def find_pages_files(self):
        """Find all Pages documents in target directory, excluding those already in 'pages' folders"""
        # Find all .pages files
        all_pages_files = list(self.target_dir.rglob("*.pages"))
        
        # Filter out files that are already in 'pages' folders
        pages_files = []
        already_processed = 0
        
        for file in all_pages_files:
            # Check if this file is in a 'pages' subfolder
            if 'pages' in file.parts[len(self.target_dir.parts):]:
                already_processed += 1
                self.logger.debug(f"Skipping (in pages folder): {file.name}")
            else:
                # Also check if a corresponding .docx already exists
                word_file = file.with_suffix('.docx')
                if word_file.exists():
                    already_processed += 1
                    self.logger.info(f"Skipping (already converted): {file.name}")
                else:
                    pages_files.append(file)
        
        self.logger.info(f"Found {len(all_pages_files)} total Pages documents")
        self.logger.info(f"Already processed: {already_processed}")
        self.logger.info(f"To be converted: {len(pages_files)}")
        
        return pages_files
    
    def restart_pages_app(self):
        """Restart Pages app for memory management"""
        self.logger.info("Restarting Pages app for memory cleanup...")
        
        # Quit Pages
        subprocess.run([
            'osascript', '-e', 
            'tell application "Pages" to quit'
        ], capture_output=True)
        time.sleep(3)
        
        # Restart Pages
        subprocess.run([
            'osascript', '-e', 
            'tell application "Pages" to activate'
        ], capture_output=True)
        time.sleep(2)
    
    def sanitize_filename_for_applescript(self, filepath):
        """Escape special characters in filepath for AppleScript"""
        # Convert Path to string and escape backslashes and quotes
        filepath_str = str(filepath)
        filepath_str = filepath_str.replace('\\', '\\\\')
        filepath_str = filepath_str.replace('"', '\\"')
        return filepath_str
    
    def convert_file(self, pages_file):
        """Convert a single Pages file to Word format"""
        # Prepare paths
        word_file = pages_file.with_suffix('.docx')
        
        # Sanitize file paths for AppleScript
        pages_path = self.sanitize_filename_for_applescript(pages_file)
        word_path = self.sanitize_filename_for_applescript(word_file)
        
        # AppleScript for conversion
        applescript = f'''
        tell application "Pages"
            try
                set theDoc to open POSIX file "{pages_path}"
                delay 1
                export theDoc to POSIX file "{word_path}" as Microsoft Word
                close theDoc saving no
                return "SUCCESS"
            on error errMsg number errNum
                return "ERROR " & errNum & ": " & errMsg
            end try
        end tell
        '''
        
        # Execute AppleScript
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True
        )
        
        # Check result
        if result.returncode == 0 and word_file.exists():
            self.logger.info(f"✓ Converted: {pages_file.name}")
            return True, word_file
        else:
            error_msg = result.stderr or result.stdout or "Unknown error"
            self.logger.error(f"✗ Failed to convert {pages_file.name}: {error_msg}")
            return False, None
    
    def move_to_pages_folder(self, pages_file):
        """Move original Pages file to 'pages' subfolder"""
        pages_folder = pages_file.parent / "pages"
        pages_folder.mkdir(exist_ok=True)
        
        # Handle duplicate filenames
        dest_file = pages_folder / pages_file.name
        if dest_file.exists():
            base = pages_file.stem
            suffix = pages_file.suffix
            counter = 1
            while (pages_folder / f"{base}_{counter}{suffix}").exists():
                counter += 1
            dest_file = pages_folder / f"{base}_{counter}{suffix}"
        
        try:
            shutil.move(str(pages_file), str(dest_file))
            self.logger.info(f"✓ Moved to: pages/{dest_file.name}")
            return True
        except Exception as e:
            self.logger.error(f"✗ Failed to move {pages_file.name}: {e}")
            return False
    
    def show_progress(self, current, total, filename):
        """Display progress bar in terminal"""
        bar_length = 40
        progress = current / total if total > 0 else 1
        filled = int(bar_length * progress)
        bar = '█' * filled + '░' * (bar_length - filled)
        
        print(f'\r[{bar}] {current}/{total} - {filename[:50]:<50}', end='', flush=True)
    
    def convert_all(self):
        """Main conversion process"""
        self.logger.info("Starting Pages to Word conversion")
        self.logger.info(f"Target directory: {self.target_dir}")
        
        # Find all Pages files (excluding already processed ones)
        pages_files = self.find_pages_files()
        
        if not pages_files:
            print("\n✅ No new Pages documents to convert!")
            print("All Pages files in this directory have already been processed.")
            self.logger.info("No new files to convert")
            return
        
        # Confirm with user
        print(f"\nThis will convert {len(pages_files)} Pages documents to Word format.")
        print("Original files will be moved to 'pages' subfolders.")
        print("Already converted files will be skipped.")
        response = input("\nContinue? (y/n): ")
        if response.lower() != 'y':
            print("Cancelled by user")
            return
        
        # Start Pages app
        self.restart_pages_app()
        
        # Process files
        total_files = len(pages_files)
        for i, pages_file in enumerate(pages_files, 1):
            self.show_progress(i, total_files, pages_file.name)
            self.files_processed += 1
            
            # Convert file
            success, word_file = self.convert_file(pages_file)
            
            if success:
                self.converted_count += 1
                # Move original to pages folder
                self.move_to_pages_folder(pages_file)
            else:
                self.error_count += 1
            
            # Restart Pages periodically for memory management
            if self.files_processed % self.batch_size == 0:
                print()  # New line after progress bar
                self.restart_pages_app()
            
            # Small delay between files
            time.sleep(0.5)
        
        # Cleanup
        print()  # New line after progress bar
        subprocess.run(['osascript', '-e', 'tell application "Pages" to quit'], 
                      capture_output=True)
        
        # Summary
        self.show_summary()
    
    def show_summary(self):
        """Display conversion summary"""
        print("\n" + "="*60)
        print(f"CONVERSION COMPLETE!")
        print(f"Total Pages documents processed: {self.files_processed}")
        print(f"Successfully converted: {self.converted_count} ✓")
        print(f"Errors: {self.error_count} ✗")
        print(f"\nLog file: {self.log_path}")
        print("="*60)
        
        # Open log file
        subprocess.run(['open', str(self.log_path)])

def main():
    """Main entry point"""
    print("Pages to Word Converter - Python Version")
    print("More stable alternative for macOS Sequoia\n")
    
    # Get target directory
    if len(sys.argv) > 1:
        target_dir = sys.argv[1]
    else:
        print("Enter the path to the folder containing Pages documents:")
        print("(You can drag and drop the folder here)")
        target_dir = input().strip().strip("'\"")
    
    # Validate directory
    if not os.path.isdir(target_dir):
        print(f"Error: Directory not found: {target_dir}")
        sys.exit(1)
    
    # Run converter
    converter = PagesToWordConverter(target_dir)
    try:
        converter.convert_all()
    except KeyboardInterrupt:
        print("\n\nConversion cancelled by user")
        subprocess.run(['osascript', '-e', 'tell application "Pages" to quit'], 
                      capture_output=True)
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        print(f"\nError: {e}")

if __name__ == "__main__":
    main()