#!/usr/bin/env python3
"""
Pages to Word Converter - Python Version
More stable alternative to AppleScript for macOS Sequoia
Requires: Python 3.6+ (included with macOS)
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
        """Find all Pages documents in target directory"""
        pages_files = list(self.target_dir.rglob("*.pages"))
        self.logger.info(f"Found {len(pages_files)} Pages documents")
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
    
    def convert_file(self, pages_file):
        """Convert a single Pages file to Word format"""
        # Prepare paths
        word_file = pages_file.with_suffix('.docx')
        
        # AppleScript for conversion
        applescript = f'''
        tell application "Pages"
            try
                set theDoc to open POSIX file "{pages_file}"
                delay 1
                export theDoc to POSIX file "{word_file}" as Microsoft Word
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
        progress = current / total
        filled = int(bar_length * progress)
        bar = '█' * filled + '░' * (bar_length - filled)
        
        print(f'\r[{bar}] {current}/{total} - {filename[:50]:<50}', end='', flush=True)
    
    def convert_all(self):
        """Main conversion process"""
        self.logger.info("Starting Pages to Word conversion")
        self.logger.info(f"Target directory: {self.target_dir}")
        
        # Find all Pages files
        pages_files = self.find_pages_files()
        if not pages_files:
            self.logger.warning("No Pages documents found")
            return
        
        # Confirm with user
        print(f"\nThis will convert {len(pages_files)} Pages documents to Word format.")
        print("Original files will be moved to 'pages' subfolders.")
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
        print(f"Total Pages documents: {self.files_processed}")
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