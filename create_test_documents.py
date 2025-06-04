#!/usr/bin/env python3
"""
Create Test Pages Documents
This script creates sample Pages documents for testing the converter
"""

import os
import subprocess
from pathlib import Path
import time

def create_test_pages_documents():
    """Create test Pages documents using AppleScript"""
    
    # Create test folder on Desktop
    test_folder = Path.home() / "Desktop" / "TestPagesDocuments"
    test_folder.mkdir(exist_ok=True)
    
    print(f"Creating test documents in: {test_folder}")
    
    # Test documents to create
    test_docs = [
        ("Simple Test Document", "This is a simple test document for conversion testing."),
        ("Document with Formatting", "This document has **bold text** and *italic text*."),
        ("Longer Document", "This is a longer document.\n\n" * 10 + "End of document.")
    ]
    
    created_count = 0
    
    for filename, content in test_docs:
        print(f"\nCreating: {filename}.pages")
        
        # Clean filename for use in paths
        safe_filename = filename.replace(" ", "_")
        file_path = test_folder / f"{safe_filename}.pages"
        
        # AppleScript to create Pages document
        applescript = f'''
        tell application "Pages"
            activate
            set newDoc to make new document
            tell newDoc
                set body text to "{content}"
                save to POSIX file "{file_path}"
                close saving yes
            end tell
        end tell
        '''
        
        try:
            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0 and file_path.exists():
                print(f"✅ Created: {filename}.pages")
                created_count += 1
            else:
                print(f"❌ Failed to create: {filename}.pages")
                if result.stderr:
                    print(f"   Error: {result.stderr.strip()}")
        except subprocess.TimeoutExpired:
            print(f"❌ Timeout creating: {filename}.pages")
        except Exception as e:
            print(f"❌ Error creating {filename}.pages: {e}")
        
        # Small delay between documents
        time.sleep(1)
    
    # Quit Pages
    subprocess.run(['osascript', '-e', 'tell application "Pages" to quit'], 
                   capture_output=True)
    
    print("\n" + "="*60)
    print(f"Created {created_count} test Pages documents")
    print(f"Location: {test_folder}")
    print("\nYou can now test the converter with this folder!")
    print("="*60)
    
    # Open the folder in Finder
    subprocess.run(['open', str(test_folder)])
    
    return test_folder

def main():
    print("Test Pages Document Creator")
    print("This will create sample Pages documents for testing")
    print("-"*60)
    
    response = input("\nCreate test Pages documents? (y/n): ")
    if response.lower() != 'y':
        print("Cancelled")
        return
    
    try:
        test_folder = create_test_pages_documents()
        
        print("\n📝 Next Steps:")
        print("1. Run the pages_converter.py script")
        print("2. When prompted, drag the TestPagesDocuments folder")
        print("3. Or run directly:")
        print(f"   python3 pages_converter.py '{test_folder}'")
        
    except KeyboardInterrupt:
        print("\n\nCancelled by user")
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    main()