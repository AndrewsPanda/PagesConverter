#!/usr/bin/env python3
"""
Permission Checker for Pages Converter
Run this BEFORE the main converter to verify all permissions are set correctly
"""

import os
import sys
import subprocess
from pathlib import Path
import platform

class PermissionChecker:
    def __init__(self):
        self.issues_found = 0
        self.warnings_found = 0
        
    def print_header(self, text):
        print("\n" + "="*60)
        print(f" {text}")
        print("="*60)
    
    def check_mark(self, status):
        return "✅" if status else "❌"
    
    def check_python_version(self):
        """Check if Python 3 is properly installed"""
        self.print_header("Python Installation")
        
        version = sys.version_info
        print(f"Python Version: {version.major}.{version.minor}.{version.micro}")
        
        if version.major >= 3 and version.minor >= 6:
            print(f"{self.check_mark(True)} Python 3.6+ is installed")
            return True
        else:
            print(f"{self.check_mark(False)} Python 3.6+ is required")
            self.issues_found += 1
            return False
    
    def check_macos_version(self):
        """Check macOS version"""
        self.print_header("macOS Version")
        
        mac_ver = platform.mac_ver()[0]
        print(f"macOS Version: {mac_ver}")
        
        # Check if it's macOS 15.x (Sequoia)
        if mac_ver.startswith("15."):
            print(f"⚠️  Warning: macOS Sequoia detected - known automation issues")
            self.warnings_found += 1
        else:
            print(f"{self.check_mark(True)} macOS version compatible")
    
    def check_pages_installed(self):
        """Check if Pages is installed"""
        self.print_header("Pages App")
        
        pages_path = "/Applications/Pages.app"
        if os.path.exists(pages_path):
            print(f"{self.check_mark(True)} Pages is installed")
            
            # Try to get Pages version
            try:
                result = subprocess.run(
                    ['osascript', '-e', 'tell application "Pages" to version'],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    print(f"Pages Version: {result.stdout.strip()}")
            except:
                print("Could not determine Pages version")
        else:
            print(f"{self.check_mark(False)} Pages is not installed")
            self.issues_found += 1
    
    def check_file_access(self):
        """Check if we can access common directories"""
        self.print_header("File System Access")
        
        test_dirs = [
            ("Desktop", Path.home() / "Desktop"),
            ("Documents", Path.home() / "Documents"),
            ("Downloads", Path.home() / "Downloads")
        ]
        
        for name, path in test_dirs:
            try:
                # Try to list directory
                list(path.iterdir())
                print(f"{self.check_mark(True)} Can access {name} folder")
            except PermissionError:
                print(f"{self.check_mark(False)} Cannot access {name} folder - Check Full Disk Access")
                self.issues_found += 1
            except:
                print(f"⚠️  {name} folder issue")
                self.warnings_found += 1
    
    def check_applescript_automation(self):
        """Check if AppleScript automation works"""
        self.print_header("AppleScript Automation")
        
        # Test basic AppleScript
        try:
            result = subprocess.run(
                ['osascript', '-e', 'return "test"'],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                print(f"{self.check_mark(True)} AppleScript is working")
            else:
                print(f"{self.check_mark(False)} AppleScript failed")
                self.issues_found += 1
        except:
            print(f"{self.check_mark(False)} AppleScript not available")
            self.issues_found += 1
        
        # Test Pages automation
        print("\nTesting Pages automation (this may trigger permission dialogs)...")
        try:
            result = subprocess.run(
                ['osascript', '-e', 'tell application "Pages" to name'],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                print(f"{self.check_mark(True)} Can automate Pages")
            else:
                print(f"{self.check_mark(False)} Cannot automate Pages - Grant permission in System Settings")
                print("Error:", result.stderr.strip())
                self.issues_found += 1
        except subprocess.TimeoutExpired:
            print(f"⚠️  Pages automation timed out - may need permissions")
            self.warnings_found += 1
    
    def check_terminal_permissions(self):
        """Check Terminal permissions"""
        self.print_header("Terminal Permissions")
        
        # Check if we're running in Terminal or VS Code
        terminal_app = os.environ.get('TERM_PROGRAM', 'Unknown')
        print(f"Running in: {terminal_app}")
        
        if terminal_app == "vscode":
            print(f"{self.check_mark(True)} VS Code terminal detected")
        elif terminal_app == "Apple_Terminal":
            print(f"{self.check_mark(True)} macOS Terminal detected")
        else:
            print(f"⚠️  Unknown terminal environment")
    
    def test_write_permissions(self):
        """Test if we can write to Desktop"""
        self.print_header("Write Permissions Test")
        
        test_file = Path.home() / "Desktop" / "pages_converter_test.txt"
        try:
            test_file.write_text("Permission test")
            print(f"{self.check_mark(True)} Can write to Desktop")
            test_file.unlink()  # Delete test file
        except:
            print(f"{self.check_mark(False)} Cannot write to Desktop")
            self.issues_found += 1
    
    def check_screen_lock_settings(self):
        """Remind about screen lock settings"""
        self.print_header("Screen Lock Settings (Manual Check Required)")
        
        print("⚠️  IMPORTANT: Check these settings manually:")
        print("1. System Settings → Lock Screen")
        print("2. Set 'Turn display off...' to Never (during conversion)")
        print("3. Set 'Require password...' to Never (during conversion)")
        print("\nScreen lock/sleep will cause the converter to hang!")
    
    def show_summary(self):
        """Show summary and recommendations"""
        self.print_header("Summary")
        
        if self.issues_found == 0 and self.warnings_found == 0:
            print("✅ All checks passed! Your system is ready for Pages conversion.")
            print("\nYou can now run the main converter script.")
        elif self.issues_found > 0:
            print(f"❌ Found {self.issues_found} critical issues that must be fixed:")
            print("\n1. Go to System Settings → Privacy & Security → Full Disk Access")
            print("   - Add Terminal.app and Visual Studio Code.app")
            print("\n2. Go to System Settings → Privacy & Security → Automation")
            print("   - Enable Terminal → Pages")
            print("\n3. Restart Terminal/VS Code after changing permissions")
        else:
            print(f"⚠️  Found {self.warnings_found} warnings")
            print("The converter should work, but monitor for issues.")
        
        print("\n" + "-"*60)
        print("Next steps:")
        print("1. Fix any ❌ issues shown above")
        print("2. Run: python3 pages_converter.py")
        print("3. If you see permission dialogs, click 'OK' for all")

def main():
    print("Pages Converter Permission Checker")
    print("This will verify your system is set up correctly")
    print("-"*60)
    
    checker = PermissionChecker()
    
    # Run all checks
    checker.check_python_version()
    checker.check_macos_version()
    checker.check_pages_installed()
    checker.check_file_access()
    checker.test_write_permissions()
    checker.check_applescript_automation()
    checker.check_terminal_permissions()
    checker.check_screen_lock_settings()
    
    # Show summary
    checker.show_summary()

if __name__ == "__main__":
    main()