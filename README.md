# Pages to Word Converter for macOS

A Python-based solution for batch converting Apple Pages documents to Microsoft Word format on macOS, with special optimizations for macOS Sequoia (15.x) compatibility issues.

## 🎯 Overview

This project provides a stable alternative to Automator-based conversions, which have significant reliability issues on macOS Sequoia. The converter handles large batches efficiently with automatic memory management and comprehensive error handling.

### Key Features
- ✅ Batch conversion of Pages to Word documents
- ✅ Smart duplicate detection (won't re-convert files)
- ✅ Automatic memory management (prevents crashes)
- ✅ Filename validation and automatic fixing
- ✅ Preserves folder structure
- ✅ Detailed logging
- ✅ Progress tracking

## 📋 System Requirements

- **macOS**: Any version (optimized for Sequoia 15.x issues)
- **Python**: 3.6+ (included with macOS)
- **Pages**: Must be installed in /Applications
- **Disk Space**: 2x the size of your Pages documents
- **RAM**: 8GB minimum recommended

## 🚀 Quick Start

### 1. Initial Setup (One Time Only)

1. **Create project folder on Desktop**:
   ```bash
   mkdir ~/Desktop/PagesConverter
   cd ~/Desktop/PagesConverter
   ```

2. **Save all Python scripts** to this folder:
   - `pages_converter.py` - Main converter
   - `check_filenames.py` - Filename validator
   - `fix_filenames.py` - Batch filename fixer
   - `permission_checker.py` - System permission checker
   - `menu_launcher.py` - Interactive menu interface
   - `create_test_documents.py` - Test file generator

3. **Make scripts executable**:
   ```bash
   chmod +x *.py
   ```

### 2. System Permissions Setup

**CRITICAL**: Configure these settings before first use:

1. **Full Disk Access**:
   - System Settings → Privacy & Security → Full Disk Access
   - Add: Terminal.app
   - Add: Visual Studio Code.app (if using)

2. **Automation Permissions** (will appear after first run):
   - System Settings → Privacy & Security → Automation
   - Enable all permissions for Terminal

3. **Screen Lock** (temporary during conversion):
   - System Settings → Lock Screen
   - Set all options to "Never"
   - **Remember to restore after conversion!**

### 3. First Run

```bash
# Check system permissions first
python3 permission_checker.py

# Use the menu interface (easiest)
python3 menu_launcher.py
```

## 📖 Usage Guide

### Method 1: Interactive Menu (Recommended)

```bash
python3 menu_launcher.py
```

Menu options:
1. Check System Permissions
2. Create Test Documents
3. Convert Pages to Word
4. View Latest Log
5. Quick Setup Guide
6. Emergency Kill Pages

### Method 2: Direct Commands

#### Check for Problematic Filenames
```bash
python3 check_filenames.py /path/to/folder
```

#### Fix Problematic Filenames
```bash
# Preview changes (dry run)
python3 fix_filenames.py /path/to/folder

# Actually fix filenames
python3 fix_filenames.py /path/to/folder --fix
```

#### Convert Pages to Word
```bash
# Interactive mode (will ask for folder)
python3 pages_converter.py

# Direct mode
python3 pages_converter.py /path/to/folder
```

## 🔄 Recommended Workflow

For best results, follow this sequence:

```bash
# 1. Check permissions (first time only)
python3 permission_checker.py

# 2. Check for filename issues
python3 check_filenames.py ~/Desktop/YourFolder

# 3. Fix any issues found
python3 fix_filenames.py ~/Desktop/YourFolder --fix

# 4. Run the converter
python3 pages_converter.py ~/Desktop/YourFolder
```

## 📁 How Files Are Organized

Before conversion:
```
YourFolder/
├── Document1.pages
├── Document2.pages
└── Subfolder/
    └── Document3.pages
```

After conversion:
```
YourFolder/
├── Document1.docx        # New Word file
├── Document2.docx        # New Word file
├── pages/               # New folder for originals
│   ├── Document1.pages  # Original moved here
│   └── Document2.pages  # Original moved here
└── Subfolder/
    ├── Document3.docx   # New Word file
    └── pages/           # New folder for originals
        └── Document3.pages
```

## ⚠️ Common Issues and Solutions

### Issue: "Permission denied" errors
**Solution**: Re-add Terminal to Full Disk Access in System Settings

### Issue: Conversion hangs or stops
**Solution**: 
- Ensure screen is unlocked
- Disable screen sleep during conversion
- Check Activity Monitor for Pages memory usage

### Issue: Some filenames cause errors
**Problem characters**:
- Backslash `\`
- Colon `:`
- Trailing spaces before `.pages`
- Double spaces
- Tab characters

**Solution**: Run `fix_filenames.py` before converting

### Issue: "Already converted" but no Word files
**Solution**: 
- Check if .docx files are hidden
- Look in 'pages' subfolders
- Check the log file on Desktop

## 🛠️ Advanced Usage

### Converting Multiple Folders
```bash
# Create a simple loop
for folder in ~/Documents/*/; do
    echo "Processing $folder"
    python3 check_filenames.py "$folder"
    python3 fix_filenames.py "$folder" --fix
    python3 pages_converter.py "$folder"
done
```

### Customizing Batch Size
Edit `pages_converter.py` line ~19:
```python
self.batch_size = 5  # Reduce to 3 for low memory systems
```

### Filtering Large Files
Add after line ~47 in `pages_converter.py`:
```python
# Skip files larger than 50MB
pages_files = [f for f in pages_files if f.stat().st_size < 50_000_000]
```

## 📊 Performance Tips

1. **Close other applications** to free memory
2. **Connect to power** (don't run on battery)
3. **Process in batches** of 100-200 files
4. **Monitor Activity Monitor** for Pages memory usage
5. **Keep screen unlocked** during conversion

## 🐛 Troubleshooting

### Check the Log Files
Log files are saved to Desktop with timestamp:
```bash
# View latest log
ls -la ~/Desktop/pages_conversion_*.log
cat ~/Desktop/pages_conversion_*.log | tail -50
```

### Force Quit Pages if Stuck
```bash
killall Pages
```

### Prevent Sleep During Long Conversions
```bash
# In a separate Terminal window
caffeinate -d

# Press Ctrl+C when done
```

## 📝 Example Real-World Usage

Converting a folder with 121 documents:
```bash
# 1. Check filenames
python3 check_filenames.py ~/Desktop/Form65s
# Found 23 files with issues

# 2. Fix the issues
python3 fix_filenames.py ~/Desktop/Form65s --fix
# Fixed 24 files (handles duplicates)

# 3. Convert
python3 pages_converter.py ~/Desktop/Form65s
# Converted 120/121 successfully in 8.5 minutes
```

## ✅ Success Indicators

You'll know it's working when you see:
- Progress bar moving in terminal
- Green checkmarks (✓) for each file
- New .docx files appearing
- 'pages' folders being created
- Log file on Desktop with details

## 🚫 What NOT to Do

- Don't run on your entire Documents folder at once
- Don't let the screen lock during conversion
- Don't use filenames with backslashes or colons
- Don't run multiple conversions simultaneously
- Don't delete log files until you've verified all conversions

## 🔧 Maintenance

### After macOS Updates
Re-run permission checker:
```bash
python3 permission_checker.py
```

### Cleaning Up Old Logs
```bash
# List all logs
ls -la ~/Desktop/pages_conversion_*.log

# Delete logs older than 30 days
find ~/Desktop -name "pages_conversion_*.log" -mtime +30 -delete
```

## 📄 License

This project is provided as-is for personal use. Feel free to modify for your needs.

## 🙏 Acknowledgments

Created as a workaround for Automator reliability issues in macOS Sequoia (15.x). Special thanks to the Python and AppleScript communities.

---

**Remember**: Always backup your important documents before bulk conversions!