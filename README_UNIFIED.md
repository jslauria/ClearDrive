# OneDrive File Organizer - UNIFIED VERSION

**All features in one application** - Scans your ENTIRE OneDrive!

Written and designed by **Jeffery Lauria**  
For support: **jeff@Brokencomputer.net**

---

## 🎯 Key Features

✅ **Full OneDrive Scan** - Recursively scans ALL folders and subfolders  
✅ **Device Code Authentication** - Simple browser-based login  
✅ **AI-Powered Categorization** - Smart file organization (optional)  
✅ **Process Downloads Folder** - Move Downloads files to root, then organize  
✅ **Archive Feature** - Auto-archive items 24+ months old  
✅ **Clean Up Empty Folders** - Remove empty folders after organizing  
✅ **Large File Support** - Handle ISO, executables, disk images  
✅ **Safe Operation** - Never deletes files, only moves  

## 🆕 What's New

### Process Downloads Folder (NEW!)
Automatically moves all files from Downloads folder to root before organizing:
- ✅ Finds all files in Downloads (including subfolders)
- ✅ Moves them to OneDrive root
- ✅ Then organizes them into proper categories
- ✅ Cleans up Downloads folder structure

### Full OneDrive Scan
Scans **your entire OneDrive**, not just root-level files:
- ✅ Scans ALL folders recursively
- ✅ Finds files buried in subdirectories
- ✅ Shows file paths during processing
- ✅ Organizes everything into clean root-level folders

### Clean Up Empty Folders
After organizing, automatically removes empty folders left behind:
- ✅ Deletes empty folders recursively
- ✅ Preserves organized category folders
- ✅ Preserves Archive folder
- ✅ Multiple passes to clean nested empty folders

## 📦 Quick Start

1. **Download all files** to one folder:
   - START_ORGANIZER.bat
   - onedrive_organizer_unified_gui.py
   - onedrive_organizer_unified.py
   - requirements.txt

2. **Double-click** `START_ORGANIZER.bat`

3. **Authenticate:**
   - Click "Connect to OneDrive"
   - Copy the device code shown
   - Visit the URL and paste code
   - Sign in with Microsoft 365 account

4. **Configure options:**
   - (Optional) Enter Anthropic API key
   - ☐ Include large files
   - ☑ **Process Downloads folder** ← NEW!
   - ☑ Archive old files (enter months)
   - ☑ Clean up empty folders

5. **Click "Start Organizing"**

## 📥 Process Downloads Folder

The Downloads processing feature runs **first**, before any other organization:

**What it does:**
1. Finds your Downloads folder
2. Scans all files (including in subfolders)
3. Moves every file to OneDrive root
4. Then proceeds with normal organization

**Example:**
```
Before Downloads Processing:
  /Downloads/
    ├── document.pdf
    ├── temp/
    │   └── invoice.pdf
    └── old/
        └── photo.jpg

After Downloads Processing (moved to root):
  /document.pdf
  /invoice.pdf
  /photo.jpg

After Full Organization:
  /Documents/document.pdf
  /Financial/invoice.pdf
  /Photos/photo.jpg
  /Downloads/  (empty - will be cleaned up if enabled)
```

## 🗂️ Archive Feature

Archives files/folders not accessed in X months (default 24) to `Archive/YEAR/`

**Archive Structure:**
```
Archive/
  ├── 2023/  ← Items last accessed in 2023
  ├── 2022/
  └── 2021/
```

## 🧹 Clean Up Empty Folders

After moving all files, the cleanup feature:

1. Scans entire OneDrive for empty folders
2. Deletes empty folders (except Archive and organized categories)
3. Repeats until no empty folders remain
4. Logs every deletion

## 🔍 How It Works - Full Workflow

### Phase 1: Downloads Processing (if enabled)
```
PROCESSING DOWNLOADS FOLDER
Found Downloads folder - scanning contents...
  Scanning: /Downloads/
  Scanning: /Downloads/temp/
  Scanning: /Downloads/old/
Found 15 files in Downloads folder

Moving files from Downloads to root...
[1/15] /Downloads/document.pdf
  ✓ Moved: document.pdf
[2/15] /Downloads/temp/invoice.pdf
  ✓ Moved: invoice.pdf
...
✓ Moved 15 files from Downloads to root
```

### Phase 2: Full OneDrive Scan
```
Scanning entire OneDrive...
  Scanning: /
  Scanning: /Projects/
  Scanning: /Documents/
...
Found 1,247 items
```

### Phase 3: Organization
```
Processing files...
[1/856] /document.pdf
  → Documents
[2/856] /invoice.pdf
  → Financial
...
```

### Phase 4: Archive (if enabled)
```
Processing folders for archiving...
Old folder (not accessed in 36 months)
  → Archive/2021/
```

### Phase 5: Cleanup (if enabled)
```
Cleaning up empty folders...
Pass 1:
  Found empty: /Downloads/temp/
  ✓ Deleted empty folder
...
```

## ⚙️ Options

**Process Downloads folder:**
- ☐ Leave Downloads alone (default)
- ☑ Move all Downloads files to root before organizing

**Include large files:**
- ☐ Skip files >10MB (default)
- ☑ Process ISO, executables, disk images

**Archive old files:**
- ☐ Organize all files normally (default)
- ☑ Move old items to Archive/YEAR/
- Enter months (12, 24, 36, etc.)

**Clean up empty folders:**
- ☐ Leave empty folders (default)
- ☑ Delete empty folders after organizing

## 📋 Requirements

- Python 3.7+
- Microsoft 365 / Business account
- Internet connection

## 📄 Files Included

- **START_ORGANIZER.bat** - Launcher
- **onedrive_organizer_unified_gui.py** - GUI
- **onedrive_organizer_unified.py** - Engine  
- **requirements.txt** - Dependencies

## 🎬 Example Complete Workflow

**Before (messy with Downloads):**
```
OneDrive/
  ├── Downloads/
  │   ├── report.pdf
  │   ├── temp/
  │   │   └── invoice.pdf
  │   └── old/
  │       └── budget.xlsx
  ├── Old Projects/
  │   └── 2020/
  │       └── photo.jpg
  └── Random/
      └── presentation.pptx
```

**After (organized with all features):**
```
OneDrive/
  ├── Archive/
  │   └── 2020/
  │       └── photo.jpg  (if old)
  ├── Documents/
  │   └── report.pdf  (from Downloads)
  ├── Financial/
  │   ├── invoice.pdf  (from Downloads/temp/)
  │   └── budget.xlsx  (from Downloads/old/)
  └── Presentations/
      └── presentation.pptx

(All empty folders removed, including Downloads)
```

## 🛡️ Safety

- **Files never deleted** - only moved
- **Empty folders removed** - if cleanup enabled
- **Downloads preserved** - files moved, not deleted
- **Full logging** - every action recorded
- **Reversible** - manually move files back if needed

## 💡 Tips

- **Enable Downloads processing** if you accumulate files there
- **First run without cleanup** to see organization results
- **Check Activity Log** for detailed progress
- **Downloads folder will be empty** after processing (will be cleaned up if cleanup enabled)

## 🎯 What Gets Processed

✅ All files from Downloads folder (if enabled)  
✅ All files from anywhere in OneDrive  
✅ Files in subdirectories at any depth  
✅ Documents, images, videos, archives  
✅ Spreadsheets, presentations, PDFs  

## 📥 Downloads Folder Benefits

**Why process Downloads first?**
- Many files accumulate in Downloads over time
- Subfolders create unnecessary nesting
- Files are scattered and hard to find
- Moving to root allows proper categorization

**Result:**
- Clean Downloads folder
- Files properly categorized
- Easier to find later
- Organized structure maintained

---

**Written and designed by Jeffery Lauria**  
For support: **jeff@Brokencomputer.net**

This unified version scans your ENTIRE OneDrive and organizes everything!
