## OneDrive File Organizer (Microsoft 365 / Business)

**OneDrive File Organizer** is a desktop GUI application designed to automatically analyze, organize, and maintain Microsoft 365 OneDrive environments. Built for business users and administrators, the tool performs a full OneDrive scan and intelligently restructures files into organized folders while preserving data integrity.

The application supports both AI assisted classification and traditional keyword based organization, enabling flexible deployment across enterprise and personal Microsoft 365 tenants.

---

### 🚀 Key Features

**Microsoft 365 Device Code Authentication**
Secure sign in using Microsoft device code authentication compatible with business tenants and modern security controls.

**Full OneDrive Organization**
• Recursively scans all folders and subfolders
• Categorizes and reorganizes files at the root level
• Moves files safely without deletion

**AI Assisted File Classification (Optional)**
Integrates with Anthropic APIs for intelligent file categorization.
If no API key is provided, the application defaults to keyword based sorting.

**Archive Mode**
Automatically moves inactive files into an Archive folder structured by year based on last access date.

**Downloads Folder Processing**
Optionally relocates files from the Downloads directory before organization to prevent clutter accumulation.

**Cleanup Automation**
Removes empty folders created during restructuring while preserving organized directories.

**Large File Handling**
Optional inclusion of disk images, executables, and other large file types.

**Unified Desktop GUI**
Built with Tkinter for a simple, guided workflow:

1. Configure AI (optional)
2. Authenticate with OneDrive
3. Select organization options
4. Run automated organization

**Persistent Configuration**
User preferences and settings are securely saved locally for reuse.

---

### 🧠 Typical Use Cases

• Microsoft 365 tenant cleanup and restructuring
• End user OneDrive organization
• Post migration file normalization
• Governance and data hygiene initiatives
• MSP or IT admin remediation workflows

---

### ⚙️ Technology Stack

Python 3
Tkinter GUI Framework
Microsoft Graph Authentication (Device Code Flow)
Threaded background processing
Optional AI classification via Anthropic API

---

### 🔐 Safety Design

The organizer:
• Moves files only, never deletes data
• Requires explicit user confirmation before execution
• Provides real time activity logging
• Maintains visibility throughout execution

---

### 📦 Author

Developed by **Jeffery Lauria**
jeff@brokencomputer.net
Release: February 2025

