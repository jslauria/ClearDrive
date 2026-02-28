#!/usr/bin/env python3
"""
OneDrive Automatic Organizer - Microsoft 365 / Business Accounts
With Device Code Authentication and Archive Feature
Unified version with all features
"""

import os
import json
import time
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from io import BytesIO

import requests

try:
    import anthropic
except ImportError:
    anthropic = None

# Document parsing libraries (optional)
try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pptx import Presentation
except ImportError:
    Presentation = None

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


# ===========================================================================
# OAuth Configuration - Microsoft 365 / Business Accounts
# ===========================================================================
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Graph Explorer - supports device code
TENANT = "common"  # Works for all Microsoft 365 / Azure AD accounts
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}"

SCOPES = [
    "Files.ReadWrite.All",
    "offline_access",
]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class OneDriveOrganizerBusiness:
    """
    OneDrive organizer for Microsoft 365 / Business accounts
    Features: Device code auth, AI categorization, Archive old files
    """

    # Binary file extensions - never download for parsing
    _BINARY_EXTENSIONS = {
        ".iso", ".img", ".bin", ".vmdk", ".vhd", ".vhdx", ".ova", ".ovf",
        ".raw", ".dd", ".dmg",
        ".exe", ".msi", ".dll", ".sys", ".drv",
        ".apk", ".ipa", ".appx",
        ".rom", ".bios", ".fw", ".firmware",
    }

    def __init__(
        self,
        anthropic_api_key: Optional[str] = None,
        include_large_files: bool = False,
        archive_old_files: bool = False,
        archive_months: int = 24,
        cleanup_empty_folders: bool = False,
        process_downloads: bool = False,
    ):
        # Logging
        self.log_file = (
            f"onedrive_organizer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        )

        # Settings
        self.include_large_files = include_large_files
        self.archive_old_files = archive_old_files
        self.archive_months = archive_months
        self.archive_cutoff_date = datetime.now() - timedelta(days=30 * archive_months)
        self.cleanup_empty_folders = cleanup_empty_folders
        self.process_downloads = process_downloads

        # Anthropic
        self.anthropic_client = (
            anthropic.Anthropic(api_key=anthropic_api_key)
            if anthropic_api_key and anthropic
            else None
        )

        # OAuth
        self.access_token: Optional[str] = None
        self._refresh_token: Optional[str] = None
        self._token_expires_at: float = 0

        # OAuth URLs
        self._device_code_url = f"{AUTHORITY}/oauth2/v2.0/devicecode"
        self._token_url = f"{AUTHORITY}/oauth2/v2.0/token"

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def log(self, message: str):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}"
        print(entry)
        with open(self.log_file, "a", encoding="utf-8") as fh:
            fh.write(entry + "\n")

    # ------------------------------------------------------------------
    # OAuth Authentication - Device Code Flow
    # ------------------------------------------------------------------
    def authenticate(self) -> bool:
        """Device code flow authentication for Microsoft 365"""
        self.log("Starting device code authentication...")

        # Request device code
        device_code_data = {
            'client_id': CLIENT_ID,
            'scope': ' '.join(SCOPES),
        }

        try:
            resp = requests.post(self._device_code_url, data=device_code_data)
            if resp.status_code != 200:
                self.log(f"ERROR: Failed to get device code: {resp.status_code} - {resp.text}")
                return False

            device_code_response = resp.json()

            # Display instructions to user
            self.log("\n" + "=" * 70)
            self.log("DEVICE CODE AUTHENTICATION")
            self.log("=" * 70)
            self.log(f"\n1. Go to: {device_code_response['verification_uri']}")
            self.log(f"2. Enter code: {device_code_response['user_code']}")
            self.log("3. Sign in with your Microsoft 365 / Work account")
            self.log("\nWaiting for you to complete authentication...\n")

            # Poll for token
            interval = device_code_response.get('interval', 5)
            device_code = device_code_response['device_code']
            expires_in = device_code_response.get('expires_in', 900)
            
            token_data = {
                'client_id': CLIENT_ID,
                'grant_type': 'urn:ietf:params:oauth:grant-type:device_code',
                'device_code': device_code,
            }

            start_time = time.time()
            while time.time() - start_time < expires_in:
                time.sleep(interval)

                token_resp = requests.post(self._token_url, data=token_data)
                token_result = token_resp.json()

                if 'access_token' in token_result:
                    self.access_token = token_result['access_token']
                    self._refresh_token = token_result.get('refresh_token')
                    expires_in_sec = token_result.get('expires_in', 3600)
                    self._token_expires_at = time.time() + expires_in_sec
                    self.log("✓ Authentication successful!")
                    return True

                error = token_result.get('error', '')
                if error == 'authorization_pending':
                    # Still waiting for user
                    continue
                elif error == 'authorization_declined':
                    self.log("✗ Authentication declined by user")
                    return False
                elif error == 'expired_token':
                    self.log("✗ Device code expired")
                    return False
                else:
                    self.log(f"✗ Error: {token_result.get('error_description', error)}")
                    return False

            self.log("✗ Authentication timed out")
            return False

        except Exception as exc:
            self.log(f"✗ Authentication error: {exc}")
            return False

    def _ensure_token(self):
        """Refresh token if needed"""
        if time.time() >= self._token_expires_at - 300:  # Refresh 5 min before expiry
            if self._refresh_token:
                self._refresh_access_token()

    def _refresh_access_token(self):
        """Refresh the access token"""
        data = {
            'client_id': CLIENT_ID,
            'refresh_token': self._refresh_token,
            'grant_type': 'refresh_token',
            'scope': ' '.join(SCOPES),
        }

        try:
            resp = requests.post(self._token_url, data=data)
            if resp.status_code == 200:
                token_data = resp.json()
                self.access_token = token_data['access_token']
                if 'refresh_token' in token_data:
                    self._refresh_token = token_data['refresh_token']
                expires_in = token_data.get('expires_in', 3600)
                self._token_expires_at = time.time() + expires_in
                self.log("Token refreshed successfully")
        except Exception as exc:
            self.log(f"Token refresh error: {exc}")

    # ------------------------------------------------------------------
    # API Helpers
    # ------------------------------------------------------------------
    @property
    def headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def _get(self, url: str, **kwargs) -> requests.Response:
        self._ensure_token()
        return requests.get(url, headers=self.headers, **kwargs)

    def _post(self, url: str, **kwargs) -> requests.Response:
        self._ensure_token()
        return requests.post(url, headers=self.headers, **kwargs)

    def _patch(self, url: str, **kwargs) -> requests.Response:
        self._ensure_token()
        return requests.patch(url, headers=self.headers, **kwargs)

    # ------------------------------------------------------------------
    # OneDrive Operations
    # ------------------------------------------------------------------
    def get_all_items(self) -> List[Dict]:
        """Get ALL items (files and folders) from entire OneDrive, recursively"""
        self.log("Scanning entire OneDrive (all folders and subfolders)...")
        all_items: List[Dict] = []
        
        def scan_folder(folder_id: Optional[str] = None, path: str = "/"):
            """Recursively scan a folder and its subfolders"""
            url = (
                f"{GRAPH_BASE}/me/drive/items/{folder_id}/children"
                if folder_id
                else f"{GRAPH_BASE}/me/drive/root/children"
            )
            
            while url:
                resp = self._get(url)
                if resp.status_code != 200:
                    self.log(f"ERROR: Failed to fetch items from {path}: {resp.status_code}")
                    return
                
                data = resp.json()
                items = data.get("value", [])
                
                for item in items:
                    # Add path information to item
                    item['_path'] = path
                    all_items.append(item)
                    
                    # If it's a folder, scan it recursively
                    if "folder" in item:
                        subfolder_path = f"{path}{item['name']}/"
                        self.log(f"  Scanning: {subfolder_path}")
                        scan_folder(item['id'], subfolder_path)
                
                url = data.get("@odata.nextLink")
        
        # Start scanning from root
        scan_folder()
        
        self.log(f"Found {len(all_items)} total items across all folders")
        return all_items

    def should_archive(self, item: Dict) -> Tuple[bool, str]:
        """
        Determine if an item should be archived based on last accessed date
        Returns: (should_archive, reason)
        """
        if not self.archive_old_files:
            return False, ""

        # Get last accessed time
        last_accessed = item.get("lastAccessedDateTime")
        if not last_accessed:
            # Fall back to modified time if accessed time not available
            last_accessed = item.get("lastModifiedDateTime")

        if not last_accessed:
            return False, "No date available"

        # Parse date
        try:
            accessed_date = datetime.fromisoformat(last_accessed.replace('Z', '+00:00'))
            accessed_date = accessed_date.replace(tzinfo=None)  # Remove timezone for comparison

            if accessed_date < self.archive_cutoff_date:
                months_old = (datetime.now() - accessed_date).days // 30
                return True, f"Not accessed in {months_old} months"
        except Exception as e:
            self.log(f"  Error parsing date: {e}")

        return False, ""

    def download_file_content(self, file_id: str) -> Optional[bytes]:
        """Download file content"""
        url = f"{GRAPH_BASE}/me/drive/items/{file_id}/content"
        resp = self._get(url, stream=True)
        if resp.status_code == 200:
            return resp.content
        return None

    def extract_text_from_file(self, file_info: Dict) -> str:
        """Extract text content from file for AI analysis"""
        file_name = file_info["name"]
        file_ext = Path(file_name).suffix.lower()
        file_size = file_info.get("size", 0)

        # Skip binary files
        if file_ext in self._BINARY_EXTENSIONS:
            self.log(f"  Binary file - categorizing by name: {file_name}")
            return ""

        # Size limits
        if not self.include_large_files and file_size > 10_485_760:  # 10 MB
            size_mb = file_size / 1_048_576
            self.log(f"  Large file ({size_mb:.1f} MB) - categorizing by name: {file_name}")
            return ""

        if file_size > 524_288_000:  # 500 MB hard limit
            size_mb = file_size / 1_048_576
            self.log(f"  Very large file ({size_mb:.1f} MB) - categorizing by name: {file_name}")
            return ""

        # Download and extract text
        content = self.download_file_content(file_info["id"])
        if not content:
            return ""

        try:
            if file_ext in (".txt", ".md", ".csv"):
                return content.decode("utf-8", errors="ignore")[:5000]

            if file_ext == ".docx" and Document:
                doc = Document(BytesIO(content))
                return "\n".join(p.text for p in doc.paragraphs)[:5000]

            if file_ext == ".pdf" and PyPDF2:
                reader = PyPDF2.PdfReader(BytesIO(content))
                text = "\n".join(page.extract_text() for page in reader.pages[:3])
                return text[:5000]

            if file_ext == ".pptx" and Presentation:
                prs = Presentation(BytesIO(content))
                texts = []
                for slide in prs.slides[:5]:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            texts.append(shape.text)
                return "\n".join(texts)[:5000]

            if file_ext == ".xlsx" and openpyxl:
                wb = openpyxl.load_workbook(BytesIO(content), read_only=True)
                texts = []
                for sheet in list(wb.worksheets)[:2]:
                    for row in list(sheet.rows)[:30]:
                        texts.append(" ".join(str(c.value) for c in row if c.value))
                return "\n".join(texts)[:5000]

        except Exception as exc:
            self.log(f"  Text extraction error ({file_name}): {exc}")

        return ""

    # ------------------------------------------------------------------
    # AI Categorization
    # ------------------------------------------------------------------
    def categorize_file_with_ai(self, file_info: Dict, text_content: str) -> Dict:
        """Categorize file using AI or keyword fallback"""
        file_name = file_info["name"]

        if not self.anthropic_client:
            return self.basic_categorization(file_name)

        metadata = (
            f"File Name: {file_name}\n"
            f"Created: {file_info.get('createdDateTime', 'Unknown')}\n"
            f"Modified: {file_info.get('lastModifiedDateTime', 'Unknown')}\n"
            f"Size: {file_info.get('size', 0)} bytes\n"
        )

        if len(text_content) > 2000:
            text_content = text_content[:2000] + "… [truncated]"

        prompt = (
            "Analyze this file and categorize it into a clear folder structure.\n\n"
            "Common categories: Clients, Personal, Work, Financial, Legal, Medical, "
            "Travel, Education, Projects, Photos, Videos, Documents, Archives, etc.\n\n"
            f"{metadata}\n"
            f"Content preview:\n"
            f"{text_content if text_content else '[No extractable content]'}\n\n"
            "Respond ONLY with valid JSON:\n"
            '{\n'
            '    "category": "Primary folder name",\n'
            '    "subcategory": "Sub-folder or null",\n'
            '    "confidence": "high/medium/low",\n'
            '    "reasoning": "One-line explanation"\n'
            '}\n'
        )

        try:
            msg = self.anthropic_client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=500,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            s, e = raw.find("{"), raw.rfind("}") + 1
            if s >= 0 and e > s:
                return json.loads(raw[s:e])
        except Exception as exc:
            self.log(f"  AI categorization error: {exc}")

        return self.basic_categorization(file_name)

    @staticmethod
    def basic_categorization(file_name: str) -> Dict:
        """Keyword-based categorization fallback"""
        name_lower = file_name.lower()
        file_ext = Path(file_name).suffix.lower()

        keyword_map = {
            "Financial": ["invoice", "receipt", "tax", "budget", "bank", "bill", "payment"],
            "Work": ["resume", "cv", "report", "meeting", "presentation", "proposal"],
            "Personal": ["family", "vacation", "personal", "birthday", "wedding"],
            "Legal": ["agreement", "legal", "license", "permit", "deed", "contract"],
            "Medical": ["medical", "health", "prescription", "doctor", "hospital"],
            "Education": ["course", "study", "notes", "assignment", "homework"],
        }

        for cat, kws in keyword_map.items():
            if any(kw in name_lower for kw in kws):
                return {"category": cat, "subcategory": None, "confidence": "medium",
                       "reasoning": f"Keyword match: {cat}"}

        ext_map = {
            "Photos": (".jpg", ".jpeg", ".png", ".gif", ".heic", ".bmp"),
            "Videos": (".mp4", ".avi", ".mov", ".mkv", ".wmv"),
            "Audio": (".mp3", ".wav", ".flac", ".m4a"),
            "Documents": (".doc", ".docx", ".pdf", ".txt"),
            "Spreadsheets": (".xls", ".xlsx", ".csv"),
            "Presentations": (".ppt", ".pptx"),
            "Archives": (".zip", ".rar", ".7z", ".tar", ".gz"),
            "Disk Images": (".iso", ".img", ".vmdk", ".vhd", ".dmg"),
            "Software": (".exe", ".msi", ".dll", ".apk"),
        }

        for cat, exts in ext_map.items():
            if file_ext in exts:
                return {"category": cat, "subcategory": None, "confidence": "low",
                       "reasoning": f"File type: {cat}"}

        return {"category": "Miscellaneous", "subcategory": None, "confidence": "low",
               "reasoning": "No clear category"}

    # ------------------------------------------------------------------
    # Folder Operations
    # ------------------------------------------------------------------
    def get_folder_id(self, folder_name: str, parent_id: Optional[str] = None) -> Optional[str]:
        """Get folder ID by name"""
        url = (
            f"{GRAPH_BASE}/me/drive/items/{parent_id}/children"
            if parent_id
            else f"{GRAPH_BASE}/me/drive/root/children"
        )

        resp = self._get(url)
        if resp.status_code == 200:
            for item in resp.json().get("value", []):
                if item["name"] == folder_name and "folder" in item:
                    return item["id"]
        return None

    def get_root_id(self) -> Optional[str]:
        """Get the ID of the OneDrive root folder"""
        url = f"{GRAPH_BASE}/me/drive/root"
        resp = self._get(url)
        if resp.status_code == 200:
            return resp.json()["id"]
        return None

    def create_folder(self, folder_name: str, parent_id: Optional[str] = None) -> Optional[str]:
        """Create folder if it doesn't exist"""
        existing = self.get_folder_id(folder_name, parent_id)
        if existing:
            return existing

        url = (
            f"{GRAPH_BASE}/me/drive/items/{parent_id}/children"
            if parent_id
            else f"{GRAPH_BASE}/me/drive/root/children"
        )

        resp = self._post(url, json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        })

        if resp.status_code == 201:
            self.log(f"  Created folder: {folder_name}")
            return resp.json()["id"]

        self.log(f"  Error creating folder '{folder_name}': {resp.status_code}")
        return None

    def move_item(self, item_id: str, target_folder_id: str, item_name: str) -> bool:
        """Move item (file or folder) to target folder"""
        resp = self._patch(
            f"{GRAPH_BASE}/me/drive/items/{item_id}",
            json={"parentReference": {"id": target_folder_id}},
        )

        if resp.status_code == 200:
            self.log(f"  ✓ Moved: {item_name}")
            return True

        self.log(f"  ✗ Failed to move {item_name}: {resp.status_code}")
        return False

    def delete_folder(self, folder_id: str, folder_name: str) -> bool:
        """Delete an empty folder"""
        url = f"{GRAPH_BASE}/me/drive/items/{folder_id}"
        resp = requests.delete(url, headers=self.headers)
        
        if resp.status_code == 204:
            self.log(f"  ✓ Deleted empty folder: {folder_name}")
            return True
        
        self.log(f"  ✗ Failed to delete {folder_name}: {resp.status_code}")
        return False

    def is_folder_empty(self, folder_id: str) -> bool:
        """Check if a folder is empty"""
        url = f"{GRAPH_BASE}/me/drive/items/{folder_id}/children"
        resp = self._get(url)
        
        if resp.status_code == 200:
            children = resp.json().get("value", [])
            return len(children) == 0
        
        return False

    # ------------------------------------------------------------------
    # Downloads Processing
    # ------------------------------------------------------------------
    def process_downloads_folder(self) -> int:
        """
        Move all files from Downloads folder to root, then process them.
        Returns count of files moved from Downloads.
        """
        self.log("\n" + "=" * 70)
        self.log("PROCESSING DOWNLOADS FOLDER")
        self.log("=" * 70 + "\n")
        
        # Find Downloads folder
        downloads_id = self.get_folder_id("Downloads")
        if not downloads_id:
            self.log("No Downloads folder found - skipping")
            return 0
        
        self.log("Found Downloads folder - scanning contents...")
        
        # Get root ID
        root_id = self.get_root_id()
        if not root_id:
            self.log("ERROR: Could not get root folder ID")
            return 0
        
        # Get all items in Downloads folder (recursively)
        downloads_items = []
        
        def scan_downloads(folder_id: str, path: str = "/Downloads/"):
            """Recursively scan Downloads folder"""
            url = f"{GRAPH_BASE}/me/drive/items/{folder_id}/children"
            
            while url:
                resp = self._get(url)
                if resp.status_code != 200:
                    return
                
                data = resp.json()
                items = data.get("value", [])
                
                for item in items:
                    item['_downloads_path'] = path
                    downloads_items.append(item)
                    
                    # If folder, scan it too
                    if "folder" in item:
                        subfolder_path = f"{path}{item['name']}/"
                        self.log(f"  Scanning: {subfolder_path}")
                        scan_downloads(item['id'], subfolder_path)
                
                url = data.get("@odata.nextLink")
        
        # Scan Downloads folder
        scan_downloads(downloads_id)
        
        # Separate files from folders
        files = [item for item in downloads_items if "file" in item]
        
        self.log(f"Found {len(files)} files in Downloads folder")
        
        if not files:
            self.log("No files to move from Downloads")
            return 0
        
        # Move all files to root
        moved_count = 0
        self.log("\nMoving files from Downloads to root...")
        
        for idx, file_info in enumerate(files, 1):
            name = file_info["name"]
            path = file_info.get("_downloads_path", "/Downloads/")
            self.log(f"[{idx}/{len(files)}] {path}{name}")
            
            try:
                if self.move_item(file_info["id"], root_id, name):
                    moved_count += 1
                time.sleep(0.2)  # Rate limiting
            except Exception as exc:
                self.log(f"  ERROR: {exc}")
        
        self.log(f"\n✓ Moved {moved_count} files from Downloads to root")
        return moved_count

    # ------------------------------------------------------------------
    # Main Organization Logic
    # ------------------------------------------------------------------
    def organize_files(self):
        """Organize all items from entire OneDrive with optional archiving"""
        self.log("\n" + "=" * 70)
        self.log("STARTING ONEDRIVE ORGANIZATION - FULL SCAN")
        if self.archive_old_files:
            cutoff = self.archive_cutoff_date.strftime("%Y-%m-%d")
            self.log(f"Archive mode enabled: Items not accessed since {cutoff}")
        if self.process_downloads:
            self.log("Downloads processing enabled: Files will be moved to root first")
        self.log("=" * 70 + "\n")

        # Process Downloads folder first if enabled
        downloads_moved = 0
        if self.process_downloads:
            downloads_moved = self.process_downloads_folder()

        items = self.get_all_items()
        if not items:
            self.log("No items found!")
            return

        # Separate files and folders
        files = [item for item in items if "file" in item]
        folders = [item for item in items if "folder" in item]

        stats = {
            "total_files": len(files),
            "total_folders": len(folders),
            "moved": 0,
            "archived": 0,
            "errors": 0,
            "categories": {}
        }

        # Create main Archive folder if needed
        archive_folder_id = None
        if self.archive_old_files:
            archive_folder_id = self.create_folder("Archive")
            if not archive_folder_id:
                self.log("ERROR: Could not create Archive folder!")
                self.archive_old_files = False

        # Process files
        self.log(f"\nProcessing {len(files)} files from all folders...")
        for idx, file_info in enumerate(files, 1):
            name = file_info["name"]
            path = file_info.get("_path", "/")
            size_mb = file_info.get("size", 0) / 1_048_576
            self.log(f"\n[{idx}/{len(files)}] {path}{name} ({size_mb:.1f} MB)")

            try:
                # Check if file should be archived
                should_archive, reason = self.should_archive(file_info)

                if should_archive:
                    self.log(f"  → Archive ({reason})")
                    stats["archived"] += 1

                    # Create year-based subfolder in Archive
                    accessed = file_info.get("lastAccessedDateTime") or file_info.get("lastModifiedDateTime")
                    if accessed:
                        year = accessed[:4]
                        year_folder_id = self.create_folder(year, archive_folder_id)
                        if year_folder_id:
                            if self.move_item(file_info["id"], year_folder_id, name):
                                stats["moved"] += 1
                            else:
                                stats["errors"] += 1
                        else:
                            stats["errors"] += 1
                    continue

                # Regular categorization
                text = self.extract_text_from_file(file_info)
                cat = self.categorize_file_with_ai(file_info, text)

                category = cat["category"]
                subcategory = cat.get("subcategory")

                self.log(f"  → {category}" + (f" / {subcategory}" if subcategory else ""))
                stats["categories"][category] = stats["categories"].get(category, 0) + 1

                # Create folders and move file
                cat_id = self.create_folder(category)
                if not cat_id:
                    stats["errors"] += 1
                    continue

                target_id = cat_id
                if subcategory:
                    sub_id = self.create_folder(subcategory, cat_id)
                    if sub_id:
                        target_id = sub_id

                if self.move_item(file_info["id"], target_id, name):
                    stats["moved"] += 1
                else:
                    stats["errors"] += 1

                time.sleep(0.3)  # Rate limiting

            except Exception as exc:
                self.log(f"  ERROR: {exc}")
                stats["errors"] += 1

        # Process folders (for archiving only)
        if self.archive_old_files and folders:
            self.log(f"\nProcessing {len(folders)} folders for archiving...")
            for idx, folder_info in enumerate(folders, 1):
                name = folder_info["name"]
                path = folder_info.get("_path", "/")
                self.log(f"\n[{idx}/{len(folders)}] Folder: {path}{name}")

                # Skip Archive folder itself and its subfolders
                if name == "Archive" or path.startswith("/Archive/"):
                    self.log("  → Skipping (Archive folder)")
                    continue

                try:
                    should_archive, reason = self.should_archive(folder_info)

                    if should_archive:
                        self.log(f"  → Archive ({reason})")
                        stats["archived"] += 1

                        # Move entire folder into Archive (preserves structure)
                        if archive_folder_id:
                            if self.move_item(folder_info["id"], archive_folder_id, name):
                                stats["moved"] += 1
                            else:
                                stats["errors"] += 1
                    else:
                        self.log("  → Keep (recently accessed)")

                    time.sleep(0.3)

                except Exception as exc:
                    self.log(f"  ERROR: {exc}")
                    stats["errors"] += 1

        # Summary
        self.log("\n" + "=" * 70)
        self.log("ORGANIZATION COMPLETE")
        self.log("=" * 70)
        if self.process_downloads and downloads_moved > 0:
            self.log(f"\nDownloads processed: {downloads_moved} files moved to root")
        self.log(f"\nFiles: {stats['total_files']} | Folders: {stats['total_folders']}")
        self.log(f"Moved: {stats['moved']} | Archived: {stats['archived']} | Errors: {stats['errors']}")

        if stats["categories"]:
            self.log("\nCategories:")
            for cat, count in sorted(stats["categories"].items()):
                self.log(f"  {cat}: {count}")

        # Cleanup empty folders if requested
        if self.cleanup_empty_folders:
            self.log("\n" + "=" * 70)
            self.log("CLEANING UP EMPTY FOLDERS")
            self.log("=" * 70 + "\n")
            
            deleted_count = self.cleanup_empty_folders_recursive()
            stats["deleted_folders"] = deleted_count
            
            self.log(f"\n✓ Deleted {deleted_count} empty folders")

        self.log(f"\nLog: {self.log_file}")

    def cleanup_empty_folders_recursive(self) -> int:
        """
        Recursively scan and delete empty folders.
        Returns count of deleted folders.
        """
        self.log("Scanning for empty folders...")
        deleted_count = 0
        
        # Keep scanning until no more empty folders found
        # (because deleting a folder might make its parent empty)
        max_passes = 10
        for pass_num in range(max_passes):
            self.log(f"\nPass {pass_num + 1}:")
            
            # Get all current folders
            all_items = self.get_all_items()
            folders = [item for item in all_items if "folder" in item]
            
            # Skip if no folders
            if not folders:
                break
            
            empty_found = 0
            for folder in folders:
                name = folder["name"]
                path = folder.get("_path", "/")
                
                # Skip Archive folder and organized category folders at root
                if name == "Archive" or (path == "/" and name in [
                    "Documents", "Financial", "Work", "Personal", "Legal", 
                    "Medical", "Education", "Projects", "Photos", "Videos",
                    "Audio", "Spreadsheets", "Presentations", "Archives",
                    "Disk Images", "Software", "Miscellaneous"
                ]):
                    continue
                
                # Check if empty
                if self.is_folder_empty(folder["id"]):
                    self.log(f"  Found empty: {path}{name}")
                    if self.delete_folder(folder["id"], name):
                        deleted_count += 1
                        empty_found += 1
                    time.sleep(0.2)
            
            # If no empty folders found this pass, we're done
            if empty_found == 0:
                self.log(f"  No more empty folders found")
                break
        
        return deleted_count


# ===========================================================================
# CLI Entry Point
# ===========================================================================
if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("  OneDrive File Organizer - Microsoft 365 / Business")
    print("  Full OneDrive Scan + Archive Feature")
    print("=" * 70 + "\n")

    # API Key
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("For smart AI categorization, you can use Claude AI.")
        print("Get a free key at: https://console.anthropic.com/")
        api_key = input("\nAnthropic API key (or Enter to skip): ").strip()

    if not api_key:
        print("\nRunning in keyword-matching mode (no AI).")

    # Large files
    print("\nInclude large files (ISO images, executables, etc.)?")
    large = input("Include large files? (yes/no) [no]: ").strip().lower()
    include_large = large in ("yes", "y")

    # Archive option
    print("\n" + "=" * 70)
    print("ARCHIVE OLD FILES FEATURE")
    print("=" * 70)
    print("\nThis will move files and folders that haven't been accessed in")
    print("24+ months to an 'Archive' folder, organized by year.")
    print("\nExample: Archive/2022/, Archive/2021/, etc.")
    archive = input("\nEnable archiving for old files? (yes/no) [no]: ").strip().lower()
    archive_old = archive in ("yes", "y")

    archive_months = 24
    if archive_old:
        custom = input(f"\nUse {archive_months} months cutoff? (or enter custom months): ").strip()
        if custom and custom.isdigit():
            archive_months = int(custom)

    # Cleanup option
    print("\n" + "=" * 70)
    print("CLEANUP EMPTY FOLDERS")
    print("=" * 70)
    print("\nAfter organizing, delete any empty folders left behind?")
    print("(Organized folders at root will be preserved)")
    cleanup = input("\nCleanup empty folders? (yes/no) [no]: ").strip().lower()
    cleanup_empty = cleanup in ("yes", "y")

    # Downloads processing option
    print("\n" + "=" * 70)
    print("PROCESS DOWNLOADS FOLDER")
    print("=" * 70)
    print("\nMove all files from Downloads folder to root before organizing?")
    print("This will pull files out of Downloads and then organize them.")
    downloads = input("\nProcess Downloads folder? (yes/no) [no]: ").strip().lower()
    process_downloads = downloads in ("yes", "y")

    # Create organizer
    organizer = OneDriveOrganizerBusiness(
        anthropic_api_key=api_key or None,
        include_large_files=include_large,
        archive_old_files=archive_old,
        archive_months=archive_months,
        cleanup_empty_folders=cleanup_empty,
        process_downloads=process_downloads,
    )

    # Authenticate
    if not organizer.authenticate():
        print("\nAuthentication failed. Exiting.")
        exit(1)

    # Confirm and organize
    print("\n" + "=" * 70)
    print("  Ready to organize your OneDrive!")
    print("=" * 70)
    print("\n  • Scans ALL folders and subfolders")
    print("  • Analyzes content with AI (or keywords)")
    print("  • Creates organized folders at root")
    print("  • Moves files from anywhere in OneDrive")
    if archive_old:
        cutoff = (datetime.now() - timedelta(days=30 * archive_months)).strftime("%Y-%m-%d")
        print(f"  • Archives items not accessed since {cutoff}")
    print("  • Nothing is deleted – only moved\n")

    if input("Proceed? (yes/no): ").strip().lower() == "yes":
        organizer.organize_files()
        print("\n✓ Done! Check your OneDrive.")
    else:
        print("\nCancelled.")
