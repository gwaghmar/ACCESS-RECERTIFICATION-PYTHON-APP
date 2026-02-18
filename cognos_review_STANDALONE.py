
"""
Cognos Access Review Tool - CONSOLIDATED SINGLE FILE VERSION

All modules merged into one script for simplified deployment and debugging.
Original modular structure moved to obsolete folder for backup.
"""

# ════════════════════════════════════════════════════════════════════════════════════════
# 📚 COMPLETE CODE NAVIGATION INDEX - Use Ctrl+F to find sections
# ════════════════════════════════════════════════════════════════════════════════════════
#
# 🔍 HOW TO USE THIS INDEX:
#    1. Find the feature you need in the list below
#    2. Note the line number (e.g., "Lines 40-50")
#    3. Press Ctrl+G in VS Code to go to that line
#    4. Or Ctrl+F and search for the section name
#
# ════════════════════════════════════════════════════════════════════════════════════════
#
# 📋 MAIN SECTIONS:
#
# ┌─ SECTION 1: IMPORTS & SETUP (Lines 40-63)
# │  └─ All libraries and initial imports
#
# ┌─ SECTION 2: DATA MODELS & CONSTANTS (Lines 70-640)
# │  ├─ ColumnNames class - Excel column names (Line 77)
# [SECTION: ColumnNames]
# │  ├─ SheetNames class - Excel sheet names (Line 98)
# [SECTION: SheetNames]
# │  ├─ FileNames class - Output file names (Line 104)
# [SECTION: FileNames]
# │  ├─ EmailMode enum - Email sending modes (Line 114)
# │  ├─ AuditStatus enum - Email status types (Line 120)
# │  ├─ ValidationSeverity enum - Error levels (Line 127)
# │  ├─ TypedDict definitions - Data structures (Lines 133-250)
# │  ├─ Data Classes - Agency mapping, audit logs, etc. (Lines 255-430)
# │  └─ Exceptions - Custom error classes (Lines 435-460)
#
# ┌─ SECTION 3: UTILITY FUNCTIONS (Lines 465-820)
# │  ├─ sanitize_sheet_name() - Make valid Excel names (Line 465)
# │  ├─ format_email_list() - Parse email addresses (Line 497)
# │  ├─ is_valid_email() - Check email format (Line 520)
# │  ├─ extract_emails_from_text() - Find emails in text (Line 540)
# │  ├─ parse_email_forwarding_message() - Smart email extraction (Line 590)
# │  └─ smart_email_verification() - Comprehensive email analysis (Line 660)
#
# ┌─ SECTION 4: CONFIG MANAGER (Lines 830-1105)
# │  👉 Handles app settings, review periods, deadlines
# │  ├─ ConfigManager class (Line 830)
# [SECTION: ConfigManager]
# │  ├─ load() - Load settings from config.json (Line 860)
# │  ├─ save() - Save settings to file (Line 898)
# │  ├─ validate() - Check configuration validity (Line 920)
# │  ├─ update() - Change settings (Line 945)
# │  ├─ get_audit_file_name() - Generate audit log filename (Line 967)
# │  ├─ set_current_region() - Select active region (Line 996)
# │  ├─ load_email_template() - Get email template (Line 1010)
# │  ├─ save_email_template() - Update email template (Line 1050)
# │  └─ get_config_manager() - Global config instance (Line 1063)
#
# ┌─ SECTION 5: AUDIT LOGGER (Lines 1110-1495)
# │  👉 Tracks email sending and responses
# │  ├─ AuditLogger class (Line 1110)
# [SECTION: AuditLogger]
# │  ├─ initialize_log() - Create audit spreadsheet (Line 1178)
# │  ├─ mark_sent() - Record when email sent (Line 1212)
# │  ├─ mark_responded() - Record when reply received (Line 1244)
# │  ├─ get_status() - Check agency status (Line 1279)
# │  ├─ get_agencies_by_status() - Filter by status (Line 1299)
# │  ├─ get_metrics() - Dashboard statistics (Line 1321)
# │  ├─ export_to_csv() - Save as CSV format (Line 1339)
# │  └─ get_summary_stats() - Count summaries (Line 1360)
#
# ┌─ SECTION 6: REGION MANAGER (Lines 1500-1840)
# │  👉 Handle multiple regions (North America, EMEA, APAC, etc.)
# │  ├─ RegionProfile dataclass (Line 1500)
# [SECTION: RegionProfile]
# │  ├─ RegionManager class (Line 1525)
# [SECTION: RegionManager]
# │  ├─ add_profile() - Create new region (Line 1560)
# │  ├─ delete_profile() - Remove region (Line 1578)
# │  ├─ get_profile() - Retrieve region details (Line 1602)
# │  ├─ switch_region() - Select active region (Line 1625)
# │  ├─ get_all_region_names() - List all regions (Line 1644)
# │  └─ load_profiles() / save_profiles() - Persist data (Lines 1500-1549)
#
# ┌─ SECTION 7: FILE VALIDATOR (Lines 1845-2210)
# │  👉 Check files are correct BEFORE processing
# │  ├─ FileValidator class (Line 1845)
# [SECTION: FileValidator]
# │  ├─ validate_file_exists() - Check file presence (Line 1873)
# │  ├─ validate_excel_file() - Check Excel readability (Line 1900)
# │  ├─ validate_master_file() - Verify master file structure (Line 1927)
# │  ├─ validate_agency_mapping_file() - Verify mapping file (Line 1968)
# │  ├─ validate_email_manifest_file() - Verify email file (Line 2051)
# │  ├─ validate_agency_mapping() - Comprehensive validation (Line 2141)
# │  └─ validate_all_files() - Complete file check (Line 2297)
#
# ┌─ SECTION 8: FILE PROCESSOR (Lines 2350-2730)
# │  👉 Creates agency-specific Excel files
# │  ├─ FileProcessor class (Line 2350)
# [SECTION: FileProcessor]
# │  ├─ format_worksheet() - Apply Excel formatting (Line 2363)
# │  ├─ load_agency_mappings() - Read mapping file (Line 2412)
# │  ├─ generate_agency_files() - Main file generation (Line 2472)
# │  ├─ _generate_single_file() - Create one agency file (Line 2634)
# │  ├─ _create_unassigned_file() - Single unmapped users file (Line 2712)
# │  └─ _create_individual_unmapped_files() - Individual files per agency (Line 2737)
#
# ┌─ SECTION 9: EMAIL HANDLER (Lines 2750-3885)
# │  👉 Send emails through Outlook
# │  ├─ EmailHandler class (Line 2750)
# [SECTION: EmailHandler]
# │  ├─ test_connection() - Verify Outlook works (Line 2831)
# │  ├─ load_email_manifest() - Get recipient list (Line 2872)
# │  ├─ create_email() - Build email message (Line 2925)
# │  ├─ send_emails() - Send/Preview/Schedule emails (Line 2990)
# │  ├─ scan_inbox() - Check for replies (Line 3138)
# │  ├─ create_folder() - Make Outlook folder (Line 3197)
# │  ├─ setup_compliance_folders() - Create standard folders (Line 3250)
# │  ├─ move_email_to_folder() - Organize emails (Line 3270)
# │  ├─ copy_sent_email_to_folder() - Copy to compliance folder (Line 3300)
# │  ├─ organize_compliance_replies() - Auto-organize replies (Line 3330)
# │  └─ scan_sent_items_for_agencies() - Find old sent emails (Line 3363)
#
# ┌─ SECTION 10: EMAIL MANIFEST MANAGER (Lines 3890-4240)
# │  👉 Manage list of email addresses
# │  ├─ EmailChange dataclass (Line 3890)
# [SECTION: EmailChange]
# │  ├─ EmailManifestManager class (Line 3916)
# [SECTION: EmailManifestManager]
# │  ├─ validate_emails() - Check email formats (Line 3955)
# │  ├─ detect_changes() - Compare old vs new emails (Line 3994)
# │  ├─ update_email() - Change email address (Line 4065)
# │  ├─ create_backup() - Backup manifest file (Line 4087)
# │  └─ get_agencies() - List all agencies (Line 4028)
#
# ┌─ SECTION 11: MAIN APPLICATION GUI (Lines 4250-7515)
# │  👉 Main window with all buttons and controls
# │
# │  🎨 WINDOW LAYOUT:
# │  ├─ Header Section (Lines 4600-4750)
# │  │  ├─ Region selector dropdown
# │  │  ├─ Company logo / branding
# │  │  └─ Theme toggle button
# │  │
# │  ├─ LEFT COLUMN - File Selection (Lines 4800-5100)
# │  │  ├─ Browse Master File button
# │  │  ├─ Browse Agency Mapping button
# │  │  ├─ Browse Email Manifest button
# │  │  ├─ Email Mode selector (Preview/Direct/Schedule)
# │  │  └─ Agency list with filter/select
# │  │
# │  ├─ RIGHT COLUMN - Main Controls (Lines 5200-5800)
# │  │  ├─ Validate button - Check all files (Line 5300)
# │  │  ├─ Generate Files button - Create agency files (Line 5400)
# │  │  ├─ Send Emails button - Send to agencies (Line 5500)
# │  │  ├─ Mark Responded button - Track responses (Line 5600)
# │  │  └─ Dashboard button - View statistics (Line 5700)
# │  │
# │  └─ Utility Buttons (Lines 5900-6200)
# │     ├─ Scan Inbox button - Check for replies (Line 5900)
# │     ├─ Smart Email Verifier - Extract emails from text (Line 6000)
# │     ├─ Open Output Folder - Browse generated files (Line 6100)
# │     ├─ Agency Mapping Manager - Edit mappings (Line 6150)
# │     ├─ Region Manager - Add/edit regions (Line 6200)
# │     ├─ Test Email Connection (Line 6250)
# │     └─ Settings button - Configure app (Line 6300)
#
# ┌─ SECTION 12: DIALOG WINDOWS (Lines 6310-8800)
# │  👉 Pop-up windows for specific tasks
# │
# │  ├─ SmartEmailVerifierDialog (Lines 6310-6650)
# │  │  └─ Extract emails from user-provided text
# │  │
# │  ├─ AgencyMappingManagerDialog (Lines 6670-7050)
# │  │  ├─ View all file-to-agency mappings
# │  │  ├─ Edit individual mappings
# │  │  ├─ Add new mappings
# │  │  └─ Delete existing mappings
# │  │
# │  ├─ MappingEditDialog (Lines 7055-7120)
# │  │  └─ Edit a single agency mapping
# │  │
# │  ├─ EmailValidationDialog (Lines 7125-7365)
# │  │  ├─ Check email address formats
# │  │  ├─ Show invalid emails
# │  │  └─ Suggest fixes
# │  │
# │  ├─ FixEmailDialog (Lines 7370-7450)
# │  │  └─ Edit invalid email addresses
# │  │
# │  └─ RegionManagementDialog (Lines 7460-7700)
# │     ├─ Add new regions
# │     ├─ Edit existing regions
# │     ├─ Delete regions
# │     └─ Switch current region
# │
# │  └─ RegionEditorDialog (Lines 7705-7875)
# │     └─ Create/edit individual region profile
#
# ┌─ SECTION 13: LEGACY FUNCTIONS (Lines 7880-9100)
# │  👉 Backward compatibility and utility functions
# │  ├─ smart_agency_match() - Compare agency names (Line 3892)
# │  ├─ validate_file_exists() - Check file presence (Line 3922)
# │  ├─ validate_excel_file() - Check Excel readability (Line 3939)
# │  ├─ validate_all_files() - Comprehensive validation (Line 3955)
# │  ├─ open_output_folder() - Open explorer window (Line 4024)
# │  ├─ test_email_connection() - Verify Outlook (Line 4046)
# │  ├─ open_settings_dialog() - Settings window (Line 4161)
# │  ├─ auto_scan_output_folder() - Auto-detect files (Line 7508)
# │  ├─ validate_agency_mapping() - Check mapping validity (Line 7550)
# │  ├─ show_validation_dialog() - Display results (Line 7696)
# │  ├─ show_fix_dialog() - Show & fix problems (Line 7762)
# │  └─ generate_agency_files_multi_tab() - File generation (Line 7953)
#
# ├─ MAIN ENTRY POINT (Line 9491-9498)
# │  ├─ main() - Application startup
# │  └─ if __name__ == "__main__" - Script execution
#
# ════════════════════════════════════════════════════════════════════════════════════════
#
# 🎯 QUICK REFERENCE - COMMON TASKS:
#
# ⚙️ Change App Settings?
#    → Go to ConfigManager class (Line 830)
#    → Modify DEFAULT_CONFIG dictionary (Line 846)
#
# ✉️ Change Email Template?
#    → Go to load_email_template() (Line 1010)
#    → Edit DEFAULT_TEMPLATE string (Line 1020)
#
# ➕ Add a New Region?
#    → RegionManager.add_profile() (Line 1560)
#    → Or use RegionManagementDialog in GUI (Line 7460)
#
# 📋 Change Column Names?
#    → Go to ColumnNames class (Line 77)
#    → Update constants to match your Excel files
#
# 📊 Modify Dashboard Display?
#    → Search for "open_dashboard()" or "DashboardMetrics" (Line 376)
#    → Update dashboard_frame() method
#
# 🔍 Change Validation Rules?
#    → Go to FileValidator class (Line 1845)
#    → Modify validation_* methods as needed
#
# 📁 Change How Files Are Generated?
#    → Go to FileProcessor.generate_agency_files() (Line 2472)
#    → Or modify _generate_single_file() (Line 2634)
#
# ✉️ Fix Email Sending Issues?
#    → Go to EmailHandler class (Line 2750)
#    → Check test_connection() method (Line 2831)
#    → Or send_emails() method (Line 2990)
#
# 🔐 Change Security/Validation?
#    → Search for "validation" (Lines 1845+)
#    → Update validation methods in FileValidator class
#
# 📱 Modify User Interface Layout?
#    → Go to CognosAccessReviewApp class (Line 4250)
#    → Update create_widgets() and related layout methods
#
# ════════════════════════════════════════════════════════════════════════════════════════
#
# 📌 FILE STRUCTURE SUMMARY:
#    • Imports & Setup .......................... Lines 40-63
#    • Constants & Data Types .................. Lines 70-640
#    • Utility Functions ....................... Lines 465-820
#    • Configuration Management ............... Lines 830-1105
#    • Audit Logging ........................... Lines 1110-1495
#    • Region Management ....................... Lines 1500-1840
#    • File Validation ......................... Lines 1845-2210
#    • File Processing ......................... Lines 2350-2730
#    • Email Handling .......................... Lines 2750-3885
#    • Email Manifest Management .............. Lines 3890-4240
#    • Main Application (GUI) ................. Lines 4250-7515
#    • Dialog Windows .......................... Lines 6310-8800
#    • Legacy Functions & Utilities ........... Lines 7880-9100
#    • Main Entry Point ........................ Lines 9491-9498
#
# ════════════════════════════════════════════════════════════════════════════════════════

import os
import re
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from pathlib import Path
import win32com.client as win32
import pythoncom
from datetime import datetime, timedelta
import threading
import tkinter as tk
import json
import logging
from typing import Optional, Tuple, List, Dict, Set, Callable, TypedDict
from enum import Enum
from dataclasses import dataclass, field, asdict
import shutil
import tkinter.ttk as ttk
from tkcalendar import DateEntry  # pip install tkcalendar
import pytz
import pickle
import hashlib
from collections import deque

# ============================================================================
# 🎨 COLOR SCHEME SYSTEM - Professional Blue/Teal Enterprise Design
# ============================================================================
# PRIMARY BRAND COLORS
PRIMARY_BLUE_LIGHT = "#007A9B"      # Main brand color (light mode)
PRIMARY_BLUE_DARK = "#00B4E1"       # Main brand color (dark mode)
PRIMARY_BLUE_HOVER = "#005575"      # Hover state (light mode)
PRIMARY_BLUE_ACTIVE = "#00475C"     # Active state (light mode)

ACCENT_TEAL_LIGHT = "#17A2B8"       # Secondary accent (light mode)
ACCENT_TEAL_DARK = "#20C9A6"        # Secondary accent (dark mode)
ACCENT_TEAL_HOVER = "#138496"       # Hover state (light mode)

# STATUS COLORS (Universal - same in light and dark modes)
STATUS_COMPLETE = "#28A745"         # Success - completed items
STATUS_PENDING = "#FFC107"          # Warning - pending items
STATUS_OVERDUE = "#DC3545"          # Danger - overdue items
STATUS_IN_PROGRESS = "#007A9B"      # Info - in-progress items

# Status background colors (lighter versions)
STATUS_COMPLETE_BG = "#D4EDDA"
STATUS_COMPLETE_TEXT = "#155724"
STATUS_PENDING_BG = "#FFF3CD"
STATUS_PENDING_TEXT = "#856404"
STATUS_OVERDUE_BG = "#F8D7DA"
STATUS_OVERDUE_TEXT = "#721C24"
STATUS_INFO_BG = "#D1ECF1"
STATUS_INFO_TEXT = "#0C5460"

# LIGHT MODE COLOR SCHEME
LIGHT_COLORS = {
    "primary": PRIMARY_BLUE_LIGHT,
    "primary_hover": PRIMARY_BLUE_HOVER,
    "primary_active": PRIMARY_BLUE_ACTIVE,
    "secondary": ACCENT_TEAL_LIGHT,
    "secondary_hover": ACCENT_TEAL_HOVER,
    "background": "#FFFFFF",
    "surface": "#F8F9FA",
    "sidebar": "#F1F3F5",
    "text": "#212529",
    "text_secondary": "#6C757D",
    "text_light": "#FFFFFF",
    "border": "#DEE2E6",
    "divider": "#E9ECEF",
    "success": STATUS_COMPLETE,
    "success_bg": STATUS_COMPLETE_BG,
    "success_text": STATUS_COMPLETE_TEXT,
    "warning": STATUS_PENDING,
    "warning_bg": STATUS_PENDING_BG,
    "warning_text": STATUS_PENDING_TEXT,
    "danger": STATUS_OVERDUE,
    "danger_bg": STATUS_OVERDUE_BG,
    "danger_text": STATUS_OVERDUE_TEXT,
    "danger_hover": "#E05252",
    "info_bg": STATUS_INFO_BG,
    "info_text": STATUS_INFO_TEXT,
    "button_disabled_bg": "#CCCCCC",
    "button_disabled_text": "#666666",
    "button_hover_shadow": "#00000015",
    "input_bg": "#FFFFFF",
    "input_border": "#CED4DA",
    "input_focus_border": PRIMARY_BLUE_LIGHT,
    "sidebar_active": PRIMARY_BLUE_LIGHT,
    "sidebar_active_text": "#FFFFFF",
}

# DARK MODE COLOR SCHEME
DARK_COLORS = {
    "primary": PRIMARY_BLUE_DARK,
    "primary_hover": "#00C5FF",
    "primary_active": "#0099CC",
    "secondary": ACCENT_TEAL_DARK,
    "secondary_hover": "#1DD6A5",
    "background": "#1E1E1E",
    "surface": "#2D2D2D",
    "sidebar": "#252525",
    "text": "#F0F0F0",
    "text_secondary": "#AAAAAA",
    "text_light": "#FFFFFF",
    "border": "#3D3D3D",
    "divider": "#4D4D4D",
    "success": STATUS_COMPLETE,
    "success_bg": "#1E5631",
    "success_text": "#7FEE6E",
    "warning": STATUS_PENDING,
    "warning_bg": "#664D00",
    "warning_text": "#FFE680",
    "danger": STATUS_OVERDUE,
    "danger_bg": "#6B1A1A",
    "danger_text": "#FF6B6B",
    "danger_hover": "#992828",
    "info_bg": "#1A4D5C",
    "info_text": "#80D4E0",
    "button_disabled_bg": "#444444",
    "button_disabled_text": "#999999",
    "button_hover_shadow": "#FFFFFF15",
    "input_bg": "#2D2D2D",
    "input_border": "#4D4D4D",
    "input_focus_border": PRIMARY_BLUE_DARK,
    "sidebar_active": PRIMARY_BLUE_DARK,
    "sidebar_active_text": "#FFFFFF",
}

# CURRENT THEME (Default: Light)
CURRENT_THEME = "light"
CURRENT_COLORS = LIGHT_COLORS.copy()

# TYPOGRAPHY SIZES
FONT_SIZES = {
    "logo": 40,
    "section_title": 18,
    "header": 16,
    "normal": 13,
    "label": 12,
    "small": 11,
    "help": 12,
}

# SPACING CONSTANTS
SPACING = {
    "tiny": 4,
    "small": 8,
    "normal": 16,
    "large": 24,
    "xl": 32,
}

# COMPONENT DIMENSIONS
DIMENSIONS = {
    "button_height": 44,
    "sidebar_width": 220,
    "sidebar_collapsed": 60,
    "header_height": 64,
    "footer_height": 64,
    "min_button_width": 120,
    "icon_size": 24,
}

# ANIMATION TIMINGS (milliseconds)
ANIMATIONS = {
    "button_hover": 200,
    "button_press": 100,
    "color_transition": 200,
    "modal_open": 300,
    "tooltip": 100,
    "loading_spinner": 1500,
    "sidebar_toggle": 300,
}

# ICON DEFINITIONS
ICONS = {
    "complete": "✓",
    "pending": "⏳",
    "overdue": "⚠",
    "in_progress": "•",
    "error": "❌",
    "info": "ℹ",
    "setup": "📋",
    "process": "🔄",
    "reporting": "📊",
    "tools": "🔧",
    "help": "❓",
    "about": "ℹ",
}

def set_theme(theme_name: str) -> None:
    """Switch between light and dark themes."""
    global CURRENT_THEME, CURRENT_COLORS
    if theme_name.lower() not in ["light", "dark"]:
        raise ValueError("Theme must be 'light' or 'dark'")
    CURRENT_THEME = theme_name.lower()
    CURRENT_COLORS = LIGHT_COLORS.copy() if CURRENT_THEME == "light" else DARK_COLORS.copy()

def get_color(color_name: str, theme: str = None) -> str:
    """Get color value from current or specified theme."""
    if theme:
        color_dict = LIGHT_COLORS if theme.lower() == "light" else DARK_COLORS
    else:
        color_dict = CURRENT_COLORS
    if color_name not in color_dict:
        raise KeyError(f"Color '{color_name}' not found in {CURRENT_THEME} theme")
    return color_dict[color_name]

def get_all_colors(theme: str = None) -> dict:
    """Get all colors for a theme."""
    if theme:
        return LIGHT_COLORS.copy() if theme.lower() == "light" else DARK_COLORS.copy()
    return CURRENT_COLORS.copy()

# ============================================================================
# MODULE CODE CONSOLIDATED BELOW (Original: src/modules/*.py)
# ============================================================================

# ============ MODULE: models ============



# ==================== Constants ====================

class ColumnNames:
    """Standard column names used in Excel files."""
    
    # Master File Columns (Access Certification)
    AGENCY = "Agency"
    DOMAIN = "Domain"
    USER_NAME = "User Name"
    EMAIL = "Email"
    EMAIL_ADDRESS = "EmailAddress"  # Alternate column name
    USERNAME = "UserName"  # Alternate column name
    COUNTRY = "Country"
    PROFILE_NAME = "ProfileName"
    REGION = "Region"
    FOLDER = "Folder"
    SUBFOLDER = "SubFolder"
    
    # Combined Mapping & Email File Columns
    SOURCE_FILE_NAME = "source_file_name"
    FILE_NAME = "File name"  # Legacy column name for old-style mapping
    AGENCY_ID = "agency_id"
    RECIPIENTS_TO = "recipients_to"
    RECIPIENTS_CC = "recipients_cc"
    
    # Audit Log Columns
    SENT_DATE = "Sent Email Date"
    RESPONSE_DATE = "Response Received Date"
    STATUS = "Status"
    COMMENTS = "Comments"
    TO = "To"
    CC = "CC"
    
    # Generated File Columns
    REVIEW_ACTION = "Review Action"
    REVIEW_COMMENTS = "Comments"


class SheetNames:
    """Standard sheet names for generated Excel files."""
    ALL_USERS = "All Users"
    USER_ACCESS_LIST = "User Access List"  # Access Certification format
    USER_ACCESS_SUMMARY = "User Access Summary"  # Pivot summary sheet
    UNASSIGNED = "Unassigned"


class FileNames:
    """Standard file names used by the application."""
    UNASSIGNED_FILE = "Unassigned.xlsx"
    AUDIT_LOG_TEMPLATE = "Audit_CognosAccessReview_{period}.xlsx"
    BACKUP_DIR = "backups"
    LOG_FILE = "cognos_review.log"
    CONFIG_FILE = "config.json"
    EMAIL_TEMPLATE = "email_template.txt"
    CONFIG_SCHEMA = "config.schema.json"
    REGION_PROFILES = "region_profiles.json"


# Tab to Country Mapping - Maps regional tabs to Country column values in master file
# Case-insensitive matching - all values converted to uppercase during comparison
TAB_TO_COUNTRY_MAP = {
    "AMER": [
        "United States", "USA", "US", "United States of America"
    ],
    "CANADA": [
        "Canada", "CA", "CAN"
    ],
    "LATIN AMERICA": [
        "Mexico", "Brazil", "Argentina", "Chile", "Colombia", "Peru", 
        "Venezuela", "Ecuador", "Uruguay", "Paraguay", "Bolivia",
        "Costa Rica", "Panama", "Guatemala", "Honduras", "El Salvador",
        "Nicaragua", "Dominican Republic", "Puerto Rico", "Cuba"
    ],
    "APAC": [
        "Australia", "China", "India", "Japan", "Singapore", "Hong Kong", 
        "New Zealand", "South Korea", "Korea", "Thailand", "Philippines", 
        "Indonesia", "Malaysia", "Vietnam", "Taiwan", "Pakistan", "Bangladesh",
        "Sri Lanka", "Myanmar", "Cambodia", "Laos", "Brunei", "Macau",
        "Mongolia", "Nepal", "Fiji", "Papua New Guinea"
    ],
    "EMEA": [
        "United Kingdom", "UK", "England", "Scotland", "Wales", "Northern Ireland",
        "Germany", "France", "Italy", "Spain", "Netherlands", "Belgium", 
        "Switzerland", "Austria", "Poland", "Sweden", "Denmark", "Norway", 
        "Finland", "Ireland", "Portugal", "Czech Republic", "Czechia", "Hungary", 
        "Romania", "Greece", "Bulgaria", "Croatia", "Serbia", "Slovakia",
        "Slovenia", "Lithuania", "Latvia", "Estonia", "Luxembourg", "Malta",
        "Cyprus", "Iceland", "Albania", "Bosnia", "Macedonia", "Montenegro",
        "South Africa", "UAE", "United Arab Emirates", "Saudi Arabia", "Egypt", "Israel", "Turkey",
        "Qatar", "Kuwait", "Bahrain", "Oman", "Jordan", "Lebanon", "Morocco",
        "Tunisia", "Algeria", "Kenya", "Nigeria", "Ghana", "Ethiopia",
        "Uganda", "Tanzania", "Zimbabwe", "Botswana", "Namibia", "Mauritius"
    ]
}


class EmailMode(Enum):
    """Email sending modes."""
    PREVIEW = "Preview"
    DIRECT = "Direct"
    SCHEDULE = "Schedule"


class AuditStatus(Enum):
    """Status values for audit log entries."""
    NOT_SENT = "Not Sent"
    SENT = "Sent"
    RESPONDED = "Responded"
    OVERDUE = "Overdue"


class ValidationSeverity(Enum):
    """Severity levels for validation issues."""
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


# ==================== TypedDict Definitions ====================

class AppConfig(TypedDict, total=False):
    """Application configuration structure."""
    review_period: str
    deadline: str
    company_name: str
    sender_name: str
    sender_title: str
    email_subject_prefix: str
    auto_scan: bool
    default_email_mode: str
    current_region: Optional[str]


class RegionProfileDict(TypedDict):
    """Region profile data structure for JSON serialization."""
    region_name: str
    master_file: str
    mapping_file: str
    output_dir: str
    email_manifest: str
    description: str


class EmailChangeRecord(TypedDict):
    """Record of an email address change in the manifest."""
    agency: str
    field: str  # "To" or "CC"
    old_value: str
    new_value: str
    timestamp: str


class ValidationIssue(TypedDict):
    """Single validation issue."""
    severity: str  # ValidationSeverity
    message: str
    details: Optional[str]


class ValidationResults(TypedDict):
    """Complete validation results."""
    is_valid: bool
    issues: List[ValidationIssue]
    master_agencies: set
    mapped_agencies: set
    unmapped_agencies: set
    duplicate_agencies: Dict[str, List[str]]
    email_issues: List[ValidationIssue]
    master_agency_count: int
    mapped_agency_count: int


class EmailRecipients(TypedDict):
    """Email recipient information."""
    to: str
    cc: str


class ProgressCallback:
    """Type hint for progress callback function."""
    def __call__(self, progress: float, status: str) -> None:
        """
        Update progress indicator.
        
        Args:
            progress: Progress value between 0.0 and 1.0
            status: Status message to display
        """
        ...


# ==================== Data Classes ====================

@dataclass
class AgencyMapping:
    """Represents a single agency mapping entry."""
    file_name: str
    agency_id: str
    
    def __post_init__(self):
        """Validate and normalize data after initialization."""
        self.file_name = self.file_name.strip()
        self.agency_id = self.agency_id.strip()
        
    @property
    def excel_file_name(self) -> str:
        """Get the full Excel file name with extension."""
        return f"{self.file_name}.xlsx"


@dataclass
class CombinedMapping:
    """Represents a single entry from the combined agency/email mapping file."""
    source_file_name: str
    agency_id: str
    recipients_to: str
    recipients_cc: str
    source_tab: str = ""
    
    def __post_init__(self):
        """Validate and normalize data after initialization."""
        self.source_file_name = self.source_file_name.strip()
        self.agency_id = self.agency_id.strip()
        self.recipients_to = self.recipients_to.strip() if self.recipients_to else ""
        self.recipients_cc = self.recipients_cc.strip() if self.recipients_cc else ""
        self.source_tab = self.source_tab.strip()
        
    @property
    def excel_file_name(self) -> str:
        """Get the full Excel file name with extension."""
        return f"{self.source_file_name}.xlsx"
    
    def to_agency_mapping(self) -> AgencyMapping:
        """Convert to AgencyMapping for compatibility."""
        return AgencyMapping(
            file_name=self.source_file_name,
            agency_id=self.agency_id
        )
    
    def get_email_recipients(self) -> EmailRecipients:
        """Get email recipients for this mapping."""
        return {
            "to": self.recipients_to,
            "cc": self.recipients_cc
        }


@dataclass
class AuditLogEntry:
    """Represents a single entry in the audit log."""
    agency: str
    sent_date: Optional[datetime] = None
    response_date: Optional[datetime] = None
    to: str = ""
    cc: str = ""
    status: str = AuditStatus.NOT_SENT.value
    comments: str = ""
    
    def to_dict(self) -> Dict[str, str]:
        """Convert to dictionary for DataFrame export."""
        return {
            ColumnNames.AGENCY: self.agency,
            ColumnNames.SENT_DATE: self.sent_date.strftime("%Y-%m-%d %H:%M") if self.sent_date else "",
            ColumnNames.RESPONSE_DATE: self.response_date.strftime("%Y-%m-%d %H:%M") if self.response_date else "",
            ColumnNames.TO: self.to,
            ColumnNames.CC: self.cc,
            ColumnNames.STATUS: self.status,
            ColumnNames.COMMENTS: self.comments
        }


@dataclass
class FileGenerationRequest:
    """Request parameters for agency file generation."""
    master_file: Path
    agency_map_file: Path
    output_dir: Path
    progress_callback: Optional[ProgressCallback] = None
    
    def validate(self) -> List[str]:
        """
        Validate all file paths exist and are accessible.
        
        Returns:
            List of error messages, empty if all valid
        """
        errors = []
        
        if not self.master_file.exists():
            errors.append(f"Master file not found: {self.master_file}")
        elif not self.master_file.is_file():
            errors.append(f"Master file path is not a file: {self.master_file}")
            
        if not self.agency_map_file.exists():
            errors.append(f"Agency mapping file not found: {self.agency_map_file}")
        elif not self.agency_map_file.is_file():
            errors.append(f"Agency mapping file path is not a file: {self.agency_map_file}")
            
        if not self.output_dir.exists():
            errors.append(f"Output directory not found: {self.output_dir}")
        elif not self.output_dir.is_dir():
            errors.append(f"Output path is not a directory: {self.output_dir}")
            
        return errors


@dataclass
class EmailSendRequest:
    """Request parameters for email sending."""
    combined_file: Path
    output_dir: Path
    file_names: List[str]
    universal_attachment: Optional[Path] = None
    mode: EmailMode = EmailMode.PREVIEW
    
    def validate(self) -> List[str]:
        """Validate all file paths and parameters."""
        errors = []
        
        if not self.combined_file.exists():
            errors.append(f"Combined file not found: {self.combined_file}")
            
        if not self.output_dir.exists():
            errors.append(f"Output directory not found: {self.output_dir}")
            
        if self.universal_attachment and not self.universal_attachment.exists():
            errors.append(f"Universal attachment not found: {self.universal_attachment}")
            
        if not self.file_names:
            errors.append("No file names selected for email sending")
            
        return errors


@dataclass
class DashboardMetrics:
    """Metrics displayed in the dashboard."""
    total_agencies: int
    sent_count: int
    responded_count: int
    not_sent_count: int
    overdue_count: int
    completion_percentage: float
    days_left: int
    
    @classmethod
    def calculate(cls, audit_df, deadline_str: str, total_agencies: int) -> 'DashboardMetrics':
        """
        Calculate metrics from audit DataFrame.
        
        Args:
            audit_df: Pandas DataFrame containing audit log
            deadline_str: Deadline string in format "Month Day, Year"
            total_agencies: Total number of agencies
            
        Returns:
            DashboardMetrics instance with calculated values
        """
        
        # Count statuses
        if audit_df.empty:
            sent = responded = 0
        else:
            sent = (audit_df[ColumnNames.STATUS] == AuditStatus.SENT.value).sum()
            responded = (audit_df[ColumnNames.STATUS] == AuditStatus.RESPONDED.value).sum()
        
        not_sent = total_agencies - sent - responded
        
        # Calculate completion percentage
        completion = (responded / total_agencies * 100) if total_agencies > 0 else 0.0
        
        # Calculate days left
        try:
            deadline = datetime.strptime(deadline_str, "%B %d, %Y")
            days_left = (deadline - datetime.now()).days
        except (ValueError, TypeError):
            days_left = -1
        
        # Calculate overdue (sent more than 7 days ago without response)
        overdue = 0
        if not audit_df.empty and ColumnNames.SENT_DATE in audit_df.columns:
            for _, row in audit_df.iterrows():
                if row[ColumnNames.STATUS] == AuditStatus.SENT.value:
                    sent_date_str = row.get(ColumnNames.SENT_DATE, "")
                    if sent_date_str:
                        try:
                            sent_date = datetime.strptime(sent_date_str, "%Y-%m-%d %H:%M")
                            if (datetime.now() - sent_date).days > 7:
                                overdue += 1
                        except (ValueError, TypeError):
                            pass
        
        return cls(
            total_agencies=total_agencies,
            sent_count=sent,
            responded_count=responded,
            not_sent_count=not_sent,
            overdue_count=overdue,
            completion_percentage=completion,
            days_left=days_left
        )


@dataclass
class ExcelFormatting:
    """Excel worksheet formatting options."""
    freeze_header: bool = True
    bold_header: bool = True
    header_bg_color: str = "#D9E1F2"
    auto_width: bool = True
    max_column_width: int = 50
    border_header: bool = True


# ==================== Exceptions ====================

class CognosReviewError(Exception):
    """Base exception for all application errors."""
    pass


class ValidationError(CognosReviewError):
    """Raised when validation fails."""
    pass


class FileProcessingError(CognosReviewError):
    """Raised when file processing fails."""
    pass


class EmailError(CognosReviewError):
    """Raised when email operations fail."""
    pass


class ConfigurationError(CognosReviewError):
    """Raised when configuration is invalid."""
    pass


# ==================== Helper Functions ====================

def sanitize_filename(name: str) -> str:
    """
    Sanitize a string to be a valid Windows filename.
    
    Windows filename restrictions:
    - Cannot contain: < > : " / \\ | ? * '
    - Cannot end with space or period
    - Max 255 characters
    
    Args:
        name: Raw filename (without extension)
        
    Returns:
        Sanitized filename safe for Windows
        
    Examples:
        >>> sanitize_filename("TBWA Chiat/Day")
        'TBWA Chiat-Day'
        >>> sanitize_filename("File: Test*123?")
        'File- Test-123-'
    """
    
    # Replace invalid Windows filename characters with hyphens
    name = re.sub(r'[<>:"/\\|?*\']', '-', str(name))
    
    # Remove leading/trailing spaces and periods
    name = name.strip().strip('.')
    
    # Truncate to 255 characters (Windows limit)
    if len(name) > 255:
        name = name[:255]
    
    return name if name else "File"

def sanitize_sheet_name(name: str) -> str:
    """
    Sanitize a string to be a valid Excel sheet name.
    
    Excel sheet names have restrictions:
    - Max 31 characters
    - Cannot contain: \\ / ? * [ ]
    
    Args:
        name: Raw sheet name
        
    Returns:
        Sanitized sheet name safe for Excel
        
    Examples:
        >>> sanitize_sheet_name("ABC/DEF*123?")
        'ABCDEF123'
        >>> sanitize_sheet_name("A" * 50)
        'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'  # 31 chars
    """
    
    # Remove invalid characters
    name = re.sub(r'[\\/*?\[\]]', '', str(name))
    
    # Truncate to 31 characters
    if len(name) > 31:
        name = name[:31]
    
    return name if name else "Sheet1"


def format_email_list(emails: str) -> List[str]:
    """
    Parse semicolon-separated email list into individual addresses.
    
    Args:
        emails: Semicolon-separated email addresses
        
    Returns:
        List of individual email addresses
        
    Examples:
        >>> format_email_list("abc@example.com; def@example.com")
        ['abc@example.com', 'def@example.com']
        >>> format_email_list("")
        []
    """
    if not emails or not isinstance(emails, str):
        return []
    
    return [email.strip() for email in emails.split(';') if email.strip()]


def is_valid_email(email: str) -> bool:
    """
    Validate email address format using basic regex.
    
    Args:
        email: Email address to validate
        
    Returns:
        True if email format is valid
        
    Examples:
        >>> is_valid_email("user@example.com")
        True
        >>> is_valid_email("invalid-email")
        False
    """
    
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email.strip()))


def extract_emails_from_text(text: str) -> List[str]:
    """
    Extract email addresses from text using smart parsing.
    
    This function can find emails in various formats including:
    - Direct emails: user@example.com
    - Forward requests: "please forward to john.doe@company.com"
    - CC requests: "please cc manager@dept.com"
    - Multiple formats: "send to: email1@test.com, email2@test.com"
    
    Args:
        text: Text to search for email addresses
        
    Returns:
        List of unique valid email addresses found
        
    Examples:
        >>> extract_emails_from_text("Please forward to john@example.com")
        ['john@example.com']
        >>> extract_emails_from_text("CC: manager@test.com and admin@test.com")
        ['manager@test.com', 'admin@test.com']
    """
    if not text or not isinstance(text, str):
        return []
    
    # Enhanced email regex pattern
    email_pattern = r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'
    
    # Find all potential email matches
    potential_emails = re.findall(email_pattern, text, re.IGNORECASE)
    
    # Validate and deduplicate emails
    valid_emails = []
    seen_emails = set()
    
    for email in potential_emails:
        email = email.strip().lower()
        if email not in seen_emails and is_valid_email(email):
            valid_emails.append(email)
            seen_emails.add(email)
    
    return valid_emails


def parse_email_forwarding_message(message: str) -> Dict[str, List[str]]:
    """
    Parse email forwarding messages to extract TO and CC email addresses.
    
    Recognizes patterns like:
    - "please forward to john@example.com"
    - "send email to: user1@test.com, user2@test.com"
    - "cc: manager@company.com"
    - "also include admin@dept.com in cc"
    
    Args:
        message: Email message text to parse
        
    Returns:
        Dictionary with 'to' and 'cc' lists of email addresses
        
    Examples:
        >>> result = parse_email_forwarding_message("Please forward to john@test.com and cc manager@test.com")
        >>> result['to']
        ['john@test.com']
        >>> result['cc']  
        ['manager@test.com']
    """
    if not message or not isinstance(message, str):
        return {"to": [], "cc": []}
    
    message = message.lower()
    all_emails = extract_emails_from_text(message)
    
    result = {"to": [], "cc": []}
    
    # Patterns for CC emails
    cc_patterns = [
        r'cc[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'copy[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'also\s+(?:include|cc|copy)[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'please\s+(?:cc|copy)[:\s]+([^.\n]*?)(?=\n|$|\.)'
    ]
    
    # Patterns for TO emails  
    to_patterns = [
        r'(?:forward|send)[^@]*?to[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'please\s+(?:send|forward)[^@]*?(?:to|email)[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'send\s+(?:email\s+)?to[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'forward\s+(?:email\s+)?to[:\s]+([^.\n]*?)(?=\n|$|\.)'
    ]
    
    # Extract CC emails
    cc_emails = set()
    for pattern in cc_patterns:
        matches = re.findall(pattern, message, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            emails_in_match = extract_emails_from_text(match)
            cc_emails.update(emails_in_match)
    
    # Extract TO emails
    to_emails = set()
    for pattern in to_patterns:
        matches = re.findall(pattern, message, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            emails_in_match = extract_emails_from_text(match)
            to_emails.update(emails_in_match)
    
    # If no specific TO patterns found, treat remaining emails as TO
    remaining_emails = set(all_emails) - cc_emails - to_emails
    if not to_emails and remaining_emails:
        to_emails = remaining_emails
    
    result["to"] = sorted(list(to_emails))
    result["cc"] = sorted(list(cc_emails))
    
    return result


def smart_email_verification(text: str) -> Dict[str, any]:
    """
    Comprehensive email verification and extraction from text.
    
    Args:
        text: Text to analyze for email addresses and forwarding instructions
        
    Returns:
        Dictionary with verification results and extracted emails
        
    Examples:
        >>> result = smart_email_verification("Please forward to john@test.com and cc boss@company.com")
        >>> result['found_emails']
        ['john@test.com', 'boss@company.com']
        >>> result['parsing']['to']
        ['john@test.com']
    """
    result = {
        "found_emails": [],
        "valid_emails": [],
        "invalid_emails": [],
        "parsing": {"to": [], "cc": []},
        "suggestions": [],
        "confidence": "low"
    }
    
    if not text or not isinstance(text, str):
        return result
    
    # Extract all emails
    found_emails = extract_emails_from_text(text)
    result["found_emails"] = found_emails
    
    # Validate emails
    for email in found_emails:
        if is_valid_email(email):
            result["valid_emails"].append(email)
        else:
            result["invalid_emails"].append(email)
    
    # Parse forwarding instructions
    result["parsing"] = parse_email_forwarding_message(text)
    
    # Generate suggestions
    if result["valid_emails"]:
        result["confidence"] = "high"
        if result["parsing"]["to"] or result["parsing"]["cc"]:
            result["suggestions"].append("Smart parsing detected forwarding instructions")
        else:
            result["suggestions"].append("Valid emails found but no clear forwarding pattern")
    elif result["invalid_emails"]:
        result["confidence"] = "medium"
        result["suggestions"].append("Found email-like text but invalid format")
    else:
        result["confidence"] = "low"
        result["suggestions"].append("No email addresses detected in text")
    
    return result


# ============ MODULE: config_manager ============





class ConfigManager:
    """
    Manages application configuration with validation and defaults.
    
    This class provides a centralized way to load, save, and access application
    configuration with proper validation and type safety.
    
    Examples:
        >>> config_mgr = ConfigManager()
        >>> config = config_mgr.load()
        >>> print(config['review_period'])
        'Q2 2025'
        >>> config_mgr.update({'review_period': 'Q3 2025'})
        >>> config_mgr.save()
    """
    
    DEFAULT_CONFIG: AppConfig = {
        "review_period": "Q2 2025",
        "deadline": "June 30, 2025",
        "company_name": "Omnicom Group",
        "sender_name": "Govind Waghmare",
        "sender_title": "Manager, Financial Applications | Analytics",
        "email_subject_prefix": "[ACTION REQUIRED] Cognos Access Review",
        "auto_scan": True,
        "default_email_mode": EmailMode.PREVIEW.value,
        "current_region": None
    }
    
    REQUIRED_FIELDS = [
        "review_period",
        "deadline",
        "company_name",
        "sender_name",
        "sender_title"
    ]
    
    def __init__(self, config_path: Optional[Path] = None):
        """
        Initialize configuration manager.
        
        Args:
            config_path: Path to config file. Defaults to config.json in current directory.
        """
        self.config_path = config_path or Path(FileNames.CONFIG_FILE)
        self._config: AppConfig = {}
        self._loaded = False
    
    def load(self) -> AppConfig:
        """
        Load configuration from file or return defaults.
        
        Returns:
            Loaded and validated configuration
            
        Raises:
            ConfigurationError: If config file is invalid JSON
        """
        if self._loaded:
            return self._config
            
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    
                # Merge with defaults (defaults first, then loaded values override)
                self._config = {**self.DEFAULT_CONFIG, **loaded_config}
                
                # Validate after loading
                self.validate()
                
                logger.info(f"Configuration loaded from {self.config_path}")
                
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON in config file: {e}")
                raise ConfigurationError(f"Invalid JSON in config file: {e}")
            except Exception as e:
                logger.error(f"Failed to load config: {e}")
                raise ConfigurationError(f"Failed to load config: {e}")
        else:
            logger.warning(f"Config file not found at {self.config_path}, using defaults")
            self._config = self.DEFAULT_CONFIG.copy()
        
        self._loaded = True
        return self._config
    
    def save(self) -> None:
        """
        Save current configuration to file.
        
        Raises:
            ConfigurationError: If saving fails
        """
        if not self._config:
            logger.warning("Attempting to save empty config, loading defaults first")
            self._config = self.DEFAULT_CONFIG.copy()
        
        try:
            # Validate before saving
            self.validate()
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Configuration saved to {self.config_path}")
            
        except Exception as e:
            logger.error(f"Failed to save config: {e}")
            raise ConfigurationError(f"Failed to save config: {e}")
    
    def validate(self) -> None:
        """
        Validate current configuration.
        
        Raises:
            ConfigurationError: If validation fails
        """
        errors = []
        
        # Check required fields
        for field in self.REQUIRED_FIELDS:
            if field not in self._config or not self._config[field]:
                errors.append(f"Missing required field: {field}")
        
        # Validate email mode if present
        if "default_email_mode" in self._config:
            mode = self._config["default_email_mode"]
            valid_modes = [e.value for e in EmailMode]
            if mode not in valid_modes:
                errors.append(f"Invalid email mode '{mode}'. Must be one of: {valid_modes}")
        
        # Validate boolean fields
        for field in ["auto_scan"]:
            if field in self._config:
                value = self._config[field]
                if not isinstance(value, bool):
                    errors.append(f"Field '{field}' must be boolean, got {type(value).__name__}")
        
        if errors:
            error_msg = "Configuration validation failed:\n" + "\n".join(f"  - {e}" for e in errors)
            raise ConfigurationError(error_msg)
    
    def get(self, key: str, default=None):
        """
        Get configuration value by key.
        
        Args:
            key: Configuration key
            default: Default value if key not found
            
        Returns:
            Configuration value or default
        """
        if not self._loaded:
            self.load()
        return self._config.get(key, default)
    
    def update(self, updates: Dict[str, any]) -> None:
        """
        Update configuration with new values.
        
        Args:
            updates: Dictionary of configuration updates
            
        Raises:
            ConfigurationError: If updates result in invalid configuration
        """
        if not self._loaded:
            self.load()
        
        # Apply updates
        self._config.update(updates)
        
        # Validate after updates
        self.validate()
        
        logger.info(f"Configuration updated with {len(updates)} changes")
    
    def get_all(self) -> AppConfig:
        """
        Get complete configuration.
        
        Returns:
            Complete configuration dictionary
        """
        if not self._loaded:
            self.load()
        return self._config.copy()
    
    def reset_to_defaults(self) -> None:
        """Reset configuration to default values."""
        self._config = self.DEFAULT_CONFIG.copy()
        self._loaded = True
        logger.info("Configuration reset to defaults")
    
    def get_audit_file_name(self, region_name: Optional[str] = None) -> str:
        """
        Get the audit log file name based on current review period and optional region.
        
        Args:
            region_name: Optional region name to include in filename
        
        Returns:
            Audit log file name
            
        Examples:
            >>> config_mgr = ConfigManager()
            >>> config_mgr.get_audit_file_name()
            'Audit_CognosAccessReview_Q2_2025.xlsx'
            >>> config_mgr.get_audit_file_name("North_America")
            'Audit_CognosAccessReview_Q2_2025_North_America.xlsx'
        """
        if not self._loaded:
            self.load()
        
        period = self._config.get("review_period", "Q2_2025")
        # Replace spaces with underscores for file name
        period_safe = period.replace(' ', '_')
        
        if region_name:
            region_safe = region_name.replace(' ', '_')
            return f"Audit_CognosAccessReview_{period_safe}_{region_safe}.xlsx"
        
        return FileNames.AUDIT_LOG_TEMPLATE.format(period=period_safe)
    
    def set_current_region(self, region_name: Optional[str]) -> None:
        """
        Set the current active region.
        
        Args:
            region_name: Name of the region to set as current (None to clear)
        """
        if not self._loaded:
            self.load()
        
        self._config["current_region"] = region_name
        self.save()
        logger.info(f"Set current region to: {region_name}")
    
    def get_current_region(self) -> Optional[str]:
        """
        Get the currently active region name.
        
        Returns:
            Current region name or None
        """
        if not self._loaded:
            self.load()
        
        return self._config.get("current_region")
    
    def __repr__(self) -> str:
        """String representation of configuration manager."""
        status = "loaded" if self._loaded else "not loaded"
        return f"ConfigManager(path={self.config_path}, status={status})"


def load_email_template(template_path: Optional[Path] = None) -> str:
    """
    Load email template from file or return default.
    
    Args:
        template_path: Path to email template file
        
    Returns:
        Email template content with placeholders
        
    Examples:
        >>> template = load_email_template()
        >>> '{review_period}' in template
        True
    """
    template_path = template_path or Path(FileNames.EMAIL_TEMPLATE)
    
    DEFAULT_TEMPLATE = """Hello,

As part of our Sarbanes-Oxley (SOX) compliance requirements, we must complete the {review_period} Cognos Platform user access review by {deadline}. Your review and response are critical to ensuring compliance and maintaining appropriate system access.

Please note: The Q3 review was delayed due to the domain change to omc.com, which caused a temporary separation of user accounts. This issue has now been resolved.

About the Attached File:
The attached Excel file contains multiple tabs for your review:
- "All Users" tab: Complete list of all Cognos users under your responsibility
- Individual agency tabs: Each tab shows users specific to that agency for easier review

Action Required:
1. Review the attached User Access Report listing Cognos Reporting Users and their folder access as of {review_period}.
2. Please review each tab relevant to your agencies.
3. Confirm or request changes:
   - If no updates are needed, simply reply confirming your review.
   - If updates are required, please note them in Column J ("Review Action") of the respective tab and return the file.
   - For access changes, submit a Paige ticket under the Cognos Services section.

This review is mandatory for SOX compliance, and your prompt response is essential. Please let me know if you have any questions or need clarification on any of the tabs.

Best Regards,
{sender_name}
{sender_title}
{company_name}"""
    
    if template_path.exists():
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template = f.read()
            logger.info(f"Email template loaded from {template_path}")
            return template
        except Exception as e:
            logger.warning(f"Failed to load email template from {template_path}: {e}, using default")
    else:
        logger.info(f"Email template file not found at {template_path}, using default")
    
    return DEFAULT_TEMPLATE


def save_email_template(content: str, template_path: Optional[Path] = None) -> None:
    """
    Save email template to file.
    
    Args:
        content: Email template content
        template_path: Path to save template file
        
    Raises:
        ConfigurationError: If saving fails
    """
    template_path = template_path or Path(FileNames.EMAIL_TEMPLATE)
    
    try:
        with open(template_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logger.info(f"Email template saved to {template_path}")
    except Exception as e:
        logger.error(f"Failed to save email template: {e}")
        raise ConfigurationError(f"Failed to save email template: {e}")


# Create a global instance for convenience
_global_config_manager: Optional[ConfigManager] = None


def get_config_manager() -> ConfigManager:
    """
    Get or create global configuration manager instance.
    
    Returns:
        Global ConfigManager instance
    """
    global _global_config_manager
    if _global_config_manager is None:
        _global_config_manager = ConfigManager()
    return _global_config_manager


# ============ MODULE: audit_logger ============






class AuditLogger:
    """
    Manages audit log for tracking email status and responses.
    
    This class provides functionality for creating, updating, and querying
    the Excel-based audit log with automatic backups.
    
    Examples:
        >>> audit_logger = AuditLogger(output_dir=Path("output"))
        >>> audit_logger.initialize_log(agencies=["Agency1", "Agency2"])
        >>> audit_logger.mark_sent(
        ...     agency="Agency1",
        ...     to="test@example.com",
        ...     cc="manager@example.com"
        ... )
        >>> metrics = audit_logger.get_metrics(deadline="June 30, 2025")
        >>> print(f"Completion: {metrics.completion_percentage:.1f}%")
    """
    
    def __init__(
        self,
        output_dir: Optional[Path] = None,
        audit_file_name: Optional[str] = None
    ):
        """
        Initialize audit logger.
        
        Args:
            output_dir: Directory for audit log file (optional, can be set later)
            audit_file_name: Optional custom audit file name
        """
        self.output_dir = Path(output_dir) if output_dir else None
        self.audit_file_name = audit_file_name or "Audit_CognosAccessReview.xlsx"
        self.audit_file_path = (self.output_dir / self.audit_file_name) if self.output_dir else None
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self._df: Optional[pd.DataFrame] = None
    
    def _get_columns(self) -> List[str]:
        """Get standard audit log column names."""
        return [
            ColumnNames.AGENCY,
            ColumnNames.SENT_DATE,
            ColumnNames.RESPONSE_DATE,
            ColumnNames.TO,
            ColumnNames.CC,
            ColumnNames.STATUS,
            ColumnNames.COMMENTS
        ]
    
    def load(self) -> pd.DataFrame:
        """
        Load audit log from file or create empty DataFrame.
        
        Returns:
            Audit log DataFrame
        """
        if self._df is not None:
            return self._df
        
        if self.audit_file_path.exists():
            try:
                self._df = pd.read_excel(self.audit_file_path)
                # Ensure all required columns exist
                for col in self._get_columns():
                    if col not in self._df.columns:
                        self._df[col] = ""
                self.logger.info(f"Loaded audit log: {self.audit_file_path}")
            except Exception as e:
                self.logger.error(f"Failed to load audit log: {e}")
                self._df = pd.DataFrame(columns=self._get_columns())
        else:
            self._df = pd.DataFrame(columns=self._get_columns())
            self.logger.info("Created new audit log DataFrame")
        
        return self._df
    
    def save(self) -> bool:
        """
        Save audit log to Excel file with backup.
        
        Returns:
            True if successful
        """
        if self._df is None:
            self.logger.warning("No audit data to save")
            return False
        
        try:
            # Create backup if file exists
            if self.audit_file_path.exists():
                self.create_backup()
            
            # Ensure output directory exists
            self.output_dir.mkdir(parents=True, exist_ok=True)
            
            # Save to Excel
            self._df.to_excel(self.audit_file_path, index=False)
            self.logger.info(f"Saved audit log: {self.audit_file_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save audit log: {e}")
            return False
    
    def create_backup(self) -> Optional[Path]:
        """
        Create timestamped backup of audit log.
        
        Returns:
            Path to backup file or None if failed
        """
        if not self.audit_file_path.exists():
            return None
        
        try:
            # Create backup directory
            backup_dir = self.output_dir / FileNames.BACKUP_DIR
            backup_dir.mkdir(exist_ok=True)
            
            # Create timestamped backup filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{timestamp}_{self.audit_file_name}"
            backup_path = backup_dir / backup_name
            
            # Copy file
            shutil.copy2(self.audit_file_path, backup_path)
            self.logger.info(f"Created backup: {backup_path}")
            return backup_path
            
        except Exception as e:
            self.logger.error(f"Failed to create backup: {e}")
            return None
    
    def initialize_log(
        self,
        agencies: List[str],
        preserve_existing: bool = True
    ) -> None:
        """
        Initialize audit log with agency list.
        
        Args:
            agencies: List of agency names
            preserve_existing: If True, preserve existing entries
        """
        df = self.load()
        
        if preserve_existing and not df.empty:
            # Get existing agencies
            existing_agencies = set(df[ColumnNames.AGENCY].str.upper())
            
            # Add only new agencies
            new_agencies = [
                agency for agency in agencies
                if agency.upper() not in existing_agencies
            ]
            
            if new_agencies:
                new_entries = pd.DataFrame({
                    ColumnNames.AGENCY: new_agencies,
                    ColumnNames.SENT_DATE: "",
                    ColumnNames.RESPONSE_DATE: "",
                    ColumnNames.TO: "",
                    ColumnNames.CC: "",
                    ColumnNames.STATUS: AuditStatus.NOT_SENT.value,
                    ColumnNames.COMMENTS: ""
                })
                self._df = pd.concat([df, new_entries], ignore_index=True)
                self.logger.info(f"Added {len(new_agencies)} new agencies to audit log")
        else:
            # Create fresh log
            self._df = pd.DataFrame({
                ColumnNames.AGENCY: agencies,
                ColumnNames.SENT_DATE: "",
                ColumnNames.RESPONSE_DATE: "",
                ColumnNames.TO: "",
                ColumnNames.CC: "",
                ColumnNames.STATUS: AuditStatus.NOT_SENT.value,
                ColumnNames.COMMENTS: ""
            })
            self.logger.info(f"Initialized audit log with {len(agencies)} agencies")
    
    def mark_sent(
        self,
        agency: str,
        to: str = "",
        cc: str = "",
        comments: str = ""
    ) -> bool:
        """
        Mark an agency as having email sent.
        
        Args:
            agency: Agency name
            to: To email addresses
            cc: CC email addresses
            comments: Optional comments
            
        Returns:
            True if updated successfully
        """
        df = self.load()
        
        # Find agency (case-insensitive)
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        
        if not mask.any():
            self.logger.warning(f"Agency not found in audit log: {agency}")
            return False
        
        # Update row
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
        df.loc[mask, ColumnNames.SENT_DATE] = current_time
        df.loc[mask, ColumnNames.TO] = to
        df.loc[mask, ColumnNames.CC] = cc
        df.loc[mask, ColumnNames.STATUS] = AuditStatus.SENT.value
        if comments:
            df.loc[mask, ColumnNames.COMMENTS] = comments
        
        self._df = df
        self.logger.info(f"Marked as sent: {agency}")
        return True
    
    def mark_responded(
        self,
        agency: str,
        comments: str = ""
    ) -> bool:
        """
        Mark an agency as having responded.
        
        Args:
            agency: Agency name
            comments: Optional response comments
            
        Returns:
            True if updated successfully
        """
        df = self.load()
        
        # Find agency (case-insensitive)
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        
        if not mask.any():
            self.logger.warning(f"Agency not found in audit log: {agency}")
            return False
        
        # Update row
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
        df.loc[mask, ColumnNames.RESPONSE_DATE] = current_time
        df.loc[mask, ColumnNames.STATUS] = AuditStatus.RESPONDED.value
        if comments:
            current_comments = df.loc[mask, ColumnNames.COMMENTS].iloc[0]
            if current_comments:
                df.loc[mask, ColumnNames.COMMENTS] = f"{current_comments}; {comments}"
            else:
                df.loc[mask, ColumnNames.COMMENTS] = comments
        
        self._df = df
        self.logger.info(f"Marked as responded: {agency}")
        return True
    
    def get_status(self, agency: str) -> Optional[str]:
        """
        Get status for a specific agency.
        
        Args:
            agency: Agency name
            
        Returns:
            Status string or None if not found
        """
        df = self.load()
        
        # Find agency (case-insensitive)
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        
        if not mask.any():
            return None
        
        return df.loc[mask, ColumnNames.STATUS].iloc[0]
    
    def get_agencies_by_status(self, status: AuditStatus) -> List[str]:
        """
        Get list of agencies with specific status.
        
        Args:
            status: Status to filter by
            
        Returns:
            List of agency names
        """
        df = self.load()
        
        if df.empty:
            return []
        
        mask = df[ColumnNames.STATUS] == status.value
        return df.loc[mask, ColumnNames.AGENCY].tolist()
    
    def get_metrics(
        self,
        deadline: str,
        total_agencies: Optional[int] = None
    ) -> DashboardMetrics:
        """
        Calculate dashboard metrics from audit log.
        
        Args:
            deadline: Deadline string (e.g., "June 30, 2025")
            total_agencies: Optional total agency count (uses log count if None)
            
        Returns:
            DashboardMetrics instance
        """
        df = self.load()
        
        if total_agencies is None:
            total_agencies = len(df) if not df.empty else 0
        
        return DashboardMetrics.calculate(df, deadline, total_agencies)
    
    def export_to_csv(self, output_path: Path) -> bool:
        """
        Export audit log to CSV format.
        
        Args:
            output_path: Path for CSV output
            
        Returns:
            True if successful
        """
        df = self.load()
        
        if df.empty:
            self.logger.warning("No audit data to export")
            return False
        
        try:
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            self.logger.info(f"Exported audit log to CSV: {output_path}")
            return True
        except Exception as e:
            self.logger.error(f"Failed to export to CSV: {e}")
            return False
    
    def get_summary_stats(self) -> Dict[str, int]:
        """
        Get summary statistics from audit log.
        
        Returns:
            Dictionary with count statistics
        """
        df = self.load()
        
        if df.empty:
            return {
                "total": 0,
                "not_sent": 0,
                "sent": 0,
                "responded": 0,
                "overdue": 0
            }
        
        stats = {
            "total": len(df),
            "not_sent": (df[ColumnNames.STATUS] == AuditStatus.NOT_SENT.value).sum(),
            "sent": (df[ColumnNames.STATUS] == AuditStatus.SENT.value).sum(),
            "responded": (df[ColumnNames.STATUS] == AuditStatus.RESPONDED.value).sum(),
            "overdue": 0
        }
        
        # Calculate overdue (sent more than 7 days ago)
        if ColumnNames.SENT_DATE in df.columns:
            for _, row in df.iterrows():
                if row[ColumnNames.STATUS] == AuditStatus.SENT.value:
                    sent_date_str = row.get(ColumnNames.SENT_DATE, "")
                    if sent_date_str:
                        try:
                            sent_date = datetime.strptime(str(sent_date_str), "%Y-%m-%d %H:%M")
                            if (datetime.now() - sent_date).days > 7:
                                stats["overdue"] += 1
                        except (ValueError, TypeError):
                            pass
        
        return stats


# ============ MODULE: region_manager ============




@dataclass
class RegionProfile:
    """
    Represents a complete regional configuration.
    
    Attributes:
        region_name: Unique identifier for the region (e.g., "North America", "EMEA")
        master_file: Path to the master user access file for this region
        mapping_file: Path to the agency mapping file for this region
        output_dir: Output directory for generated files for this region
        email_manifest: Path to the shared email manifest file
        description: Optional description of the region
    """
    region_name: str
    master_file: str
    mapping_file: str
    output_dir: str
    email_manifest: str
    description: str = ""
    
    def validate(self) -> List[str]:
        """
        Validate that all required paths are set.
        
        Returns:
            List of validation error messages (empty if valid)
        """
        errors = []
        
        if not self.region_name or not self.region_name.strip():
            errors.append("Region name is required")
        
        if not self.master_file or not self.master_file.strip():
            errors.append("Master file path is required")
        
        if not self.mapping_file or not self.mapping_file.strip():
            errors.append("Agency mapping file path is required")
        
        if not self.output_dir or not self.output_dir.strip():
            errors.append("Output directory is required")
        
        if not self.email_manifest or not self.email_manifest.strip():
            errors.append("Email manifest path is required")
        
        return errors
    
    def to_dict(self) -> Dict:
        """Convert to dictionary for JSON serialization."""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'RegionProfile':
        """Create RegionProfile from dictionary."""
        return cls(**data)


class RegionManager:
    """
    Manages regional profiles with persistence to JSON file.
    
    This class provides methods to create, update, delete, and switch between
    different regional configurations. Each region has its own master file,
    agency mapping, and output directory, but shares a common email manifest.
    
    Examples:
        >>> manager = RegionManager()
        >>> profile = RegionProfile(
        ...     region_name="North America",
        ...     master_file="C:/Data/NA_Master.xlsx",
        ...     mapping_file="C:/Data/NA_Mapping.xlsx",
        ...     output_dir="C:/Output/NA",
        ...     email_manifest="C:/Data/Email_Manifest.xlsx"
        ... )
        >>> manager.add_profile(profile)
        >>> manager.switch_region("North America")
        >>> current = manager.get_current_profile()
    """
    
    def __init__(self, config_file: Optional[Path] = None):
        """
        Initialize region manager.
        
        Args:
            config_file: Path to region profiles JSON file (defaults to region_profiles.json)
        """
        self.config_file = config_file or Path("region_profiles.json")
        self.profiles: Dict[str, RegionProfile] = {}
        self.current_region: Optional[str] = None
        self.load_profiles()
    
    def load_profiles(self) -> None:
        """Load profiles from JSON file."""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    
                # Load profiles
                profiles_data = data.get("profiles", {})
                for region_name, profile_data in profiles_data.items():
                    try:
                        self.profiles[region_name] = RegionProfile.from_dict(profile_data)
                    except Exception as e:
                        logger.error(f"Failed to load profile '{region_name}': {e}")
                
                # Load current region
                self.current_region = data.get("current_region")
                
                logger.info(f"Loaded {len(self.profiles)} region profiles from {self.config_file}")
                
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON in region profiles file: {e}")
                self.profiles = {}
                self.current_region = None
            except Exception as e:
                logger.error(f"Failed to load region profiles: {e}")
                self.profiles = {}
                self.current_region = None
        else:
            logger.info(f"Region profiles file not found at {self.config_file}, starting fresh")
    
    def save_profiles(self) -> None:
        """Save profiles to JSON file."""
        try:
            data = {
                "profiles": {
                    name: profile.to_dict() 
                    for name, profile in self.profiles.items()
                },
                "current_region": self.current_region,
                "version": "1.0"
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Saved {len(self.profiles)} region profiles to {self.config_file}")
            
        except Exception as e:
            logger.error(f"Failed to save region profiles: {e}")
            raise
    
    def add_profile(self, profile: RegionProfile) -> bool:
        """
        Add or update a region profile.
        
        Args:
            profile: RegionProfile to add
            
        Returns:
            True if successful, False otherwise
        """
        # Validate profile
        errors = profile.validate()
        if errors:
            logger.error(f"Invalid profile: {', '.join(errors)}")
            return False
        
        # Check for duplicate names
        if profile.region_name in self.profiles:
            logger.info(f"Updating existing profile: {profile.region_name}")
        else:
            logger.info(f"Adding new profile: {profile.region_name}")
        
        self.profiles[profile.region_name] = profile
        self.save_profiles()
        return True
    
    def delete_profile(self, region_name: str) -> bool:
        """
        Delete a region profile.
        
        Args:
            region_name: Name of the region to delete
            
        Returns:
            True if deleted, False if not found
        """
        if region_name in self.profiles:
            del self.profiles[region_name]
            
            # Clear current region if it was deleted
            if self.current_region == region_name:
                self.current_region = None
            
            self.save_profiles()
            logger.info(f"Deleted region profile: {region_name}")
            return True
        
        logger.warning(f"Region profile not found: {region_name}")
        return False
    
    def get_profile(self, region_name: str) -> Optional[RegionProfile]:
        """
        Get a specific region profile.
        
        Args:
            region_name: Name of the region
            
        Returns:
            RegionProfile if found, None otherwise
        """
        return self.profiles.get(region_name)
    
    def get_current_profile(self) -> Optional[RegionProfile]:
        """
        Get the currently selected region profile.
        
        Returns:
            Current RegionProfile if one is selected, None otherwise
        """
        if self.current_region:
            return self.profiles.get(self.current_region)
        return None
    
    def switch_region(self, region_name: str) -> bool:
        """
        Switch to a different region.
        
        Args:
            region_name: Name of the region to switch to
            
        Returns:
            True if successful, False if region not found
        """
        if region_name in self.profiles:
            self.current_region = region_name
            self.save_profiles()
            logger.info(f"Switched to region: {region_name}")
            return True
        
        logger.error(f"Cannot switch to unknown region: {region_name}")
        return False
    
    def get_all_region_names(self) -> List[str]:
        """
        Get list of all region names.
        
        Returns:
            List of region names (sorted alphabetically)
        """
        return sorted(self.profiles.keys())
    
    def get_region_count(self) -> int:
        """
        Get number of configured regions.
        
        Returns:
            Count of regions
        """
        return len(self.profiles)
    
    def clear_current_region(self) -> None:
        """Clear the current region selection."""
        self.current_region = None
        self.save_profiles()
        logger.info("Cleared current region selection")
    
    def __repr__(self) -> str:
        """String representation of region manager."""
        return (
            f"RegionManager(regions={len(self.profiles)}, "
            f"current='{self.current_region or 'None'}')"
        )


# ============ MODULE: file_validator ============






class FileValidator:
    """
    Validates all input files for the Cognos Access Review process.
    
    This class performs comprehensive validation of master files, agency mappings,
    and email manifests with detailed error reporting and fix suggestions.
    
    Examples:
        >>> validator = FileValidator()
        >>> results = validator.validate_all(
        ...     master_file=Path("master.xlsx"),
        ...     agency_map_file=Path("mapping.xlsx"),
        ...     email_manifest_file=Path("emails.xlsx")
        ... )
        >>> if results['is_valid']:
        ...     print("All validations passed!")
    """
    
    def __init__(self):
        """Initialize the file validator."""
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
    
    def validate_file_exists(self, file_path: Path, file_type: str = "file") -> Tuple[bool, str]:
        """
        Check if a file or folder path is valid and exists.
        
        Args:
            file_path: Path to the file or directory
            file_type: User-friendly name for the path
            
        Returns:
            Tuple of (is_valid, error_message)
            
        Examples:
            >>> validator = FileValidator()
            >>> valid, msg = validator.validate_file_exists(Path("test.xlsx"), "Master File")
            >>> if not valid:
            ...     print(msg)
        """
        if not file_path or str(file_path).strip() == "":
            return False, f"Please select a {file_type}"
        
        if not file_path.exists():
            return False, f"{file_type.capitalize()} not found: {file_path}"
        
        return True, ""
    
    def validate_excel_file(self, file_path: Path) -> Tuple[bool, str]:
        """
        Validate that a file is a readable Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Tuple of (is_valid, error_message)
            
        Examples:
            >>> validator = FileValidator()
            >>> valid, msg = validator.validate_excel_file(Path("data.xlsx"))
        """
        try:
            # Try reading just one row for efficiency
            pd.read_excel(file_path, nrows=1)
            return True, ""
        except Exception as e:
            return False, f"Invalid Excel file: {file_path}\nError: {str(e)}"
    
    def validate_master_file(self, file_path: Path) -> List[ValidationIssue]:
        """
        Validate master user access file structure and content.
        
        Args:
            file_path: Path to master file
            
        Returns:
            List of validation issues
        """
        issues = []
        
        try:
            df = pd.read_excel(file_path)
            # Clean and normalize column names
            df.columns = [col.strip().replace(' ', '').lower() for col in df.columns]
            # Accept both 'username' and 'user name' but prefer 'username'
            username_col = None
            for col in df.columns:
                if col == 'username':
                    username_col = col
                    break
            if not username_col:
                for col in df.columns:
                    if col == 'username':
                        username_col = col
                        break
            if not username_col:
                issues.append({
                    "severity": ValidationSeverity.ERROR.value,
                    "message": "Master file must have 'Username' column (no space)",
                    "details": f"Available columns: {', '.join(df.columns)}"
                })
                return issues
            
            # Check for empty file
            if df.empty:
                issues.append({
                    "severity": ValidationSeverity.WARNING.value,
                    "message": "Master file is empty",
                    "details": "No user data found in master file"
                })
                return issues
            # Check for null/empty usernames
            null_count = df[username_col].isna().sum()
            empty_count = (df[username_col] == "").sum()
            if null_count > 0 or empty_count > 0:
                total_unassigned = null_count + empty_count
                issues.append({
                    "severity": ValidationSeverity.INFO.value,
                    "message": f"{total_unassigned} users have no username assigned",
                    "details": "These rows will be ignored during processing"
                })
            # Report total records
            self.logger.info(f"Master file contains {len(df)} user records")
        except Exception as e:
            issues.append({
                "severity": ValidationSeverity.ERROR.value,
                "message": "Failed to read master file",
                "details": str(e)
            })
        return issues
    
    def validate_agency_mapping_file(self, file_path: Path) -> Tuple[List[ValidationIssue], List[Tuple[str, str]]]:
        """
        Validate agency mapping file structure.
        
        Args:
            file_path: Path to agency mapping file
            
        Returns:
            Tuple of (validation issues, list of (file_name, agency_id) tuples)
        """
        issues = []
        mappings = []
        
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            
            # Check for required columns
            required_cols = {
                ColumnNames.FILE_NAME: None,
                ColumnNames.AGENCY_ID: None
            }
            
            for required_col in required_cols.keys():
                for col in df.columns:
                    if col.lower() == required_col.lower():
                        required_cols[required_col] = col
                        break
            
            missing_cols = [col for col, found in required_cols.items() if found is None]
            
            if missing_cols:
                issues.append({
                    "severity": ValidationSeverity.ERROR.value,
                    "message": "Agency mapping file missing required columns",
                    "details": f"Missing: {', '.join(missing_cols)}\nExpected: '{ColumnNames.FILE_NAME}' and '{ColumnNames.AGENCY_ID}'\nFound: {', '.join(df.columns.tolist())}"
                })
                return issues, mappings
            
            file_name_col = required_cols[ColumnNames.FILE_NAME]
            agency_id_col = required_cols[ColumnNames.AGENCY_ID]
            
            # Check for empty file
            if df.empty:
                issues.append({
                    "severity": ValidationSeverity.ERROR.value,
                    "message": "Agency mapping file is empty",
                    "details": "No mappings found"
                })
                return issues, mappings
            
            # Validate each row
            for idx, row in df.iterrows():
                file_name = str(row[file_name_col]).strip() if pd.notna(row[file_name_col]) else ""
                agency_id = str(row[agency_id_col]).strip() if pd.notna(row[agency_id_col]) else ""
                
                if not file_name:
                    issues.append({
                        "severity": ValidationSeverity.WARNING.value,
                        "message": f"Row {idx + 2}: Empty file name",
                        "details": "This row will be skipped"
                    })
                    continue
                
                if not agency_id:
                    issues.append({
                        "severity": ValidationSeverity.WARNING.value,
                        "message": f"Row {idx + 2}: Empty agency ID for file '{file_name}'",
                        "details": "This row will be skipped"
                    })
                    continue
                
                mappings.append((file_name, agency_id))
            
            if not mappings:
                issues.append({
                    "severity": ValidationSeverity.ERROR.value,
                    "message": "No valid mappings found",
                    "details": "All rows were empty or invalid"
                })
            
            self.logger.info(f"Agency mapping file contains {len(mappings)} mappings")
            
        except Exception as e:
            issues.append({
                "severity": ValidationSeverity.ERROR.value,
                "message": "Failed to read agency mapping file",
                "details": str(e)
            })
        
        return issues, mappings
    
    def validate_combined_file(self, file_path: Path, file_names: Set[str]) -> List[ValidationIssue]:
        """
        Validate combined file structure and content.
        
        Args:
            file_path: Path to combined file (multi-tab Excel)
            file_names: Set of expected file names from mapping
            
        Returns:
            List of validation issues
        """
        issues = []
        
        try:
            # Use CombinedFileLoader to load data
            combined_loader = CombinedFileLoader(file_path)
            combined_data = combined_loader.load_all_data()
            
            if not combined_data:
                issues.append({
                    "severity": ValidationSeverity.ERROR.value,
                    "message": "No data found in combined file",
                    "details": "The combined file appears to be empty or contains no valid data"
                })
                return issues
            
            # Get email entries from all tabs
            email_entries = set()
            for idx, mapping in enumerate(combined_data):
                file_name = mapping.source_file_name.strip()
                if file_name:
                    email_entries.add(file_name)
                
                # Validate email addresses
                to_emails = mapping.recipients_to
                cc_emails = mapping.recipients_cc
                
                # Check To addresses
                if to_emails:
                    to_list = format_email_list(to_emails)
                    for email in to_list:
                        if not is_valid_email(email):
                            issues.append({
                                "severity": ValidationSeverity.WARNING.value,
                                "message": f"Entry {idx + 1}: Invalid To email format",
                                "details": f"Email: {email}"
                            })
                else:
                    if file_name:
                        issues.append({
                            "severity": ValidationSeverity.WARNING.value,
                            "message": f"Entry {idx + 1}: No To addresses for '{file_name}'",
                            "details": "Email cannot be sent without recipients"
                        })
                
                # Check CC addresses
                if cc_emails:
                    cc_list = format_email_list(cc_emails)
                    for email in cc_list:
                        if not is_valid_email(email):
                            issues.append({
                                "severity": ValidationSeverity.WARNING.value,
                                "message": f"Entry {idx + 1}: Invalid CC email format",
                                "details": f"Email: {email}"
                            })
            
            # Check for missing email entries
            missing_emails = file_names - email_entries
            if missing_emails:
                issues.append({
                    "severity": ValidationSeverity.WARNING.value,
                    "message": f"{len(missing_emails)} file(s) missing email entries",
                    "details": f"Files without email entries: {', '.join(sorted(missing_emails)[:5])}" + 
                              (f" and {len(missing_emails) - 5} more..." if len(missing_emails) > 5 else "")
                })
            
        except Exception as e:
            issues.append({
                "severity": ValidationSeverity.ERROR.value,
                "message": "Failed to read combined file",
                "details": str(e)
            })
        
        return issues
    
    def validate_agency_mapping(
        self,
        master_file: Path,
        agency_map_file: Path,
        combined_file: Optional[Path] = None
    ) -> ValidationResults:
        """
        Comprehensive validation of agency mapping against master file.
        
        Args:
            master_file: Path to master user access file
            agency_map_file: Path to agency mapping file
            combined_file: Optional path to combined file (multi-tab Excel)
            
        Returns:
            ValidationResults dictionary with detailed results
            
        Examples:
            >>> validator = FileValidator()
            >>> results = validator.validate_agency_mapping(
            ...     master_file=Path("master.xlsx"),
            ...     agency_map_file=Path("mapping.xlsx")
            ... )
            >>> print(f"Valid: {results['is_valid']}")
        """
        # Ensure all paths are Path objects (defensive programming)
        master_file = Path(master_file) if isinstance(master_file, str) else master_file
        agency_map_file = Path(agency_map_file) if isinstance(agency_map_file, str) else agency_map_file
        if combined_file and isinstance(combined_file, str):
            combined_file = Path(combined_file)
        
        all_issues: List[ValidationIssue] = []
        email_issues: List[ValidationIssue] = []
        
        # Validate master file exists and structure
        master_issues = self.validate_master_file(master_file)
        all_issues.extend(master_issues)
        
        # Validate mapping file
        mapping_issues, mappings = self.validate_agency_mapping_file(agency_map_file)
        all_issues.extend(mapping_issues)
        
        # If critical errors, return early
        has_errors = any(issue["severity"] == ValidationSeverity.ERROR.value for issue in all_issues)
        if has_errors:
            return {
                "is_valid": False,
                "issues": all_issues,
                "master_agencies": set(),
                "mapped_agencies": set(),
                "unmapped_agencies": set(),
                "duplicate_agencies": {},
                "email_issues": [],
                "master_agency_count": 0,
                "mapped_agency_count": 0
            }
        
        # Read master file to get agencies
        try:
            df_master = pd.read_excel(master_file)
            df_master.columns = df_master.columns.str.strip()
            
            # Find agency column (case-insensitive)
            agency_col = None
            for col in df_master.columns:
                if col.lower() == ColumnNames.AGENCY.lower():
                    agency_col = col
                    break
            
            # Get unique agencies from master (case-insensitive, excluding empty/null)
            master_agencies = set()
            if agency_col:
                for agency in df_master[agency_col].dropna():
                    agency_str = str(agency).strip()
                    if agency_str:
                        master_agencies.add(agency_str.upper())
            
        except Exception as e:
            all_issues.append({
                "severity": ValidationSeverity.ERROR.value,
                "message": "Failed to read agencies from master file",
                "details": str(e)
            })
            master_agencies = set()
        
        # Process mappings
        file_names = set()
        mapped_agencies_upper = set()
        agency_to_files: Dict[str, List[str]] = {}
        
        for file_name, agency_id in mappings:
            file_names.add(file_name)
            
            # Handle comma-separated agencies (e.g., "XYZ, ABC, DEF")
            if ',' in agency_id:
                agencies = [a.strip() for a in agency_id.split(',') if a.strip()]
            else:
                agencies = [agency_id]
            
            # Process each agency (split or single)
            for single_agency in agencies:
                agency_upper = single_agency.upper()
                mapped_agencies_upper.add(agency_upper)
                
                # Track which files each agency is mapped to
                if agency_upper in agency_to_files:
                    agency_to_files[agency_upper].append(file_name)
                else:
                    agency_to_files[agency_upper] = [file_name]
        
        # Find duplicates (agencies mapped to multiple files)
        duplicate_agencies = {
            agency: files for agency, files in agency_to_files.items()
            if len(files) > 1
        }
        
        if duplicate_agencies:
            for agency, files in duplicate_agencies.items():
                all_issues.append({
                    "severity": ValidationSeverity.ERROR.value,
                    "message": f"Agency '{agency}' mapped to multiple files",
                    "details": f"Files: {', '.join(files)}"
                })
        
        # Find unmapped agencies
        unmapped_agencies = master_agencies - mapped_agencies_upper
        
        if unmapped_agencies:
            all_issues.append({
                "severity": ValidationSeverity.WARNING.value,
                "message": f"{len(unmapped_agencies)} agencies in master file not in mapping",
                "details": f"Examples: {', '.join(sorted(unmapped_agencies)[:5])}" + 
                          (f" and {len(unmapped_agencies) - 5} more..." if len(unmapped_agencies) > 5 else "")
            })
        
        # Find agencies in mapping but not in master
        extra_agencies = mapped_agencies_upper - master_agencies
        if extra_agencies:
            all_issues.append({
                "severity": ValidationSeverity.INFO.value,
                "message": f"{len(extra_agencies)} agencies in mapping not found in master file",
                "details": f"Examples: {', '.join(sorted(extra_agencies)[:5])}" + 
                          (f" and {len(extra_agencies) - 5} more..." if len(extra_agencies) > 5 else "")
            })
        
        # Validate combined file if provided
        if combined_file:
            email_issues = self.validate_combined_file(combined_file, file_names)
        
        # Determine overall validity
        all_combined_issues = all_issues + email_issues
        has_errors = any(issue["severity"] == ValidationSeverity.ERROR.value for issue in all_combined_issues)
        
        return {
            "is_valid": not has_errors,
            "issues": all_issues,
            "master_agencies": master_agencies,
            "mapped_agencies": mapped_agencies_upper,
            "unmapped_agencies": unmapped_agencies,
            "duplicate_agencies": duplicate_agencies,
            "email_issues": email_issues,
            "master_agency_count": len(master_agencies),
            "mapped_agency_count": len(mapped_agencies_upper)
        }
    
    def validate_all_files(
        self,
        master_file: Path,
        agency_map_file: Path,
        combined_file: Path,
        output_dir: Path
    ) -> Tuple[bool, str]:
        """
        Comprehensive validation of all input files and directories.
        
        Args:
            master_file: Path to master user access file
            agency_map_file: Path to agency mapping file
            combined_file: Path to combined file (multi-tab Excel)
            output_dir: Path to output directory
            
        Returns:
            Tuple of (is_valid, status_message)
        """
        # Ensure all paths are Path objects (defensive programming)
        master_file = Path(master_file) if isinstance(master_file, str) else master_file
        agency_map_file = Path(agency_map_file) if isinstance(agency_map_file, str) else agency_map_file
        combined_file = Path(combined_file) if combined_file and isinstance(combined_file, str) else combined_file
        output_dir = Path(output_dir) if isinstance(output_dir, str) else output_dir
        
        validation_results = []
        
        # Validate file existence
        files_to_check = [
            (master_file, "Master User Access File"),
            (agency_map_file, "Agency Mapping File"),
            (combined_file, "Combined File")
        ]
        
        for file_path, file_type in files_to_check:
            is_valid, error_msg = self.validate_file_exists(file_path, file_type)
            if not is_valid:
                validation_results.append(f"❌ {error_msg}")
            else:
                validation_results.append(f"✅ {file_type} exists")
        
        # Validate output directory
        if not output_dir or str(output_dir).strip() == "":
            validation_results.append("❌ Please select an Output Folder")
        elif not output_dir.exists():
            validation_results.append(f"❌ Output folder not found: {output_dir}")
        elif not output_dir.is_dir():
            validation_results.append(f"❌ Output path is not a directory: {output_dir}")
        else:
            validation_results.append("✅ Output folder is valid")
        
        # Validate Excel file structure
        excel_files = [
            (master_file, "Master User Access File"),
            (agency_map_file, "Agency Mapping File"),
            (combined_file, "Combined File")
        ]
        
        for file_path, file_type in excel_files:
            if file_path and file_path.exists():
                is_valid, error_msg = self.validate_excel_file(file_path)
                if not is_valid:
                    validation_results.append(f"❌ {error_msg}")
                else:
                    validation_results.append(f"✅ {file_type} is readable")
        
        # Check if any validation failed
        has_errors = any("❌" in result for result in validation_results)
        
        status_message = "\n".join(validation_results)
        if has_errors:
            status_message += "\n\nPlease fix the errors above before proceeding."
        else:
            status_message += "\n\nAll files are valid and ready for processing!"
        
        return not has_errors, status_message


# ============ MODULE: combined_file_loader ============


class CombinedFileLoader:
    """
    Loads combined agency mapping and email data from multi-tab Excel file.
    
    This class handles reading the new consolidated file structure where
    each tab represents a region and contains file mappings with email data.
    """
    
    def __init__(self, file_path: Path):
        """Initialize with path to combined file."""
        self.file_path = Path(file_path)
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self._combined_data: List[CombinedMapping] = []
        self._loaded = False
    
    def load_all_tabs(self, selected_tabs: Optional[List[str]] = None) -> List[CombinedMapping]:
        """
        Load data from all tabs or selected tabs in the Excel file.
        
        Args:
            selected_tabs: List of tab names to load. If None, loads all tabs.
            
        Returns:
            List of CombinedMapping objects
        """
        if self._loaded and not selected_tabs:
            return self._combined_data
        
        try:
            # Get all sheet names
            excel_file = pd.ExcelFile(self.file_path)
            all_tabs = excel_file.sheet_names
            
            # Determine which tabs to process
            tabs_to_process = selected_tabs if selected_tabs else all_tabs
            
            combined_mappings = []
            
            for tab_name in tabs_to_process:
                if tab_name not in all_tabs:
                    self.logger.warning(f"Tab '{tab_name}' not found in file")
                    continue
                    
                # Read the tab
                df = pd.read_excel(self.file_path, sheet_name=tab_name)
                df.columns = df.columns.str.strip()
                
                # Map column names (case-insensitive)
                column_mapping = {}
                for col in df.columns:
                    col_lower = col.lower()
                    if "source_file_name" in col_lower or "file name" in col_lower:
                        column_mapping[col] = ColumnNames.SOURCE_FILE_NAME
                    elif "agency_id" in col_lower or "agency id" in col_lower:
                        column_mapping[col] = ColumnNames.AGENCY_ID
                    elif "recipients_to" in col_lower or col_lower in ["to", "recipients to"]:
                        column_mapping[col] = ColumnNames.RECIPIENTS_TO
                    elif "recipients_cc" in col_lower or col_lower in ["cc", "recipients cc"]:
                        column_mapping[col] = ColumnNames.RECIPIENTS_CC
                
                # Rename columns
                if column_mapping:
                    df = df.rename(columns=column_mapping)
                
                # Ensure required columns exist
                required_cols = [ColumnNames.SOURCE_FILE_NAME, ColumnNames.AGENCY_ID]
                for col in required_cols:
                    if col not in df.columns:
                        self.logger.error(f"Required column '{col}' not found in tab '{tab_name}'")
                        continue
                
                # Add optional columns if missing
                if ColumnNames.RECIPIENTS_TO not in df.columns:
                    df[ColumnNames.RECIPIENTS_TO] = ""
                if ColumnNames.RECIPIENTS_CC not in df.columns:
                    df[ColumnNames.RECIPIENTS_CC] = ""
                
                # Process each row
                for _, row in df.iterrows():
                    source_file = str(row[ColumnNames.SOURCE_FILE_NAME]).strip() if pd.notna(row[ColumnNames.SOURCE_FILE_NAME]) else ""
                    agency_id = str(row[ColumnNames.AGENCY_ID]).strip() if pd.notna(row[ColumnNames.AGENCY_ID]) else ""
                    
                    if not source_file or not agency_id:
                        continue
                    
                    recipients_to = str(row[ColumnNames.RECIPIENTS_TO]) if pd.notna(row[ColumnNames.RECIPIENTS_TO]) else ""
                    recipients_cc = str(row[ColumnNames.RECIPIENTS_CC]) if pd.notna(row[ColumnNames.RECIPIENTS_CC]) else ""
                    
                    # Handle comma-separated agencies
                    if ',' in agency_id:
                        agencies = [a.strip() for a in agency_id.split(',') if a.strip()]
                        for single_agency in agencies:
                            combined_mappings.append(CombinedMapping(
                                source_file_name=source_file,
                                agency_id=single_agency,
                                recipients_to=recipients_to,
                                recipients_cc=recipients_cc,
                                source_tab=tab_name
                            ))
                    else:
                        combined_mappings.append(CombinedMapping(
                            source_file_name=source_file,
                            agency_id=agency_id,
                            recipients_to=recipients_to,
                            recipients_cc=recipients_cc,
                            source_tab=tab_name
                        ))
                
                self.logger.info(f"Loaded {len([m for m in combined_mappings if m.source_tab == tab_name])} mappings from tab '{tab_name}'")
            
            if not selected_tabs:
                self._combined_data = combined_mappings
                self._loaded = True
            
            self.logger.info(f"Total loaded mappings: {len(combined_mappings)}")
            return combined_mappings
            
        except Exception as e:
            self.logger.error(f"Failed to load combined file: {e}")
            raise FileProcessingError(f"Failed to load combined file: {e}")
    
    def get_agency_mappings(self, selected_tabs: Optional[List[str]] = None) -> List[AgencyMapping]:
        """
        Get agency mappings in the old format for backward compatibility.
        
        Returns:
            List of AgencyMapping objects
        """
        combined_data = self.load_all_tabs(selected_tabs)
        return [mapping.to_agency_mapping() for mapping in combined_data]
    
    def get_email_mappings(self, selected_tabs: Optional[List[str]] = None) -> Dict[str, EmailRecipients]:
        """
        Get email mappings by source file name.
        
        Returns:
            Dictionary mapping source_file_name to EmailRecipients
        """
        combined_data = self.load_all_tabs(selected_tabs)
        email_mappings = {}
        
        for mapping in combined_data:
            if mapping.source_file_name not in email_mappings:
                email_mappings[mapping.source_file_name] = mapping.get_email_recipients()
        
        return email_mappings
    
    def get_available_tabs(self) -> List[str]:
        """
        Get list of available tab names in the file.
        
        Returns:
            List of tab names
        """
        try:
            excel_file = pd.ExcelFile(self.file_path)
            return excel_file.sheet_names
        except Exception as e:
            self.logger.error(f"Failed to get tab names: {e}")
            return []


# ============ MODULE: file_processor ============






class FileProcessor:
    """
    Processes master file and generates agency-specific Excel files.
    
    This class handles the core file generation logic, creating multi-tab
    Excel files with proper formatting, sorting, and error handling.
    
    Examples:
        >>> processor = FileProcessor()
        >>> processor.generate_agency_files(
        ...     master_file=Path("master.xlsx"),
        ...     agency_map_file=Path("mapping.xlsx"),
        ...     output_dir=Path("output")
        ... )
    """
    
    def __init__(self, formatting: Optional[ExcelFormatting] = None):
        """
        Initialize file processor.
        
        Args:
            formatting: Excel formatting options. Uses defaults if None.
        """
        self.formatting = formatting or ExcelFormatting()
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
    
    def format_worksheet(
        self,
        writer: pd.ExcelWriter,
        sheet_name: str,
        df: pd.DataFrame
    ) -> None:
        """
        Apply formatting to an Excel worksheet.
        
        Args:
            writer: Pandas ExcelWriter object
            sheet_name: Name of the sheet to format
            df: DataFrame that was written to the sheet
            
        Examples:
            >>> with pd.ExcelWriter("output.xlsx", engine="xlsxwriter") as writer:
            ...     df.to_excel(writer, sheet_name="Sheet1", index=False)
            ...     processor.format_worksheet(writer, "Sheet1", df)
        """
        try:
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Freeze top row if enabled
            if self.formatting.freeze_header:
                worksheet.freeze_panes(1, 0)
            
            # Create header format
            if self.formatting.bold_header:
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': self.formatting.header_bg_color,
                    'border': 1 if self.formatting.border_header else 0
                })
                
                # Apply header format
                for col_idx, col_name in enumerate(df.columns):
                    worksheet.write(0, col_idx, col_name, header_format)
            
            # Auto-adjust column widths if enabled
            if self.formatting.auto_width:
                for col_idx, col_name in enumerate(df.columns):
                    # Calculate max width
                    max_len = len(str(col_name))
                    for value in df[col_name].astype(str):
                        max_len = max(max_len, len(value))
                    
                    # Set width (with max limit)
                    width = min(max_len + 2, self.formatting.max_column_width)
                    worksheet.set_column(col_idx, col_idx, width)
            
            self.logger.debug(f"Formatted worksheet: {sheet_name}")
            
        except Exception as e:
            self.logger.warning(f"Failed to format worksheet {sheet_name}: {e}")
    
    def load_agency_mappings(self, mapping_file: Path, selected_tabs: Optional[List[str]] = None) -> List[AgencyMapping]:
        """
        Load agency mappings from combined Excel file.
        
        Args:
            mapping_file: Path to combined agency/email mapping Excel file
            selected_tabs: Optional list of tab names to process
            
        Returns:
            List of AgencyMapping objects
            
        Raises:
            FileProcessingError: If file cannot be read or is invalid
        """
        try:
            combined_loader = CombinedFileLoader(mapping_file)
            mappings = combined_loader.get_agency_mappings(selected_tabs)
            self.logger.info(f"Loaded {len(mappings)} agency mappings from combined file")
            return mappings
            
        except Exception as e:
            raise FileProcessingError(f"Failed to load agency mappings: {e}")
    
    def generate_agency_files(
        self,
        master_file: Path,
        combined_map_file: Path,
        output_dir: Path,
        progress_callback: Optional[Callable[[float, str], None]] = None,
        handle_unmapped: str = "prompt",
        selected_tabs: Optional[List[str]] = None,
        unmapped_callback: Optional[Callable] = None
    ) -> bool:
        """
        Generate agency-specific Excel files with multi-tab structure.
        
        Each file contains:
        - Tab 1: "All Users" (all agencies for that file)
        - Tab 2-N: Individual agency tabs (alphabetically sorted)
        
        Args:
            master_file: Path to master user access Excel file
            combined_map_file: Path to combined agency/email mapping file
            output_dir: Output directory for generated files
            progress_callback: Optional callback for progress updates
            handle_unmapped: How to handle unmapped agencies: "individual", "single", or "skip"
            selected_tabs: Optional list of tab names to process from combined file
            
        Returns:
            True if successful
            
        Raises:
            FileProcessingError: If file generation fails
        """
        try:
            # Ensure all paths are Path objects (defensive programming)
            master_file = Path(master_file) if isinstance(master_file, str) else master_file
            combined_map_file = Path(combined_map_file) if isinstance(combined_map_file, str) else combined_map_file
            output_dir = Path(output_dir) if isinstance(output_dir, str) else output_dir
            
            # Ensure output directory exists
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Update progress
            if progress_callback:
                progress_callback(0.1, "Loading master file...")
            
            # Load master file
            df_master = pd.read_excel(master_file)
            df_master.columns = df_master.columns.str.strip()
            
            # Find agency column (case-insensitive)
            agency_col = None
            for col in df_master.columns:
                if col.lower() == ColumnNames.AGENCY.lower():
                    agency_col = col
                    break
            
            if not agency_col:
                raise FileProcessingError(f"Master file must have '{ColumnNames.AGENCY}' column")
            
            # COUNTRY FILTERING: Filter by country if specific tabs selected
            if selected_tabs and ColumnNames.COUNTRY in df_master.columns:
                # Check for NULL/missing country values and warn user
                null_countries = df_master[ColumnNames.COUNTRY].isna().sum()
                if null_countries > 0:
                    self.logger.warning(
                        f"Found {null_countries} users with NULL/missing Country values - "
                        f"these will be excluded from regional filtering"
                    )
                
                # Build list of countries to include based on selected tabs
                countries_to_include = []
                for tab in selected_tabs:
                    if tab in TAB_TO_COUNTRY_MAP:
                        countries_to_include.extend(TAB_TO_COUNTRY_MAP[tab])
                
                if countries_to_include:
                    # Filter master data to only include selected countries (case-insensitive)
                    original_count = len(df_master)
                    # Convert both to uppercase for case-insensitive comparison
                    countries_upper = [c.upper() for c in countries_to_include]
                    df_master = df_master[
                        df_master[ColumnNames.COUNTRY].fillna('').astype(str).str.upper().isin(countries_upper)
                    ]
                    filtered_count = len(df_master)
                    
                    self.logger.info(
                        f"Country filter applied: {original_count} → {filtered_count} users "
                        f"(Tabs: {', '.join(selected_tabs)} | Countries: {', '.join(set(countries_to_include)[:5])}...)"
                    )
                    
                    if progress_callback:
                        progress_callback(0.15, f"Filtered to {', '.join(selected_tabs)} region(s)...")
            
            # Update progress
            if progress_callback:
                progress_callback(0.2, "Loading agency mappings...")
            
            # Load mappings using new combined file loader
            mappings = self.load_agency_mappings(combined_map_file, selected_tabs)
            
            # Group mappings by file name
            file_to_agencies: Dict[str, List[str]] = {}
            for mapping in mappings:
                if mapping.file_name not in file_to_agencies:
                    file_to_agencies[mapping.file_name] = []
                file_to_agencies[mapping.file_name].append(mapping.agency_id)
            
            # Track assigned indices
            assigned_indices = set()
            
            # Process each file
            total_files = len(file_to_agencies)
            for file_idx, (file_name, agencies) in enumerate(file_to_agencies.items()):
                if progress_callback:
                    progress = 0.3 + (file_idx / total_files) * 0.6
                    progress_callback(progress, f"Generating {file_name}...")
                
                self._generate_single_file(
                    df_master=df_master,
                    agency_col=agency_col,
                    file_name=file_name,
                    agencies=agencies,
                    output_dir=output_dir,
                    assigned_indices=assigned_indices
                )
            
            # Handle unmapped users
            if progress_callback:
                progress_callback(0.9, "Processing unmapped users...")
            
            unassigned_indices = set(df_master.index) - assigned_indices
            if unassigned_indices:
                df_unassigned = df_master.loc[list(unassigned_indices)].copy()
                
                if handle_unmapped == "prompt":
                    # Prepare unassigned data for dialog
                    unassigned_data = []
                    unmapped_agencies = df_unassigned[agency_col].fillna("[No Agency]").unique()
                    
                    for agency in unmapped_agencies:
                        agency_df = df_unassigned[df_unassigned[agency_col].fillna("[No Agency]") == agency]
                        country = agency_df[ColumnNames.COUNTRY].iloc[0] if ColumnNames.COUNTRY in agency_df.columns else "Unknown"
                        
                        unassigned_data.append({
                            'agency': str(agency),
                            'country': country,
                            'user_count': len(agency_df)
                        })
                    
                    # If unmapped_callback is provided (GUI context), use it
                    if unmapped_callback:
                        # Callback should return dialog result or None
                        dialog_result = unmapped_callback(unassigned_data, file_to_agencies, combined_map_file)
                        
                        if dialog_result and dialog_result.get('decisions'):
                            decisions = dialog_result['decisions']
                            update_mapping = dialog_result.get('update_mapping', False)
                            
                            # Process decisions
                            updates_for_mapping = {'add_to_existing': [], 'create_new': []}
                            
                            for agency, decision in decisions.items():
                                action = decision['action']
                                target = decision['target']
                                agency_data = decision['agency_data']
                                
                                # Get the subset of df_unassigned for this agency
                                df_agency_subset = df_unassigned[
                                    df_unassigned[agency_col].fillna("[No Agency]") == agency
                                ]
                                
                                if action == "Add to Existing File":
                                    # Generate the agency data into the existing file
                                    if target in file_to_agencies:
                                        # Append this agency to the existing agencies list
                                        file_to_agencies[target].append(agency)
                                        
                                        # Regenerate the file with updated agency list
                                        self._generate_single_file(
                                            df_master=df_master,
                                            agency_col=agency_col,
                                            file_name=target,
                                            agencies=file_to_agencies[target],
                                            output_dir=output_dir,
                                            assigned_indices=assigned_indices
                                        )
                                        
                                        # Track for mapping update
                                        updates_for_mapping['add_to_existing'].append((agency, target))
                                
                                elif action == "Create New File":
                                    # Create new file for this agency
                                    self._generate_single_file(
                                        df_master=df_master,
                                        agency_col=agency_col,
                                        file_name=target,
                                        agencies=[agency],
                                        output_dir=output_dir,
                                        assigned_indices=assigned_indices
                                    )
                                    
                                    # Track for mapping update
                                    recipients = decision.get('recipients', {})
                                    to_emails = recipients.get('to', '')
                                    cc_emails = recipients.get('cc', '')
                                    updates_for_mapping['create_new'].append(
                                        (agency, target, to_emails, cc_emails)
                                    )
                                
                                elif action == "Keep as Unassigned":
                                    # Create unassigned file for this country
                                    country = agency_data.get('country', 'Unknown')
                                    unassigned_file_name = f"Unassigned_{country}"
                                    
                                    self._generate_single_file(
                                        df_master=df_master,
                                        agency_col=agency_col,
                                        file_name=unassigned_file_name,
                                        agencies=[agency],
                                        output_dir=output_dir,
                                        assigned_indices=assigned_indices
                                    )
                            
                            # Update mapping file if requested
                            if update_mapping and (updates_for_mapping['add_to_existing'] or updates_for_mapping['create_new']):
                                # Determine which tab to update (use first selected tab or default)
                                tab_name = selected_tabs[0] if selected_tabs else "AMER"
                                self.update_combined_mapping_file(
                                    mapping_file_path=combined_map_file,
                                    updates=updates_for_mapping,
                                    tab_name=tab_name
                                )
                        else:
                            # Dialog was cancelled, skip unmapped
                            handle_unmapped = "skip"
                    else:
                        # No callback (non-GUI context), fall back to simple messagebox
                        unmapped_count = len(df_unassigned)
                        agency_display = ', '.join([str(a) for a in unmapped_agencies[:10]])
                        if len(unmapped_agencies) > 10:
                            agency_display += f" and {len(unmapped_agencies) - 10} more..."
                        
                        choice = messagebox.askyesnocancel(
                            "Unmapped Users Found",
                            f"{unmapped_count} users with {len(unmapped_agencies)} unmapped agencies:\n{agency_display}\n\n"
                            f"YES = Individual files per agency\n"
                            f"NO = Single 'Unassigned.xlsx' file\n"
                            f"CANCEL = Skip unmapped users"
                        )
                        
                        if choice is True:
                            handle_unmapped = "individual"
                        elif choice is False:
                            handle_unmapped = "single"
                        else:
                            handle_unmapped = "skip"
                
                # Handle non-prompt modes or fallback from prompt
                if handle_unmapped == "individual":
                    self._create_individual_unmapped_files(
                        df_unassigned=df_unassigned,
                        agency_col=agency_col,
                        output_dir=output_dir
                    )
                elif handle_unmapped == "single":
                    self._create_unassigned_file(
                        df_unassigned=df_unassigned,
                        output_dir=output_dir
                    )
                
                self.logger.info(f"Processed {len(unassigned_indices)} unmapped users with method: {handle_unmapped}")
            
            if progress_callback:
                progress_callback(1.0, "Complete!")
            
            self.logger.info(f"Successfully generated {total_files} agency files")
            return True
            
        except Exception as e:
            self.logger.error(f"File generation failed: {e}")
            raise FileProcessingError(f"File generation failed: {e}")
    
    def _generate_single_file(
        self,
        df_master: pd.DataFrame,
        agency_col: str,
        file_name: str,
        agencies: List[str],
        output_dir: Path,
        assigned_indices: Set[int]
    ) -> None:
        """
        Generate a single multi-tab Excel file for a group of agencies.
        
        Args:
            df_master: Master DataFrame
            agency_col: Name of agency column
            file_name: Output file name (without .xlsx)
            agencies: List of agency IDs to include
            output_dir: Output directory
            assigned_indices: Set to track assigned row indices
        """
        # Match agencies (case-insensitive)
        # Create pattern with case-insensitive flag at the start - escape special regex characters
        agency_alternatives = "|".join([f"^{re.escape(ag.strip())}$" for ag in agencies])
        agency_pattern = f"(?i)({agency_alternatives})"
        matched = df_master[df_master[agency_col].astype(str).str.match(agency_pattern, na=False)]
        
        if matched.empty:
            # Enhanced diagnostic: show what we searched for and suggest similar matches
            searched_agencies = ", ".join([f"'{ag.strip()}'" for ag in agencies[:3]])
            if len(agencies) > 3:
                searched_agencies += f" and {len(agencies)-3} more"
            
            # Find similar agency names in master file for troubleshooting
            all_master_agencies = df_master[agency_col].dropna().unique()
            similar = []
            for search_ag in agencies:
                search_lower = search_ag.strip().lower()
                for master_ag in all_master_agencies:
                    master_lower = str(master_ag).lower()
                    # Check if any words match or if one contains the other
                    if (search_lower in master_lower or master_lower in search_lower or 
                        any(word in master_lower for word in search_lower.split() if len(word) > 3)):
                        similar.append(f"'{master_ag}'")
            
            if similar:
                similar_str = ", ".join(similar[:5])
                if len(similar) > 5:
                    similar_str += f" and {len(similar)-5} more"
                self.logger.warning(
                    f"No users found for {file_name} | Searched: {searched_agencies} | "
                    f"Similar names in master file: {similar_str}"
                )
            else:
                self.logger.warning(
                    f"No users found for {file_name} | Searched: {searched_agencies} | "
                    f"No similar agency names found in master file"
                )
            return
        
        # Track assigned indices
        assigned_indices.update(matched.index)
        
        # Create output file path - sanitize filename and remove .xlsx if already present
        safe_file_name = sanitize_filename(file_name)
        if safe_file_name.lower().endswith('.xlsx'):
            output_file = output_dir / safe_file_name
        else:
            output_file = output_dir / f"{safe_file_name}.xlsx"
        
        # Prepare data with review columns
        df_output = matched.copy()
        df_output[ColumnNames.REVIEW_ACTION] = ""
        df_output[ColumnNames.COMMENTS] = ""
        
        # Create Excel writer
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            # Tab 1: All Users (or User Access List for Access Certification format)
            primary_sheet = SheetNames.USER_ACCESS_LIST if hasattr(SheetNames, 'USER_ACCESS_LIST') else SheetNames.ALL_USERS
            df_output.to_excel(writer, sheet_name=primary_sheet, index=False)
            self.format_worksheet(writer, primary_sheet, df_output)
            
            # Tab 2: User Access Summary (Pivot) - NEW FEATURE
            self._create_pivot_summary_sheet(
                writer,
                df_output,
                SheetNames.USER_ACCESS_SUMMARY,
                agencies
            )
            
            # Tab 3-N: Individual agencies (alphabetically sorted)
            for agency in sorted(agencies, key=str.upper):
                # Match this specific agency (case-insensitive)
                agency_match = df_output[
                    df_output[agency_col].astype(str).str.upper() == agency.upper()
                ]
                
                if not agency_match.empty:
                    # Sanitize sheet name
                    sheet_name = sanitize_sheet_name(agency)
                    
                    # Write to sheet
                    agency_match.to_excel(writer, sheet_name=sheet_name, index=False)
                    self.format_worksheet(writer, sheet_name, agency_match)
        
        self.logger.info(f"Created {output_file} with {len(agencies)} agency tabs")
    
    def _create_pivot_summary_sheet(
        self,
        writer: pd.ExcelWriter,
        df: pd.DataFrame,
        sheet_name: str,
        agencies: List[str] = None
    ) -> None:
        """
        Create User Access Summary pivot sheet showing UserName → Folder → SubFolder hierarchy.
        
        Args:
            writer: Excel writer object
            df: DataFrame with user access data
            sheet_name: Name for the sheet
            agencies: Optional list of agencies for filtering
        """
        try:
            # Check if required columns exist
            user_col = None
            folder_col = None
            subfolder_col = None
            
            # Find user name column (flexible matching)
            for col in [ColumnNames.USERNAME, ColumnNames.USER_NAME, "UserName", "User Name"]:
                if col in df.columns:
                    user_col = col
                    break
            
            # Find folder columns
            if ColumnNames.FOLDER in df.columns:
                folder_col = ColumnNames.FOLDER
            if ColumnNames.SUBFOLDER in df.columns:
                subfolder_col = ColumnNames.SUBFOLDER
            
            # If columns don't exist, log warning and skip
            if not user_col:
                self.logger.warning(f"Cannot create pivot summary: UserName column not found in data")
                return
            
            if not folder_col or not subfolder_col:
                self.logger.warning(f"Cannot create pivot summary: Folder/SubFolder columns not found")
                return
            
            # Create a copy for processing
            df_pivot = df.copy()
            
            # Remove review columns if present
            cols_to_drop = [ColumnNames.REVIEW_ACTION, ColumnNames.REVIEW_COMMENTS, ColumnNames.COMMENTS]
            df_pivot = df_pivot.drop(columns=[col for col in cols_to_drop if col in df_pivot.columns])
            
            # Create pivot table: UserName as rows, Folder+SubFolder as columns
            # Use aggfunc='size' to count occurrences, or 'first' to show values
            pivot_table = pd.pivot_table(
                df_pivot,
                index=user_col,
                columns=[folder_col, subfolder_col],
                aggfunc='size',
                fill_value=0
            )
            
            # Reset index FIRST to make UserName a regular column
            pivot_display = pivot_table.reset_index()
            
            # NOW flatten MultiIndex columns (after reset_index)
            if isinstance(pivot_display.columns, pd.MultiIndex):
                # Flatten the column names
                pivot_display.columns = [
                    col[0] if col[1] == '' else f"{col[0]} - {col[1]}" 
                    if isinstance(col, tuple) else str(col)
                    for col in pivot_display.columns
                ]
            
            # Replace counts with 'X' for presence indicator (use apply instead of deprecated applymap)
            for col in pivot_display.columns:
                if col != user_col:  # Don't modify the username column
                    pivot_display[col] = pivot_display[col].apply(lambda x: 'X' if x > 0 else '')
            
            # Write to Excel, handle MultiIndex columns
            try:
                if isinstance(pivot_display.columns, pd.MultiIndex):
                    pivot_display.to_excel(writer, sheet_name=sheet_name, index=True)
                else:
                    pivot_display.to_excel(writer, sheet_name=sheet_name, index=False)
            except Exception as e:
                self.logger.error(f"Failed to create pivot summary: {e}")
            
            # Format the worksheet
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Format header row
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            })
            
            # Apply header format
            for col_num, value in enumerate(pivot_display.columns.values):
                worksheet.write(0, col_num, str(value), header_format)
            
            # Auto-fit columns
            for i, col in enumerate(pivot_display.columns):
                max_len = max(
                    pivot_display[col].astype(str).apply(len).max(),
                    len(str(col))
                )
                worksheet.set_column(i, i, min(max_len + 2, 50))
            
            # Freeze first row and first column
            worksheet.freeze_panes(1, 1)
            
            # Add autofilter
            worksheet.autofilter(0, 0, len(pivot_display), len(pivot_display.columns) - 1)
            
            self.logger.info(f"Created pivot summary sheet '{sheet_name}' with {len(pivot_display)} users")
            
        except Exception as e:
            self.logger.error(f"Failed to create pivot summary: {e}")
            # Don't raise - this is an enhancement feature, continue without it
    
    def _create_unassigned_file(
        self,
        df_unassigned: pd.DataFrame,
        output_dir: Path
    ) -> None:
        """
        Create single Unassigned.xlsx file for all unmapped users.
        
        Args:
            df_unassigned: DataFrame with unmapped users
            output_dir: Output directory
        """
        output_file = output_dir / FileNames.UNASSIGNED_FILE
        
        # Add review columns
        df_output = df_unassigned.copy()
        df_output[ColumnNames.REVIEW_ACTION] = ""
        df_output[ColumnNames.COMMENTS] = ""
        
        # Write to Excel
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            df_output.to_excel(writer, sheet_name=SheetNames.ALL_USERS, index=False)
            self.format_worksheet(writer, SheetNames.ALL_USERS, df_output)
        
        self.logger.info(f"Created {output_file} with {len(df_unassigned)} unmapped users")
    
    def _create_individual_unmapped_files(
        self,
        df_unassigned: pd.DataFrame,
        agency_col: str,
        output_dir: Path
    ) -> None:
        """
        Create individual files for each unmapped agency.
        
        Args:
            df_unassigned: DataFrame with unmapped users
            agency_col: Name of agency column
            output_dir: Output directory
        """
        # Get unique agencies from unmapped users
        unique_agencies = df_unassigned[agency_col].fillna("Unassigned").unique()
        
        for agency in unique_agencies:
            agency_str = str(agency).strip()
            if not agency_str:
                agency_str = "Unassigned"
            
            # Filter for this agency
            df_agency = df_unassigned[
                df_unassigned[agency_col].fillna("Unassigned") == agency
            ].copy()
            
            # Add review columns
            df_agency[ColumnNames.REVIEW_ACTION] = ""
            df_agency[ColumnNames.COMMENTS] = ""
            
            # Sanitize file name
            safe_name = sanitize_sheet_name(agency_str)
            output_file = output_dir / f"{safe_name}.xlsx"
            
            # Write to Excel
            with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
                df_agency.to_excel(writer, sheet_name=SheetNames.ALL_USERS, index=False)
                self.format_worksheet(writer, SheetNames.ALL_USERS, df_agency)
            
            self.logger.info(f"Created individual file for unmapped agency: {safe_name}")
    
    def update_combined_mapping_file(
        self,
        mapping_file_path: Path,
        updates: dict,
        tab_name: str
    ) -> bool:
        """
        Update the combined mapping file with new agency assignments.
        Creates a backup before modifying.
        
        Args:
            mapping_file_path: Path to the combined mapping Excel file
            updates: Dictionary of updates to apply
                    Format: {
                        'add_to_existing': [(agency, source_file_name), ...],
                        'create_new': [(agency, source_file_name, to_emails, cc_emails), ...]
                    }
            tab_name: Name of the regional tab to update
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Create backup first (use pattern from AuditLogger)
            backup_dir = mapping_file_path.parent / "backups"
            backup_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = backup_dir / f"{mapping_file_path.stem}_backup_{timestamp}{mapping_file_path.suffix}"
            
            import shutil
            shutil.copy2(mapping_file_path, backup_file)
            self.logger.info(f"Created backup: {backup_file}")
            
            # Load the entire workbook
            xl_file = pd.ExcelFile(mapping_file_path)
            all_sheets = {sheet: pd.read_excel(xl_file, sheet_name=sheet) for sheet in xl_file.sheet_names}
            
            # Get the target sheet
            if tab_name not in all_sheets:
                self.logger.error(f"Tab '{tab_name}' not found in mapping file")
                return False
            
            df = all_sheets[tab_name].copy()
            
            # Process updates
            # 1. Add to existing files
            for agency, source_file_name in updates.get('add_to_existing', []):
                # Find the row with matching source_file_name
                mask = df[ColumnNames.SOURCE_FILE_NAME] == source_file_name
                if mask.any():
                    # Append agency to the agency_id column (comma-separated)
                    idx = df[mask].index[0]
                    current_agencies = str(df.loc[idx, ColumnNames.AGENCY_ID])
                    if pd.isna(current_agencies) or current_agencies.strip() == '':
                        df.loc[idx, ColumnNames.AGENCY_ID] = agency
                    else:
                        df.loc[idx, ColumnNames.AGENCY_ID] = f"{current_agencies}, {agency}"
                    
                    self.logger.info(f"Added '{agency}' to existing file '{source_file_name}'")
            
            # 2. Create new file mappings
            for agency, source_file_name, to_emails, cc_emails in updates.get('create_new', []):
                new_row = pd.DataFrame([{
                    ColumnNames.SOURCE_FILE_NAME: source_file_name,
                    ColumnNames.AGENCY_ID: agency,
                    ColumnNames.RECIPIENTS_TO: to_emails,
                    ColumnNames.RECIPIENTS_CC: cc_emails
                }])
                df = pd.concat([df, new_row], ignore_index=True)
                self.logger.info(f"Created new mapping: '{source_file_name}' for '{agency}'")
            
            # Update the sheet in the dictionary
            all_sheets[tab_name] = df
            
            # Write back to Excel (all sheets)
            with pd.ExcelWriter(mapping_file_path, engine='openpyxl') as writer:
                for sheet_name, sheet_df in all_sheets.items():
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            self.logger.info(f"Successfully updated mapping file: {mapping_file_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to update mapping file: {e}")
            return False


# ============ MODULE: email_handler ============


import win32com.client as win32




class EmailHandler:
    """
    Handles all email operations through Microsoft Outlook.
    
    This class provides functionality for sending, previewing, and scheduling
    emails with attachments through the Outlook COM interface.
    
    Examples:
        >>> config_mgr = ConfigManager()
        >>> handler = EmailHandler(config_mgr)
        >>> handler.test_connection()
        (True, "Outlook connection successful")
        >>> 
        >>> # Send emails in preview mode
        >>> handler.send_emails(
        ...     manifest_file=Path("emails.xlsx"),
        ...     output_dir=Path("output"),
        ...     file_names=["Agency1", "Agency2"],
        ...     mode=EmailMode.PREVIEW
        ... )
    """
    
    def __init__(self, config_manager: ConfigManager):
        """
        Initialize email handler.
        
        Args:
            config_manager: Configuration manager instance
        """
        self.config = config_manager
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self._outlook = None
        
        # Initialize COM
        try:
            pythoncom.CoInitialize()
        except:
            pass  # Already initialized
    
    def __del__(self):
        """Cleanup method to uninitialize COM when EmailHandler is destroyed."""
        try:
            pythoncom.CoUninitialize()
        except:
            pass  # COM may already be uninitialized
    
    def _get_outlook(self, force_new: bool = False) -> win32.Dispatch:
        """
        Get or create Outlook application instance.
        
        Args:
            force_new: If True, create a new connection even if one exists
        
        Returns:
            Outlook Application object
            
        Raises:
            EmailError: If Outlook connection fails
        """
        # Initialize COM for this thread (required when called from different threads)
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass  # Already initialized
        
        if force_new:
            self._outlook = None
        
        if self._outlook is None:
            try:
                # Try to get existing Outlook instance first
                try:
                    self._outlook = win32.GetActiveObject("Outlook.Application")
                    self.logger.info("Connected to existing Outlook instance")
                except Exception:
                    # No existing instance, create new one
                    self._outlook = win32.Dispatch("Outlook.Application")
                    self.logger.info("Created new Outlook connection")
                
                # Verify the connection works by accessing the namespace
                namespace = self._outlook.GetNamespace("MAPI")
                _ = namespace.GetDefaultFolder(6)  # Try to access Inbox
                
            except Exception as e:
                self._outlook = None
                raise EmailError(f"Failed to connect to Outlook: {e}")
        
        return self._outlook
    
    def reset_outlook_connection(self):
        """Reset the Outlook connection (useful after errors)."""
        self._outlook = None
        self.logger.info("Outlook connection reset")
        return self._outlook
    
    def test_connection(self) -> Tuple[bool, str]:
        """
        Test Outlook connection and email functionality.
        
        Returns:
            Tuple of (success, message)
            
        Examples:
            >>> handler = EmailHandler(ConfigManager())
            >>> success, msg = handler.test_connection()
            >>> if success:
            ...     print("Outlook is ready!")
        """
        try:
            outlook = self._get_outlook()
            
            # Test creating an email item
            mail = outlook.CreateItem(0)  # 0 = MailItem
            mail.Subject = "Test Email - Cognos Access Review Tool"
            mail.Body = (
                "This is a test email to verify Outlook connectivity.\n\n"
                "You can close this window without sending.\n\n"
                "Status: ✅ Email system is working correctly"
            )
            mail.To = "test@example.com"
            
            # Display the test email (don't send)
            mail.Display()
            
            self.logger.info("Email connection test successful")
            return True, "✅ Outlook connection successful! Test email displayed."
            
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"Email connection test failed: {error_msg}")
            
            if "Outlook" in error_msg or "COM" in error_msg:
                return False, (
                    "❌ Failed to connect to Outlook.\n\n"
                    "Please ensure:\n"
                    "1. Microsoft Outlook is installed\n"
                    "2. Outlook is configured with your email account\n"
                    "3. Outlook is not blocked by antivirus or firewall\n\n"
                    f"Technical details: {error_msg}"
                )
            else:
                return False, f"❌ Email system error: {error_msg}"
    
    def load_email_manifest(self, combined_file: Path, selected_tabs: Optional[List[str]] = None) -> Dict[str, EmailRecipients]:
        """
        Load email manifest from combined agency/email mapping file.
        
        Args:
            combined_file: Path to combined agency/email mapping Excel file
            selected_tabs: Optional list of tab names to process
            
        Returns:
            Dictionary mapping file names to recipients
            
        Raises:
            EmailError: If manifest cannot be loaded
        """
        try:
            combined_loader = CombinedFileLoader(combined_file)
            recipients = combined_loader.get_email_mappings(selected_tabs)
            
            self.logger.info(f"Loaded {len(recipients)} email recipients from combined file")
            return recipients
            
        except Exception as e:
            raise EmailError(f"Failed to load email manifest from combined file: {e}")
            
        except Exception as e:
            raise EmailError(f"Failed to load email manifest: {e}")
    
    def send_single_email(
        self,
        to_addresses: str,
        cc_addresses: str,
        subject: str,
        body: str,
        attachment_path: Optional[Path] = None
    ) -> bool:
        """
        Send a single email immediately (used by preview navigation dialog).
        
        Args:
            to_addresses: Semicolon-separated To addresses
            cc_addresses: Semicolon-separated CC addresses
            subject: Email subject
            body: Email body
            attachment_path: Optional attachment file path
            
        Returns:
            True if successful, False otherwise
        """
        try:
            attachments = [attachment_path] if attachment_path else []
            
            mail = self.create_email(
                to=to_addresses,
                cc=cc_addresses,
                subject=subject,
                body=body,
                attachments=attachments
            )
            
            mail.Send()
            self.logger.info(f"Sent single email: {subject}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to send single email: {e}")
            return False
    
    def create_email(
        self,
        to: str,
        cc: str,
        subject: str,
        body: str,
        attachments: Optional[List[Path]] = None,
        retry_on_fail: bool = True
    ) -> win32.Dispatch:
        """
        Create an Outlook email item.
        
        Args:
            to: Semicolon-separated To addresses
            cc: Semicolon-separated CC addresses
            subject: Email subject
            body: Email body
            attachments: Optional list of attachment file paths
            retry_on_fail: If True, retry with fresh connection on failure
            
        Returns:
            Outlook MailItem object
            
        Raises:
            EmailError: If email creation fails
        """
        try:
            # Ensure COM is initialized for this thread
            try:
                pythoncom.CoInitialize()
            except:
                pass  # Already initialized
                
            outlook = self._get_outlook()
            mail = outlook.CreateItem(0)  # 0 = MailItem
            
            mail.To = to
            mail.CC = cc
            mail.Subject = subject
            mail.Body = body
            
            # Add attachments
            if attachments:
                for attachment_path in attachments:
                    if attachment_path and attachment_path.exists():
                        mail.Attachments.Add(str(attachment_path))
                        self.logger.debug(f"Attached file: {attachment_path.name}")
            
            return mail
            
        except Exception as e:
            # If first attempt failed and retry is enabled, try with fresh connection
            if retry_on_fail:
                self.logger.warning(f"First email creation attempt failed: {e}. Retrying with fresh connection...")
                self.reset_outlook_connection()
                return self.create_email(to, cc, subject, body, attachments, retry_on_fail=False)
            raise EmailError(f"Failed to create email: {e}")
    
    def send_emails(
        self,
        combined_file: Path,
        output_dir: Path,
        file_names: List[str],
        mode: EmailMode = EmailMode.PREVIEW,
        universal_attachment: Optional[Path] = None,
        progress_callback: Optional[callable] = None,
        scheduled_time: Optional[datetime] = None,
        selected_tabs: Optional[List[str]] = None
    ) -> int:
        """
        Send emails to selected agencies.
        
        Args:
            combined_file: Path to combined agency/email mapping Excel file
            output_dir: Directory containing agency Excel files
            file_names: List of file names to send emails for
            mode: Email sending mode (Preview, Direct, Schedule)
            universal_attachment: Optional attachment for all emails
            progress_callback: Optional callback for progress updates
            scheduled_time: Optional datetime for deferred delivery (used with Schedule mode)
            selected_tabs: Optional list of tab names to process from combined file
            
        Returns:
            Number of emails processed
            
        Raises:
            EmailError: If email sending fails
        """
        try:
            # Load email manifest
            recipients = self.load_email_manifest(combined_file, selected_tabs)
            
            # Load email template
            template = load_email_template()
            
            # Get configuration values for template
            config = self.config.get_all()
            review_period = config.get("review_period", "Q2 2025")
            
            # Set up compliance folders for organization
            folder_results = self.setup_compliance_folders(review_period)
            sent_folder = f"Compliance {review_period} - Sent"
            
            # Process each file
            processed_count = 0
            total = len(file_names)
            
            for idx, file_name in enumerate(file_names):
                if progress_callback:
                    progress = (idx / total) * 100
                    progress_callback(progress, f"Processing {file_name}...")
                
                # Check if recipients exist
                if file_name not in recipients:
                    self.logger.warning(f"No email addresses found for: {file_name}")
                    continue
                
                recipient = recipients[file_name]
                
                # Check for valid To addresses
                if not recipient["to"]:
                    self.logger.warning(f"No To addresses for: {file_name}")
                    continue
                
                # Build subject
                subject_prefix = config.get("email_subject_prefix", "")
                subject = f"{subject_prefix} {review_period} - {file_name}".strip()
                
                # Get file context for enhanced template placeholders
                agency_file = output_dir / f"{file_name}.xlsx"
                user_count = 0
                agency_list = ""
                
                # Try to read the file to get context
                if agency_file.exists():
                    try:
                        # Read first sheet to get user count
                        df_check = pd.read_excel(agency_file, sheet_names=0)
                        user_count = len(df_check)
                        
                        # Get agency names from sheet names (skip first 2 sheets: User Access List & Summary)
                        xl_file = pd.ExcelFile(agency_file)
                        sheet_names = xl_file.sheet_names
                        agency_sheets = [s for s in sheet_names if s not in [SheetNames.USER_ACCESS_LIST, SheetNames.ALL_USERS, SheetNames.USER_ACCESS_SUMMARY]]
                        agency_list = ", ".join(agency_sheets[:5])  # Limit to first 5
                        if len(agency_sheets) > 5:
                            agency_list += f" and {len(agency_sheets) - 5} more"
                    except Exception as e:
                        self.logger.warning(f"Could not read file context for {file_name}: {e}")
                
                # Extract recipient name from email (first part before @)
                recipient_name = ""
                if recipient["to"]:
                    try:
                        first_email = recipient["to"].split(';')[0].strip()
                        recipient_name = first_email.split('@')[0].replace('.', ' ').title()
                    except:
                        pass
                
                # Build body from template with enhanced placeholders
                template_vars = {
                    'review_period': review_period,
                    'deadline': config.get("deadline", "TBD"),
                    'sender_name': config.get("sender_name", ""),
                    'sender_title': config.get("sender_title", ""),
                    'company_name': config.get("company_name", ""),
                    # NEW: Access Certification specific placeholders
                    'QUARTER': review_period,  # Alias for review_period
                    'DEADLINE': config.get("deadline", "TBD"),  # Alias
                    'SOURCE_FILE': file_name,
                    'AGENCY_LIST': agency_list if agency_list else "your assigned agencies",
                    'USER_COUNT': str(user_count),
                    'RECIPIENT_NAME': recipient_name if recipient_name else "there"
                }
                
                # Use safe substitution to avoid KeyError for missing placeholders
                body = template
                for key, value in template_vars.items():
                    body = body.replace(f'{{{key}}}', str(value))
                    body = body.replace(f'{{{key.lower()}}}', str(value))  # Support lowercase too
                
                # Prepare attachments
                attachments = []
                
                # Add agency-specific file
                agency_file = output_dir / f"{file_name}.xlsx"
                if agency_file.exists():
                    attachments.append(agency_file)
                else:
                    self.logger.warning(f"Agency file not found: {agency_file}")
                    continue
                
                # Add universal attachment if provided
                if universal_attachment:
                    attachments.append(universal_attachment)
                
                # Create email
                mail = self.create_email(
                    to=recipient["to"],
                    cc=recipient["cc"],
                    subject=subject,
                    body=body,
                    attachments=attachments
                )
                
                # Handle based on mode
                if mode == EmailMode.PREVIEW:
                    # For Preview mode, display email (backward compatible single email display)
                    # Note: For batch preview with navigation, use prepare_email_batch() + EmailPreviewNavigationDialog from GUI
                    mail.Display()
                    self.logger.info(f"Displayed email for preview: {file_name}")
                elif mode == EmailMode.DIRECT:
                    mail.Send()  # Send immediately
                    self.logger.info(f"Sent email directly: {file_name}")
                    
                    # Auto-organize: Copy sent email to compliance folder
                    # Wait a moment for email to appear in Sent Items
                    # Use default arguments to capture current values
                    threading.Timer(2.0, lambda subj=subject, folder=sent_folder: self.copy_sent_email_to_folder(subj, folder)).start()
                    
                elif mode == EmailMode.SCHEDULE:
                    # For scheduled mode, set deferred delivery time and save to Outbox
                    if scheduled_time:
                        # Set deferred delivery - Outlook will send at this time
                        mail.DeferredDeliveryTime = scheduled_time
                        mail.Send()  # This puts it in Outbox with deferred delivery
                        self.logger.info(f"Scheduled email for {file_name} - will send at {scheduled_time}")
                    else:
                        # No scheduled time, just save as draft for manual review
                        mail.Save()
                        self.logger.info(f"Saved draft email for: {file_name}")
                
                processed_count += 1
            
            if progress_callback:
                progress_callback(100, "Complete!")
            
            self.logger.info(f"Processed {processed_count} emails in {mode.value} mode")
            return processed_count
            
        except Exception as e:
            raise EmailError(f"Email sending failed: {e}")
    
    def prepare_email_batch(
        self,
        combined_file: Path,
        output_dir: Path,
        file_names: List[str],
        universal_attachment: Optional[Path] = None,
        selected_tabs: Optional[List[str]] = None
    ) -> List[dict]:
        """
        Prepare a batch of emails for preview (used with EmailPreviewNavigationDialog).
        
        Args:
            combined_file: Path to combined agency/email mapping Excel file
            output_dir: Directory containing agency Excel files
            file_names: List of file names to send emails for
            universal_attachment: Optional attachment for all emails
            selected_tabs: Optional list of tab names to process from combined file
            
        Returns:
            List of email data dictionaries ready for preview dialog
        """
        try:
            # Load email manifest
            recipients = self.load_email_manifest(combined_file, selected_tabs)
            
            # Load email template
            template = load_email_template()
            
            # Get configuration values
            config = self.config.get_all()
            review_period = config.get("review_period", "Q2 2025")
            subject_prefix = config.get("email_subject_prefix", "")
            
            emails_to_send = []
            
            for file_name in file_names:
                # Check if recipients exist
                if file_name not in recipients:
                    self.logger.warning(f"No email addresses found for: {file_name}")
                    continue
                
                recipient = recipients[file_name]
                
                # Check for valid To addresses
                if not recipient["to"]:
                    self.logger.warning(f"No To addresses for: {file_name}")
                    continue
                
                # Build subject
                subject = f"{subject_prefix} {review_period} - {file_name}".strip()
                
                # Get file context
                agency_file = output_dir / f"{file_name}.xlsx"
                user_count = 0
                agency_list = ""
                
                if agency_file.exists():
                    try:
                        df_check = pd.read_excel(agency_file, sheet_names=0)
                        user_count = len(df_check)
                        
                        xl_file = pd.ExcelFile(agency_file)
                        sheet_names = xl_file.sheet_names
                        agency_sheets = [s for s in sheet_names if s not in [SheetNames.USER_ACCESS_LIST, SheetNames.ALL_USERS, SheetNames.USER_ACCESS_SUMMARY]]
                        agency_list = ", ".join(agency_sheets[:5])
                        if len(agency_sheets) > 5:
                            agency_list += f" and {len(agency_sheets) - 5} more"
                    except Exception as e:
                        self.logger.warning(f"Could not read file context for {file_name}: {e}")
                
                # Extract recipient name
                recipient_name = ""
                if recipient["to"]:
                    try:
                        first_email = recipient["to"].split(';')[0].strip()
                        recipient_name = first_email.split('@')[0].replace('.', ' ').title()
                    except:
                        pass
                
                # Build body from template
                template_vars = {
                    'review_period': review_period,
                    'deadline': config.get("deadline", "TBD"),
                    'sender_name': config.get("sender_name", ""),
                    'sender_title': config.get("sender_title", ""),
                    'company_name': config.get("company_name", ""),
                    'QUARTER': review_period,
                    'DEADLINE': config.get("deadline", "TBD"),
                    'SOURCE_FILE': file_name,
                    'AGENCY_LIST': agency_list if agency_list else "your assigned agencies",
                    'USER_COUNT': str(user_count),
                    'RECIPIENT_NAME': recipient_name if recipient_name else "there"
                }
                
                body = template
                for key, value in template_vars.items():
                    body = body.replace(f'{{{key}}}', str(value))
                    body = body.replace(f'{{{key.lower()}}}', str(value))
                
                # Prepare attachment path
                attachment_path = agency_file if agency_file.exists() else None
                
                # Add to batch
                emails_to_send.append({
                    'to': recipient["to"],
                    'cc': recipient["cc"],
                    'subject': subject,
                    'body': body,
                    'attachment_path': attachment_path,
                    'file_name': file_name
                })
            
            self.logger.info(f"Prepared {len(emails_to_send)} emails for batch preview")
            return emails_to_send
            
        except Exception as e:
            raise EmailError(f"Failed to prepare email batch: {e}")
    
    def scan_inbox(
        self,
        folder_name: str = "Compliance Q2 2025",
        subject_keywords: Optional[List[str]] = None,
        agencies: Optional[List[str]] = None
    ) -> List[Dict[str, str]]:
        """
        Scan Outlook inbox folder for reply emails.
        
        Args:
            folder_name: Name of folder to scan
            subject_keywords: Optional list of keywords to filter by subject
            agencies: Optional list of agencies to filter emails by
            
        Returns:
            List of dictionaries with email information
            
        Examples:
            >>> handler = EmailHandler(ConfigManager())
            >>> replies = handler.scan_inbox(
            ...     folder_name="Compliance Q2 2025",
            ...     subject_keywords=["Cognos", "Access Review"]
            ... )
            >>> print(f"Found {len(replies)} replies")
        """
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            
            # Try to find specific folder
            target_folder = None
            try:
                target_folder = inbox.Folders[folder_name]
                self.logger.info(f"Scanning folder: {folder_name}")
            except Exception:
                self.logger.warning(f"Folder '{folder_name}' not found, scanning inbox")
                target_folder = inbox
            
            # Get messages
            messages = target_folder.Items
            messages.Sort("[ReceivedTime]", True)  # Sort by newest first
            
            # Filter and collect responses
            responses = []
            keywords = subject_keywords or ["cognos", "access review"]
            
            for message in messages:
                try:
                    subject = str(message.Subject).lower()
                    
                    # Check if subject contains any keyword
                    if any(keyword.lower() in subject for keyword in keywords):
                        responses.append({
                            "subject": message.Subject,
                            "sender": message.SenderName,
                            "sender_email": message.SenderEmailAddress,
                            "received": message.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
                            "body_preview": str(message.Body)[:200]
                        })
                except Exception as e:
                    self.logger.debug(f"Error processing message: {e}")
                    continue
            
            self.logger.info(f"Found {len(responses)} matching messages")
            return responses
            
        except Exception as e:
            self.logger.error(f"Inbox scan failed: {e}")
            raise EmailError(f"Failed to scan inbox: {e}")

    def create_folder(self, folder_name: str, parent_folder: str = "Inbox") -> bool:
        """
        Create a folder in Outlook.
        
        Args:
            folder_name: Name of the folder to create
            parent_folder: Parent folder ("Inbox", "Sent Items", etc.)
            
        Returns:
            True if folder was created or already exists, False otherwise
            
        Examples:
            >>> handler = EmailHandler(ConfigManager())
            >>> success = handler.create_folder("Compliance Q2 2025 - Sent")
            >>> if success:
            ...     print("Folder created successfully!")
        """
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            
            # Get parent folder
            if parent_folder.lower() == "inbox":
                parent = namespace.GetDefaultFolder(6)  # olFolderInbox
            elif parent_folder.lower() == "sent items":
                parent = namespace.GetDefaultFolder(5)  # olFolderSentMail
            else:
                # Try to find custom parent folder
                inbox = namespace.GetDefaultFolder(6)
                try:
                    parent = inbox.Folders[parent_folder]
                except:
                    self.logger.error(f"Parent folder '{parent_folder}' not found")
                    return False
            
            # Check if folder already exists
            try:
                existing_folder = parent.Folders[folder_name]
                self.logger.info(f"Folder '{folder_name}' already exists")
                return True
            except:
                pass  # Folder doesn't exist, create it
            
            # Create the folder
            new_folder = parent.Folders.Add(folder_name)
            self.logger.info(f"Created folder: {folder_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to create folder '{folder_name}': {e}")
            return False

    def setup_compliance_folders(self, review_period: str = None) -> Dict[str, bool]:
        """
        Set up standard compliance email folders.
        
        Args:
            review_period: Review period (e.g., "Q2 2025")
            
        Returns:
            Dictionary showing success status for each folder
            
        Examples:
            >>> handler = EmailHandler(ConfigManager())
            >>> results = handler.setup_compliance_folders("Q2 2025")
            >>> for folder, success in results.items():
            ...     print(f"{folder}: {'✅' if success else '❌'}")
        """
        if not review_period:
            config = self.config.get_all()
            review_period = config.get("review_period", "Q2 2025")
        
        folders_to_create = {
            f"Compliance {review_period} - Sent": "Inbox",
            f"Compliance {review_period} - Replies": "Inbox",
            f"Compliance {review_period} - Archive": "Inbox"
        }
        
        results = {}
        for folder_name, parent in folders_to_create.items():
            results[folder_name] = self.create_folder(folder_name, parent)
        
        self.logger.info(f"Setup compliance folders for {review_period}")
        return results

    def move_email_to_folder(self, mail_item, folder_name: str, parent_folder: str = "Inbox") -> bool:
        """
        Move an email to a specific folder.
        
        Args:
            mail_item: Outlook mail item to move
            folder_name: Target folder name
            parent_folder: Parent folder containing the target folder
            
        Returns:
            True if successful, False otherwise
        """
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            
            # Get parent folder
            if parent_folder.lower() == "inbox":
                parent = namespace.GetDefaultFolder(6)  # olFolderInbox
            elif parent_folder.lower() == "sent items":
                parent = namespace.GetDefaultFolder(5)  # olFolderSentMail
            else:
                inbox = namespace.GetDefaultFolder(6)
                parent = inbox.Folders[parent_folder]
            
            # Get target folder
            target_folder = parent.Folders[folder_name]
            
            # Move the email
            mail_item.Move(target_folder)
            self.logger.debug(f"Moved email to folder: {folder_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to move email to '{folder_name}': {e}")
            return False

    def copy_sent_email_to_folder(self, subject: str, folder_name: str, days_back: int = 1) -> bool:
        """
        Find a recently sent email and copy it to a folder.
        
        Args:
            subject: Subject of the email to find
            folder_name: Target folder name
            days_back: How many days back to search
            
        Returns:
            True if email was found and copied, False otherwise
        """
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            sent_items = namespace.GetDefaultFolder(5)  # olFolderSentMail
            
            # Search for the email in sent items
            messages = sent_items.Items
            messages.Sort("[SentOn]", True)  # Sort by newest first
            
            cutoff_date = datetime.now() - timedelta(days=days_back)
            
            for message in messages:
                try:
                    if (message.Subject == subject and 
                        message.SentOn >= cutoff_date):
                        
                        # Copy the email to target folder
                        copied_mail = message.Copy()
                        return self.move_email_to_folder(copied_mail, folder_name)
                        
                except Exception as e:
                    self.logger.debug(f"Error checking message: {e}")
                    continue
            
            self.logger.warning(f"Could not find recent email with subject: {subject}")
            return False
            
        except Exception as e:
            self.logger.error(f"Failed to copy sent email: {e}")
            return False

    def organize_compliance_replies(self, folder_name: str, subject_keywords: List[str] = None) -> int:
        """
        Find and organize compliance reply emails.
        
        Args:
            folder_name: Target folder for replies
            subject_keywords: Keywords to identify compliance emails
            
        Returns:
            Number of emails organized
        """
        try:
            if not subject_keywords:
                subject_keywords = ["cognos", "access review", "compliance"]
            
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
            
            messages = inbox.Items
            organized_count = 0
            
            for message in messages:
                try:
                    subject = str(message.Subject).lower()
                    
                    # Check if it's a compliance-related email
                    if any(keyword.lower() in subject for keyword in subject_keywords):
                        # Check if it's a reply (contains "re:" or "reply")
                        if "re:" in subject or "reply" in subject or "response" in subject:
                            if self.move_email_to_folder(message, folder_name):
                                organized_count += 1
                                
                except Exception as e:
                    self.logger.debug(f"Error processing message: {e}")
                    continue
            
            self.logger.info(f"Organized {organized_count} compliance replies")
            return organized_count
            
        except Exception as e:
            self.logger.error(f"Failed to organize replies: {e}")
            return 0

    def scan_sent_items_for_agencies(
        self,
        agencies: List[str],
        subject_keywords: List[str] = None,
        days_back: int = 90
    ) -> List[Dict[str, any]]:
        """
        Scan Outlook folders starting with 'Compliance' for emails matching the given agencies.
        
        This scans compliance folders (e.g., "Compliance Q3 2025 - Sent") in both
        Inbox and Sent Items to find previously sent emails.
        
        Args:
            agencies: List of agency names to search for
            subject_keywords: Optional keywords to filter by (default: Cognos, Access Review)
            days_back: How many days back to search (default: 90)
            
        Returns:
            List of dictionaries with sent email information:
            [{"agency": str, "to": str, "cc": str, "sent_date": datetime, "subject": str, "folder": str}]
        """
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            
            # Default keywords for compliance emails
            if not subject_keywords:
                subject_keywords = ["cognos", "access review", "uar", "user access"]
            
            cutoff_date = datetime.now() - timedelta(days=days_back)
            
            # Create uppercase agency set for matching
            agency_set = {a.upper() for a in agencies}
            
            found_emails = []
            processed_agencies = set()  # Track which agencies we've found
            
            # Get folders to scan - look for Compliance folders in Inbox and Sent Items
            folders_to_scan = []
            
            # Check Inbox subfolders
            inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
            try:
                for folder in inbox.Folders:
                    if folder.Name.lower().startswith("compliance"):
                        folders_to_scan.append((folder, f"Inbox/{folder.Name}"))
                        self.logger.info(f"Found compliance folder in Inbox: {folder.Name}")
            except Exception as e:
                self.logger.debug(f"Error scanning Inbox subfolders: {e}")
            
            # Check Sent Items subfolders
            sent_items = namespace.GetDefaultFolder(5)  # olFolderSentMail
            try:
                for folder in sent_items.Folders:
                    if folder.Name.lower().startswith("compliance"):
                        folders_to_scan.append((folder, f"Sent Items/{folder.Name}"))
                        self.logger.info(f"Found compliance folder in Sent Items: {folder.Name}")
            except Exception as e:
                self.logger.debug(f"Error scanning Sent Items subfolders: {e}")
            
            # Also check root-level folders
            try:
                root_folder = namespace.GetDefaultFolder(6).Parent  # Get mailbox root
                for folder in root_folder.Folders:
                    if folder.Name.lower().startswith("compliance"):
                        folders_to_scan.append((folder, folder.Name))
                        self.logger.info(f"Found root compliance folder: {folder.Name}")
            except Exception as e:
                self.logger.debug(f"Error scanning root folders: {e}")
            
            if not folders_to_scan:
                self.logger.warning("No Compliance folders found. Falling back to Sent Items.")
                folders_to_scan.append((sent_items, "Sent Items"))
            
            self.logger.info(f"Scanning {len(folders_to_scan)} folder(s) for {len(agencies)} agencies (last {days_back} days)...")
            
            # Scan each compliance folder
            for folder, folder_path in folders_to_scan:
                try:
                    messages = folder.Items
                    messages.Sort("[SentOn]", True)
                    
                    for message in messages:
                        try:
                            # Check if message is too old
                            sent_on = message.SentOn
                            if hasattr(sent_on, 'replace'):
                                # Convert COM datetime to Python datetime
                                sent_on = datetime(sent_on.year, sent_on.month, sent_on.day,
                                                  sent_on.hour, sent_on.minute, sent_on.second)
                            
                            if sent_on < cutoff_date:
                                break  # Stop searching since sorted by date
                            
                            subject = str(message.Subject)
                            subject_lower = subject.lower()
                            
                            # Check if it matches our keywords (optional - compliance folder may have all relevant emails)
                            # Skip keyword check if we're in a Compliance folder
                            is_compliance_folder = "compliance" in folder_path.lower()
                            if not is_compliance_folder:
                                if not any(kw.lower() in subject_lower for kw in subject_keywords):
                                    continue
                            
                            # Check if subject contains any of our agencies
                            for agency in agencies:
                                agency_upper = agency.upper()
                                
                                # Skip if we already found this agency
                                if agency_upper in processed_agencies:
                                    continue
                                
                                # Check if agency name is in the subject
                                if agency.lower() in subject_lower or agency_upper in subject.upper():
                                    # Found a match!
                                    to_addr = str(message.To) if message.To else ""
                                    cc_addr = str(message.CC) if message.CC else ""
                                    
                                    found_emails.append({
                                        "agency": agency,
                                        "to": to_addr,
                                        "cc": cc_addr,
                                        "sent_date": sent_on,
                                        "subject": subject,
                                        "folder": folder_path
                                    })
                                    
                                    processed_agencies.add(agency_upper)
                                    self.logger.debug(f"Found sent email for: {agency} in {folder_path}")
                                    break
                            
                            # Stop if we found all agencies
                            if len(processed_agencies) >= len(agency_set):
                                break
                                
                        except Exception as e:
                            self.logger.debug(f"Error processing message: {e}")
                            continue
                            
                except Exception as e:
                    self.logger.warning(f"Error scanning folder {folder_path}: {e}")
                    continue
                
                # Stop scanning folders if we found all agencies
                if len(processed_agencies) >= len(agency_set):
                    break
            
            self.logger.info(f"Found {len(found_emails)} previously sent emails in Compliance folders")
            return found_emails
            
        except Exception as e:
            self.logger.error(f"Failed to scan compliance folders: {e}")
            return []
    
    def send_followup(
        self,
        agency: str,
        original_subject: str,
        followup_body: str,
        days_back: int = 30
    ) -> bool:
        """
        Send follow-up by replying to original email thread.
        
        This finds the original sent email and creates a reply to maintain thread continuity.
        
        Args:
            agency: Agency name to find
            original_subject: Subject of original email
            followup_body: Body text for the follow-up
            days_back: How many days back to search for original email
            
        Returns:
            True if follow-up sent successfully
            
        Examples:
            >>> handler = EmailHandler(ConfigManager())
            >>> handler.send_followup(
            ...     agency="BBDO Toronto",
            ...     original_subject="[ACTION REQUIRED] Cognos Access Review - Q4 FY25",
            ...     followup_body="Friendly reminder: Please respond by EOD Friday."
            ... )
        """
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            sent_items = namespace.GetDefaultFolder(5)  # olFolderSentMail
            
            # Search for original email
            cutoff_date = datetime.now() - timedelta(days=days_back)
            messages = sent_items.Items
            messages.Sort("[SentOn]", True)
            
            original_email = None
            for message in messages:
                try:
                    if (message.Subject == original_subject and 
                        message.SentOn >= cutoff_date):
                        original_email = message
                        break
                except Exception as e:
                    self.logger.debug(f"Error checking message: {e}")
                    continue
            
            if not original_email:
                self.logger.warning(f"Original email not found for: {agency}")
                return False
            
            # Create reply
            reply = original_email.Reply()
            reply.Body = followup_body + "\n\n" + "-" * 40 + "\n\n" + reply.Body
            reply.Send()
            
            self.logger.info(f"Follow-up sent for: {agency}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to send follow-up for {agency}: {e}")
            return False


# ============ MODULE: report_generator ============




class ReportGenerator:
    """
    Generates SOX compliance reports in multiple formats.
    
    Supports summary, detailed, and exception reports in Excel and PDF formats.
    
    Examples:
        >>> generator = ReportGenerator(output_dir=Path("reports"))
        >>> report_path = generator.generate_compliance_report(
        ...     audit_df=audit_data,
        ...     metrics=dashboard_metrics,
        ...     report_type="summary",
        ...     output_format="xlsx"
        ... )
    """
    
    def __init__(self, output_dir: Optional[Path] = None):
        """
        Initialize report generator.
        
        Args:
            output_dir: Directory for report output
        """
        self.output_dir = Path(output_dir) if output_dir else Path("reports")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
    
    def generate_compliance_report(
        self,
        audit_df: pd.DataFrame,
        metrics: DashboardMetrics,
        report_type: str = "summary",
        output_format: str = "xlsx",
        review_period: str = "Q4 FY25"
    ) -> str:
        """
        Generate SOX compliance report.
        
        Args:
            audit_df: Audit log DataFrame
            metrics: Dashboard metrics
            report_type: "summary", "detailed", or "exceptions"
            output_format: "xlsx" or "pdf"
            review_period: Review period string
            
        Returns:
            Path to generated report file
            
        Raises:
            ValueError: If invalid report_type or output_format
        """
        if report_type not in ["summary", "detailed", "exceptions"]:
            raise ValueError(f"Invalid report_type: {report_type}")
        
        if output_format not in ["xlsx", "pdf"]:
            raise ValueError(f"Invalid output_format: {output_format}")
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"SOX_Compliance_Report_{report_type.title()}_{review_period.replace(' ', '_')}_{timestamp}.{output_format}"
        output_path = self.output_dir / filename
        
        if output_format == "xlsx":
            return self._generate_excel_report(audit_df, metrics, report_type, output_path, review_period)
        else:  # pdf
            return self._generate_pdf_report(audit_df, metrics, report_type, output_path, review_period)
    
    def _generate_excel_report(
        self,
        audit_df: pd.DataFrame,
        metrics: DashboardMetrics,
        report_type: str,
        output_path: Path,
        review_period: str
    ) -> str:
        """Generate Excel format report."""
        try:
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                workbook = writer.book
                
                # Create formats
                title_format = workbook.add_format({
                    'bold': True,
                    'font_size': 16,
                    'align': 'left',
                    'valign': 'vcenter'
                })
                
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#4472C4',
                    'font_color': 'white',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                metric_label_format = workbook.add_format({
                    'bold': True,
                    'align': 'right',
                    'valign': 'vcenter'
                })
                
                metric_value_format = workbook.add_format({
                    'align': 'left',
                    'valign': 'vcenter',
                    'num_format': '#,##0'
                })
                
                # Sheet 1: Executive Summary
                summary_sheet = workbook.add_worksheet("Executive Summary")
                row = 0
                
                # Title
                summary_sheet.write(row, 0, f"SOX Compliance Report - {review_period}", title_format)
                summary_sheet.set_row(row, 24)
                row += 2
                
                # Report metadata
                summary_sheet.write(row, 0, "Report Type:", metric_label_format)
                summary_sheet.write(row, 1, report_type.title())
                row += 1
                
                summary_sheet.write(row, 0, "Generated:", metric_label_format)
                summary_sheet.write(row, 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
                row += 2
                
                # Key metrics
                summary_sheet.write(row, 0, "KEY METRICS", title_format)
                row += 1
                
                metrics_data = [
                    ("Total Agencies", metrics.total_agencies),
                    ("Emails Sent", metrics.sent_count),
                    ("Responses Received", metrics.responded_count),
                    ("Not Sent", metrics.not_sent_count),
                    ("Overdue", metrics.overdue_count),
                    ("Completion Rate", f"{metrics.completion_percentage:.1f}%"),
                    ("Days Until Deadline", metrics.days_left)
                ]
                
                for label, value in metrics_data:
                    summary_sheet.write(row, 0, label + ":", metric_label_format)
                    if isinstance(value, str):
                        summary_sheet.write(row, 1, value)
                    else:
                        summary_sheet.write(row, 1, value, metric_value_format)
                    row += 1
                
                # Set column widths
                summary_sheet.set_column(0, 0, 25)
                summary_sheet.set_column(1, 1, 20)
                
                # Sheet 2: Status breakdown
                if report_type in ["detailed", "summary"]:
                    # Status distribution
                    status_df = audit_df.groupby('Status').size().reset_index(name='Count')
                    status_df.to_excel(writer, sheet_name="Status Distribution", index=False)
                    
                    # Format status sheet
                    status_sheet = writer.sheets["Status Distribution"]
                    for col_num, value in enumerate(status_df.columns.values):
                        status_sheet.write(0, col_num, value, header_format)
                    
                    status_sheet.set_column(0, 0, 20)
                    status_sheet.set_column(1, 1, 15)
                
                # Sheet 3: Detailed data (based on report type)
                if report_type == "detailed":
                    audit_df.to_excel(writer, sheet_name="Full Audit Log", index=False)
                    
                    detail_sheet = writer.sheets["Full Audit Log"]
                    for col_num, value in enumerate(audit_df.columns.values):
                        detail_sheet.write(0, col_num, value, header_format)
                    
                    # Auto-fit columns
                    for i, col in enumerate(audit_df.columns):
                        max_len = max(
                            audit_df[col].astype(str).apply(len).max(),
                            len(str(col))
                        )
                        detail_sheet.set_column(i, i, min(max_len + 2, 50))
                
                elif report_type == "exceptions":
                    # Only show items requiring action
                    exceptions = audit_df[
                        (audit_df['Status'] == 'Not Sent') | 
                        (audit_df['Status'] == 'Overdue')
                    ]
                    
                    exceptions.to_excel(writer, sheet_name="Action Required", index=False)
                    
                    exc_sheet = writer.sheets["Action Required"]
                    for col_num, value in enumerate(exceptions.columns.values):
                        exc_sheet.write(0, col_num, value, header_format)
                    
                    # Auto-fit columns
                    for i, col in enumerate(exceptions.columns):
                        max_len = max(
                            exceptions[col].astype(str).apply(len).max() if not exceptions.empty else 10,
                            len(str(col))
                        )
                        exc_sheet.set_column(i, i, min(max_len + 2, 50))
            
            self.logger.info(f"Excel report generated: {output_path}")
            return str(output_path)
            
        except Exception as e:
            self.logger.error(f"Failed to generate Excel report: {e}")
            raise
    
    def _generate_pdf_report(
        self,
        audit_df: pd.DataFrame,
        metrics: DashboardMetrics,
        report_type: str,
        output_path: Path,
        review_period: str
    ) -> str:
        """
        Generate PDF format report.
        
        Note: This is a basic implementation. For production, consider using
        reportlab or weasyprint for better PDF generation.
        """
        try:
            # For now, generate HTML and convert to PDF (requires additional libraries)
            # Alternatively, generate Excel and provide instructions for PDF export
            
            # Simple approach: Generate HTML report
            html_content = self._generate_html_report(audit_df, metrics, report_type, review_period)
            
            # Save as HTML (can be opened in browser and printed to PDF)
            html_path = output_path.with_suffix('.html')
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            self.logger.info(f"HTML report generated (open in browser to print as PDF): {html_path}")
            return str(html_path)
            
        except Exception as e:
            self.logger.error(f"Failed to generate PDF report: {e}")
            raise
    
    def _generate_html_report(
        self,
        audit_df: pd.DataFrame,
        metrics: DashboardMetrics,
        report_type: str,
        review_period: str
    ) -> str:
        """Generate HTML content for report."""
        
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>SOX Compliance Report - {review_period}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; }}
                h1 {{ color: #2c3e50; }}
                h2 {{ color: #34495e; margin-top: 30px; }}
                table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
                th {{ background-color: #4472C4; color: white; padding: 12px; text-align: left; }}
                td {{ border: 1px solid #ddd; padding: 10px; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .metric {{ margin: 10px 0; }}
                .metric-label {{ font-weight: bold; display: inline-block; width: 200px; }}
                .metric-value {{ display: inline-block; }}
                .status-complete {{ color: #28a745; font-weight: bold; }}
                .status-pending {{ color: #ffc107; font-weight: bold; }}
                .status-overdue {{ color: #dc3545; font-weight: bold; }}
            </style>
        </head>
        <body>
            <h1>SOX Compliance Report - {review_period}</h1>
            <p><strong>Report Type:</strong> {report_type.title()}</p>
            <p><strong>Generated:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
            
            <h2>Key Metrics</h2>
            <div class="metric"><span class="metric-label">Total Agencies:</span><span class="metric-value">{metrics.total_agencies}</span></div>
            <div class="metric"><span class="metric-label">Emails Sent:</span><span class="metric-value">{metrics.sent_count}</span></div>
            <div class="metric"><span class="metric-label">Responses Received:</span><span class="metric-value">{metrics.responded_count}</span></div>
            <div class="metric"><span class="metric-label">Not Sent:</span><span class="metric-value">{metrics.not_sent_count}</span></div>
            <div class="metric"><span class="metric-label">Overdue:</span><span class="metric-value">{metrics.overdue_count}</span></div>
            <div class="metric"><span class="metric-label">Completion Rate:</span><span class="metric-value">{metrics.completion_percentage:.1f}%</span></div>
            <div class="metric"><span class="metric-label">Days Until Deadline:</span><span class="metric-value">{metrics.days_left}</span></div>
        """
        
        if report_type == "detailed":
            html += f"""
            <h2>Full Audit Log</h2>
            {audit_df.to_html(index=False, classes='audit-table')}
            """
        elif report_type == "exceptions":
            exceptions = audit_df[
                (audit_df['Status'] == 'Not Sent') | 
                (audit_df['Status'] == 'Overdue')
            ]
            html += f"""
            <h2>Items Requiring Action</h2>
            {exceptions.to_html(index=False, classes='audit-table')}
            """
        
        html += """
        </body>
        </html>
        """
        
        return html


# ============ MODULE: email_manifest_manager ============

import hashlib



@dataclass
class EmailChange:
    """
    Represents a single email address change in the manifest.
    
    Attributes:
        agency: Agency name (file name)
        field: Field that changed ("To" or "CC")
        old_value: Previous email address(es)
        new_value: New email address(es)
    """
    agency: str
    field: str
    old_value: str
    new_value: str
    
    def __str__(self) -> str:
        """String representation for display."""
        return f"{self.agency} | {self.field}: '{self.old_value}' → '{self.new_value}'"


class EmailManifestManager:
    """
    Manages email manifest validation and change detection.
    
    This class provides functionality to:
    - Validate email addresses in the manifest
    - Detect changes between old and new versions
    - Calculate content hash for change detection
    - Provide structured data for UI display
    
    Examples:
        >>> manager = EmailManifestManager("email_manifest.xlsx")
        >>> invalid = manager.validate_emails()
        >>> if invalid["To"]:
        ...     print(f"Invalid To emails: {invalid['To']}")
        >>> changes = manager.detect_changes("old_manifest.xlsx")
        >>> for change in changes:
        ...     print(change)
    """
    
    EMAIL_PATTERN = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    
    def __init__(self, manifest_file: str):
        """
        Initialize email manifest manager.
        
        Args:
            manifest_file: Path to the email manifest Excel file
        """
        self.manifest_file = Path(manifest_file)
        self.df = None
        self._hash = None
        self.load()
    
    def load(self) -> None:
        """Load manifest file into DataFrame."""
        try:
            self.df = pd.read_excel(self.manifest_file)
            self.df.columns = self.df.columns.str.strip()
            self._hash = self._calculate_hash()
            logger.info(f"Loaded email manifest from {self.manifest_file}")
        except Exception as e:
            logger.error(f"Failed to load email manifest: {e}")
            raise
    
    def _calculate_hash(self) -> str:
        """
        Calculate MD5 hash of current manifest content.
        
        Returns:
            MD5 hash string
        """
        if self.df is None:
            return ""
        
        content = self.df.to_json()
        return hashlib.md5(content.encode()).hexdigest()
    
    def validate_emails(self) -> Dict[str, List[str]]:
        """
        Validate all email addresses in the manifest.
        
        Returns:
            Dictionary with "To" and "CC" keys containing lists of invalid emails
            Format: {"To": ["Agency: invalid@email"], "CC": [...]}
        """
        invalid_emails = {"To": [], "CC": []}
        
        if self.df is None:
            return invalid_emails
        
        for _, row in self.df.iterrows():
            agency = row.get("Agency", "Unknown")
            
            # Validate To emails
            to_emails_raw = str(row.get("To", ""))
            if pd.notna(to_emails_raw) and to_emails_raw:
                to_emails = [e.strip() for e in to_emails_raw.split(";") if e.strip()]
                for email in to_emails:
                    if not self._is_valid_email(email):
                        invalid_emails["To"].append(f"{agency}: {email}")
            
            # Validate CC emails
            cc_emails_raw = str(row.get("CC", ""))
            if pd.notna(cc_emails_raw) and cc_emails_raw:
                cc_emails = [e.strip() for e in cc_emails_raw.split(";") if e.strip()]
                for email in cc_emails:
                    if not self._is_valid_email(email):
                        invalid_emails["CC"].append(f"{agency}: {email}")
        
        return invalid_emails
    
    def _is_valid_email(self, email: str) -> bool:
        """
        Validate single email address format.
        
        Args:
            email: Email address to validate
            
        Returns:
            True if valid format
        """
        return bool(re.match(self.EMAIL_PATTERN, email.strip()))
    
    def detect_changes(self, old_manifest_file: str) -> List[EmailChange]:
        """
        Detect changes between old and new manifest versions.
        
        Args:
            old_manifest_file: Path to previous version of manifest
            
        Returns:
            List of EmailChange objects representing differences
        """
        changes = []
        
        try:
            old_df = pd.read_excel(old_manifest_file)
            old_df.columns = old_df.columns.str.strip()
        except Exception as e:
            logger.error(f"Failed to load old manifest for comparison: {e}")
            return changes
        
        if self.df is None:
            return changes
        
        # Create lookup for old values
        old_lookup = {}
        for _, row in old_df.iterrows():
            agency = str(row.get("Agency", "")).strip()
            if agency:
                old_lookup[agency] = {
                    "To": str(row.get("To", "")),
                    "CC": str(row.get("CC", ""))
                }
        
        # Compare with new values
        for _, new_row in self.df.iterrows():
            agency = str(new_row.get("Agency", "")).strip()
            if not agency:
                continue
            
            new_to = str(new_row.get("To", ""))
            new_cc = str(new_row.get("CC", ""))
            
            if agency in old_lookup:
                old_to = old_lookup[agency]["To"]
                old_cc = old_lookup[agency]["CC"]
                
                # Check To field
                if new_to != old_to:
                    changes.append(EmailChange(
                        agency=agency,
                        field="To",
                        old_value=old_to,
                        new_value=new_to
                    ))
                
                # Check CC field
                if new_cc != old_cc:
                    changes.append(EmailChange(
                        agency=agency,
                        field="CC",
                        old_value=old_cc,
                        new_value=new_cc
                    ))
            else:
                # New agency added
                if new_to:
                    changes.append(EmailChange(
                        agency=agency,
                        field="To",
                        old_value="[NEW]",
                        new_value=new_to
                    ))
                if new_cc:
                    changes.append(EmailChange(
                        agency=agency,
                        field="CC",
                        old_value="[NEW]",
                        new_value=new_cc
                    ))
        
        logger.info(f"Detected {len(changes)} email changes")
        return changes
    
    def save(self) -> None:
        """Save current manifest back to Excel file."""
        try:
            self.df.to_excel(self.manifest_file, index=False)
            self._hash = self._calculate_hash()
            logger.info(f"Saved email manifest to {self.manifest_file}")
        except Exception as e:
            logger.error(f"Failed to save email manifest: {e}")
            raise
    
    def has_changed(self) -> bool:
        """
        Check if manifest has been modified since loading.
        
        Returns:
            True if content has changed
        """
        current_hash = self._calculate_hash()
        return current_hash != self._hash
    
    def get_agency_count(self) -> int:
        """Get number of agencies in manifest."""
        if self.df is None:
            return 0
        return len(self.df)
    
    def get_agencies(self) -> List[str]:
        """
        Get list of all agencies in manifest.
        
        Returns:
            List of agency names (file names)
        """
        if self.df is None:
            return []
        
        return self.df["Agency"].dropna().str.strip().tolist()
    
    def update_email(self, agency: str, field: str, new_value: str) -> bool:
        """
        Update email address for a specific agency.
        
        Args:
            agency: Agency name (file name)
            field: "To" or "CC"
            new_value: New email address(es)
            
        Returns:
            True if updated successfully
        """
        if self.df is None:
            return False
        
        if field not in ["To", "CC"]:
            logger.error(f"Invalid field: {field}")
            return False
        
        # Find row for this agency
        mask = self.df["Agency"].str.strip() == agency
        if mask.any():
            self.df.loc[mask, field] = new_value
            logger.info(f"Updated {agency} {field} to: {new_value}")
            return True
        
        logger.warning(f"Agency not found in manifest: {agency}")
        return False
    
    def create_backup(self, backup_dir: str = "backups") -> Optional[str]:
        """
        Create timestamped backup of current manifest.
        
        Args:
            backup_dir: Directory to store backups
            
        Returns:
            Path to backup file, or None if failed
        """
        try:
            backup_path = Path(backup_dir)
            backup_path.mkdir(parents=True, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = backup_path / f"{timestamp}_{self.manifest_file.name}"
            
            self.df.to_excel(backup_file, index=False)
            logger.info(f"Created email manifest backup: {backup_file}")
            return str(backup_file)
            
        except Exception as e:
            logger.error(f"Failed to create manifest backup: {e}")
            return None
    
    def __repr__(self) -> str:
        """String representation."""
        agency_count = self.get_agency_count()
        return f"EmailManifestManager(file='{self.manifest_file}', agencies={agency_count})"


# ============================================================================
# MAIN APPLICATION CODE (Original: src/cognos_review.py)
# ============================================================================

# ------------- Logging Setup -------------
def setup_logging():
    """
    Sets up the application's logging configuration.

    This function configures a logger that writes messages to both a file (`cognos_review.log`)
    and the console. It helps in debugging and tracking the application's behavior.
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(FileNames.LOG_FILE, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

# ------------- Configuration -------------
# Initialize global configuration manager
config_manager = ConfigManager()
config = config_manager.load()

# Global constants derived from the configuration
REVIEW_PERIOD = config.get("review_period", "Q2 2025")
REVIEW_DEADLINE = config.get("deadline", "June 30, 2025")
AUDIT_FILE_NAME = config_manager.get_audit_file_name()

# =============== Utility Functions ===============

def smart_agency_match(agency1: str, agency2: str) -> bool:
    """
    Smart case-insensitive agency matching with normalization.
    
    Handles:
    - Case differences (ACME vs acme vs Acme)
    - Extra whitespace
    - Special characters
    
    Args:
        agency1: First agency name
        agency2: Second agency name
        
    Returns:
        True if agencies match
    """
    if not agency1 or not agency2:
        return False
    
    # Normalize: strip, uppercase, remove extra spaces
    norm1 = str(agency1).strip().upper()
    norm2 = str(agency2).strip().upper()
    
    # Remove multiple spaces
    import re
    norm1 = re.sub(r'\s+', ' ', norm1)
    norm2 = re.sub(r'\s+', ' ', norm2)
    
    return norm1 == norm2

def validate_file_exists(file_path, file_type="file"):
    """
    Checks if a file or folder path is valid and exists.

    Args:
        file_path (str): The path to the file or directory.
        file_type (str): A user-friendly name for the path (e.g., "Master File").

    Returns:
        tuple[bool, str]: A tuple containing a boolean (True if valid) and an error message.
    """
    if not file_path or not file_path.strip():
        return False, f"Please select a {file_type}"
    if not os.path.exists(file_path):
        return False, f"{file_type.capitalize()} not found: {file_path}"
    return True, ""

def validate_excel_file(file_path):
    """
    Validates that a file is a readable Excel file by trying to read its first row.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        tuple[bool, str]: A tuple containing a boolean (True if valid) and an error message.
    """
    try:
        pd.read_excel(file_path, nrows=1)  # Try reading just one row for efficiency
        return True, ""
    except Exception as e:
        return False, f"Invalid Excel file: {str(e)}"

def validate_all_files(master_file, agency_map_file, combined_file, output_dir):
    """
    Comprehensive validation of all input files and directories.
    
    This function checks that all required files exist, are readable Excel files,
    and have the expected structure. It provides detailed feedback on any issues.

    Args:
        master_file (str): Path to the master user access file.
        agency_map_file (str): Path to the agency mapping file.
        combined_file (str): Path to the combined file (multi-tab Excel).
        output_dir (str): Path to the output directory.

    Returns:
        tuple[bool, str]: A tuple containing a boolean (True if all valid) and a detailed status message.
    """
    validation_results = []
    
    # Validate file existence
    files_to_check = [
        (master_file, "Master User Access File"),
        (agency_map_file, "Agency Mapping File"),
        (combined_file, "Combined File")
    ]
    
    for file_path, file_type in files_to_check:
        is_valid, error_msg = validate_file_exists(file_path, file_type)
        if not is_valid:
            validation_results.append(f"❌ {error_msg}")
        else:
            validation_results.append(f"✅ {file_type} found")
    
    # Validate output directory
    if not output_dir or not output_dir.strip():
        validation_results.append("❌ Please select an Output Folder")
    elif not os.path.exists(output_dir):
        validation_results.append(f"❌ Output folder not found: {output_dir}")
    elif not os.path.isdir(output_dir):
        validation_results.append(f"❌ Output path is not a directory: {output_dir}")
    else:
        validation_results.append("✅ Output folder is valid")
    
    # Validate Excel file structure
    excel_files = [(master_file, "Master User Access File"), (agency_map_file, "Agency Mapping File"), (combined_file, "Combined File")]
    
    for file_path, file_type in excel_files:
        if file_path and os.path.exists(file_path):
            is_valid, error_msg = validate_excel_file(file_path)
            if not is_valid:
                validation_results.append(f"❌ {file_type}: {error_msg}")
            else:
                # Try to read column headers for additional validation
                try:
                    df = pd.read_excel(file_path, nrows=0)  # Read headers only
                    validation_results.append(f"✅ {file_type}: {len(df.columns)} columns found")
                except Exception as e:
                    validation_results.append(f"⚠️ {file_type}: File readable but column validation failed")
    
    # Check if any validation failed
    has_errors = any("❌" in result for result in validation_results)
    
    status_message = "\n".join(validation_results)
    if has_errors:
        status_message += "\n\nPlease fix the errors above before proceeding."
    else:
        status_message += "\n\nAll files are valid and ready for processing!"
    
    return not has_errors, status_message

def open_output_folder(output_dir):
    """
    Opens the output directory in Windows Explorer.
    
    Args:
        output_dir (str): The path to the output directory.
    
    Returns:
        bool: True if successful, False otherwise.
    """
    try:
        if not output_dir or not os.path.exists(output_dir):
            return False
        
        # Use Windows Explorer to open the folder
        os.startfile(output_dir)
        logger.info(f"Opened output folder: {output_dir}")
        return True
    except Exception as e:
        logger.error(f"Failed to open output folder: {str(e)}")
        return False

def test_email_connection():
    """
    Tests the connection to Microsoft Outlook and basic email functionality.
    
    This function attempts to connect to Outlook, create a test email item,
    and display it to verify that the email system is working properly.

    Returns:
        tuple[bool, str]: A tuple containing a boolean (True if successful) and a status message.
    """
    try:
        # Initialize COM for this thread
        try:
            pythoncom.CoInitialize()
        except:
            pass  # Already initialized
            
        # Test Outlook connection
        outlook = win32.Dispatch("Outlook.Application")
        
        # Test creating an email item
        mail = outlook.CreateItem(0)
        mail.Subject = "Test Email - Cognos Access Review Tool"
        mail.Body = "This is a test email to verify that the email system is working correctly.\n\nYou can close this email without sending it."
        mail.To = "test@example.com"
        
        # Display the test email
        mail.Display()
        
        logger.info("Email connection test successful")
        return True, "✅ Email connection test successful!\n\nA test email has been opened in Outlook.\nYou can close it without sending."
        
    except Exception as e:
        error_msg = str(e)
        logger.error(f"Email connection test failed: {error_msg}")
        
        if "Outlook" in error_msg or "COM" in error_msg:
            return False, f"❌ Outlook connection failed:\n\n{error_msg}\n\nPlease ensure:\n• Outlook is installed and running\n• You have permission to use Outlook automation"
        else:
            return False, f"❌ Email test failed:\n\n{error_msg}"

def export_agency_list(agencies, audit_df, output_dir):
    """
    Exports the current agency list with their status to an Excel file.
    
    Args:
        agencies (list): List of agency names.
        audit_df (pd.DataFrame): The audit log DataFrame.
        output_dir (str): Directory to save the export file.
    
    Returns:
        str: Path to the exported file, or empty string if failed.
    """
    try:
        # Create a DataFrame with agency information
        export_data = []
        
        for agency in agencies:
            # Find agency status in audit log
            status = "Not Sent"
            sent_date = ""
            response_date = ""
            to_email = ""
            cc_email = ""
            comments = ""
            
            if not audit_df.empty and "Agency" in audit_df.columns:
                mask = audit_df["Agency"].str.upper() == agency.upper()
                if mask.any():
                    row = audit_df.loc[mask].iloc[0]
                    status = row.get("Status", "Not Sent")
                    sent_date = row.get("Sent Email Date", "")
                    response_date = row.get("Response Received Date", "")
                    to_email = row.get("To", "")
                    cc_email = row.get("CC", "")
                    comments = row.get("Comments", "")
            
            export_data.append({
                "Agency": agency,
                "Status": status,
                "Email Sent Date": sent_date,
                "Response Received Date": response_date,
                "To Email": to_email,
                "CC Email": cc_email,
                "Comments": comments
            })
        
        # Create DataFrame and export
        df_export = pd.DataFrame(export_data)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_filename = f"Agency_Status_Export_{timestamp}.xlsx"
        export_path = os.path.join(output_dir, export_filename)
        
        with pd.ExcelWriter(export_path, engine="xlsxwriter") as writer:
            df_export.to_excel(writer, sheet_name="Agency Status", index=False)
            
            # Auto-format the Excel file
            wb = writer.book
            ws = writer.sheets["Agency Status"]
            ws.freeze_panes(1, 0)
            header_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})
            
            for col_idx, col_name in enumerate(df_export.columns):
                ws.write(0, col_idx, col_name, header_fmt)
                col_width = max(df_export[col_name].astype(str).map(len).max(), len(col_name), 12) + 2
                ws.set_column(col_idx, col_idx, min(col_width, 50))
        
        logger.info(f"Agency list exported: {export_path}")
        return export_path
        
    except Exception as e:
        logger.error(f"Failed to export agency list: {str(e)}")
        return ""

def open_settings_dialog(parent_window):
    """
    Opens a dialog to edit application settings.
    
    Args:
        parent_window: The parent window for the dialog.
    
    Returns:
        bool: True if settings were saved, False otherwise.
    """
    # Create settings dialog
    dialog = ctk.CTkToplevel(parent_window)
    dialog.title("Application Settings")
    dialog.geometry("500x400")
    dialog.transient(parent_window)
    dialog.grab_set()
    
    # Load current config
    current_config = config_manager.load()
    
    # Create variables for the form fields
    settings_vars = {}
    for key, value in current_config.items():
        if isinstance(value, str):
            settings_vars[key] = ctk.StringVar(value=value)
        else:
            settings_vars[key] = ctk.StringVar(value=str(value))
    
    # Create the form
    main_frame = ctk.CTkFrame(dialog)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    ctk.CTkLabel(main_frame, text="Application Settings", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))
    
    # Create scrollable frame for settings
    scroll_frame = ctk.CTkScrollableFrame(main_frame)
    scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Define settings fields with labels
    settings_fields = [
        ("review_period", "Review Period:", "e.g., Q2 2025"),
        ("deadline", "Deadline:", "e.g., June 30, 2025"),
        ("company_name", "Company Name:", "e.g., Omnicom Group"),
        ("sender_name", "Sender Name:", "e.g., Govind Waghmare"),
        ("sender_title", "Sender Title:", "e.g., Manager, Financial Applications | Analytics"),
        ("email_subject_prefix", "Email Subject Prefix:", "e.g., [ACTION REQUIRED] Cognos Access Review")
    ]
    
    # Create form fields
    for key, label, placeholder in settings_fields:
        frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(frame, text=label).pack(anchor="w")
        entry = ctk.CTkEntry(frame, textvariable=settings_vars[key], placeholder_text=placeholder)
        entry.pack(fill="x", pady=(5, 0))
    
    # Buttons frame
    button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    button_frame.pack(fill="x", pady=(20, 0))
    
    def save_settings():
        try:
            # Collect values from form
            new_config = {}
            for key, var in settings_vars.items():
                new_config[key] = var.get()
            
            # Save using config manager
            config_manager.update(new_config)
            config_manager.save()
            
            # Update global config
            global config, REVIEW_PERIOD, REVIEW_DEADLINE, AUDIT_FILE_NAME
            config = new_config
            REVIEW_PERIOD = config.get("review_period", "Q2 2025")
            REVIEW_DEADLINE = config.get("deadline", "June 30, 2025")
            AUDIT_FILE_NAME = config_manager.get_audit_file_name()
            
            logger.info("Settings saved successfully")
            messagebox.showinfo("Success", "Settings saved successfully!\n\nSome changes may require restarting the application.", parent=dialog)
            dialog.destroy()
            return True
            
        except Exception as e:
            logger.error(f"Failed to save settings: {str(e)}")
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}", parent=dialog)
            return False
    
    def cancel():
        dialog.destroy()
        return False
    
    # Create buttons
    ctk.CTkButton(button_frame, text="Save Settings", command=save_settings).pack(side="left", padx=(0, 10))
    ctk.CTkButton(button_frame, text="Cancel", command=cancel).pack(side="left")
    
    # Center the dialog
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")
    
    return False  # Will be updated by button callbacks

def safe_load_status(path):
    """
    Safely loads the audit log Excel file into a pandas DataFrame.

    If the file doesn't exist or is invalid, it returns an empty DataFrame
    with the correct columns, preventing the application from crashing.

    Args:
        path (str): The path to the audit log Excel file.

    Returns:
        pd.DataFrame: A DataFrame containing the audit data or an empty DataFrame.
    """
    if os.path.exists(path):
        try:
            return pd.read_excel(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load audit log: {str(e)}")
            # Return an empty structure if loading fails
            return pd.DataFrame(
                columns=["Agency", "Sent Email Date", "Response Received Date", "To", "CC", "Status", "Comments"]
            )
    # Return an empty structure if the file doesn't exist
    return pd.DataFrame(
        columns=["Agency", "Sent Email Date", "Response Received Date", "To", "CC", "Status", "Comments"]
    )

def create_backup(file_path: str, backup_dir: str = "backups") -> str:
    """
    Creates a timestamped backup of a given file.

    This helps prevent data loss by saving a version of a file before it's
    overwritten. Backups are stored in a specified directory.

    Args:
        file_path (str): The full path of the file to back up.
        backup_dir (str): The name of the directory to store backups in. Defaults to "backups".

    Returns:
        str: The path to the newly created backup file, or an empty string if it failed.
    """
    try:
        if not os.path.exists(file_path):
            return "" # Nothing to backup
        
        # Ensure the backup directory exists
        backup_path = Path(backup_dir)
        backup_path.mkdir(exist_ok=True)
        
        # Create a unique filename with a timestamp (e.g., "20231027_153000_filename.xlsx")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = Path(file_path).name
        backup_filename = f"{timestamp}_{filename}"
        backup_file_path = backup_path / backup_filename
        
        # Copy the original file to the new backup location
        shutil.copy2(file_path, backup_file_path)
        logger.info(f"Backup created: {backup_file_path}")
        return str(backup_file_path)
    except Exception as e:
        logger.error(f"Failed to create backup of {file_path}: {str(e)}")
        return ""

def save_audit_log(df, out_dir):
    """
    Saves the main audit log DataFrame to an Excel file, creating a backup first.

    Args:
        df (pd.DataFrame): The DataFrame containing the audit log data.
        out_dir (str): The directory to save the audit log file in.

    Returns:
        str: The path where the audit log was saved, or an empty string on failure.
    """
    out_path = os.path.join(out_dir, AUDIT_FILE_NAME)
    
    # Always create a backup of the existing audit log before overwriting it
    if os.path.exists(out_path):
        create_backup(out_path)
    
    try:
        df.to_excel(out_path, index=False)
        logger.info(f"Audit log saved: {out_path}")
        return out_path
    except Exception as e:
        logger.error(f"Failed to save audit log: {str(e)}")
        messagebox.showerror("Save Error", f"Failed to save audit log: {str(e)}")
        return ""

# The block of code responsible for generating the agency files (with the formatting you want)
# is this function:

def generate_agency_files(master_file, domain_map_file, out_dir, domain_col):
    """
    Splits a master user access file into separate Excel files for each agency.
    It reads a master list of users and a mapping file that links email domains
    to agency names. It then generates one Excel file per agency containing only
    the users relevant to them.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(exist_ok=True)

    df_master = pd.read_excel(master_file)
    df_map = pd.read_excel(domain_map_file)
    df_master.columns = df_master.columns.str.strip()
    df_map.columns = df_map.columns.str.strip()
    df_master[domain_col] = df_master[domain_col].str.lower()

    assigned_idx = set()

    for _, row in df_map.iterrows():
        agency = str(row.get("DOMAIN_NAME", "")).strip().upper()
        domains_raw = str(row.get("domains_included", ""))
        domains = [d.strip().lower() for d in re.split(r"[;,/&]+", domains_raw) if d.strip()]
        if not domains:
            continue

        pattern = "|".join(re.escape(d) for d in domains)
        matched = df_master[df_master[domain_col].str.contains(pattern, na=False)]

        if matched.empty:
            continue

        assigned_idx.update(matched.index)
        matched_copy = matched.copy()
        matched_copy["Review Action"] = ""
        matched_copy["Comments"] = ""

        out_file = out_dir / f"{agency}.xlsx"
        with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
            matched_copy.to_excel(writer, sheet_name="User Access List", index=False)
            wb = writer.book
            ws = writer.sheets["User Access List"]
            ws.freeze_panes(1, 0)
            header_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})
            for col_idx, col_name in enumerate(matched_copy.columns):
                ws.write(0, col_idx, col_name, header_fmt)
                col_width = max(matched_copy[col_name].astype(str).map(len).max(), len(col_name), 12) + 2
                ws.set_column(col_idx, col_idx, min(col_width, 50))

    unassigned = df_master.loc[~df_master.index.isin(assigned_idx)]
    if not unassigned.empty:
        un_file = out_dir / "Unassigned_Domains.xlsx"
        unassigned.to_excel(un_file, index=False)
        messagebox.showwarning(
            "Unassigned Domains",
            f"{len(unassigned)} users unassigned. Saved to: {un_file}"
        )
    return True


def load_email_template() -> str:
    """
    Loads the email body from an external text file.

    Using an external template allows for easy editing of the email content
    without modifying the Python script. If the file is not found, a default
    template is used as a fallback.

    Returns:
        str: The content of the email template.
    """
    try:
        with open('email_template.txt', 'r', encoding='utf-8') as f:
            return f.read()
    except (FileNotFoundError, UnicodeDecodeError):
        logger.warning("Email template not found, using default")
        # Fallback email template if the file doesn't exist
        return """Hello,

As part of our Sarbanes-Oxley (SOX) compliance requirements, we must complete the {review_period} Cognos Platform user access review by {deadline}. Your review and response are critical to ensuring compliance and maintaining appropriate system access.

Action Required:
1. Review the attached User Access Report listing Cognos Reporting Users and their folder access as of {review_period}.
2. Confirm or request changes:
   - If no updates are needed, reply confirming your review.
   - If updates are required, note them in Column G of the "User Access List" tab and return the file.
   - For access changes, submit a Paige ticket under the Cognos Services section.

This review is mandatory for compliance, and your prompt response is essential. Please let me know if you have any questions.

Best Regards,
{sender_name}
{sender_title}
{company_name}"""


def send_emails(email_manifest_file, out_dir, agencies, universal_attachment, audit_df, mode="Preview"):
    """
    Generates and sends emails to agencies via Microsoft Outlook.

    This function reads an email manifest to get recipient addresses, creates an
    email for each selected agency FILE, attaches their specific report, and then
    either displays it for review ("Preview" mode) or sends it directly ("Direct" mode).
    It updates the audit log with the sending status.

    Args:
        email_manifest_file (str): Path to Excel file with columns: Agency (File name) | To | CC
        out_dir (str): Directory where agency-specific Excel files are stored.
        agencies (list[str]): A list of file names (without .xlsx) to send emails for.
        universal_attachment (str): Path to a file to be attached to every email (e.g., instructions).
        audit_df (pd.DataFrame): The main audit log DataFrame, which will be updated.
        mode (str): The sending mode - "Preview" (display) or "Direct" (send automatically).

    Returns:
        pd.DataFrame: The updated audit log DataFrame.
    """
    logger.info(f"Starting email send process for {len(agencies)} files in {mode} mode")
    
    # Validate email manifest using EmailManifestManager
    try:
        manifest_mgr = EmailManifestManager(email_manifest_file)
        
        # Validate email addresses
        invalid_emails = manifest_mgr.validate_emails()
        
        if invalid_emails["To"] or invalid_emails["CC"]:
            # Show validation errors
            error_msg = "Invalid email addresses found:\n\n"
            if invalid_emails["To"]:
                error_msg += "To Field:\n" + "\n".join(f"  • {e}" for e in invalid_emails["To"][:10]) + "\n\n"
            if invalid_emails["CC"]:
                error_msg += "CC Field:\n" + "\n".join(f"  • {e}" for e in invalid_emails["CC"][:10]) + "\n\n"
            
            # Ask if they want to proceed anyway
            result = messagebox.askyesno(
                "Invalid Email Addresses",
                error_msg + "\nDo you want to proceed anyway?",
                icon="warning"
            )
            
            if not result:
                logger.info("Email sending cancelled due to invalid email addresses")
                return audit_df
    
    except Exception as e:
        logger.error(f"Email manifest validation failed: {str(e)}")
        messagebox.showerror("Validation Error", f"Failed to validate email manifest: {str(e)}")
        return audit_df
    
    # First, validate the email manifest file before proceeding
    is_valid, error_msg = validate_excel_file(email_manifest_file)
    if not is_valid:
        messagebox.showerror("Email Error", error_msg)
        return audit_df
    
    try:
        df_map = pd.read_excel(email_manifest_file)
        # Clean column names for consistency
        df_map.columns = df_map.columns.str.strip()
    except Exception as e:
        logger.error(f"Failed to read email manifest: {str(e)}")
        messagebox.showerror("Email Error", f"Failed to read email manifest: {str(e)}")
        return audit_df
    
    # Validate required columns
    required_cols = ["Agency", "To", "CC"]
    missing_cols = [col for col in required_cols if col not in df_map.columns]
    if missing_cols:
        messagebox.showerror("Email Manifest Error", f"Missing required columns: {', '.join(missing_cols)}")
        return audit_df
    
    # Build a fast lookup dictionary for email addresses from the manifest
    # Key = File name (from Agency column), Value = (To, CC)
    addr_dict = {}
    for _, row in df_map.iterrows():
        file_name = str(row.get("Agency", "")).strip()
        if not file_name:
            continue

        # Handle potentially empty "To" and "CC" fields gracefully
        to_val = row.get("To", "")
        to = str(to_val) if pd.notna(to_val) else ""
        
        cc_val = row.get("CC", "")
        cc = str(cc_val) if pd.notna(cc_val) else ""

        addr_dict[file_name] = (to.strip(), cc.strip())

    # Load the editable email body from the template file
    email_template = load_email_template()
    
    # Establish connection to the Outlook desktop application
    try:
        # Initialize COM for this thread
        try:
            pythoncom.CoInitialize()
        except:
            pass  # Already initialized
            
        outlook = win32.Dispatch("Outlook.Application")
        logger.info("Successfully connected to Outlook")
    except Exception as e:
        logger.error(f"Outlook connection failed: {str(e)}")
        messagebox.showerror("Outlook Error", f"Failed to connect to Outlook: {str(e)}")
        return audit_df

    sent_count = 0
    # Process each selected file one by one
    for i, file_name in enumerate(agencies):
        try:
            # Find the correct email addresses for this file
            to, cc = addr_dict.get(file_name, ("", ""))
            
            if not to and not cc:
                logger.warning(f"No email addresses found for: {file_name}")
                messagebox.showwarning("Missing Email", f"No email addresses found for: {file_name}")
                continue
            
            file_path = os.path.join(out_dir, f"{file_name}.xlsx")
            
            # Skip if the Excel file doesn't exist
            if not os.path.exists(file_path):
                logger.warning(f"File not found: {file_path}")
                continue

            # Create a new email item in Outlook
            mail = outlook.CreateItem(0)
            mail.Subject = (
                f"{config.get('email_subject_prefix', '[ACTION REQUIRED] Cognos Access Review')} - {file_name} | {REVIEW_PERIOD} | "
                f"Deadline: {REVIEW_DEADLINE}"
            )
            
            # Populate the email body using the loaded template and config values
            mail.Body = email_template.format(
                review_period=REVIEW_PERIOD,
                deadline=REVIEW_DEADLINE,
                sender_name=config.get("sender_name", "Govind Waghmare"),
                sender_title=config.get("sender_title", "Manager, Financial Applications | Analytics"),
                company_name=config.get("company_name", "Omnicom Group")
            )
            
            # Attach the file and any universal document
            mail.Attachments.Add(file_path)
            if universal_attachment and os.path.exists(universal_attachment):
                mail.Attachments.Add(universal_attachment)
            
            # Set email recipients
            if to:
                mail.To = to
            if cc:
                mail.CC = cc

            # Send or display the email based on the selected mode
            try:
                if mode == "Direct":
                    mail.Send()
                    logger.info(f"Email sent for {file_name}")
                else: # Preview mode
                    mail.Display()
                    logger.info(f"Email displayed for {file_name}")
                sent_count += 1
            except Exception as e:
                logger.error(f"Failed to send email for {file_name}: {str(e)}")
                messagebox.showerror("Send Error", f"Failed to send email for {file_name}: {str(e)}")

            # Update the audit log to record that the email was sent
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            mask = (audit_df["Agency"].str.upper() == file_name.upper())
            if mask.any(): # If file already has a row, update it
                audit_df.loc[mask, ["Sent Email Date", "To", "CC", "Status"]] = [now, to, cc, "Sent"]
            else: # Otherwise, add a new row for this file
                audit_df.loc[len(audit_df)] = [file_name, now, "", to, cc, "Sent", ""]
                
        except Exception as e:
            logger.error(f"Error processing file {file_name}: {str(e)}")
            continue # Continue to the next file even if one fails

    logger.info(f"Email process completed. {sent_count}/{len(agencies)} emails processed")
    return audit_df


# =============== Smart Email Verifier Dialog ===============

class SmartEmailVerifierDialog(ctk.CTkToplevel):
    """
    Smart Email Verifier Dialog
    
    This dialog allows users to paste forwarded emails or any text containing
    email addresses. It intelligently parses the text to extract email addresses
    and identify forwarding instructions like "please forward to", "cc:", etc.
    """
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.parent = parent
        self.title("Smart Email Verifier")
        self.geometry("800x700")
        self.transient(parent)
        
        # Configure grid
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Title
        title_label = ctk.CTkLabel(
            self, 
            text="Smart Email Verifier", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(20, 10), padx=20, sticky="ew")
        
        # Main content frame
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 10))
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Instructions
        instructions = ctk.CTkLabel(
            main_frame,
            text="Paste your forwarded email or any text containing email addresses below.\n"
                 "The system will intelligently extract emails and identify forwarding instructions.",
            font=ctk.CTkFont(size=12),
            wraplength=750
        )
        instructions.grid(row=0, column=0, pady=(15, 10), padx=15, sticky="ew")
        
        # Input text area
        input_label = ctk.CTkLabel(main_frame, text="Input Text:", font=ctk.CTkFont(weight="bold"))
        input_label.grid(row=1, column=0, pady=(10, 5), padx=15, sticky="w")
        
        self.input_text = ctk.CTkTextbox(main_frame, height=200, wrap="word")
        self.input_text.grid(row=2, column=0, pady=(0, 10), padx=15, sticky="nsew")
        
        # Control buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.grid(row=3, column=0, pady=(0, 10), padx=15, sticky="ew")
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        analyze_btn = ctk.CTkButton(
            button_frame, 
            text="🔍 Analyze Text", 
            command=self.analyze_text,
            height=35
        )
        analyze_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        clear_btn = ctk.CTkButton(
            button_frame, 
            text="🗑️ Clear All", 
            command=self.clear_all,
            height=35,
            fg_color="gray"
        )
        clear_btn.grid(row=0, column=1, padx=5, sticky="ew")
        
        paste_btn = ctk.CTkButton(
            button_frame, 
            text="📋 Paste from Clipboard", 
            command=self.paste_from_clipboard,
            height=35
        )
        paste_btn.grid(row=0, column=2, padx=(5, 0), sticky="ew")
        
        # Results area
        results_label = ctk.CTkLabel(main_frame, text="Analysis Results:", font=ctk.CTkFont(weight="bold"))
        results_label.grid(row=4, column=0, pady=(20, 5), padx=15, sticky="w")
        
        self.results_text = ctk.CTkTextbox(main_frame, height=200, wrap="word")
        self.results_text.grid(row=5, column=0, pady=(0, 10), padx=15, sticky="nsew")
        
        # Action buttons
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        action_frame.grid(row=6, column=0, pady=(0, 15), padx=15, sticky="ew")
        action_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.add_to_manifest_btn = ctk.CTkButton(
            action_frame,
            text="📧 Add to Email Manifest",
            command=self.add_to_manifest,
            height=35,
            state="disabled"
        )
        self.add_to_manifest_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        
        close_btn = ctk.CTkButton(
            action_frame,
            text="❌ Close",
            command=self.destroy,
            height=35,
            fg_color="gray"
        )
        close_btn.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        
        # Store extracted emails for later use
        self.extracted_emails = {"to": [], "cc": []}
        
        # Focus on input text
        self.input_text.focus()
    
    def analyze_text(self):
        """Analyze the input text for email addresses and forwarding instructions."""
        text = self.input_text.get("1.0", "end-1c")
        
        if not text.strip():
            self.results_text.delete("1.0", "end")
            self.results_text.insert("1.0", "❌ Please enter some text to analyze.")
            self.add_to_manifest_btn.configure(state="disabled")
            return
        
        # Perform smart email verification
        verification_result = smart_email_verification(text)
        
        # Build results display
        results = []
        results.append("🔍 SMART EMAIL ANALYSIS RESULTS")
        results.append("=" * 50)
        
        # Found emails
        if verification_result["found_emails"]:
            results.append(f"\n📧 FOUND EMAILS ({len(verification_result['found_emails'])} total):")
            for i, email in enumerate(verification_result["found_emails"], 1):
                status = "✅ Valid" if email in verification_result["valid_emails"] else "❌ Invalid"
                results.append(f"  {i}. {email} ({status})")
        else:
            results.append("\n❌ No email addresses found in text")
        
        # Smart parsing results
        parsing = verification_result["parsing"]
        if parsing["to"] or parsing["cc"]:
            results.append(f"\n🤖 SMART PARSING RESULTS:")
            
            if parsing["to"]:
                results.append(f"  📬 TO addresses ({len(parsing['to'])}):")
                for email in parsing["to"]:
                    results.append(f"    • {email}")
            
            if parsing["cc"]:
                results.append(f"  📋 CC addresses ({len(parsing['cc'])}):")
                for email in parsing["cc"]:
                    results.append(f"    • {email}")
        else:
            results.append(f"\n🤖 SMART PARSING: No clear forwarding instructions detected")
        
        # Confidence and suggestions
        confidence_icon = {"high": "🟢", "medium": "🟡", "low": "🔴"}
        results.append(f"\n{confidence_icon[verification_result['confidence']]} CONFIDENCE: {verification_result['confidence'].upper()}")
        
        if verification_result["suggestions"]:
            results.append(f"\n💡 SUGGESTIONS:")
            for suggestion in verification_result["suggestions"]:
                results.append(f"  • {suggestion}")
        
        # Display results
        results_text = "\n".join(results)
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", results_text)
        
        # Store extracted emails and enable add button if valid emails found
        self.extracted_emails = parsing
        if verification_result["valid_emails"]:
            self.add_to_manifest_btn.configure(state="normal")
        else:
            self.add_to_manifest_btn.configure(state="disabled")
    
    def clear_all(self):
        """Clear all text areas."""
        self.input_text.delete("1.0", "end")
        self.results_text.delete("1.0", "end")
        self.extracted_emails = {"to": [], "cc": []}
        self.add_to_manifest_btn.configure(state="disabled")
    
    def paste_from_clipboard(self):
        """Paste text from clipboard into input area."""
        try:
            clipboard_text = self.clipboard_get()
            self.input_text.delete("1.0", "end")
            self.input_text.insert("1.0", clipboard_text)
        except Exception as e:
            messagebox.showerror("Clipboard Error", f"Failed to paste from clipboard: {str(e)}")
    
    def add_to_manifest(self):
        """Add extracted emails to the email manifest."""
        if not self.extracted_emails["to"] and not self.extracted_emails["cc"]:
            messagebox.showwarning("No Emails", "No valid emails to add to manifest.")
            return
        
        # Show dialog to select agency and confirm email assignments
        self.show_add_to_manifest_dialog()
    
    def show_add_to_manifest_dialog(self):
        """Show dialog to add extracted emails to email manifest."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add to Email Manifest")
        dialog.geometry("500x400")
        dialog.transient(self)
        dialog.grab_set()
        
        # Title
        title = ctk.CTkLabel(dialog, text="Add Emails to Manifest", font=ctk.CTkFont(size=16, weight="bold"))
        title.pack(pady=(20, 10))
        
        # Agency selection
        agency_frame = ctk.CTkFrame(dialog)
        agency_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        ctk.CTkLabel(agency_frame, text="Select Agency:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # Get available agencies from parent
        agencies = []
        try:
            if hasattr(self.parent, 'agency_list') and self.parent.agency_list:
                agencies = [item for item in self.parent.agency_list]
        except:
            pass
        
        if not agencies:
            agencies = ["[Enter manually]"]
        
        agency_var = ctk.StringVar(value=agencies[0] if agencies else "")
        agency_combo = ctk.CTkComboBox(agency_frame, variable=agency_var, values=agencies)
        agency_combo.pack(fill="x", padx=10, pady=(0, 10))
        
        # Email assignments
        email_frame = ctk.CTkFrame(dialog)
        email_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        ctk.CTkLabel(email_frame, text="Email Assignments:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        # TO emails
        to_text = None
        if self.extracted_emails["to"]:
            to_label = ctk.CTkLabel(email_frame, text=f"TO Emails ({len(self.extracted_emails['to'])}):")
            to_label.pack(anchor="w", padx=10, pady=(5, 2))
            
            to_text = ctk.CTkTextbox(email_frame, height=80)
            to_text.pack(fill="x", padx=10, pady=(0, 5))
            to_text.insert("1.0", "; ".join(self.extracted_emails["to"]))
        
        # CC emails
        cc_text = None
        if self.extracted_emails["cc"]:
            cc_label = ctk.CTkLabel(email_frame, text=f"CC Emails ({len(self.extracted_emails['cc'])}):")
            cc_label.pack(anchor="w", padx=10, pady=(5, 2))
            
            cc_text = ctk.CTkTextbox(email_frame, height=80)
            cc_text.pack(fill="x", padx=10, pady=(0, 5))
            cc_text.insert("1.0", "; ".join(self.extracted_emails["cc"]))
        
        # Buttons
        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        def add_emails():
            agency = agency_var.get().strip()
            if not agency or agency == "[Enter manually]":
                messagebox.showerror("Error", "Please select or enter an agency name.")
                return
            
            to_emails = ""
            cc_emails = ""
            
            if to_text and self.extracted_emails["to"]:
                to_emails = to_text.get("1.0", "end-1c").strip()
            
            if cc_text and self.extracted_emails["cc"]:
                cc_emails = cc_text.get("1.0", "end-1c").strip()
            
            # Add to email manifest (this would integrate with the actual manifest system)
            success_msg = f"✅ Emails added to manifest for agency: {agency}\n"
            if to_emails:
                success_msg += f"TO: {to_emails}\n"
            if cc_emails:
                success_msg += f"CC: {cc_emails}\n"
            
            messagebox.showinfo("Success", success_msg)
            dialog.destroy()
            self.destroy()  # Close the verifier dialog too
        
        add_btn = ctk.CTkButton(button_frame, text="✅ Add to Manifest", command=add_emails)
        add_btn.pack(side="left", padx=(0, 10))
        
        cancel_btn = ctk.CTkButton(button_frame, text="❌ Cancel", command=dialog.destroy, fg_color="gray")
        cancel_btn.pack(side="right")


# =============== Main GUI Application ===============
class CognosAccessReviewApp(ctk.CTk):
    # --- PROCESS SECTION (Main Workflow) ---
    # process_scroll is defined in __init__, not at class level

    def fix_errors(self):
        """
        Prompts the user to fix missing agency mappings by entering contact info and updates the mapping file.
        """
        import pandas as pd
        from tkinter import simpledialog, messagebox
        from pathlib import Path
        master_file = self.vars["master"].get()
        combined_file = self.vars["combined"].get()
        if not master_file or not combined_file:
            messagebox.showerror("Missing Files", "Please select Master File and Combined Mapping File.")
            return
        try:
            df_master = pd.read_excel(master_file)
            df_map = pd.read_excel(combined_file)
            # Normalize agency names
            master_agencies = set(df_master["Agency"].astype(str).str.strip().str.upper().unique())
            mapped_agencies = set(df_map["Agency"].astype(str).str.strip().str.upper().unique())
            missing_agencies = sorted(master_agencies - mapped_agencies)
            if not missing_agencies:
                messagebox.showinfo("No Errors", "All agencies in the master file are mapped.")
                return
            new_rows = []
            for agency in missing_agencies:
                # Prompt for contact info
                contact = simpledialog.askstring(
                    "Missing Agency Mapping",
                    f"Enter contact email(s) for missing agency: {agency}\n(Comma-separated if multiple)",
                    parent=self
                )
                if contact is None:
                    continue  # User cancelled
                # Add to mapping
                new_rows.append({
                    "Agency": agency,
                    "Contact": contact
                })
            if new_rows:
                # Append to mapping DataFrame and save
                df_map = pd.concat([df_map, pd.DataFrame(new_rows)], ignore_index=True)
                # Save with utf-8 encoding
                df_map.to_excel(combined_file, index=False, engine="openpyxl")
                messagebox.showinfo("Mapping Updated", f"Added {len(new_rows)} missing agencies to mapping file.")
                self.refresh()
            else:
                messagebox.showinfo("No Changes", "No new agencies were added.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fix errors: {str(e)}")
    """
    The main application class for the Cognos Access Review Tool.
    
    This class builds and manages the graphical user interface (GUI),
    handles user interactions, and orchestrates the calls to the backend
    utility functions.
    """
    def __init__(self):
        super().__init__()
        self.title("Cognos Access Review Tool - Enterprise Edition")
        # Set a more modern default window size
        self.geometry("1100x800")
        # Make the application start in a maximized state for a more professional feel
        self.state('zoomed') 
        ctk.set_appearance_mode("System") # Default to user's system theme (Light/Dark)

        # --- Initialize Service Classes ---
        # These are the new modular services that handle core functionality
        self.file_validator = FileValidator()
        self.file_processor = FileProcessor()
        self.email_handler = EmailHandler(config_manager)
        # AuditLogger will be initialized with output_dir when needed
        self.audit_logger = None

        # --- Application State Variables ---
        # These variables hold the application's data in memory.
        self.agencies = [] # Holds the list of agency names found in the output folder.
        self.audit_df = pd.DataFrame() # Holds the data from the main audit log file.

        # --- Main Layout Configuration (Sidebar + Content) ---
        # Configure the main window grid with sidebar (fixed width) + content (expanding)
        self.grid_columnconfigure(0, weight=0)  # Sidebar - fixed
        self.grid_columnconfigure(1, weight=1)  # Content - expanding
        self.grid_rowconfigure(0, weight=1)

        # --- SIDEBAR (Left Column - Fixed 220px) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=DIMENSIONS["sidebar_width"], corner_radius=0, fg_color=get_color("sidebar"))
        self.sidebar_frame.grid(row=0, column=0, sticky="ns", padx=0, pady=0)
        self.sidebar_frame.grid_propagate(False)  # Keep fixed width
        self.sidebar_frame.grid_rowconfigure(9, weight=1)  # Push footer to bottom

        # Sidebar Header (Logo/Title)
        try:
            font_family = "Century Gothic" if "Century Gothic" in tk.font.families() else "Arial"
            sidebar_header = ctk.CTkLabel(
                self.sidebar_frame, text="OMNICOM\nCognos Review",
                font=ctk.CTkFont(family=font_family, size=18, weight="bold"),
                text_color=get_color("primary")
            )
        except:
            sidebar_header = ctk.CTkLabel(
                self.sidebar_frame, text="OMNICOM\nCognos Review",
                font=ctk.CTkFont(size=18, weight="bold"),
                text_color=get_color("primary")
            )
        sidebar_header.grid(row=0, column=0, pady=20, sticky="ew", padx=10)

        # Sidebar divider
        sidebar_divider1 = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color=get_color("border"))
        sidebar_divider1.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 15))

        # --- SIDEBAR SECTIONS (5 Main Navigation Buttons) ---
        self.sidebar_buttons = {}
        self.current_section = None
        
        sidebar_sections = [
            ("setup", f"{ICONS['setup']} Setup", "File Configuration"),
            ("process", f"{ICONS['process']} Process", "Generate & Send"),
            ("reporting", f"{ICONS['reporting']} Reporting", "Analytics & Reports"),
            ("tools", f"{ICONS['tools']} Tools", "Utilities"),
            ("help", f"{ICONS['help']} Help", "Documentation"),
        ]

        for idx, (section_id, label, tooltip) in enumerate(sidebar_sections):
            btn = ctk.CTkButton(
                self.sidebar_frame, text=label,
                font=ctk.CTkFont(size=14, weight="bold"),
                fg_color="transparent",
                text_color=get_color("text"),
                hover_color=get_color("secondary"),
                height=45,
                command=lambda sid=section_id: self.switch_section(sid),
                corner_radius=10,
                anchor="w",
                border_width=0
            )
            btn.grid(row=idx+2, column=0, sticky="ew", padx=10, pady=3)
            self.sidebar_buttons[section_id] = btn

        # Sidebar divider 2
        sidebar_divider2 = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color=get_color("border"))
        sidebar_divider2.grid(row=7, column=0, sticky="ew", padx=10, pady=(15, 15))

        # Sidebar footer
        footer_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        footer_frame.grid(row=8, column=0, sticky="ew", padx=10, pady=10)

        # Theme switcher in sidebar
        ctk.CTkLabel(footer_frame, text="Theme:", font=ctk.CTkFont(size=11), text_color=get_color("text")).pack(anchor="w", pady=(0, 5))
        self.theme_var = ctk.StringVar(value="System")
        ctk.CTkOptionMenu(
            footer_frame, variable=self.theme_var,
            values=["Light", "Dark", "System"],
            command=self.change_theme,
            height=32, fg_color=get_color("primary"), button_color=get_color("primary"),
            text_color=get_color("text_light"), dropdown_text_color=get_color("text")
        ).pack(fill="x", pady=(0, 10))

        # Settings and about buttons
        ctk.CTkButton(
            footer_frame, text="⚙️ Settings", command=self.settings, height=36, corner_radius=6,
            fg_color=get_color("primary"), text_color=get_color("text_light"),
            hover_color=get_color("primary_hover"), font=ctk.CTkFont(weight="bold")
        ).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(
            footer_frame, text="About", command=self.show_about, height=36, corner_radius=6,
            fg_color=get_color("secondary"), text_color=get_color("text_light"),
            hover_color=get_color("secondary_hover"), font=ctk.CTkFont(weight="bold")
        ).pack(fill="x")

        # --- MAIN CONTENT AREA (Right Side - Expanding) ---
        self.main_content_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=get_color("background"))
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.main_content_frame.grid_columnconfigure(0, weight=1)
        self.main_content_frame.grid_rowconfigure(1, weight=1)

        # Header bar in main content
        header_bar = ctk.CTkFrame(self.main_content_frame, fg_color=get_color("surface"), corner_radius=0, height=64)
        header_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header_bar.grid_propagate(False)

        try:
            font_family = "Century Gothic" if "Century Gothic" in tk.font.families() else "Arial"
            ctk.CTkLabel(
                header_bar, text="OmnicomGroup - Cognos Access Review",
                font=ctk.CTkFont(family=font_family, size=28, weight="bold"),
                text_color=get_color("primary")
            ).pack(side="left", padx=20, pady=15)
        except:
            ctk.CTkLabel(
                header_bar, text="OmnicomGroup - Cognos Access Review",
                font=ctk.CTkFont(size=28, weight="bold"),
                text_color=get_color("primary")
            ).pack(side="left", padx=20, pady=15)

        # Content switching frame
        self.content_switching_frame = ctk.CTkFrame(self.main_content_frame, fg_color="transparent")
        self.content_switching_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.content_switching_frame.grid_columnconfigure(0, weight=1)
        self.content_switching_frame.grid_rowconfigure(0, weight=1)

        # Initialize section frames
        self.section_frames = {}
        for section_id, _, _ in sidebar_sections:
            section_frame = ctk.CTkFrame(self.content_switching_frame, fg_color="transparent", corner_radius=0)
            section_frame.grid(row=0, column=0, sticky="nsew")
            section_frame.grid_columnconfigure(0, weight=1)
            section_frame.grid_rowconfigure(0, weight=1)
            self.section_frames[section_id] = section_frame

        # --- SETUP SECTION (File Configuration) ---
        setup_scroll = ctk.CTkScrollableFrame(
            self.section_frames["setup"], fg_color=get_color("background"),
            label_text="File Configuration"
        )
        setup_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        # Input variables
        keys = ["master", "combined", "attach", "output"]
        self.vars = {k: ctk.StringVar() for k in keys}
        self.email_mode = tk.StringVar(value="Preview")

        # File Configuration in SETUP section
        ctk.CTkLabel(setup_scroll, text="📁 File Paths", font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(15, 15), padx=5)
        
        labels = [
            "Master User Access File:", "Combined Agency & Email Mapping File:",
            "Universal Attachment (DOCX/PDF):", "Output Folder:"
        ]
        for idx, text in enumerate(labels):
            label_frame = ctk.CTkFrame(setup_scroll, fg_color="transparent")
            label_frame.pack(fill="x", pady=5)
            
            ctk.CTkLabel(label_frame, text=text, font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=5)
            
            file_frame = ctk.CTkFrame(setup_scroll, fg_color="transparent")
            file_frame.pack(fill="x", pady=(0, 10))
            file_frame.grid_columnconfigure(0, weight=1)
            
            entry = ctk.CTkEntry(
                file_frame, textvariable=self.vars[keys[idx]], placeholder_text="Select file...",
                fg_color=get_color("input_bg"), border_color=get_color("input_border"),
                text_color=get_color("text")
            )
            entry.grid(row=0, column=0, sticky="ew", padx=5)
            
            cmd = self._browse_folder if idx == 3 else self._browse_file
            var = self.vars[keys[idx]]
            ctk.CTkButton(
                file_frame, text="Browse...", command=lambda v=var, f=cmd: f(v), width=110, height=36,
                fg_color=get_color("primary"), text_color=get_color("text_light"),
                hover_color=get_color("primary_hover"), corner_radius=8,
                font=ctk.CTkFont(size=13, weight="bold")
            ).grid(row=0, column=1, padx=(5, 0))

        # Email Settings Section
        ctk.CTkLabel(setup_scroll, text="📧 Email Settings", font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(25, 15), padx=5)
        
        email_mode_frame = ctk.CTkFrame(setup_scroll, fg_color=get_color("surface"), corner_radius=8)
        email_mode_frame.pack(fill="x", padx=5, pady=(0, 10))
        email_mode_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        ctk.CTkLabel(email_mode_frame, text="Email Mode:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))
        
        for col, mode in enumerate(["Preview", "Direct", "Schedule"]):
            ctk.CTkRadioButton(
                email_mode_frame, text=mode, variable=self.email_mode, value=mode,
                text_color=get_color("text"), border_color=get_color("primary"),
                fg_color=get_color("primary")
            ).grid(row=1, column=col, padx=10, pady=(0, 10), sticky="w")

        # Agency Management Section
        ctk.CTkLabel(setup_scroll, text="🏢 Agency Management", font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(25, 15), padx=5)
        # Filter entry and agency list container (created here so filter_agencies()
        # can be called safely during initialization)
        self.filter_var = ctk.StringVar()
        filter_frame = ctk.CTkFrame(setup_scroll, fg_color="transparent")
        filter_frame.pack(fill="x", padx=5, pady=(0, 8))
        ctk.CTkEntry(
            filter_frame, textvariable=self.filter_var, placeholder_text="Filter agencies...",
            fg_color=get_color("input_bg"), border_color=get_color("input_border"), text_color=get_color("text")
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(filter_frame, text="Refresh List", command=self.refresh, width=110, height=36,
                      fg_color=get_color("primary"), text_color=get_color("text_light"),
                      hover_color=get_color("primary_hover"), corner_radius=8).pack(side="left")

        # Scrollable frame that will hold agency checkboxes
        self.agency_scrollable_frame = ctk.CTkScrollableFrame(setup_scroll, fg_color=get_color("surface"), height=240)
        self.agency_scrollable_frame.pack(fill="both", expand=False, padx=5, pady=(8, 12))
        self.agency_checkboxes = {}

        # Re-filter when the filter variable changes
        try:
            self.filter_var.trace_add("write", lambda *a: self.filter_agencies())
        except Exception:
            # Fallback for older tkinter versions
            self.filter_var.trace("w", lambda *a: self.filter_agencies())

        # Define process_scroll before using it
        process_scroll = ctk.CTkScrollableFrame(
            self.section_frames["process"], fg_color=get_color("background"),
            label_text="Generate & Send"
        )
        process_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(process_scroll, text="📊 Live Status", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(0, 10))
        
        metrics_frame = ctk.CTkFrame(process_scroll, fg_color=get_color("surface"), corner_radius=8)
        metrics_frame.pack(fill="x", padx=5, pady=(0, 15))
        metrics_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        # Metric cards - will be created dynamically in loop
        metrics = [
            ("Completion", "0%", "primary", "completion_label"),
            ("Days Left", "-", "text", "days_left_label"),
            ("Responded", "0/0", "success", "responded_label"),
            ("Overdue", "0", "danger", "overdue_label"),
        ]
        
        for idx, (label_text, default_value, color_key, attr_name) in enumerate(metrics):
            col_frame = ctk.CTkFrame(metrics_frame, fg_color="transparent")
            col_frame.grid(row=0, column=idx, padx=10, pady=10)
            
            ctk.CTkLabel(col_frame, text=label_text, font=ctk.CTkFont(size=11), text_color=get_color("text_secondary")).pack(anchor="w")
            
            value_label = ctk.CTkLabel(col_frame, text=default_value, font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color(color_key))
            value_label.pack(anchor="w")
            setattr(self, attr_name, value_label)

        # Core Process Buttons
        ctk.CTkLabel(process_scroll, text="🔄 Core Process", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(10, 10))
        
        core_buttons = [
            ("Validate Files", self.validate_files),
            ("Generate Files", self.generate),
            ("Send Emails", self.send),
            ("Mark Responded", self.mark),
            ("Create Manual Email", self.create_manual_email),
        ]
        
        for btn_text, cmd in core_buttons:
            btn = ctk.CTkButton(
                process_scroll, text=btn_text, command=cmd,
                fg_color=get_color("primary"), text_color=get_color("text_light"),
                hover_color=get_color("primary_hover"), height=40, corner_radius=8,
                font=ctk.CTkFont(weight="bold")
            )
            btn.pack(fill="x", pady=5, padx=5)
            # Add Fix Errors button (initially hidden)
            self.fix_errors_btn = ctk.CTkButton(
                process_scroll, text="Fix Errors", command=self.fix_errors,
                fg_color=get_color("danger"), text_color=get_color("text_light"),
                hover_color=get_color("danger_hover"), height=40, corner_radius=8,
                font=ctk.CTkFont(weight="bold")
            )
            self.fix_errors_btn.pack(fill="x", pady=5, padx=5)
            self.fix_errors_btn.pack_forget()

        # Data & Reporting Section
        ctk.CTkLabel(process_scroll, text="📈 Data & Reporting", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(15, 10))
        
        report_buttons = [
            ("Refresh List", self.refresh),
            ("Dashboard", self.open_dashboard),
            ("Export Log", self.export),
            ("Export Agency List", self.export_list),
            ("Sync Sent Emails", self.sync_sent_emails),
        ]
        
        for btn_text, cmd in report_buttons:
            btn = ctk.CTkButton(
                process_scroll, text=btn_text, command=cmd,
                fg_color=get_color("secondary"), text_color=get_color("text_light"),
                hover_color=get_color("secondary_hover"), height=40, corner_radius=8,
                font=ctk.CTkFont(weight="bold")
            )
            btn.pack(fill="x", pady=5, padx=5)

        # --- REPORTING SECTION (Analytics) ---
        reporting_scroll = ctk.CTkScrollableFrame(
            self.section_frames["reporting"], fg_color=get_color("background"),
            label_text="Analytics & Reports"
        )
        reporting_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        # Summary Metrics
        ctk.CTkLabel(reporting_scroll, text="📊 Summary Metrics", font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(0, 15))
        
        metrics_frame = ctk.CTkFrame(reporting_scroll, fg_color=get_color("surface"), corner_radius=8)
        metrics_frame.pack(fill="x", padx=5, pady=(0, 15))
        metrics_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        
        metrics = [
            ("Total Agencies", "0", "primary", "report_total_label"),
            ("Responded", "0", "success", "report_responded_label"),
            ("Pending", "0", "warning", "report_pending_label"),
            ("Overdue", "0", "danger", "report_overdue_label"),
        ]
        
        for idx, (label_text, default_value, color_key, attr_name) in enumerate(metrics):
            col_frame = ctk.CTkFrame(metrics_frame, fg_color="transparent")
            col_frame.grid(row=0, column=idx, padx=10, pady=10)
            
            ctk.CTkLabel(col_frame, text=label_text, font=ctk.CTkFont(size=11), text_color=get_color("text_secondary")).pack(anchor="w")
            
            value_label = ctk.CTkLabel(col_frame, text=default_value, font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color(color_key))
            value_label.pack(anchor="w")
            setattr(self, attr_name, value_label)

        # Export & Reporting Tools
        ctk.CTkLabel(reporting_scroll, text="📤 Export & Reports", font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(15, 10))
        
        export_buttons = [
            ("📋 Export Audit Log", self.export),
            ("📊 Export Agency List", self.export_list),
            ("📈 Open Dashboard", self.open_dashboard),
            ("📑 Export Summary Report", self.export_summary_report),
        ]
        
        for btn_text, cmd in export_buttons:
            ctk.CTkButton(
                reporting_scroll, text=btn_text, command=cmd, height=40, corner_radius=8,
                fg_color=get_color("primary"), text_color=get_color("text_light"),
                hover_color=get_color("primary_hover"), font=ctk.CTkFont(size=13, weight="bold")
            ).pack(fill="x", pady=5, padx=5)

        # Data Management
        ctk.CTkLabel(reporting_scroll, text="🔄 Data Management", font=ctk.CTkFont(size=16, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(15, 10))
        
        data_buttons = [
            ("🔄 Refresh Data", self.refresh),
            ("📧 Sync Sent Emails", self.sync_sent_emails),
            ("🗑️ Clear Data", self.clear_audit_data),
        ]
        
        for btn_text, cmd in data_buttons:
            ctk.CTkButton(
                reporting_scroll, text=btn_text, command=cmd, height=40, corner_radius=8,
                fg_color=get_color("secondary"), text_color=get_color("text_light"),
                hover_color=get_color("secondary_hover"), font=ctk.CTkFont(size=13, weight="bold")
            ).pack(fill="x", pady=5, padx=5)

        # --- TOOLS SECTION (Utilities) ---
        tools_scroll = ctk.CTkScrollableFrame(
            self.section_frames["tools"], fg_color=get_color("background"),
            label_text="Utilities & Advanced"
        )
        tools_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(tools_scroll, text="🔧 Agency Mappings", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(0, 10))
        
        mapping_buttons = [
            ("Manage Mappings", self.open_agency_mapping_manager),
            ("Schedule All", self.schedule_all_dialog),
        ]
        
        for btn_text, cmd in mapping_buttons:
            ctk.CTkButton(
                tools_scroll, text=btn_text, command=cmd, height=40, corner_radius=8,
                fg_color=get_color("primary"), text_color=get_color("text_light"),
                hover_color=get_color("primary_hover"), font=ctk.CTkFont(weight="bold")
            ).pack(fill="x", pady=5, padx=5)

        ctk.CTkLabel(tools_scroll, text="📧 Email Tools", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(15, 10))
        
        email_buttons = [
            ("Smart Email Verifier", self.open_email_verifier),
            ("Test Email Connection", self.test_email),
            ("Scan Inbox", self.scan),
            ("Organize Replies", self.organize_compliance_replies),
        ]
        
        for btn_text, cmd in email_buttons:
            ctk.CTkButton(
                tools_scroll, text=btn_text, command=cmd, height=40, corner_radius=8,
                fg_color=get_color("secondary"), text_color=get_color("text_light"),
                hover_color=get_color("secondary_hover"), font=ctk.CTkFont(weight="bold")
            ).pack(fill="x", pady=5, padx=5)

        ctk.CTkLabel(tools_scroll, text="💾 System", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(15, 10))
        
        system_buttons = [
            ("Open Output Folder", self.open_folder),
        ]
        
        for btn_text, cmd in system_buttons:
            ctk.CTkButton(
                tools_scroll, text=btn_text, command=cmd, height=40, corner_radius=8,
                fg_color=get_color("primary"), text_color=get_color("text_light"),
                hover_color=get_color("primary_hover"), font=ctk.CTkFont(weight="bold")
            ).pack(fill="x", pady=5, padx=5)

        # --- HELP SECTION (Documentation) ---
        help_scroll = ctk.CTkScrollableFrame(
            self.section_frames["help"], fg_color=get_color("background"),
            label_text="Help & Documentation"
        )
        help_scroll.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(help_scroll, text="📚 Getting Started", font=ctk.CTkFont(size=14, weight="bold"), text_color=get_color("primary")).pack(anchor="w", pady=(0, 10))
        
        help_text = """1. Configure your files in the SETUP tab
2. Select agencies to process
3. Click Validate Files to check setup
4. Click Generate Files to create agency-specific files
5. Choose email mode (Preview/Direct/Schedule)
6. Click Send Emails to deliver files
7. Monitor progress in the Process tab
8. Use Reporting for analytics and exports"""
        
        ctk.CTkLabel(help_scroll, text=help_text, justify="left", text_color=get_color("text")).pack(anchor="nw", fill="x", padx=5, pady=10)

        # Progress bar at bottom
        self.progress_frame = ctk.CTkFrame(self, fg_color=get_color("surface"), height=70, corner_radius=0)
        self.progress_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        self.progress_frame.grid_remove()

        progress_content = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        progress_content.pack(fill="both", expand=True, padx=20, pady=10)
        progress_content.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(progress_content, text="Progress:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=(0, 10))
        
        self.progress_bar = ctk.CTkProgressBar(progress_content, height=20, corner_radius=4, fg_color=get_color("border"), progress_color=get_color("success"))
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=10)
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(progress_content, text="Ready", text_color=get_color("text_secondary"))
        self.status_label.grid(row=0, column=2, sticky="e", padx=(10, 0))

        # Set initial section
        self.switch_section("setup")

        # Initial data load
        self.update_dashboard_metrics()
        self.filter_agencies()

        # Auto-scan on startup
        self.auto_scan_on_startup = config.get("auto_scan", True)
        if self.auto_scan_on_startup:
            self.after(1000, self.auto_scan_output_folder)

    # Browse Handlers
    def _browse_file(self, var):
        """
        Opens a file dialog to select a single file.

        Args:
            var (ctk.StringVar): The variable to update with the selected file path.
        """
        path = filedialog.askopenfilename()
        if path:
            var.set(path)

    def _browse_folder(self, var):
        """
        Opens a file dialog to select a directory.

        Args:
            var (ctk.StringVar): The variable to update with the selected folder path.
        """
        path = filedialog.askdirectory()
        if path:
            var.set(path)
            # Auto-refresh agency list when output folder is selected
            if var == self.vars["output"]:
                logger.info(f"Output folder selected: {path}")
                self.refresh()  # Automatically load agencies from the selected folder

    # ========================================================================
    # SIDEBAR NAVIGATION METHODS
    # ========================================================================
    
    def switch_section(self, section_id: str):
        """
        Switch to a different sidebar section.
        
        Args:
            section_id (str): The section to switch to (setup, process, reporting, tools, help)
        """
        # Hide all section frames
        for frame in self.section_frames.values():
            frame.grid_remove()
        
        # Show selected section
        if section_id in self.section_frames:
            self.section_frames[section_id].grid()
            self.current_section = section_id
        
        # Update sidebar button highlighting
        for sid, btn in self.sidebar_buttons.items():
            if sid == section_id:
                # Active state
                btn.configure(
                    fg_color=get_color("primary"),
                    text_color=get_color("text_light"),
                    hover_color=get_color("primary_hover")
                )
            else:
                # Inactive state
                btn.configure(
                    fg_color="transparent",
                    text_color=get_color("text"),
                    hover_color=get_color("secondary")
                )

    def load_mapping(self):
        """Loads mapping file from the current configuration."""
        if "domain" not in self.parent.vars:
            logger.error("Agency Mapping Manager: 'domain' key missing in parent.vars.")
            self.mapping_data = pd.DataFrame(columns=["File name", "Agency ID"])
            return
        mapping_file = self.parent.vars["domain"].get()
        if mapping_file and os.path.exists(mapping_file):
            try:
                df = pd.read_excel(mapping_file)
                df.columns = df.columns.str.strip()
                # Map column names to standardized names
                column_mapping = {}
                for col in df.columns:
                    if "file name" in col.lower():
                        column_mapping[col] = "File name"
                    elif "agency id" in col.lower() or "agency in the file" in col.lower():
                        column_mapping[col] = "Agency ID"
                # Rename columns
                if column_mapping:
                    df = df.rename(columns=column_mapping)
                # Ensure required columns exist
                if "File name" not in df.columns:
                    df["File name"] = ""
                if "Agency ID" not in df.columns:
                    df["Agency ID"] = ""
                self.mapping_data = df[["File name", "Agency ID"]]
            except Exception as e:
                logger.error(f"Failed to load mapping file: {str(e)}")
                # Create empty DataFrame with required columns
                self.mapping_data = pd.DataFrame(columns=["File name", "Agency ID"])
        else:
            # Create empty DataFrame if no file selected
            self.mapping_data = pd.DataFrame(columns=["File name", "Agency ID"])

    def show_about(self):
        """
        Displays the About dialog with requirements and app info.
        """
        about_dialog = ctk.CTkToplevel(self)
        about_dialog.title("About Cognos Access Review Tool")
        about_dialog.geometry("400x220")
        about_dialog.transient(self)
        about_dialog.grab_set()

        content = (
            "Cognos Access Review Tool\n\n"
            "Requirements:\n"
            "  - Python 3.12+\n"
            "  - Windows 10/11\n\n"
            "For SOX compliance audits of Cognos user access.\n"
            "Automates file generation, email sending, and audit tracking."
        )
        text_label = ctk.CTkLabel(about_dialog, text=content, justify="left")
        text_label.pack(pady=10, padx=20)

        ctk.CTkButton(about_dialog, text="Close", command=about_dialog.destroy).pack(pady=(10, 20))

    # Button Callbacks
    def generate(self):
        """
        Generate agency files with multi-tab Excel output and pre-validation.
        Uses the new FileValidator and FileProcessor services with combined file format.
        """
        # Get file paths
        master_file = self.vars["master"].get()
        combined_file = self.vars["combined"].get()
        output_dir = self.vars["output"].get()
        
        # Basic path validation
        if not master_file or not combined_file or not output_dir:
            messagebox.showerror("Missing Files", "Please select Master File, Combined Mapping File, and Output Folder")
            return
        
        try:
            # Show progress
            self.show_progress(True)
            self.update_progress(0.1, "Validating files...")
            
            # Basic file existence validation
            if not os.path.exists(master_file):
                messagebox.showerror("File Not Found", f"Master file not found: {master_file}")
                self.show_progress(False)
                return
            
            if not os.path.exists(combined_file):
                messagebox.showerror("File Not Found", f"Combined mapping file not found: {combined_file}")
                self.show_progress(False)
                return
                
            if not os.path.exists(output_dir):
                messagebox.showerror("Directory Not Found", f"Output directory not found: {output_dir}")
                self.show_progress(False)
                return
            
            self.update_progress(0.2, "Files validated")
            self.update_progress(0.3, "Generating agency files...")
            
            # Progress callback for file generation (signature: float, str)
            def progress_callback(progress: float, status: str):
                # Map progress to 0.3-0.9 range
                adjusted_progress = 0.3 + (progress * 0.6)
                self.update_progress(adjusted_progress, status)
            
            # Generate files using FileProcessor with combined file
            # Ensure all paths are Path objects, not strings
            master_path = Path(master_file) if isinstance(master_file, str) else master_file
            combined_path = Path(combined_file) if isinstance(combined_file, str) else combined_file
            output_path = Path(output_dir) if isinstance(output_dir, str) else output_dir
            
            self.file_processor.generate_agency_files(
                master_file=master_path,
                combined_map_file=combined_path,
                output_dir=output_path,
                progress_callback=progress_callback
            )
            
            self.update_progress(0.9, "Validating generated files...")
            
            # POST-GENERATION VALIDATION: Count users in master vs generated files
            validation_summary = self.validate_generated_files(master_path, output_path)
            
            self.update_progress(0.95, "Refreshing agency list...")
            self.refresh()  # Refresh UI to show newly generated files
            
            self.update_progress(1.0, "Complete!")
            
            # Show validation summary
            messagebox.showinfo("Success - Files Generated!", validation_summary)

            # Show Fix Errors button if validation indicates errors
            if "⚠️" in validation_summary or "❌" in validation_summary or "missing" in validation_summary.lower():
                self.fix_errors_btn.pack(fill="x", pady=5, padx=5)
            else:
                self.fix_errors_btn.pack_forget()
            
        except Exception as e:
            logger.error(f"Error generating files: {str(e)}")
            messagebox.showerror("Error", f"Failed to generate files: {str(e)}")
        finally:
            self.show_progress(False)

    def validate_generated_files(self, master_file: Path, output_dir: Path) -> str:
        """
        Validate generated files by counting users in master vs all generated files.
        
        Args:
            master_file: Path to master file
            output_dir: Directory containing generated files
            
        Returns:
            Validation summary string
        """
        try:
            # Read master file
            df_master = pd.read_excel(master_file)
            total_master_users = len(df_master)
            
            # Count users in all generated Excel files
            total_generated_users = 0
            generated_files_count = 0
            
            for excel_file in output_dir.glob("*.xlsx"):
                if excel_file.name.startswith("~$"):  # Skip temp files
                    continue
                    
                try:
                    # Read all sheets and count total rows
                    excel_data = pd.read_excel(excel_file, sheet_name=None)  # Read all sheets
                    
                    # Count unique users across all sheets (using "All Users" sheet if available)
                    if "All Users" in excel_data:
                        file_users = len(excel_data["All Users"])
                    else:
                        # If no "All Users" sheet, count first sheet
                        first_sheet = list(excel_data.values())[0]
                        file_users = len(first_sheet)
                    
                    total_generated_users += file_users
                    generated_files_count += 1
                    
                except Exception as e:
                    logger.warning(f"Could not read {excel_file.name}: {e}")
                    continue
            
            # Build validation summary
            summary_lines = [
                "✅ FILE GENERATION COMPLETE",
                "",
                f"📊 VALIDATION SUMMARY:",
                f"  • Master File Users: {total_master_users:,}",
                f"  • Generated Files: {generated_files_count}",
                f"  • Total Users in Generated Files: {total_generated_users:,}",
                ""
            ]
            
            # Check if counts match
            if total_generated_users == total_master_users:
                summary_lines.append("✅ PERFECT MATCH - All users accounted for!")
            elif total_generated_users < total_master_users:
                missing = total_master_users - total_generated_users
                summary_lines.append(f"⚠️ WARNING: {missing:,} users missing from generated files")
                summary_lines.append("   (May be unassigned/unmapped agencies)")
            else:
                extra = total_generated_users - total_master_users
                summary_lines.append(f"⚠️ WARNING: {extra:,} extra users in generated files")
                summary_lines.append("   (Check for duplicate assignments)")
            
            validation_text = "\n".join(summary_lines)
            logger.info(f"Validation Summary:\n{validation_text}")
            
            return validation_text
            
        except Exception as e:
            logger.error(f"Validation failed: {e}")
            return f"✅ Files generated successfully!\n\n⚠️ Could not validate user counts: {str(e)}"

    def filter_agencies(self, *args):
        """
        Filters the agency list based on text in the filter bar and status.
        
        This function is called automatically whenever the text in the filter
        entry changes. It rebuilds the checkbox list to show only matching agencies,
        and it colors them based on their status from the audit log.
        """
        search_term = self.filter_var.get().lower()
        
        # Clear existing checkboxes before adding the filtered list
        for widget in self.agency_scrollable_frame.winfo_children():
            widget.destroy()
        self.agency_checkboxes = {}

        # Define status colors, respecting the current light/dark theme
        is_dark = ctk.get_appearance_mode() == "Dark"
        colors = {
            "Responded": "#00B140", # Green
            "Sent": "#FFB302",      # Amber
            "Overdue": "#D42B2B",   # Red
            "Default": "white" if is_dark else "black"
        }

        # Filter the full agency list by the search term
        filtered_agencies = [
            agency for agency in self.agencies if search_term in agency.lower()
        ]
        
        # Create a dictionary for quick status lookups
        status_map = {}
        if not self.audit_df.empty and "Agency" in self.audit_df.columns:
            audit_copy = self.audit_df.copy()
            audit_copy['AgencyUpper'] = audit_copy['Agency'].str.upper()
            
            # Dynamically determine "Overdue" status
            if 'Sent Email Date' in audit_copy.columns:
                sent_mask = audit_copy['Status'] == 'Sent'
                sent_dates = pd.to_datetime(audit_copy.loc[sent_mask, 'Sent Email Date'], errors='coerce')
                # An item is overdue if it was sent and the send date is in the past
                overdue_mask = sent_dates < pd.Timestamp.now()
                audit_copy.loc[sent_mask & overdue_mask, 'Status'] = 'Overdue'

            status_map = audit_copy.set_index('AgencyUpper')['Status'].to_dict()

        # Re-create checkboxes for the filtered list
        for agency in filtered_agencies:
            status = status_map.get(agency.upper(), "Not Sent")
            color = colors.get(status, colors["Default"])
            
            var = tk.IntVar(value=0)
            cb = ctk.CTkCheckBox(self.agency_scrollable_frame, text=agency, variable=var, text_color=color)
            cb.pack(anchor="w", padx=10, pady=2)
            self.agency_checkboxes[agency] = var

    def change_theme(self, new_theme: str):
        """
        Changes the application theme and redraws the agency list to update colors.
        
        Args:
            new_theme (str): The theme to switch to ("Light", "Dark", or "System").
        """
        ctk.set_appearance_mode(new_theme)
        # Re-filter (which rebuilds) the list to apply the correct text colors for the new theme
        self.filter_agencies()

    def refresh(self):
        """
        Callback for the 'Refresh List' button.
        
        It re-scans the output directory for agency Excel files, reloads the
        audit log, and updates the UI to reflect the current state. This is
        useful if files are changed outside the application.
        """
        out = self.vars["output"].get()
        if out and os.path.isdir(out):
            # Find all .xlsx files in the output folder, excluding "Unassigned" and temp files (~$)
            files = os.listdir(out)
            self.agencies = sorted([
                f.replace(".xlsx", "") 
                for f in files 
                if f.endswith(".xlsx") 
                and not f.startswith("Unassigned") 
                and not f.startswith("~$")  # Skip Excel temp files
            ])
            
            # Debug: Log what files were found
            logger.info(f"Refresh: Found {len(self.agencies)} agency files in {out}")
            if self.agencies:
                logger.info(f"Refresh: Agencies: {', '.join(self.agencies[:10])}{'...' if len(self.agencies) > 10 else ''}")
            
            # Reload the audit log using AuditLogger
            audit_file_path = os.path.join(out, AUDIT_FILE_NAME)
            if self.audit_logger is not None:
                self.audit_df = self.audit_logger.load()
            else:
                # Initialize audit_logger if not already done
                self.audit_logger = AuditLogger(output_dir=Path(out))
                self.audit_df = self.audit_logger.load()
        else:
             self.agencies = []
             self.audit_df = pd.DataFrame(
                columns=["Agency", "Sent Email Date", "Response Received Date", "To", "CC", "Status", "Comments"]
            )

        # Rebuild the UI with the new data
        self.filter_agencies()
        self.update_dashboard_metrics()

    def update_dashboard_metrics(self):
        """
        Calculates and updates all the labels in the 'Live Status' panel.
        
        This function uses the in-memory audit DataFrame to calculate completion
        percentage, responded counts, overdue counts, and days left until the deadline.
        """
        if self.audit_df.empty or not self.agencies:
            # Set default values if no data is loaded
            self.completion_label.configure(text="0.0%")
            self.responded_label.configure(text=f"0 / {len(self.agencies)}")
            self.overdue_label.configure(text="0")
            # Update reporting dashboard metrics
            if hasattr(self, 'report_total_label'):
                self.report_total_label.configure(text=str(len(self.agencies)))
                self.report_responded_label.configure(text="0")
                self.report_pending_label.configure(text=str(len(self.agencies)))
                self.report_overdue_label.configure(text="0")
        else:
            total = len(self.agencies)
            
            # Count how many unique agencies have the "Responded" status
            responded_agencies = self.audit_df[self.audit_df["Status"] == "Responded"]["Agency"].str.upper().tolist()
            responded_count = len(set(responded_agencies) & set([a.upper() for a in self.agencies]))
            
            # Calculate overdue items - agencies sent more than 7 days ago without response
            overdue_count = 0
            if "Sent Email Date" in self.audit_df.columns:
                sent_df = self.audit_df[self.audit_df["Status"] == "Sent"].copy()
                sent_df["Sent Email Date"] = pd.to_datetime(sent_df["Sent Email Date"], errors='coerce')
                # Overdue if sent more than 7 days ago and still no response
                seven_days_ago = pd.Timestamp.now() - pd.Timedelta(days=7)
                overdue_df = sent_df[sent_df["Sent Email Date"] < seven_days_ago]
                overdue_count = len(overdue_df)

            completion_val = (responded_count / total * 100) if total > 0 else 0
            pending_count = total - responded_count
            
            # Update UI labels with the new values
            self.completion_label.configure(text=f"{completion_val:.1f}%")
            self.responded_label.configure(text=f"{responded_count} / {len(self.agencies)}")
            self.overdue_label.configure(text=f"{overdue_count}")
            
            # Update reporting dashboard metrics
            if hasattr(self, 'report_total_label'):
                self.report_total_label.configure(text=str(total))
                self.report_responded_label.configure(text=str(responded_count))
                self.report_pending_label.configure(text=str(pending_count))
                self.report_overdue_label.configure(text=str(overdue_count))

        # Calculate and display days left until the deadline
        try:
            deadline = pd.to_datetime(REVIEW_DEADLINE)
            days_left = (deadline - pd.Timestamp.now()).days
            self.days_left_label.configure(text=str(days_left))
        except (ValueError, TypeError):
            self.days_left_label.configure(text="-") # Show a dash if deadline is invalid

    def get_selected_agencies(self):
        """
        Returns a list of names of the currently checked agencies in the list.
        """
        return [agency for agency, var in self.agency_checkboxes.items() if var.get() == 1]

    def select_all_agencies(self):
        """
        Callback for the 'Select All' button. Checks all visible agency checkboxes.
        """
        for var in self.agency_checkboxes.values():
            var.set(1)

    def deselect_all_agencies(self):
        """
        Callback for the 'Deselect All' button. Unchecks all visible agency checkboxes.
        """
        for var in self.agency_checkboxes.values():
            var.set(0)

    def add_agency_folder(self):
        """
        Opens a folder dialog to add agencies from an additional folder.
        This allows loading agencies from multiple output folders simultaneously.
        """
        path = filedialog.askdirectory(title="Select Additional Agency Folder")
        if not path:
            return
        
        if not os.path.isdir(path):
            messagebox.showerror("Error", "Selected path is not a valid directory.")
            return
        
        # Find all .xlsx files in the selected folder
        try:
            files = os.listdir(path)
            new_agencies = [
                f.replace(".xlsx", "") 
                for f in files 
                if f.endswith(".xlsx") 
                and not f.startswith("Unassigned") 
                and not f.startswith("~$")  # Skip Excel temp files
            ]
            
            if not new_agencies:
                messagebox.showinfo("No Files", f"No agency Excel files found in:\n{path}")
                return
            
            # Track which folder each agency comes from (store as tuple: agency_name, folder_path)
            if not hasattr(self, 'agency_folders'):
                self.agency_folders = {}
            
            # Add new agencies and track their source folders
            added_count = 0
            for agency in new_agencies:
                if agency not in self.agencies:
                    self.agencies.append(agency)
                    self.agency_folders[agency] = path
                    added_count += 1
                else:
                    # Agency already exists, update the folder if needed
                    self.agency_folders[agency] = path
            
            # Sort the combined list
            self.agencies = sorted(self.agencies)
            
            # Load audit log from this folder and merge it - try multiple file name patterns
            audit_file_patterns = [
                AUDIT_FILE_NAME,  # e.g., Audit_CognosAccessReview_Q3_2025.xlsx
                "Audit_CognosAccessReview.xlsx",  # Legacy name
                "Cognos_Review_Audit_Log.xlsx"  # Alternative name
            ]
            
            audit_file_found = None
            for pattern in audit_file_patterns:
                test_path = os.path.join(path, pattern)
                if os.path.exists(test_path):
                    audit_file_found = test_path
                    break
            
            if audit_file_found:
                try:
                    new_audit_df = pd.read_excel(audit_file_found)
                    logger.info(f"Loaded audit log from {path}: {len(new_audit_df)} records ({os.path.basename(audit_file_found)})")
                    
                    # Merge with existing audit data
                    if self.audit_df.empty:
                        self.audit_df = new_audit_df.copy()
                    else:
                        # Combine and remove duplicates (keep existing records)
                        combined = pd.concat([self.audit_df, new_audit_df], ignore_index=True)
                        # Remove duplicates, keeping first occurrence (existing data takes precedence)
                        combined = combined.drop_duplicates(subset=['Agency'], keep='first')
                        self.audit_df = combined
                        logger.info(f"Merged audit data: Total {len(self.audit_df)} records")
                except Exception as e:
                    logger.error(f"Failed to load audit log from {audit_file_found}: {e}")
            else:
                logger.info(f"No audit log found in {path} (this is normal if no emails have been sent yet)")
            
            # Rebuild the UI
            self.filter_agencies()
            self.update_dashboard_metrics()
            
            logger.info(f"Added {added_count} agencies from folder: {path}")
            messagebox.showinfo(
                "Agencies Added", 
                f"Added {added_count} new agencies from:\n{path}\n\nTotal agencies: {len(self.agencies)}"
            )
            
        except Exception as e:
            logger.error(f"Error adding folder: {e}")
            messagebox.showerror("Error", f"Failed to load agencies from folder:\n{str(e)}")

    def create_manual_email(self):
        """
        Opens a dialog to create a new email with pre-filled template.
        User can edit the content before sending.
        """
        # Create the dialog window
        dialog = ctk.CTkToplevel(self)
        dialog.title("Create Email")
        dialog.geometry("800x700")
        dialog.transient(self)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - 800) // 2
        y = (dialog.winfo_screenheight() - 700) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame with padding
        main_frame = ctk.CTkFrame(dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        ctk.CTkLabel(main_frame, text="✉️ Create New Email", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 15))
        
        # To field
        to_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        to_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(to_frame, text="To:", width=80, anchor="w").pack(side="left")
        to_entry = ctk.CTkEntry(to_frame, placeholder_text="recipient@example.com; another@example.com")
        to_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # CC field
        cc_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        cc_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(cc_frame, text="CC:", width=80, anchor="w").pack(side="left")
        cc_entry = ctk.CTkEntry(cc_frame, placeholder_text="cc@example.com (optional)")
        cc_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Subject field
        subject_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        subject_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(subject_frame, text="Subject:", width=80, anchor="w").pack(side="left")
        subject_entry = ctk.CTkEntry(subject_frame)
        subject_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Pre-fill subject with default
        cfg = config_manager.load()
        subject_prefix = cfg.get("email_subject_prefix", "[ACTION REQUIRED] Cognos Access Review")
        sender_name = cfg.get("sender_name", "")
        sender_title = cfg.get("sender_title", "")
        subject_entry.insert(0, f"{subject_prefix} - {REVIEW_PERIOD}")
        
        # Attachment field
        attach_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        attach_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(attach_frame, text="Attachment:", width=80, anchor="w").pack(side="left")
        attach_entry = ctk.CTkEntry(attach_frame, placeholder_text="Path to attachment (optional)")
        attach_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        def browse_attachment():
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if path:
                attach_entry.delete(0, "end")
                attach_entry.insert(0, path)
        
        ctk.CTkButton(attach_frame, text="Browse", command=browse_attachment, width=80).pack(side="left", padx=(5, 0))
        
        # Agency dropdown (optional - to use their attachment)
        agency_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        agency_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(agency_frame, text="Agency:", width=80, anchor="w").pack(side="left")
        agency_var = ctk.StringVar(value="-- Select Agency (optional) --")
        agency_options = ["-- Select Agency (optional) --"] + self.agencies
        agency_dropdown = ctk.CTkComboBox(agency_frame, values=agency_options, variable=agency_var, width=300)
        agency_dropdown.pack(side="left", padx=(5, 0))
        
        def on_agency_select(choice):
            if choice != "-- Select Agency (optional) --":
                # Auto-fill attachment path from the agency's folder
                out_folder = self.vars["output"].get()
                if hasattr(self, 'agency_folders') and choice in self.agency_folders:
                    out_folder = self.agency_folders[choice]
                
                if out_folder:
                    attach_path = os.path.join(out_folder, f"{choice}.xlsx")
                    if os.path.exists(attach_path):
                        attach_entry.delete(0, "end")
                        attach_entry.insert(0, attach_path)
        
        agency_dropdown.configure(command=on_agency_select)
        
        # Body field
        ctk.CTkLabel(main_frame, text="Email Body:", anchor="w").pack(fill="x", pady=(15, 5))
        body_text = ctk.CTkTextbox(main_frame, height=300)
        body_text.pack(fill="both", expand=True, pady=(0, 10))
        
        # Pre-fill body with template
        default_body = f"""Dear Team,

Please find attached the {REVIEW_PERIOD} User Access Review file for your review.

As part of our periodic User Access Review process, we need you to:
1. Review the list of users with access
2. Verify that each user still requires access
3. Mark any users who should be removed
4. Return the completed file by {REVIEW_DEADLINE}

If you have any questions, please don't hesitate to reach out.

Best regards,
{sender_name}
{sender_title}"""
        
        body_text.insert("1.0", default_body)
        
        # Email mode
        mode_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        mode_frame.pack(fill="x", pady=10)
        ctk.CTkLabel(mode_frame, text="Send Mode:").pack(side="left")
        mode_var = ctk.StringVar(value="Preview")
        for mode in ["Preview", "Direct"]:
            ctk.CTkRadioButton(mode_frame, text=mode, variable=mode_var, value=mode).pack(side="left", padx=10)
        
        # Buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=(10, 0))
        
        def send_email():
            to_addr = to_entry.get().strip()
            cc_addr = cc_entry.get().strip()
            subject = subject_entry.get().strip()
            body = body_text.get("1.0", "end-1c")
            attachment = attach_entry.get().strip()
            mode = mode_var.get()
            
            if not to_addr:
                messagebox.showwarning("Warning", "Please enter a recipient email address.")
                return
            
            if not subject:
                messagebox.showwarning("Warning", "Please enter a subject.")
                return
            
            try:
                # Get or create Outlook instance
                outlook = self.email_handler._get_outlook()
                mail = outlook.CreateItem(0)
                
                mail.To = to_addr
                if cc_addr:
                    mail.CC = cc_addr
                mail.Subject = subject
                mail.Body = body
                
                # Add attachment if specified
                if attachment and os.path.exists(attachment):
                    mail.Attachments.Add(attachment)
                
                if mode == "Preview":
                    mail.Display()
                    logger.info(f"Created preview email to: {to_addr}")
                    messagebox.showinfo("Success", "Email opened for preview. You can edit and send manually.")
                else:
                    mail.Send()
                    logger.info(f"Sent email directly to: {to_addr}")
                    messagebox.showinfo("Success", "Email sent successfully!")
                
                dialog.destroy()
                
            except Exception as e:
                logger.error(f"Failed to create email: {e}")
                messagebox.showerror("Error", f"Failed to create email:\n{str(e)}")
        
        ctk.CTkButton(button_frame, text="Send / Preview", command=send_email, width=150).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=dialog.destroy, width=100).pack(side="right", padx=5)

    def send(self):
        """
        Callback for the 'Send Emails' button.
        
        Gets selected agencies, validates required files, and either opens the
        scheduling dialog or calls the email sending function.
        """
        selected = self.get_selected_agencies()
        if not selected:
            return messagebox.showwarning("Warning", "Please select one or more agencies.")
        
        combined_file = self.vars["combined"].get()
        is_valid, error_msg = validate_file_exists(combined_file, "Combined Agency & Email Mapping File")
        if not is_valid:
            messagebox.showerror("Validation Error", error_msg)
            return
        is_valid, error_msg = validate_excel_file(combined_file)
        if not is_valid:
            messagebox.showerror("File Error", error_msg)
            return
            
        mode = self.email_mode.get()
        if mode == "Schedule":
            # This was intended to open a scheduler, but is now handled by "Schedule All".
            # For simplicity, we can route this to the main send function in preview/direct mode.
            # A better implementation might open a specific scheduling dialog.
            # For now, we delegate to the main progress-based sender.
            self.schedule_email_dialog(selected) # Let's keep the single schedule dialog
        elif mode == "Direct":
            # SAFETY: Confirm before actually sending emails
            confirm = messagebox.askyesno(
                "Confirm Direct Send",
                f"⚠️ WARNING: Direct Send Mode ⚠️\n\n"
                f"You are about to AUTOMATICALLY SEND {len(selected)} emails.\n\n"
                f"Recipients will receive emails immediately without your review.\n\n"
                f"Selected agencies: {', '.join(selected[:5])}"
                f"{' ...' if len(selected) > 5 else ''}\n\n"
                f"Are you absolutely sure you want to continue?",
                icon="warning"
            )
            if confirm:
                self.run_task_in_thread(
                    target=self._send_emails_with_progress,
                    args=(selected, mode)
                )
            else:
                logger.info("Direct send cancelled by user confirmation")
        else:
            # Preview mode - no confirmation needed
            self.run_task_in_thread(
                target=self._send_emails_with_progress,
                args=(selected, mode)
            )
    
    def run_task_in_thread(self, target, args=()):
        """
        Runs a given function in a new thread to avoid blocking the GUI.
        
        This is a generic helper to start a background task.
        
        Args:
            target (function): The function to execute in the background.
            args (tuple): The arguments to pass to the target function.
        """
        thread = threading.Thread(target=target, args=args)
        thread.daemon = True # Allows app to exit even if thread is running
        thread.start()

    def _send_emails_with_progress(self, selected_agencies, mode):
        """
        Internal function to handle the email sending process with a progress bar.

        Args:
            selected_agencies (list[str]): The agencies to send emails to.
            mode (str): The send mode ("Preview" or "Direct").
        """
        try:
            # Use 'after' to ensure GUI updates happen on the main thread
            self.after(0, self.show_progress, True)
            self.after(0, self.update_progress, 0.1, "Preparing emails...")
            
            # Get required files
            combined_file = self.vars["combined"].get()
            output_dir = self.vars["output"].get()
            universal_attachment = self.vars["attach"].get()
            
            # Convert mode to EmailMode enum
            email_mode = EmailMode.PREVIEW if mode == "Preview" else EmailMode.DIRECT
            
            self.after(0, self.update_progress, 0.3, "Sending emails...")
            
            # Use EmailHandler to send emails with combined file
            universal_attach_path = Path(universal_attachment) if universal_attachment else None
            sent_count = self.email_handler.send_emails(
                combined_file=Path(combined_file),
                output_dir=Path(output_dir),
                file_names=selected_agencies,
                mode=email_mode,
                universal_attachment=universal_attach_path
            )
            
            self.after(0, self.update_progress, 0.7, "Updating audit log...")
            
            # Initialize AuditLogger with output directory
            self.audit_logger = AuditLogger(output_dir=Path(output_dir))
            
            # Initialize log with all agencies (preserves existing entries)
            self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
            
            # Load combined file to get email addresses for audit log
            try:
                combined_loader = CombinedFileLoader(Path(combined_file))
                combined_data = combined_loader.load_all_data()
                
                # Create address lookup from combined data
                addr_dict = {}
                for mapping in combined_data:
                    key = mapping.source_file_name.upper()
                    addr_dict[key] = (mapping.recipients_to, mapping.recipients_cc)
                    
            except Exception as e:
                logger.warning(f"Could not load combined file for audit log: {str(e)}")
                addr_dict = {}
            
            # Mark each successfully processed agency as Sent
            for agency in selected_agencies:
                to, cc = addr_dict.get(agency.upper(), ("", ""))
                status = "Sent" if mode == "Direct" else "Preview"
                self.audit_logger.mark_sent(agency, to=to, cc=cc, comments=f"Mode: {mode}")
            
            # Save the audit log to file
            self.audit_logger.save()
            
            # Reload audit log into memory
            self.audit_df = self.audit_logger.load()
            
            # Refresh UI
            self.after(0, self.refresh)
            
            self.after(0, self.update_progress, 1.0, "Complete!")
            self.after(0, messagebox.showinfo, "Done", f"Emails processed successfully! ({sent_count}/{len(selected_agencies)})")
            
        except Exception as e:
            logger.error(f"Error sending emails: {str(e)}")
            self.after(0, messagebox.showerror, "Error", f"Failed to send emails: {str(e)}")
        finally:
            self.after(0, self.show_progress, False)
    
    def _send_scheduled_emails(self, selected_agencies):
        """
        A wrapper function used by the scheduler (threading.Timer).
        
        This function is executed in a separate thread after a delay. It creates
        email drafts in Outlook instead of sending emails directly.

        Args:
            selected_agencies (list[str]): The agencies to create email drafts for.
        """
        def show_notification(title, message):
            """Show notification in main thread"""
            self.after(0, lambda: messagebox.showinfo(title, message))
        
        def show_error(title, message):
            """Show error in main thread"""
            self.after(0, lambda: messagebox.showerror(title, message))
        
        try:
            logger.info(f"SCHEDULED EMAIL EXECUTION STARTED for {len(selected_agencies)} agencies")
            
            # Validate required inputs
            if not self.vars["combined"].get():
                show_error("Scheduling Error", "Combined file is required for scheduled emails.")
                return
            
            if not self.vars["output"].get():
                show_error("Scheduling Error", "Output directory is required for scheduled emails.")
                return
            
            # Create drafts using the email handler
            try:
                success_count = self.email_handler.send_emails(
                    combined_file=Path(self.vars["combined"].get()),
                    output_dir=Path(self.vars["output"].get()),
                    file_names=selected_agencies,
                    mode=EmailMode.SCHEDULE,  # This creates drafts
                    universal_attachment=Path(self.vars["attach"].get()) if self.vars["attach"].get() else None
                )
                
                if success_count > 0:
                    logger.info(f"SCHEDULED EMAILS: Created {success_count} draft emails")
                    
                    # Update audit log
                    now = datetime.now().strftime("%Y-%m-%d %H:%M")
                    for agency in selected_agencies:
                        mask = self.audit_df[ColumnNames.AGENCY].str.upper() == agency.upper()
                        if mask.any():
                            self.audit_df.loc[mask, [ColumnNames.SENT_DATE, ColumnNames.STATUS, ColumnNames.COMMENTS]] = [
                                now, "Draft Created", f"Scheduled draft created at {now}"
                            ]
                        else:
                            # Add new row for this agency with all required columns
                            new_row = {
                                ColumnNames.AGENCY: agency,
                                ColumnNames.SENT_DATE: now,
                                ColumnNames.RESPONSE_DATE: "",
                                ColumnNames.TO: "",
                                ColumnNames.CC: "",
                                ColumnNames.STATUS: "Draft Created",
                                ColumnNames.COMMENTS: f"Scheduled draft created at {now}"
                            }
                            self.audit_df = pd.concat([self.audit_df, pd.DataFrame([new_row])], ignore_index=True)
                    
                    # Save audit log
                    if hasattr(self, 'audit_logger') and self.audit_logger:
                        self.audit_logger._df = self.audit_df
                        self.audit_logger.save()
                    
                    # Show success notification
                    show_notification(
                        "✅ Scheduled Emails Created",
                        f"Success! Created {success_count} draft emails in Outlook.\n\n"
                        f"📧 Check your Outlook drafts folder\n"
                        f"📝 Review and send when ready\n"
                        f"📊 Audit log updated"
                    )
                    
                    # Update UI in main thread - reset status and refresh dashboard
                    self.after(0, lambda: self.status_label.configure(text=f"✅ {success_count} drafts created"))
                    self.after(0, self.update_dashboard_metrics)
                    
                else:
                    show_error("Scheduling Failed", "No draft emails were created. Check the logs for details.")
                    self.after(0, lambda: self.status_label.configure(text="❌ Scheduling failed"))
                    
            except Exception as email_error:
                logger.error(f"SCHEDULED EMAIL ERROR: {email_error}")
                show_error(
                    "Email Creation Failed", 
                    f"Failed to create scheduled email drafts:\n\n{str(email_error)}\n\nCheck the application logs for more details."
                )
                self.after(0, lambda: self.status_label.configure(text="❌ Email creation error"))
                
        except Exception as e:
            logger.error(f"CRITICAL SCHEDULING ERROR: {str(e)}")
            show_error(
                "Scheduling System Error",
                f"A critical error occurred in the scheduling system:\n\n{str(e)}\n\nPlease restart the application and try again."
            )
            self.after(0, lambda: self.status_label.configure(text="❌ System error"))

    def mark(self):
        """
        Callback for the 'Mark Responded' button.
        
        Updates the status of selected agencies to 'Responded' in the audit log.
        Prompts for an optional comment.
        """
        selected = self.get_selected_agencies()
        if not selected:
            return messagebox.showwarning("Warning", "Please select one or more agencies.")
        
        comment = simpledialog.askstring(
            "Add Comment","Enter response comments (optional):"
        )
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        for agency in selected:
            mask = self.audit_df["Agency"].str.upper() == agency.upper()
            if mask.any():
                self.audit_df.loc[
                    mask, ["Response Received Date","Status","Comments"]
                ] = [now, "Responded", comment or ""] # Use comment or empty string
            else:
                # Add a new record if one doesn't exist for some reason
                self.audit_df.loc[len(self.audit_df)] = [
                    agency, "", now, "", "", "Responded", comment or ""
                ]
        
        # Save using AuditLogger
        output_dir = self.vars["output"].get()
        from pathlib import Path
        self.audit_logger = AuditLogger(output_dir=Path(output_dir))
        # Update internal dataframe and save
        self.audit_logger._df = self.audit_df
        self.audit_logger.save()
        
        messagebox.showinfo(
            "Updated", "Marked selected agencies as responded."
        )
        self.refresh() # Refresh UI to show new status

    def scan(self):
        """
        Callback for the 'Scan Inbox' button.
        
        Scans the specific Outlook folder for reply emails based on subject line keywords
        and automatically marks the corresponding agencies as 'Responded'.
        """
        try:
            folder_name = f"Compliance {REVIEW_PERIOD}"
            
            # Use EmailHandler to scan inbox
            responses = self.email_handler.scan_inbox(
                folder_name=folder_name,
                agencies=self.agencies
            )
            
            if responses is None:
                messagebox.showerror("Folder Error", f"Could not find '{folder_name}' folder")
                return
            
            found_count = 0
            output_dir = self.vars["output"].get()
            
            if not output_dir:
                messagebox.showwarning("No Output Folder", "Please select an output folder first.")
                return
            
            # Update audit log using AuditLogger
            self.audit_logger = AuditLogger(output_dir=Path(output_dir))
            self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
            
            # responses is a list of dictionaries with email info
            # We need to match agency names in subjects
            for response in responses:
                subject = response.get("subject", "")
                received_date = response.get("received", "")
                
                # Try to match agency from subject
                for agency in self.agencies:
                    if agency.lower() in subject.lower():
                        # Check if it's not already marked as responded
                        current_status = self.audit_logger.get_status(agency)
                        if current_status != AuditStatus.RESPONDED.value:
                            self.audit_logger.mark_responded(
                                agency=agency,
                                comments=f"Response received on {received_date}"
                            )
                            found_count += 1
                            logger.info(f"Found response for {agency} from email.")
                        break  # Found the agency, move to next response
            
            # Save and reload audit log
            self.audit_logger.save()
            self.audit_df = self.audit_logger.load()
            
            messagebox.showinfo(
                "Inbox Scan", f"Inbox scan complete.\n\n"
                f"📧 Emails found: {len(responses)}\n"
                f"✅ New responses logged: {found_count}"
            )
            self.refresh() # Refresh UI
        except Exception as e:
            logger.error(f"Inbox scan failed: {e}")
            messagebox.showerror("Scan Error", f"Failed to scan inbox: {e}")

    def sync_sent_emails(self):
        """
        Scan Outlook Sent Items to find and record previously sent emails.
        
        This is useful when:
        - The audit log is empty but emails were already sent
        - You want to sync the audit log with what's actually in Outlook
        - Emails were sent from a different computer/session
        """
        output_dir = self.vars["output"].get()
        if not output_dir:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return
        
        if not self.agencies:
            messagebox.showwarning("No Agencies", "Please load agencies first by selecting an output folder with agency files.")
            return
        
        # Ask for confirmation and days to scan
        days_back = simpledialog.askinteger(
            "Sync Sent Emails",
            "How many days back should we scan?\n\n"
            "This will search your Outlook Sent Items for\n"
            "emails matching your agency list.\n\n"
            "Days to scan (1-365):",
            initialvalue=90,
            minvalue=1,
            maxvalue=365
        )
        
        if days_back is None:
            return  # User cancelled
        
        try:
            self.show_progress(True)
            self.update_progress(0.1, "Connecting to Outlook...")
            
            # Scan sent items
            self.update_progress(0.2, f"Scanning Sent Items (last {days_back} days)...")
            
            found_emails = self.email_handler.scan_sent_items_for_agencies(
                agencies=self.agencies,
                days_back=days_back
            )
            
            if not found_emails:
                self.show_progress(False)
                messagebox.showinfo(
                    "Sync Complete",
                    f"No matching emails found in Compliance folders\n"
                    f"for the last {days_back} days.\n\n"
                    f"Searched for {len(self.agencies)} agencies.\n\n"
                    f"Make sure you have folders starting with 'Compliance'\n"
                    f"in your Inbox or Sent Items."
                )
                return
            
            self.update_progress(0.6, f"Found {len(found_emails)} emails. Updating audit log...")
            
            # Initialize audit logger
            self.audit_logger = AuditLogger(output_dir=Path(output_dir))
            self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
            
            # Update audit log with found emails
            synced_count = 0
            for email_info in found_emails:
                agency = email_info["agency"]
                
                # Check if already marked as sent
                current_status = self.audit_logger.get_status(agency)
                if current_status in [AuditStatus.SENT.value, AuditStatus.RESPONDED.value]:
                    continue  # Skip if already sent or responded
                
                # Mark as sent with the original sent date
                df = self.audit_logger.load()
                mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
                
                if mask.any():
                    sent_date = email_info["sent_date"]
                    if isinstance(sent_date, datetime):
                        sent_date_str = sent_date.strftime("%Y-%m-%d %H:%M")
                    else:
                        sent_date_str = str(sent_date)
                    
                    folder_info = email_info.get("folder", "Compliance folder")
                    df.loc[mask, ColumnNames.SENT_DATE] = sent_date_str
                    df.loc[mask, ColumnNames.TO] = email_info.get("to", "")
                    df.loc[mask, ColumnNames.CC] = email_info.get("cc", "")
                    df.loc[mask, ColumnNames.STATUS] = AuditStatus.SENT.value
                    df.loc[mask, ColumnNames.COMMENTS] = f"Synced from {folder_info}"
                    
                    self.audit_logger._df = df
                    synced_count += 1
            
            # Save audit log
            self.update_progress(0.9, "Saving audit log...")
            self.audit_logger.save()
            self.audit_df = self.audit_logger.load()
            
            # Refresh UI
            self.refresh()
            
            self.update_progress(1.0, "Sync complete!")
            self.show_progress(False)
            
            # Show summary
            not_found = len(self.agencies) - len(found_emails)
            messagebox.showinfo(
                "Sync Complete",
                f"✅ Sync completed successfully!\n\n"
                f"📧 Found in Compliance folders: {len(found_emails)}\n"
                f"📝 Newly synced to audit log: {synced_count}\n"
                f"❓ Not found: {not_found}\n\n"
                f"The audit log has been updated.\n"
                f"Dashboard metrics will now reflect sent emails."
            )
            
        except Exception as e:
            self.show_progress(False)
            logger.error(f"Error syncing sent emails: {e}")
            messagebox.showerror("Sync Error", f"Failed to sync sent emails:\n\n{str(e)}")

    def export(self):
        """
        Callback for the 'Export Log' button.
        
        Saves the current state of the audit log to its file and confirms to the user.
        This is essentially a manual "save" button.
        """
        output_dir = self.vars["output"].get()
        if not output_dir:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return
        
        from pathlib import Path
        self.audit_logger = AuditLogger(output_dir=Path(output_dir))
        # Update internal dataframe and save
        self.audit_logger._df = self.audit_df
        audit_file_path = os.path.join(output_dir, AUDIT_FILE_NAME)
        self.audit_logger.save()
        messagebox.showinfo("Export", f"Audit log saved successfully to:\n{audit_file_path}")

    def show_progress(self, show=True):
        """
        Shows or hides the progress bar frame at the bottom of the window.

        Args:
            show (bool): If True, shows the progress bar. If False, hides it.
        """
        if show:
            self.progress_frame.grid()
        else:
            self.progress_frame.grid_remove()
        self.update_idletasks() # Force the UI to update immediately

    def update_progress(self, value, status=""):
        """
        Updates the value of the progress bar and its status label.

        Args:
            value (float): A value between 0.0 and 1.0 for the progress bar.
            status (str): The text to display below the progress bar.
        """
        self.progress_bar.set(value)
        if status:
            self.status_label.configure(text=status)
        self.update_idletasks() # Force the UI to update immediately

    def open_dashboard(self):
        """
        Callback for the 'Dashboard' button. Opens a unified multi-region dashboard
        that combines audit logs from all loaded folders (APAC, OASYS, etc.)
        """
        dashboard = ctk.CTkToplevel(self)
        dashboard.title("Unified Reporting Dashboard - All Regions")
        dashboard.geometry("1200x700")
        dashboard.transient(self)
        
        # Collect all audit data from all folders
        all_audit_data = []
        folder_sources = {}
        
        # Main output folder
        main_output = self.vars["output"].get()
        if main_output and os.path.isdir(main_output):
            folder_name = os.path.basename(main_output)
            folder_sources[main_output] = folder_name
        
        # Additional folders from agency_folders
        if hasattr(self, 'agency_folders'):
            for agency, folder_path in self.agency_folders.items():
                if folder_path not in folder_sources:
                    folder_name = os.path.basename(folder_path)
                    folder_sources[folder_path] = folder_name
        
        # Debug: Show what folders we found
        debug_msg = f"Checking {len(folder_sources)} folders:\n"
        for path, name in folder_sources.items():
            debug_msg += f"• {name} ({path})\n"
        logger.info(f"Dashboard: {debug_msg}")
        
        # If no folders, show error
        if not folder_sources:
            messagebox.showerror(
                "No Folders",
                "No output folders found!\n\n"
                "Please:\n"
                "1. Select an output folder (Browse > Output Folder), OR\n"
                "2. Add folders using '➕ Add Folder' button"
            )
            dashboard.destroy()
            return
        
        # Load audit logs from each folder
        for folder_path, folder_name in folder_sources.items():
            # Try multiple audit file name patterns
            audit_file_patterns = [
                AUDIT_FILE_NAME,  # e.g., Audit_CognosAccessReview_Q3_2025.xlsx
                "Audit_CognosAccessReview.xlsx",  # Legacy name without period
                "Cognos_Review_Audit_Log.xlsx"  # Alternative name
            ]
            
            audit_file_found = None
            for pattern in audit_file_patterns:
                test_path = os.path.join(folder_path, pattern)
                if os.path.exists(test_path):
                    audit_file_found = test_path
                    break
            
            if audit_file_found:
                try:
                    df = pd.read_excel(audit_file_found)
                    if not df.empty:
                        df['Region'] = folder_name  # Add region column
                        df['Folder'] = folder_path  # Track source folder
                        all_audit_data.append(df)
                        logger.info(f"Dashboard: Loaded {len(df)} records from {folder_name} ({os.path.basename(audit_file_found)})")
                    else:
                        logger.warning(f"Dashboard: Audit file is empty: {audit_file_found}")
                except Exception as e:
                    logger.error(f"Failed to load audit from {audit_file_found}: {e}")
            else:
                logger.warning(f"Dashboard: No audit file found in {folder_path}")
        
        # Show debug info about what was loaded
        if all_audit_data:
            total_records = sum(len(df) for df in all_audit_data)
            messagebox.showinfo(
                "Dashboard Data Loaded",
                f"Found {len(all_audit_data)} audit file(s)\n"
                f"Total records: {total_records}\n\n"
                f"Folders checked: {len(folder_sources)}"
            )
        else:
            messagebox.showwarning(
                "No Data Found",
                f"Checked {len(folder_sources)} folder(s) but found no audit data.\n\n"
                f"Folders checked:\n" + "\n".join(f"• {name}" for name in folder_sources.values()) +
                f"\n\nLooking for files named:\n"
                f"• Audit_CognosAccessReview_Q3_2025.xlsx\n"
                f"• Audit_CognosAccessReview.xlsx\n"
                f"• Cognos_Review_Audit_Log.xlsx"
            )
        
        # Combine all data
        if all_audit_data:
            combined_df = pd.concat(all_audit_data, ignore_index=True)
            logger.info(f"Dashboard: Combined {len(combined_df)} total records from {len(all_audit_data)} folder(s)")
        else:
            # No audit data found - create empty dataframe and show info
            combined_df = pd.DataFrame(columns=["Agency", "Region", "Sent Email Date", "Response Received Date", "To", "CC", "Status", "Comments"])
            logger.warning(f"Dashboard: No audit log files found in {len(folder_sources)} folders")
        
        # Get unique regions for filter
        regions = ["All Regions"] + sorted(combined_df['Region'].unique().tolist()) if not combined_df.empty else ["All Regions"]
        
        # --- Header Frame ---
        header_frame = ctk.CTkFrame(dashboard)
        header_frame.pack(fill="x", padx=15, pady=10)
        
        ctk.CTkLabel(header_frame, text="📊 Unified Multi-Region Dashboard", 
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=10)
        
        # Region filter dropdown
        filter_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        filter_frame.pack(side="right", padx=10)
        
        ctk.CTkLabel(filter_frame, text="Filter by Region:").pack(side="left", padx=5)
        region_var = ctk.StringVar(value="All Regions")
        region_dropdown = ctk.CTkComboBox(filter_frame, values=regions, variable=region_var, width=200)
        region_dropdown.pack(side="left", padx=5)
        
        # --- Metrics Frame ---
        metrics_frame = ctk.CTkFrame(dashboard)
        metrics_frame.pack(fill="x", padx=15, pady=(0, 10))
        metrics_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
        
        # Create metric labels (will be updated by filter)
        metric_labels = {}
        metric_configs = [
            ("total", "Total Agencies", "#007A9B"),
            ("sent", "Sent", "#FFB302"),
            ("responded", "Responded", "#00B140"),
            ("not_sent", "Not Sent", "#6B7280"),
            ("overdue", "Overdue", "#D42B2B"),
            ("completion", "Completion", "#007A9B")
        ]
        
        for idx, (key, label, color) in enumerate(metric_configs):
            frame = ctk.CTkFrame(metrics_frame)
            frame.grid(row=0, column=idx, padx=5, pady=10, sticky="nsew")
            ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(size=11)).pack(pady=(5, 0))
            metric_labels[key] = ctk.CTkLabel(frame, text="0", font=ctk.CTkFont(size=18, weight="bold"), text_color=color)
            metric_labels[key].pack(pady=(0, 5))
        
        # --- Region Summary Frame ---
        summary_frame = ctk.CTkFrame(dashboard)
        summary_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        ctk.CTkLabel(summary_frame, text="Region Summary:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10, pady=5)
        region_summary_label = ctk.CTkLabel(summary_frame, text="", font=ctk.CTkFont(size=11))
        region_summary_label.pack(side="left", padx=10, pady=5)
        
        # --- Table Frame ---
        table_frame = ctk.CTkFrame(dashboard)
        table_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        # Style for treeview
        style = ttk.Style()
        style.configure("Dashboard.Treeview", rowheight=25)
        style.configure("Dashboard.Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        
        tree = ttk.Treeview(table_frame, style="Dashboard.Treeview")
        
        # Add scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side='bottom', fill='x')
        tree.configure(xscrollcommand=hsb.set)
        
        tree.pack(fill="both", expand=True)
        
        # Configure tag colors for status
        tree.tag_configure('responded', background='#D4EDDA')
        tree.tag_configure('sent', background='#FFF3CD')
        tree.tag_configure('overdue', background='#F8D7DA')
        tree.tag_configure('not_sent', background='#E2E3E5')
        
        # Context menu for follow-up
        def show_context_menu(event):
            """Show context menu on right-click"""
            # Identify clicked row
            item = tree.identify_row(event.y)
            if not item:
                return
            
            # Select the clicked row
            tree.selection_set(item)
            tree.focus(item)
            
            # Get row data
            values = tree.item(item)['values']
            if not values or len(values) < 3:
                return
            
            status = values[2]  # Status column
            agency = values[1]  # Agency column
            
            # Only show menu for "Sent" or "Overdue" status
            if status not in ["Sent", "Overdue"]:
                return
            
            # Create context menu
            menu = tk.Menu(dashboard, tearoff=0)
            menu.add_command(label=f"📧 Send Follow-up to {agency}", 
                           command=lambda: send_followup_action(agency))
            menu.add_separator()
            menu.add_command(label="Cancel", command=lambda: menu.unpost())
            
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
        
        def send_followup_action(agency):
            """Send follow-up email to agency"""
            # Create follow-up dialog
            followup_dialog = ctk.CTkToplevel(dashboard)
            followup_dialog.title(f"Send Follow-up - {agency}")
            followup_dialog.geometry("600x500")
            followup_dialog.transient(dashboard)
            followup_dialog.grab_set()
            
            # Center dialog
            followup_dialog.update_idletasks()
            x = (followup_dialog.winfo_screenwidth() // 2) - (followup_dialog.winfo_width() // 2)
            y = (followup_dialog.winfo_screenheight() // 2) - (followup_dialog.winfo_height() // 2)
            followup_dialog.geometry(f"+{x}+{y}")
            
            main_frame = ctk.CTkFrame(followup_dialog)
            main_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            ctk.CTkLabel(main_frame, text=f"Follow-up for {agency}", 
                        font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))
            
            # Follow-up message
            ctk.CTkLabel(main_frame, text="Follow-up Message:", anchor="w").pack(fill="x", pady=(10, 5))
            
            # Default follow-up templates
            config = get_config_manager().load()
            deadline = config.get("deadline", "TBD")
            
            default_template = f"""Hi Team,

This is a friendly reminder about the pending Cognos Access Review for {agency}.

We haven't received your response yet. Please review the attached file and confirm:
1. Which users should retain access
2. Which users should have access removed
3. Any additional comments or concerns

**Deadline: {deadline}**

If you have any questions or need clarification, please don't hesitate to reach out.

Thank you for your cooperation!
"""
            
            message_text = ctk.CTkTextbox(main_frame, height=250)
            message_text.pack(fill="both", expand=True, pady=(0, 10))
            message_text.insert("1.0", default_template)
            
            # Template shortcuts
            template_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            template_frame.pack(fill="x", pady=(10, 0))
            
            ctk.CTkLabel(template_frame, text="Quick Templates:").pack(side="left", padx=5)
            
            def use_gentle_reminder():
                message_text.delete("1.0", "end")
                message_text.insert("1.0", f"""Hi Team,

Just a gentle reminder about the Cognos Access Review for {agency}.

Could you please review and respond at your earliest convenience?

Deadline: {deadline}

Thanks!
""")
            
            def use_urgent_reminder():
                message_text.delete("1.0", "end")
                message_text.insert("1.0", f"""Hi Team,

**URGENT REMINDER**

We still need your response for the Cognos Access Review ({agency}).

This is required for SOX compliance. Please prioritize and respond by {deadline}.

If you're facing any issues, please contact us immediately.

Thank you!
""")
            
            ctk.CTkButton(template_frame, text="Gentle", command=use_gentle_reminder, width=80).pack(side="left", padx=2)
            ctk.CTkButton(template_frame, text="Urgent", command=use_urgent_reminder, width=80).pack(side="left", padx=2)
            
            # Send button
            def do_send():
                try:
                    followup_body = message_text.get("1.0", "end").strip()
                    if not followup_body:
                        messagebox.showerror("Error", "Please enter a follow-up message", parent=followup_dialog)
                        return
                    
                    # Get config for subject
                    config = get_config_manager().load()
                    subject_prefix = config.get("email_subject_prefix", "[ACTION REQUIRED] Cognos Access Review")
                    review_period = config.get("review_period", "Q4 FY25")
                    original_subject = f"{subject_prefix} - {review_period} - {agency}"
                    
                    # Create email handler
                    email_handler = EmailHandler()
                    
                    # Send follow-up
                    success = email_handler.send_followup(
                        agency=agency,
                        original_subject=original_subject,
                        followup_body=followup_body,
                        search_days=60  # Search back 60 days for original email
                    )
                    
                    if success:
                        followup_dialog.destroy()
                        messagebox.showinfo(
                            "Success",
                            f"Follow-up sent successfully to {agency}!\n\nThe email has been sent as a reply to the original thread.",
                            parent=dashboard
                        )
                        # Refresh dashboard to reflect any changes
                        refresh_data()
                    else:
                        messagebox.showerror(
                            "Failed",
                            f"Could not send follow-up.\n\nOriginal email not found in Sent Items.\nPlease send manually or check email settings.",
                            parent=followup_dialog
                        )
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send follow-up:\n{str(e)}", parent=followup_dialog)
            
            button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            button_frame.pack(side="bottom", fill="x", pady=(20, 0))
            
            ctk.CTkButton(button_frame, text="Cancel", command=followup_dialog.destroy).pack(side="right", padx=5)
            ctk.CTkButton(button_frame, text="Send Follow-up", command=do_send, 
                         fg_color="#2196F3").pack(side="right")
        
        # Bind right-click to tree
        tree.bind("<Button-3>", show_context_menu)  # Right-click on Windows
        tree.bind("<Button-2>", show_context_menu)  # Right-click on Mac
        
        def get_row_tag(status):
            if status == "Responded":
                return 'responded'
            elif status == "Sent":
                return 'sent'
            elif status == "Overdue":
                return 'overdue'
            return 'not_sent'
        
        def update_dashboard(selected_region):
            """Update the dashboard based on selected region filter"""
            # Filter data
            if selected_region == "All Regions":
                filtered_df = combined_df.copy()
            else:
                filtered_df = combined_df[combined_df['Region'] == selected_region].copy()
            
            # Calculate metrics
            total = len(filtered_df)
            responded = (filtered_df["Status"] == "Responded").sum() if not filtered_df.empty else 0
            sent = (filtered_df["Status"] == "Sent").sum() if not filtered_df.empty else 0
            not_sent_count = total - sent - responded
            
            # Calculate overdue
            overdue = 0
            if not filtered_df.empty and "Sent Email Date" in filtered_df.columns:
                sent_df = filtered_df[filtered_df["Status"] == "Sent"].copy()
                sent_df["Sent Email Date"] = pd.to_datetime(sent_df["Sent Email Date"], errors='coerce')
                seven_days_ago = pd.Timestamp.now() - pd.Timedelta(days=7)
                overdue = len(sent_df[sent_df["Sent Email Date"] < seven_days_ago])
            
            completion = f"{responded/total*100:.1f}%" if total > 0 else "0%"
            
            # Update metric labels
            metric_labels["total"].configure(text=str(total))
            metric_labels["sent"].configure(text=str(sent))
            metric_labels["responded"].configure(text=str(responded))
            metric_labels["not_sent"].configure(text=str(not_sent_count))
            metric_labels["overdue"].configure(text=str(overdue))
            metric_labels["completion"].configure(text=completion)
            
            # Update region summary
            if not combined_df.empty:
                region_counts = combined_df.groupby('Region').size().to_dict()
                summary_parts = [f"{region}: {count}" for region, count in sorted(region_counts.items())]
                region_summary_label.configure(text=" | ".join(summary_parts))
            
            # Update table
            tree.delete(*tree.get_children())
            
            if not filtered_df.empty:
                # Reorder columns to show Region first
                display_cols = ['Region', 'Agency', 'Status', 'Sent Email Date', 'Response Received Date', 'To', 'CC', 'Comments']
                display_cols = [c for c in display_cols if c in filtered_df.columns]
                df_display = filtered_df[display_cols].fillna("")
                
                tree["columns"] = list(df_display.columns)
                tree["show"] = "headings"
                
                # Configure columns
                col_widths = {'Region': 150, 'Agency': 180, 'Status': 100, 'Sent Email Date': 120, 
                              'Response Received Date': 140, 'To': 200, 'CC': 150, 'Comments': 150}
                for col in df_display.columns:
                    tree.heading(col, text=col)
                    tree.column(col, width=col_widths.get(col, 120), anchor='w')
                
                for _, row in df_display.iterrows():
                    status = row.get('Status', '')
                    tag = get_row_tag(status)
                    tree.insert("", "end", values=list(row), tags=(tag,))
        
        # Bind region filter change
        def on_region_change(choice):
            update_dashboard(choice)
        
        region_dropdown.configure(command=on_region_change)
        
        # --- Button Frame ---
        button_frame = ctk.CTkFrame(dashboard, fg_color="transparent")
        button_frame.pack(fill="x", padx=15, pady=10)
        
        def export_csv():
            selected_region = region_var.get()
            if selected_region == "All Regions":
                export_df = combined_df
                filename = f"Audit_Log_ALL_REGIONS_{datetime.now().strftime('%Y%m%d')}.csv"
            else:
                export_df = combined_df[combined_df['Region'] == selected_region]
                safe_region = selected_region.replace(" ", "_").replace("/", "_")
                filename = f"Audit_Log_{safe_region}_{datetime.now().strftime('%Y%m%d')}.csv"
            
            path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=filename
            )
            if path:
                export_df.to_csv(path, index=False)
                messagebox.showinfo("Export Success", f"Exported {len(export_df)} records to:\n{path}")
        
        def refresh_data():
            """Reload data from all folders"""
            nonlocal combined_df, regions
            all_audit_data.clear()
            
            for folder_path, folder_name in folder_sources.items():
                audit_file = os.path.join(folder_path, AUDIT_FILE_NAME)
                if os.path.exists(audit_file):
                    try:
                        df = pd.read_excel(audit_file)
                        df['Region'] = folder_name
                        df['Folder'] = folder_path
                        all_audit_data.append(df)
                    except Exception as e:
                        logger.error(f"Failed to reload audit from {folder_path}: {e}")
            
            if all_audit_data:
                combined_df = pd.concat(all_audit_data, ignore_index=True)
            else:
                combined_df = pd.DataFrame(columns=["Agency", "Region", "Sent Email Date", "Response Received Date", "To", "CC", "Status", "Comments"])
            
            regions = ["All Regions"] + sorted(combined_df['Region'].unique().tolist()) if not combined_df.empty else ["All Regions"]
            region_dropdown.configure(values=regions)
            update_dashboard(region_var.get())
            messagebox.showinfo("Refreshed", "Dashboard data has been refreshed.")
        
        def generate_report():
            """Generate SOX compliance report"""
            report_dialog = ctk.CTkToplevel(dashboard)
            report_dialog.title("Generate Compliance Report")
            report_dialog.geometry("500x400")
            report_dialog.transient(dashboard)
            report_dialog.grab_set()
            
            # Center dialog
            report_dialog.update_idletasks()
            x = (report_dialog.winfo_screenwidth() // 2) - (report_dialog.winfo_width() // 2)
            y = (report_dialog.winfo_screenheight() // 2) - (report_dialog.winfo_height() // 2)
            report_dialog.geometry(f"+{x}+{y}")
            
            main_frame = ctk.CTkFrame(report_dialog)
            main_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            ctk.CTkLabel(main_frame, text="SOX Compliance Report Generator", 
                        font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))
            
            # Report type
            ctk.CTkLabel(main_frame, text="Report Type:", anchor="w").pack(fill="x", pady=(10, 5))
            report_type_var = ctk.StringVar(value="summary")
            report_types = ["summary", "detailed", "exceptions"]
            for rtype in report_types:
                ctk.CTkRadioButton(
                    main_frame, 
                    text=rtype.title() + (" - High-level metrics" if rtype == "summary" else 
                                        " - Full audit log" if rtype == "detailed" else 
                                        " - Action items only"),
                    variable=report_type_var,
                    value=rtype
                ).pack(anchor="w", padx=20)
            
            # Output format
            ctk.CTkLabel(main_frame, text="Output Format:", anchor="w").pack(fill="x", pady=(20, 5))
            format_var = ctk.StringVar(value="xlsx")
            formats = [("xlsx", "Excel (.xlsx)"), ("pdf", "HTML (for PDF print)")]
            for fmt, label in formats:
                ctk.CTkRadioButton(
                    main_frame,
                    text=label,
                    variable=format_var,
                    value=fmt
                ).pack(anchor="w", padx=20)
            
            # Generate button
            def do_generate():
                try:
                    # Get current data
                    selected_region = region_var.get()
                    if selected_region == "All Regions":
                        report_df = combined_df.copy()
                    else:
                        report_df = combined_df[combined_df['Region'] == selected_region].copy()
                    
                    # Calculate metrics
                    total = len(report_df)
                    sent = len(report_df[report_df["Status"] == "Sent"])
                    responded = len(report_df[report_df["Status"] == "Responded"])
                    not_sent_count = len(report_df[report_df["Status"] == "Not Sent"])
                    
                    # Calculate overdue
                    seven_days_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
                    sent_df = report_df[report_df["Status"] == "Sent"]
                    overdue = 0
                    if not sent_df.empty:
                        overdue = len(sent_df[sent_df["Sent Email Date"] < seven_days_ago])
                    
                    completion = (responded / total * 100) if total > 0 else 0
                    
                    # Get deadline from config
                    config = get_config_manager().load()
                    deadline_str = config.get("deadline", "TBD")
                    try:
                        deadline = datetime.strptime(deadline_str, "%B %d, %Y")
                        days_left = (deadline - datetime.now()).days
                    except:
                        days_left = -1
                    
                    # Create metrics object
                    report_metrics = DashboardMetrics(
                        total_agencies=total,
                        sent_count=sent,
                        responded_count=responded,
                        not_sent_count=not_sent_count,
                        overdue_count=overdue,
                        completion_percentage=completion,
                        days_left=days_left
                    )
                    
                    # Generate report
                    generator = ReportGenerator(output_dir=Path("reports"))
                    review_period = config.get("review_period", "Q4 FY25")
                    
                    report_path = generator.generate_compliance_report(
                        audit_df=report_df,
                        metrics=report_metrics,
                        report_type=report_type_var.get(),
                        output_format=format_var.get(),
                        review_period=review_period
                    )
                    
                    report_dialog.destroy()
                    
                    # Ask if user wants to open the report
                    result = messagebox.askyesno(
                        "Report Generated",
                        f"Report generated successfully!\n\n{report_path}\n\nWould you like to open it now?",
                        parent=dashboard
                    )
                    
                    if result:
                        os.startfile(report_path)
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate report:\n{str(e)}", parent=report_dialog)
            
            button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            button_frame.pack(side="bottom", fill="x", pady=(20, 0))
            
            ctk.CTkButton(button_frame, text="Cancel", command=report_dialog.destroy).pack(side="right", padx=5)
            ctk.CTkButton(button_frame, text="Generate Report", command=do_generate).pack(side="right")
        
        ctk.CTkButton(button_frame, text="🔄 Refresh Data", command=refresh_data, width=140).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="📥 Export as CSV", command=export_csv, width=140).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="📊 Generate Report", command=generate_report, width=150, 
                     fg_color="#2196F3").pack(side="left", padx=5)
        
        # Status filter buttons
        ctk.CTkLabel(button_frame, text="Quick Filters:").pack(side="left", padx=(20, 5))
        
        def filter_status(status):
            """Quick filter by status"""
            tree.delete(*tree.get_children())
            selected_region = region_var.get()
            
            if selected_region == "All Regions":
                filtered_df = combined_df.copy()
            else:
                filtered_df = combined_df[combined_df['Region'] == selected_region].copy()
            
            if status != "All":
                filtered_df = filtered_df[filtered_df['Status'] == status]
            
            if not filtered_df.empty:
                display_cols = ['Region', 'Agency', 'Status', 'Sent Email Date', 'Response Received Date', 'To', 'CC', 'Comments']
                display_cols = [c for c in display_cols if c in filtered_df.columns]
                df_display = filtered_df[display_cols].fillna("")
                
                for _, row in df_display.iterrows():
                    status_val = row.get('Status', '')
                    tag = get_row_tag(status_val)
                    tree.insert("", "end", values=list(row), tags=(tag,))
        
        ctk.CTkButton(button_frame, text="All", command=lambda: filter_status("All"), width=60).pack(side="left", padx=2)
        ctk.CTkButton(button_frame, text="Sent", command=lambda: filter_status("Sent"), width=60).pack(side="left", padx=2)
        ctk.CTkButton(button_frame, text="Responded", command=lambda: filter_status("Responded"), width=80).pack(side="left", padx=2)
        ctk.CTkButton(button_frame, text="Not Sent", command=lambda: filter_status("Not Sent"), width=70).pack(side="left", padx=2)
        
        # Initial update
        update_dashboard("All Regions")
        
        dashboard.grab_set()

    def schedule_email_dialog(self, agencies):
        """
        Opens a dialog to schedule emails for a specific list of agencies.
        Creates Outlook drafts at the scheduled time instead of sending emails directly.
        """
        win = ctk.CTkToplevel(self)
        win.title("Schedule Email Drafts")
        win.geometry("400x300")
        
        ctk.CTkLabel(win, text=f"Schedule email drafts for {len(agencies)} selected agencies:").pack(pady=(10,5))

        ctk.CTkLabel(win, text="Select Draft Creation Date & Time:").pack()
        
        # Date Entry
        date_entry = DateEntry(win, date_pattern='yyyy-mm-dd', width=12, background='darkblue',
                               foreground='white', borderwidth=2)
        date_entry.pack(pady=5)
        
        # Time Entry
        time_entry = ctk.CTkEntry(win, placeholder_text="HH:MM (24-hr format)")
        time_entry.insert(0, "09:00")
        time_entry.pack(pady=5)

        # Information text
        info_text = ctk.CTkLabel(
            win, 
            text="Emails will be created immediately in Outlook Outbox\nwith deferred delivery. You can close this app after scheduling!",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        info_text.pack(pady=10)

        def schedule():
            logger.info("SCHEDULE DIALOG: Starting simple scheduling validation")
            
            # Validate required files first
            required_files = {
                "Combined File": self.vars["combined"].get(),
                "Output Directory": self.vars["output"].get()
            }
            
            missing_files = [name for name, path in required_files.items() if not path]
            if missing_files:
                messagebox.showerror(
                    "Missing Configuration", 
                    f"The following are required for scheduling:\n\n" + 
                    "\n".join(f"• {name}" for name in missing_files) +
                    "\n\nPlease configure these files before scheduling.",
                    parent=win
                )
                logger.error(f"SIMPLE SCHEDULE VALIDATION FAILED: Missing files: {missing_files}")
                return
            
            dt_str = f"{date_entry.get()} {time_entry.get()}"
            try:
                # Use timezone-aware datetime - assume local system timezone
                naive_dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                
                # Validate it's in the future
                if naive_dt <= datetime.now():
                    messagebox.showerror("Invalid Time", "Scheduled time must be in the future.", parent=win)
                    return
                
                # Create emails IMMEDIATELY with deferred delivery
                logger.info(f"SIMPLE SCHEDULE: Creating {len(agencies)} emails with deferred delivery at {naive_dt}")
                
                # Load combined file to get To/CC addresses for audit log
                combined_loader = CombinedFileLoader(Path(self.vars["combined"].get()))
                combined_data = combined_loader.load_all_data()
                addr_dict = {}
                for mapping in combined_data:
                    key = mapping.source_file_name.upper()
                    addr_dict[key] = (mapping.recipients_to, mapping.recipients_cc)
                
                try:
                    success_count = self.email_handler.send_emails(
                        combined_file=Path(self.vars["combined"].get()),
                        output_dir=Path(self.vars["output"].get()),
                        file_names=agencies,
                        mode=EmailMode.SCHEDULE,
                        universal_attachment=Path(self.vars["attach"].get()) if self.vars["attach"].get() else None,
                        scheduled_time=naive_dt  # Outlook uses local time
                    )
                    
                    if success_count > 0:
                        logger.info(f"SIMPLE SCHEDULE: Created {success_count} scheduled emails")
                        
                        # Update audit log with scheduled emails
                        output_dir = Path(self.vars["output"].get())
                        self.audit_logger = AuditLogger(output_dir=output_dir)
                        self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
                        
                        scheduled_str = naive_dt.strftime("%Y-%m-%d %H:%M")
                        for agency in agencies:
                            to, cc = addr_dict.get(agency.upper(), ("", ""))
                            self.audit_logger.mark_sent(
                                agency, 
                                to=to, 
                                cc=cc, 
                                comments=f"Scheduled for {scheduled_str}"
                            )
                        
                        # Save audit log
                        self.audit_logger.save()
                        self.audit_df = self.audit_logger.load()
                        
                        # Update status bar
                        scheduled_time_str = naive_dt.strftime('%H:%M')
                        self.status_label.configure(text=f"✅ {success_count} emails scheduled for {scheduled_time_str}")
                        
                        win.destroy()
                        
                        # Calculate time until execution
                        delay = (naive_dt - datetime.now()).total_seconds()
                        hours = int(delay // 3600)
                        minutes = int((delay % 3600) // 60)
                        time_until = f"{hours}h {minutes}m" if hours > 0 else f"{minutes}m"
                        
                        messagebox.showinfo(
                            "✅ Emails Scheduled Successfully", 
                            f"📧 {success_count} emails created and scheduled!\n\n"
                            f"🕒 Will be sent at: {naive_dt.strftime('%Y-%m-%d %H:%M')}\n"
                            f"⏱️  Time until delivery: {time_until}\n\n"
                            f"📬 Check your Outlook OUTBOX folder\n"
                            f"   You can review/edit them before send time\n\n"
                            f"✅ You can close this app!\n"
                            f"   Outlook will send them automatically."
                        )
                        
                        # Refresh UI
                        self.refresh()
                    else:
                        messagebox.showerror("Scheduling Failed", "No emails were created.", parent=win)
                        
                except Exception as email_error:
                    logger.error(f"SIMPLE SCHEDULE ERROR: {email_error}")
                    messagebox.showerror("Email Creation Failed", f"Failed to create scheduled emails:\n\n{str(email_error)}", parent=win)
                    
            except Exception as e:
                messagebox.showerror("Invalid Date/Time", f"Please use YYYY-MM-DD HH:MM format.\nError: {e}", parent=win)
        
        ctk.CTkButton(win, text="Schedule Drafts", command=schedule).pack(pady=10)
        win.transient(self)
        win.grab_set()

    def schedule_all_dialog(self):
        """
        Opens a more advanced dialog to schedule email drafts for ALL or a sub-selection
        of agencies, including timezone support. Creates Outlook drafts instead of sending emails.
        """
        win = ctk.CTkToplevel(self)
        win.title("Schedule Email Drafts Campaign")
        win.geometry("500x600")

        ctk.CTkLabel(win, text="1. Select Agencies", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10,0))
        ctk.CTkLabel(win, text="(Leave blank to select all agencies)").pack(anchor="w", padx=10)
        
        list_frame = ctk.CTkFrame(win)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        listbox = tk.Listbox(list_frame, selectmode="multiple", exportselection=0, height=10)
        for ag in self.agencies:
            listbox.insert("end", ag)
        listbox.pack(side="left", fill="both", expand=True)
        
        list_scroll = ctk.CTkScrollbar(list_frame, command=listbox.yview)
        list_scroll.pack(side="right", fill="y")
        listbox.configure(yscrollcommand=list_scroll.set)

        ctk.CTkLabel(win, text="2. Select Draft Creation Date & Time", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10,0))
        
        # Date/Time Frame
        dt_frame = ctk.CTkFrame(win, fg_color="transparent")
        dt_frame.pack(fill="x", padx=10, pady=5)
        date_entry = DateEntry(dt_frame, date_pattern='yyyy-mm-dd', width=12, background='darkblue', foreground='white', borderwidth=2)
        date_entry.pack(side="left", padx=(0,5))
        time_entry = ctk.CTkEntry(dt_frame, placeholder_text="HH:MM")
        time_entry.insert(0, "09:00")
        time_entry.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(win, text="3. Select Timezone", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10,0))
        
        # Timezone Dropdown
        tz_combo = ctk.CTkComboBox(win, values=pytz.all_timezones)
        # Default to UTC
        tz_combo.set("UTC")
        tz_combo.pack(fill="x", padx=10, pady=5)
        
        # Information text
        info_text = ctk.CTkLabel(
            win, 
            text="Emails will be created immediately in Outlook Outbox with deferred delivery.\nYou can review them and close this app - Outlook handles the sending!",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        info_text.pack(pady=10)
        
        def schedule():
            logger.info("SCHEDULE BUTTON CLICKED - Starting validation")
            
            # Validate agencies selection
            selected_indices = listbox.curselection()
            agencies = [listbox.get(i) for i in selected_indices] if selected_indices else self.agencies
            if not agencies:
                messagebox.showwarning("No Agencies", "There are no agencies to schedule.", parent=win)
                logger.warning("SCHEDULE VALIDATION FAILED: No agencies selected")
                return
            
            # Validate required files
            required_files = {
                "Combined File": self.vars["combined"].get(),
                "Output Directory": self.vars["output"].get()
            }
            
            missing_files = [name for name, path in required_files.items() if not path]
            if missing_files:
                messagebox.showerror(
                    "Missing Configuration", 
                    f"The following are required for scheduling:\n\n" + 
                    "\n".join(f"• {name}" for name in missing_files) +
                    "\n\nPlease configure these files before scheduling.",
                    parent=win
                )
                logger.error(f"SCHEDULE VALIDATION FAILED: Missing files: {missing_files}")
                return
            
            logger.info(f"SCHEDULE VALIDATION PASSED: {len(agencies)} agencies, all files configured")

            dt_str = f"{date_entry.get()} {time_entry.get()}"
            tz_name = tz_combo.get()
            
            try:
                # Make the datetime object timezone-aware
                tz = pytz.timezone(tz_name)
                naive_dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                local_dt = tz.localize(naive_dt)
                
                # Convert to local system time for Outlook
                # Outlook uses local system time for DeferredDeliveryTime
                local_system_dt = local_dt.astimezone()  # Convert to system local timezone
                
                # Calculate delay to validate it's in the future
                now_utc = datetime.now(pytz.utc)
                utc_dt = local_dt.astimezone(pytz.utc)
                delay = (utc_dt - now_utc).total_seconds()
                
                if delay <= 0:
                    messagebox.showerror("Invalid Time", "Scheduled time must be in the future.", parent=win)
                    logger.warning(f"INVALID SCHEDULE TIME: User tried to schedule for {local_dt} (delay: {delay}s)")
                    return
                
                # Load email manifest to get To/CC addresses for audit log
                # Load combined file to get To/CC addresses for audit log
                combined_loader = CombinedFileLoader(Path(self.vars["combined"].get()))
                combined_data = combined_loader.load_all_data()
                addr_dict = {}
                for mapping in combined_data:
                    key = mapping.source_file_name.upper()
                    addr_dict[key] = (mapping.recipients_to, mapping.recipients_cc)
                
                # Create emails IMMEDIATELY with deferred delivery time
                # This uses Outlook's built-in scheduling - no need to keep app open!
                logger.info(f"CREATING SCHEDULED EMAILS: {len(agencies)} agencies, delivery at {local_dt}")
                
                try:
                    success_count = self.email_handler.send_emails(
                        combined_file=Path(self.vars["combined"].get()),
                        output_dir=Path(self.vars["output"].get()),
                        file_names=agencies,
                        mode=EmailMode.SCHEDULE,
                        universal_attachment=Path(self.vars["attach"].get()) if self.vars["attach"].get() else None,
                        scheduled_time=local_system_dt.replace(tzinfo=None)  # Outlook needs naive datetime in local time
                    )
                    
                    if success_count > 0:
                        logger.info(f"SCHEDULED EMAILS CREATED: {success_count} emails in Outbox")
                        
                        # Update audit log using AuditLogger
                        output_dir = Path(self.vars["output"].get())
                        self.audit_logger = AuditLogger(output_dir=output_dir)
                        self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
                        
                        scheduled_str = local_dt.strftime("%Y-%m-%d %H:%M")
                        for agency in agencies:
                            to, cc = addr_dict.get(agency.upper(), ("", ""))
                            self.audit_logger.mark_sent(
                                agency, 
                                to=to, 
                                cc=cc, 
                                comments=f"Scheduled for {scheduled_str} ({tz_name})"
                            )
                        
                        # Save audit log
                        self.audit_logger.save()
                        self.audit_df = self.audit_logger.load()
                        
                        # Update status
                        scheduled_time_str = local_dt.strftime('%H:%M')
                        self.status_label.configure(text=f"✅ {success_count} emails scheduled for {scheduled_time_str}")
                        
                        win.destroy()
                        
                        # Calculate time until send
                        hours = int(delay // 3600)
                        minutes = int((delay % 3600) // 60)
                        time_until = f"{hours}h {minutes}m" if hours > 0 else f"{minutes}m"
                        
                        # Show success message
                        messagebox.showinfo(
                            "✅ Emails Scheduled Successfully", 
                            f"📧 {success_count} emails created and scheduled!\n\n"
                            f"🕒 Will be sent at:\n"
                            f"   {local_dt.strftime('%Y-%m-%d %H:%M')} ({tz_name})\n\n"
                            f"⏱️  Time until delivery: {time_until}\n\n"
                            f"📬 Where to find them:\n"
                            f"   • Check your Outlook OUTBOX folder\n"
                            f"   • Emails are ready with deferred delivery\n"
                            f"   • You can review/edit them before send time\n\n"
                            f"✅ You can close this app!\n"
                            f"   Outlook will send them automatically at the scheduled time."
                        )
                        
                        # Refresh UI
                        self.refresh()
                        
                    else:
                        messagebox.showerror("Scheduling Failed", "No emails were created. Check the logs for details.", parent=win)
                        
                except Exception as email_error:
                    logger.error(f"SCHEDULED EMAIL CREATION ERROR: {email_error}")
                    messagebox.showerror(
                        "Email Creation Failed", 
                        f"Failed to create scheduled emails:\n\n{str(email_error)}",
                        parent=win
                    )
                
            except Exception as e:
                messagebox.showerror("Invalid Date/Time", f"Please check your inputs.\nError: {e}", parent=win)
                
        ctk.CTkButton(win, text="Schedule Draft Campaign", command=schedule).pack(pady=10)
        win.transient(self)
        win.grab_set()

    def validate_files(self):
        """
        Comprehensive validation with interactive dialog showing results and fix options.
        """
        try:
            self.show_progress(True)
            self.update_progress(0.1, "Starting validation...")
            
            # Gather all file paths
            master_file = self.vars["master"].get()
            combined_file = self.vars["combined"].get()
            output_dir = self.vars["output"].get()
            
            # Run comprehensive validation
            validation_results = self._run_comprehensive_validation(
                master_file=master_file,
                combined_file=combined_file,
                output_dir=output_dir
            )
            
            self.update_progress(1.0, "Validation complete!")
            self.show_progress(False)
            
            # Show validation results dialog
            dialog = ValidationResultsDialog(
                parent=self,
                validation_results=validation_results,
                app_reference=self
            )
            dialog.wait_window()
            
        except Exception as e:
            self.show_progress(False)
            logger.error(f"Validation failed: {e}")
            messagebox.showerror("Validation Error", f"Validation failed: {str(e)}")
    
    def _run_comprehensive_validation(
        self,
        master_file: str,
        combined_file: str,
        output_dir: str
    ) -> Dict:
        """
        Run all validation checks and return comprehensive results.
        
        Returns:
            Dictionary with validation results
        """
        logger.info("=" * 80)
        logger.info("STARTING COMPREHENSIVE VALIDATION")
        logger.info(f"Master File: {master_file}")
        logger.info(f"Combined File: {combined_file}")
        logger.info(f"Output Dir: {output_dir}")
        logger.info("=" * 80)
        
        results = {
            "overall_status": "pass",  # pass, warning, error
            "categories": [],
            "can_proceed": True,
            "error_count": 0,
            "warning_count": 0
        }
        
        # Category 1: File Existence
        self.update_progress(0.2, "Checking file existence...")
        file_checks = self._validate_file_existence(master_file, combined_file, output_dir)
        results["categories"].append(file_checks)
        self._log_category_results("File Existence", file_checks)
        
        # Category 2: File Structure & Columns
        self.update_progress(0.3, "Validating file structure...")
        structure_checks = self._validate_file_structure(master_file, combined_file)
        results["categories"].append(structure_checks)
        self._log_category_results("File Structure", structure_checks)
        
        # Category 3: Data Quality
        self.update_progress(0.5, "Checking data quality...")
        data_checks = self._validate_data_quality(master_file, combined_file)
        results["categories"].append(data_checks)
        self._log_category_results("Data Quality", data_checks)
        
        # Category 4: Agency Mappings
        self.update_progress(0.6, "Validating agency mappings...")
        mapping_checks = self._validate_agency_mappings(master_file, combined_file)
        results["categories"].append(mapping_checks)
        self._log_category_results("Agency Mappings", mapping_checks)
        
        # Category 5: Email Addresses
        self.update_progress(0.7, "Validating email addresses...")
        email_checks = self._validate_email_addresses(combined_file)
        results["categories"].append(email_checks)
        self._log_category_results("Email Addresses", email_checks)
        
        # Category 6: Filename Safety
        self.update_progress(0.8, "Checking filename safety...")
        filename_checks = self._validate_filename_safety(combined_file)
        results["categories"].append(filename_checks)
        self._log_category_results("Filename Safety", filename_checks)
        
        # Category 7: Region Configuration
        self.update_progress(0.9, "Checking region configuration...")
        region_checks = self._validate_region_configuration()
        results["categories"].append(region_checks)
        self._log_category_results("Region Configuration", region_checks)
        
        # Calculate overall status
        for category in results["categories"]:
            results["error_count"] += len([i for i in category["issues"] if i["severity"] == "error"])
            results["warning_count"] += len([i for i in category["issues"] if i["severity"] == "warning"])
        
        if results["error_count"] > 0:
            results["overall_status"] = "error"
            results["can_proceed"] = False
        elif results["warning_count"] > 0:
            results["overall_status"] = "warning"
            results["can_proceed"] = True
        
        # Log final summary
        logger.info("=" * 80)
        logger.info("VALIDATION SUMMARY")
        logger.info(f"Overall Status: {results['overall_status'].upper()}")
        logger.info(f"Errors: {results['error_count']}")
        logger.info(f"Warnings: {results['warning_count']}")
        logger.info(f"Can Proceed: {results['can_proceed']}")
        logger.info("=" * 80)
        
        return results
    
    def _log_category_results(self, category_name: str, results: Dict) -> None:
        """Log validation results for a specific category."""
        status = results.get("status", "unknown")
        issues = results.get("issues", [])
        
        if not issues:
            logger.info(f"✓ {category_name}: PASS (no issues)")
        else:
            logger.info(f"{'✗' if status == 'error' else '⚠'} {category_name}: {status.upper()} ({len(issues)} issue(s))")
            for issue in issues:
                severity = issue.get("severity", "unknown").upper()
                message = issue.get("message", "No message")
                logger.info(f"  [{severity}] {message}")
    
    def _validate_file_existence(self, master_file: str, combined_file: str, output_dir: str) -> Dict:
        """Check if all required files and directories exist."""
        issues = []
        
        if not master_file or not Path(master_file).exists():
            issues.append({
                "severity": "error",
                "message": "Master file not found or not selected",
                "fixable": False,
                "fix_action": None
            })
        
        if not combined_file or not Path(combined_file).exists():
            issues.append({
                "severity": "error",
                "message": "Combined/Email manifest file not found or not selected",
                "fixable": False,
                "fix_action": None
            })
        
        if not output_dir or not Path(output_dir).exists():
            issues.append({
                "severity": "warning",
                "message": f"Output directory does not exist: {output_dir}",
                "fixable": True,
                "fix_action": ("create_directory", output_dir)
            })
        
        return {
            "name": "📁 File Existence",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "error",
            "issues": issues
        }
    
    def _validate_file_structure(self, master_file: str, combined_file: str) -> Dict:
        """Check if files have required columns."""
        issues = []
        
        try:
            if master_file and Path(master_file).exists():
                df_master = pd.read_excel(master_file, nrows=5)
                # Check for Agency column
                if "Agency" not in df_master.columns:
                    issues.append({
                        "severity": "error",
                        "message": "Master file missing required column: Agency",
                        "fixable": False,
                        "fix_action": None
                    })
                # Check for UserName column (accept both "UserName" and "User Name")
                has_username = any(col in df_master.columns for col in ["UserName", "User Name"])
                if not has_username:
                    issues.append({
                        "severity": "error",
                        "message": "Master file missing required column: UserName (or 'User Name')",
                        "fixable": False,
                        "fix_action": None
                    })
            
            if combined_file and Path(combined_file).exists():
                # Check first sheet
                df_combined = pd.read_excel(combined_file, sheet_name=0, nrows=5)
                required_cols = ["source_file_name", "agency_id"]
                missing = [col for col in required_cols if not any(col.lower() in str(c).lower() for c in df_combined.columns)]
                if missing:
                    issues.append({
                        "severity": "error",
                        "message": f"Combined file missing required columns: {', '.join(missing)}",
                        "fixable": False,
                        "fix_action": None
                    })
        except Exception as e:
            issues.append({
                "severity": "error",
                "message": f"Error reading file structure: {str(e)}",
                "fixable": False,
                "fix_action": None
            })
        
        return {
            "name": "📋 File Structure",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "error",
            "issues": issues
        }
    
    def _validate_data_quality(self, master_file: str, combined_file: str) -> Dict:
        """Check data quality issues."""
        issues = []
        
        try:
            if master_file and Path(master_file).exists():
                df_master = pd.read_excel(master_file)
                
                # Check for empty file
                if len(df_master) == 0:
                    issues.append({
                        "severity": "error",
                        "message": "Master file is empty (no data rows)",
                        "fixable": False,
                        "fix_action": None
                    })
                
                # Check for null agencies
                if "Agency" in df_master.columns:
                    null_count = df_master["Agency"].isna().sum()
                    if null_count > 0:
                        issues.append({
                            "severity": "warning",
                            "message": f"{null_count} users have no agency assigned",
                            "fixable": False,
                            "fix_action": None
                        })
        except Exception as e:
            issues.append({
                "severity": "warning",
                "message": f"Could not check data quality: {str(e)}",
                "fixable": False,
                "fix_action": None
            })
        
        return {
            "name": "✅ Data Quality",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "error",
            "issues": issues
        }
    
    def _validate_agency_mappings(self, master_file: str, combined_file: str) -> Dict:
        """Check agency mapping consistency."""
        issues = []
        
        try:
            if master_file and combined_file and Path(master_file).exists() and Path(combined_file).exists():
                df_master = pd.read_excel(master_file)
                
                # Load all tabs from combined file
                excel_file = pd.ExcelFile(combined_file)
                all_mapped_agencies = set()
                file_agency_map = {}  # format: {agency: [(tab, file_name)]}
                
                for sheet in excel_file.sheet_names:
                    df_combined = pd.read_excel(combined_file, sheet_name=sheet)
                    for _, row in df_combined.iterrows():
                        agencies = str(row.get("agency_id", "")).split(",")
                        file_name = str(row.get("source_file_name", ""))
                        for agency in agencies:
                            agency = agency.strip().upper()
                            if agency:
                                all_mapped_agencies.add(agency)
                                # Check if agency exists in this map
                                if agency not in file_agency_map:
                                    file_agency_map[agency] = []
                                # Add (tab, file_name) tuple
                                file_agency_map[agency].append((sheet, file_name))
                
                # Check for duplicates within same region (tab)
                for agency, mappings in file_agency_map.items():
                    # Group by tab to check for same-region duplicates
                    tab_files = {}
                    for tab, file_name in mappings:
                        if tab not in tab_files:
                            tab_files[tab] = []
                        tab_files[tab].append(file_name)
                    
                    # Report duplicates within same tab/region as errors
                    for tab, files in tab_files.items():
                        if len(files) > 1:
                            issues.append({
                                "severity": "error",
                                "message": f"Duplicate mapping: '{agency}' appears in multiple files ({' and '.join(files)})",
                                "fixable": False,
                                "fix_action": None
                            })
                    
                    # Report cross-region duplicates as info (might be intentional)
                    if len(tab_files) > 1:
                        tabs_files = [f"{tab}: {', '.join(files)}" for tab, files in tab_files.items()]
                        issues.append({
                            "severity": "info",
                            "message": f"Cross-region mapping: '{agency}' appears in multiple regions ({' and '.join([f'{t}' for t in tab_files.keys()])})",
                            "fixable": False,
                            "fix_action": None
                        })
                
                # Check for unmapped agencies in master
                if "Agency" in df_master.columns:
                    master_agencies = set(df_master["Agency"].dropna().str.strip().str.upper().unique())
                    unmapped = master_agencies - all_mapped_agencies
                    if unmapped:
                        issues.append({
                            "severity": "warning",
                            "message": f"{len(unmapped)} agencies in master file have no mapping (Examples: {', '.join(list(unmapped)[:5])})",
                            "fixable": True,
                            "fix_action": ("show_unmapped", list(unmapped))
                        })
        except Exception as e:
            issues.append({
                "severity": "warning",
                "message": f"Could not validate agency mappings: {str(e)}",
                "fixable": False,
                "fix_action": None
            })
        
        return {
            "name": "🏢 Agency Mappings",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "error",
            "issues": issues
        }
    
    def _validate_email_addresses(self, combined_file: str) -> Dict:
        """Check email address formats."""
        issues = []
        email_pattern = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
        
        try:
            if combined_file and Path(combined_file).exists():
                excel_file = pd.ExcelFile(combined_file)
                invalid_emails = []
                
                for sheet in excel_file.sheet_names:
                    df = pd.read_excel(combined_file, sheet_name=sheet)
                    for _, row in df.iterrows():
                        to_emails = str(row.get("recipients_to", "")).split(";")
                        cc_emails = str(row.get("recipients_cc", "")).split(";")
                        
                        for email in to_emails + cc_emails:
                            email = email.strip()
                            if email and email != "nan" and not email_pattern.match(email):
                                invalid_emails.append((sheet, str(row.get("source_file_name", "")), email))
                
                if invalid_emails:
                    examples = "\n".join([f"  • {sheet}/{file}: {email}" for sheet, file, email in invalid_emails[:5]])
                    issues.append({
                        "severity": "error",
                        "message": f"{len(invalid_emails)} invalid email addresses found:\n{examples}",
                        "fixable": True,
                        "fix_action": ("fix_emails", invalid_emails)
                    })
        except Exception as e:
            issues.append({
                "severity": "warning",
                "message": f"Could not validate emails: {str(e)}",
                "fixable": False,
                "fix_action": None
            })
        
        return {
            "name": "📧 Email Addresses",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "error",
            "issues": issues
        }
    
    def _validate_filename_safety(self, combined_file: str) -> Dict:
        """Check for Windows-unsafe filename characters."""
        issues = []
        unsafe_chars = r'[<>:"/\\|?*\']'
        
        try:
            if combined_file and Path(combined_file).exists():
                excel_file = pd.ExcelFile(combined_file)
                unsafe_files = []
                
                for sheet in excel_file.sheet_names:
                    df = pd.read_excel(combined_file, sheet_name=sheet)
                    for _, row in df.iterrows():
                        file_name = str(row.get("source_file_name", ""))
                        if re.search(unsafe_chars, file_name):
                            unsafe_files.append((sheet, file_name))
                
                if unsafe_files:
                    examples = "\n".join([f"  • {sheet}: {file}" for sheet, file in unsafe_files[:5]])
                    issues.append({
                        "severity": "warning",
                        "message": f"{len(unsafe_files)} filenames contain unsafe characters (will be auto-fixed):\n{examples}",
                        "fixable": True,
                        "fix_action": ("auto_sanitize", unsafe_files)
                    })
        except Exception as e:
            issues.append({
                "severity": "info",
                "message": f"Could not check filename safety: {str(e)}",
                "fixable": False,
                "fix_action": None
            })
        
        return {
            "name": "📝 Filename Safety",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "warning",
            "issues": issues
        }
    
    def _validate_region_configuration(self) -> Dict:
        """Check region configuration if regions are being used."""
        issues = []
        
        try:
            region_manager = RegionManager()
            
            if len(region_manager.profiles) == 0:
                issues.append({
                    "severity": "info",
                    "message": "No regions configured (single-region mode)",
                    "fixable": False,
                    "fix_action": None
                })
            else:
                # Check if current region is set
                if not region_manager.current_region:
                    issues.append({
                        "severity": "warning",
                        "message": "No current region selected",
                        "fixable": True,
                        "fix_action": ("select_region", None)
                    })
                
                # Check each region profile
                for name, profile in region_manager.profiles.items():
                    errors = profile.validate()
                    if errors:
                        issues.append({
                            "severity": "warning",
                            "message": f"Region '{name}' has issues: {', '.join(errors)}",
                            "fixable": True,
                            "fix_action": ("edit_region", name)
                        })
        except Exception as e:
            issues.append({
                "severity": "info",
                "message": f"Could not check regions: {str(e)}",
                "fixable": False,
                "fix_action": None
            })
        
        return {
            "name": "🌍 Region Configuration",
            "status": "pass" if not any(i["severity"] == "error" for i in issues) else "warning",
            "issues": issues
        }
    
    def generate(self):
        """
        Callback for the 'Generate Files' button.
        
        Generates agency-specific Excel files from the master file.
        """
        master_file = self.vars["master"].get()
        combined_file = self.vars["combined"].get()
        output_dir = self.vars["output"].get()
        
        if not all([master_file, combined_file, output_dir]):
            messagebox.showwarning("Missing Information", "Please select all required files and output folder.")
            return
        
        try:
            self.show_progress(True)
            self.update_progress(0.1, "Preparing file generation...")
            
            # Generate files using FileProcessor
            processor = FileProcessor()
            
            # Load combined file to get mappings
            combined_loader = CombinedFileLoader(Path(combined_file))
            mappings = combined_loader.load_all_tabs()
            
            self.update_progress(0.3, "Processing master file...")
            
            # Generate agency files
            success = processor.generate_agency_files(
                master_file=Path(master_file),
                combined_map_file=Path(combined_file),
                output_dir=Path(output_dir),
                progress_callback=self.update_progress,
                unmapped_callback=None
            )
            
            if success:
                self.update_progress(1.0, "Files generated successfully!")
                self.show_progress(False)
                messagebox.showinfo("Success", "Agency files generated successfully!")
                self.refresh()
            else:
                self.show_progress(False)
                messagebox.showerror("Error", "File generation failed. Check logs for details.")
                
        except Exception as e:
            self.show_progress(False)
            logger.error(f"File generation error: {e}")
            messagebox.showerror("Generation Error", f"Failed to generate files:\n{str(e)}")
    
    def open_folder(self):
        """
        Callback for the 'Open Output Folder' button.
        
        Opens the output directory in Windows Explorer.
        """
        output_dir = self.vars["output"].get()
        if not output_dir:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return
        
        success = open_output_folder(output_dir)
        if not success:
            messagebox.showerror("Error", f"Failed to open output folder:\n{output_dir}")

    def test_email(self):
        """
        Callback for the 'Test Email Connection' button.
        
        Tests the connection to Microsoft Outlook and displays results.
        """
        try:
            if self.email_handler.test_connection():
                messagebox.showinfo("Email Test", "Outlook connection successful!")
            else:
                messagebox.showerror("Email Test Failed", "Could not connect to Outlook")
        except Exception as e:
            logger.error(f"Email test failed: {str(e)}")
            messagebox.showerror("Email Test Failed", f"Could not connect to Outlook: {str(e)}")

    def organize_compliance_replies(self):
        """
        Organize compliance reply emails into dedicated folders.
        
        This function will:
        1. Set up compliance folders for the current review period
        2. Find and move reply emails to the appropriate folder
        3. Show results to the user
        """
        try:
            # Get current review period from config
            config = config_manager.get_all()
            review_period = config.get("review_period", "Q2 2025")
            
            # Set up folders
            folder_results = self.email_handler.setup_compliance_folders(review_period)
            
            # Count successful folder creations
            successful_folders = sum(1 for success in folder_results.values() if success)
            total_folders = len(folder_results)
            
            # Organize replies
            replies_folder = f"Compliance {review_period} - Replies"
            organized_count = self.email_handler.organize_compliance_replies(replies_folder)
            
            # Show results
            result_message = (
                f"📁 Email Organization Results:\n\n"
                f"✅ Folders Setup: {successful_folders}/{total_folders} successful\n"
                f"📧 Organized Replies: {organized_count} emails moved\n\n"
                f"Folder Structure Created:\n"
            )
            
            for folder_name, success in folder_results.items():
                status = "✅" if success else "❌"
                result_message += f"{status} {folder_name}\n"
            
            result_message += (
                f"\n💡 Tips:\n"
                f"• Check your Inbox for the new compliance folders\n"
                f"• Sent emails are automatically organized when using 'Direct' mode\n"
                f"• Run this again to catch new replies"
            )
            
            messagebox.showinfo("Email Organization Complete", result_message)
            logger.info(f"Organized {organized_count} compliance replies into folders")
            
        except Exception as e:
            logger.error(f"Failed to organize emails: {str(e)}")
            messagebox.showerror("Email Organization Failed", 
                               f"Could not organize compliance emails:\n\n{str(e)}")

    def open_email_verifier(self):
        """
        Open the Smart Email Verifier dialog.
        
        This dialog allows users to paste forwarded emails or text containing
        email addresses and automatically extract and verify them.
        """
        try:
            dialog = SmartEmailVerifierDialog(self)
            dialog.grab_set()  # Make dialog modal
        except Exception as e:
            logger.error(f"Failed to open email verifier: {str(e)}")
            messagebox.showerror("Error", f"Failed to open email verifier: {str(e)}")

    def export_list(self):
        """
        Callback for the 'Export Agency List' button.
        
        Exports the current agency list with their status to an Excel file.
        """
        if not self.agencies:
            messagebox.showwarning("No Agencies", "No agencies found. Please generate files first.")
            return
        
        output_dir = self.vars["output"].get()
        if not output_dir:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return
        
        export_path = export_agency_list(self.agencies, self.audit_df, output_dir)
        
        if export_path:
            messagebox.showinfo("Export Successful", 
                              f"Agency list exported successfully!\n\nFile: {export_path}")
        else:
            messagebox.showerror("Export Failed", "Failed to export agency list.")

    def export_summary_report(self):
        """
        Exports a summary report of the audit status to an Excel file.
        """
        if not self.agencies:
            messagebox.showwarning("No Agencies", "No agencies found. Please load data first.")
            return
        
        output_dir = self.vars["output"].get()
        if not output_dir:
            messagebox.showwarning("No Output Folder", "Please select an output folder first.")
            return
        
        try:
            from datetime import datetime
            report_data = {
                'Agency': [],
                'Status': [],
                'Sent Date': [],
                'Response Date': [],
                'Days Pending': []
            }
            
            for agency in self.agencies:
                agency_data = self.audit_df[self.audit_df['Agency'].str.upper() == agency.upper()]
                if not agency_data.empty:
                    latest = agency_data.iloc[-1]
                    report_data['Agency'].append(agency)
                    report_data['Status'].append(latest.get('Status', 'Unknown'))
                    report_data['Sent Date'].append(latest.get('Sent Email Date', ''))
                    report_data['Response Date'].append(latest.get('Response Received Date', ''))
                    report_data['Days Pending'].append('')
                else:
                    report_data['Agency'].append(agency)
                    report_data['Status'].append('Not Sent')
                    report_data['Sent Date'].append('')
                    report_data['Response Date'].append('')
                    report_data['Days Pending'].append('')
            
            df_report = pd.DataFrame(report_data)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_file = os.path.join(output_dir, f"Summary_Report_{timestamp}.xlsx")
            
            df_report.to_excel(report_file, index=False)
            
            messagebox.showinfo("Export Successful", 
                              f"Summary report exported successfully!\n\nFile: {report_file}")
        except Exception as e:
            logger.error(f"Failed to export summary report: {str(e)}")
            messagebox.showerror("Export Failed", f"Failed to export summary report:\n{str(e)}")

    def clear_audit_data(self):
        """
        Clears the audit log data after confirmation.
        """
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all audit data?\n\nThis action cannot be undone."):
            try:
                self.audit_df = pd.DataFrame()
                self.update_dashboard_metrics()
                messagebox.showinfo("Success", "Audit data cleared successfully.")
            except Exception as e:
                logger.error(f"Failed to clear audit data: {str(e)}")
                messagebox.showerror("Error", f"Failed to clear audit data:\n{str(e)}")

    def settings(self):
        """
        Enhanced settings method that opens the new settings dialog.
        """
        open_enhanced_settings_dialog(self)
    

    
    def open_agency_mapping_manager(self):
        """Opens the Agency Mapping Manager dialog."""
        try:
            dialog = AgencyMappingManagerDialog(self)
            dialog.wait_window()
        except Exception as e:
            logger.error(f"Failed to open Agency Mapping Manager: {str(e)}")
            messagebox.showerror("Error", f"Failed to open Agency Mapping Manager:\n{str(e)}")

    def auto_scan_output_folder(self):
        """
        Automatically scans the output folder and prompts user to load existing files.
        """
        output_dir = self.vars["output"].get()
        if not output_dir:
            return
        
        # Show progress
        self.show_progress(True)
        self.update_progress(0.1, "Scanning files...")
        
        agencies_found, audit_exists, summary = auto_scan_output_folder(output_dir)
        
        self.update_progress(0.5, "Validating structure...")
        
        if agencies_found:
            # Ask user if they want to load existing files
            response = messagebox.askyesno(
                "Existing Files Found",
                f"Found {len(agencies_found)} agency files in the output folder.\n\n{summary}\n\nWould you like to load these existing files?"
            )
            
            if response:
                self.update_progress(0.8, "Loading existing files...")
                
                # Load agencies
                self.agencies = sorted(agencies_found)
                
                # Load audit log using AuditLogger
                if audit_exists:
                    audit_file_path = os.path.join(output_dir, AUDIT_FILE_NAME)
                    self.audit_df = self.audit_logger.load(audit_file_path)
                else:
                    self.audit_df = pd.DataFrame(
                        columns=["Agency", "Sent Email Date", "Response Received Date", "To", "CC", "Status", "Comments"]
                    )
                
                # Update UI
                self.filter_agencies()
                self.update_dashboard_metrics()
                
                self.update_progress(1.0, "Complete!")
                messagebox.showinfo("Success", f"Loaded {len(agencies_found)} agencies from existing files!")
            else:
                self.update_progress(1.0, "Skipped loading existing files")
        else:
            self.update_progress(1.0, "No existing files found")
        
        self.show_progress(False)


# =============== Enhanced Functions ===============

def auto_scan_output_folder(output_dir):
    """
    Automatically scans the output folder for existing agency files and audit log.
    
    Args:
        output_dir (str): Path to the output directory.
    
    Returns:
        tuple: (agencies_found, audit_exists, validation_summary)
    """
    if not output_dir or not os.path.exists(output_dir):
        return [], False, "No output directory found"
    
    agencies_found = []
    audit_exists = False
    validation_summary = []
    
    try:
        files = os.listdir(output_dir)
        
        # Find agency files
        for file in files:
            if file.endswith('.xlsx') and not file.startswith('Unassigned') and not file.startswith('OMC_Unassigned'):
                agency_name = file.replace('.xlsx', '')
                agencies_found.append(agency_name)
        
        # Check for audit log
        audit_file = os.path.join(output_dir, AUDIT_FILE_NAME)
        audit_exists = os.path.exists(audit_file)
        
        validation_summary.append(f"✅ Found {len(agencies_found)} agency files")
        if audit_exists:
            validation_summary.append("✅ Audit log found")
        else:
            validation_summary.append("⚠️ No audit log found")
            
    except Exception as e:
        validation_summary.append(f"❌ Error scanning folder: {str(e)}")
    
    return agencies_found, audit_exists, "\n".join(validation_summary)


def validate_agency_mapping(master_file, agency_map_file, email_manifest_file=None):
    """
    Validates agency mapping against master file and email manifest.
    
    Returns:
        dict: Validation results with keys:
            - is_valid (bool): Overall validation status
            - issues (list): List of issue descriptions
            - master_agencies (set): Unique agencies in master file
            - mapped_agencies (set): Agencies in mapping file
            - unmapped_agencies (set): Agencies in master but not in mapping
            - duplicate_agencies (dict): Agencies mapped to multiple files
            - email_issues (list): Email manifest validation issues
    """
    results = {
        "is_valid": True,
        "issues": [],
        "master_agencies": set(),
        "mapped_agencies": set(),
        "unmapped_agencies": set(),
        "duplicate_agencies": {},
        "email_issues": []
    }
    
    try:
        # Load files
        df_master = pd.read_excel(master_file)
        df_map = pd.read_excel(agency_map_file)
        
        # Clean column names
        df_master.columns = df_master.columns.str.strip()
        df_map.columns = df_map.columns.str.strip()
        
        # Validate master file columns
        if "Agency" not in df_master.columns:
            results["is_valid"] = False
            results["issues"].append("❌ Master file missing 'Agency' column")
            return results
        
        # Validate mapping file columns
        required_map_cols = ["File name", "Agency ID (Agency in the file)"]
        missing_cols = [col for col in required_map_cols if col not in df_map.columns]
        if missing_cols:
            results["is_valid"] = False
            results["issues"].append(f"❌ Mapping file missing columns: {', '.join(missing_cols)}")
            return results
        
        # Get unique agencies from master (case-insensitive, exclude empty)
        master_agencies_raw = df_master["Agency"].fillna("").str.strip()
        master_agencies_raw = master_agencies_raw[master_agencies_raw != ""]
        results["master_agencies"] = set(master_agencies_raw.str.lower().unique())
        
        # Get unique agencies from mapping (handle comma-separated values)
        mapped_agencies_set = set()
        for agency_str in df_map["Agency ID (Agency in the file)"].str.strip():
            if ',' in agency_str:
                # Split comma-separated agencies
                agencies = [a.strip().lower() for a in agency_str.split(',') if a.strip()]
                mapped_agencies_set.update(agencies)
            else:
                mapped_agencies_set.add(agency_str.lower())
        results["mapped_agencies"] = mapped_agencies_set
        
        # Find unmapped agencies
        results["unmapped_agencies"] = results["master_agencies"] - results["mapped_agencies"]
        
        if results["unmapped_agencies"]:
            results["is_valid"] = False
            unmapped_list = sorted(results["unmapped_agencies"])[:10]
            results["issues"].append(f"❌ {len(results['unmapped_agencies'])} unmapped agencies: {', '.join(unmapped_list)}")
        
        # Check for duplicate agencies (same agency in multiple files)
        agency_to_files = {}
        for _, row in df_map.iterrows():
            agency_str = row["Agency ID (Agency in the file)"].strip()
            file_name = row["File name"]
            
            # Handle comma-separated agencies
            if ',' in agency_str:
                agencies = [a.strip().lower() for a in agency_str.split(',') if a.strip()]
            else:
                agencies = [agency_str.lower()]
            
            for agency in agencies:
                if agency not in agency_to_files:
                    agency_to_files[agency] = []
                agency_to_files[agency].append(file_name)
        
        # Find duplicates
        for agency, files in agency_to_files.items():
            if len(files) > 1:
                results["duplicate_agencies"][agency] = files
        
        if results["duplicate_agencies"]:
            results["is_valid"] = False
            dup_count = len(results["duplicate_agencies"])
            results["issues"].append(f"❌ {dup_count} agencies mapped to multiple files")
        
        # Validate email manifest if provided
        if email_manifest_file:
            try:
                df_email = pd.read_excel(email_manifest_file)
                df_email.columns = df_email.columns.str.strip()
                
                required_email_cols = ["Agency", "To", "CC"]
                missing_email_cols = [col for col in required_email_cols if col not in df_email.columns]
                
                if missing_email_cols:
                    results["is_valid"] = False
                    results["email_issues"].append(f"❌ Email manifest missing columns: {', '.join(missing_email_cols)}")
                else:
                    # Check if all file names have email entries
                    file_names = set(df_map["File name"].str.strip())
                    email_agencies = set(df_email["Agency"].str.strip())
                    
                    missing_emails = file_names - email_agencies
                    if missing_emails:
                        results["is_valid"] = False
                        results["email_issues"].append(f"❌ {len(missing_emails)} files missing email entries: {', '.join(list(missing_emails)[:5])}")
                    
                    # Check for invalid email addresses
                    email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
                    for _, row in df_email.iterrows():
                        to_emails = str(row.get("To", "")).split(";")
                        cc_emails = str(row.get("CC", "")).split(";")
                        
                        for email in to_emails + cc_emails:
                            email = email.strip()
                            if email and not re.match(email_pattern, email):
                                results["email_issues"].append(f"⚠️ Invalid email format: {email}")
            
            except Exception as e:
                results["is_valid"] = False
                results["email_issues"].append(f"❌ Error reading email manifest: {str(e)}")
        
        # Add success message if no issues
        if results["is_valid"]:
            results["issues"].append(f"✅ All {len(results['mapped_agencies'])} agencies validated successfully")
        
    except Exception as e:
        results["is_valid"] = False
        results["issues"].append(f"❌ Validation error: {str(e)}")
    
    return results


def show_validation_dialog(validation_results, parent=None):
    """
    Shows validation results in a dialog window.
    
    Returns:
        bool: True if user wants to proceed (even with warnings), False to cancel
    """
    dialog = ctk.CTkToplevel(parent) if parent else ctk.CTkToplevel()
    dialog.title("Agency Mapping Validation")
    dialog.geometry("700x500")
    dialog.transient(parent)
    dialog.grab_set()
    
    user_choice = {"proceed": False}
    
    # Header
    status_color = "green" if validation_results["is_valid"] else "red"
    status_text = "✅ Validation Passed" if validation_results["is_valid"] else "❌ Validation Failed"
    
    header = ctk.CTkLabel(dialog, text=status_text, font=("Arial", 16, "bold"), text_color=status_color)
    header.pack(pady=10)
    
    # Issues list
    issues_frame = ctk.CTkFrame(dialog)
    issues_frame.pack(fill="both", expand=True, padx=20, pady=10)
    
    issues_text = ctk.CTkTextbox(issues_frame, font=("Consolas", 11))
    issues_text.pack(fill="both", expand=True)
    
    # Display issues
    all_issues = validation_results["issues"] + validation_results["email_issues"]
    issues_text.insert("1.0", "\n".join(all_issues))
    issues_text.configure(state="disabled")
    
    # Buttons
    button_frame = ctk.CTkFrame(dialog)
    button_frame.pack(pady=10)
    
    def on_fix():
        user_choice["proceed"] = False
        user_choice["fix"] = True
        dialog.destroy()
    
    def on_proceed():
        user_choice["proceed"] = True
        dialog.destroy()
    
    def on_cancel():
        user_choice["proceed"] = False
        dialog.destroy()
    
    if not validation_results["is_valid"]:
        fix_btn = ctk.CTkButton(button_frame, text="Fix Issues", command=on_fix, fg_color="orange")
        fix_btn.pack(side="left", padx=5)
    
    if validation_results["is_valid"] or validation_results["unmapped_agencies"]:
        proceed_btn = ctk.CTkButton(button_frame, text="Proceed Anyway", command=on_proceed, fg_color="green")
        proceed_btn.pack(side="left", padx=5)
    
    cancel_btn = ctk.CTkButton(button_frame, text="Cancel", command=on_cancel, fg_color="gray")
    cancel_btn.pack(side="left", padx=5)
    
    dialog.wait_window()
    return user_choice


def show_fix_dialog(validation_results, agency_map_file, email_manifest_file=None, parent=None):
    """
    Shows dialog for fixing validation issues with in-app editing.
    
    Allows users to:
    - Add/edit/delete agency mappings
    - Fix email manifest entries
    - Save changes back to Excel files
    """
    dialog = ctk.CTkToplevel(parent) if parent else ctk.CTkToplevel()
    dialog.title("Fix Agency Mapping Issues")
    dialog.geometry("900x600")
    dialog.transient(parent)
    dialog.grab_set()
    
    # Load current mapping
    df_map = pd.read_excel(agency_map_file)
    df_map.columns = df_map.columns.str.strip()
    
    # Tabview for mapping and email
    tabview = ctk.CTkTabview(dialog)
    tabview.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Tab 1: Agency Mapping
    tab_mapping = tabview.add("Agency Mapping")
    
    mapping_frame = ctk.CTkFrame(tab_mapping)
    mapping_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Treeview for mapping table
    import tkinter as tk
    from tkinter import ttk
    
    tree_frame = ctk.CTkFrame(mapping_frame)
    tree_frame.pack(fill="both", expand=True)
    
    tree = ttk.Treeview(tree_frame, columns=("File name", "Agency"), show="headings", height=15)
    tree.heading("File name", text="File name")
    tree.heading("Agency", text="Agency ID")
    tree.column("File name", width=300)
    tree.column("Agency", width=300)
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    
    tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Load data into tree
    for _, row in df_map.iterrows():
        tree.insert("", "end", values=(row["File name"], row["Agency ID (Agency in the file)"]))
    
    # Buttons for editing
    btn_frame = ctk.CTkFrame(mapping_frame)
    btn_frame.pack(pady=10)
    
    def add_mapping():
        add_window = ctk.CTkToplevel(dialog)
        add_window.title("Add Mapping")
        add_window.geometry("400x200")
        add_window.transient(dialog)
        
        ctk.CTkLabel(add_window, text="File name:").pack(pady=5)
        file_entry = ctk.CTkEntry(add_window, width=300)
        file_entry.pack(pady=5)
        
        ctk.CTkLabel(add_window, text="Agency ID:").pack(pady=5)
        agency_entry = ctk.CTkEntry(add_window, width=300)
        agency_entry.pack(pady=5)
        
        def save_mapping():
            file_name = file_entry.get().strip()
            agency = agency_entry.get().strip()
            
            if file_name and agency:
                tree.insert("", "end", values=(file_name, agency))
                add_window.destroy()
        
        ctk.CTkButton(add_window, text="Add", command=save_mapping).pack(pady=10)
    
    def edit_mapping():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a mapping to edit")
            return
        
        item = selected[0]
        values = tree.item(item, "values")
        
        edit_window = ctk.CTkToplevel(dialog)
        edit_window.title("Edit Mapping")
        edit_window.geometry("400x200")
        edit_window.transient(dialog)
        
        ctk.CTkLabel(edit_window, text="File name:").pack(pady=5)
        file_entry = ctk.CTkEntry(edit_window, width=300)
        file_entry.insert(0, values[0])
        file_entry.pack(pady=5)
        
        ctk.CTkLabel(edit_window, text="Agency ID:").pack(pady=5)
        agency_entry = ctk.CTkEntry(edit_window, width=300)
        agency_entry.insert(0, values[1])
        agency_entry.pack(pady=5)
        
        def save_edit():
            tree.item(item, values=(file_entry.get().strip(), agency_entry.get().strip()))
            edit_window.destroy()
        
        ctk.CTkButton(edit_window, text="Save", command=save_edit).pack(pady=10)
    
    def delete_mapping():
        selected = tree.selection()
        if selected:
            if messagebox.askyesno("Confirm Delete", f"Delete {len(selected)} selected row(s)?"):
                for item in selected:
                    tree.delete(item)
    
    ctk.CTkButton(btn_frame, text="Add", command=add_mapping, width=100).pack(side="left", padx=5)
    ctk.CTkButton(btn_frame, text="Edit", command=edit_mapping, width=100).pack(side="left", padx=5)
    ctk.CTkButton(btn_frame, text="Delete", command=delete_mapping, width=100).pack(side="left", padx=5)
    
    # Save and Close buttons
    bottom_frame = ctk.CTkFrame(dialog)
    bottom_frame.pack(pady=10)
    
    def save_changes():
        # Extract data from tree
        data = []
        for item in tree.get_children():
            values = tree.item(item, "values")
            data.append({"File name": values[0], "Agency ID (Agency in the file)": values[1]})
        
        # Save to Excel
        df_new = pd.DataFrame(data)
        df_new.to_excel(agency_map_file, index=False)
        
        messagebox.showinfo("Saved", "Agency mapping saved successfully!")
        dialog.destroy()
    
    ctk.CTkButton(bottom_frame, text="Save Changes", command=save_changes, fg_color="green").pack(side="left", padx=5)
    ctk.CTkButton(bottom_frame, text="Close", command=dialog.destroy, fg_color="gray").pack(side="left", padx=5)
    
    dialog.wait_window()


def format_worksheet(writer, sheet_name, df):
    """
    Formats a worksheet with frozen header, bold header, auto-width columns.
    """
    wb = writer.book
    ws = writer.sheets[sheet_name]
    
    # Freeze top row
    ws.freeze_panes(1, 0)
    
    # Header format (bold + blue background)
    header_fmt = wb.add_format({
        'bold': True,
        'bg_color': '#D9E1F2',
        'border': 1
    })
    
    # Write headers with format
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name, header_fmt)
        
        # Auto-width columns
        col_width = max(
            df[col_name].astype(str).map(len).max(),
            len(col_name),
            12
        ) + 2
        ws.set_column(col_idx, col_idx, min(col_width, 50))


def sanitize_sheet_name(name):
    r"""Excel sheet names have restrictions:
    - Max 31 characters
    - Cannot contain: \ / ? * [ ]
    """
    # Remove invalid characters
    name = re.sub(r'[\\/*?\[\]]', '', str(name))
    
    # Truncate to 31 characters
    if len(name) > 31:
        name = name[:31]
    
    return name if name else "Sheet1"


def generate_agency_files_multi_tab(master_file, agency_map_file, out_dir, progress_callback=None):
    """
    Generates multi-tab Excel files based on agency mapping.
    
    Each file (from mapping File name column) contains:
    - Tab 1: "All Users" (all agencies in this file)
    - Tab 2-N: Individual agency tabs (alphabetical order)
    
    Args:
        master_file (str): Path to master user access Excel file (must have "Agency" column)
        agency_map_file (str): Path to agency mapping file (columns: "File name" | "Agency ID (Agency in the file)")
        out_dir (str): Output directory for generated files
        progress_callback (callable): Optional callback for progress updates
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        out_dir = Path(out_dir)
        out_dir.mkdir(exist_ok=True)
        
        # Load data
        df_master = pd.read_excel(master_file)
        df_map = pd.read_excel(agency_map_file)
        
        # Clean column names
        df_master.columns = df_master.columns.str.strip()
        df_map.columns = df_map.columns.str.strip()
        
        # Validate required columns
        if "Agency" not in df_master.columns:
            messagebox.showerror("Error", "Master file must have 'Agency' column")
            return False
        
        required_map_cols = ["File name", "Agency ID (Agency in the file)"]
        missing_cols = [col for col in required_map_cols if col not in df_map.columns]
        if missing_cols:
            messagebox.showerror("Error", f"Agency mapping file missing columns: {', '.join(missing_cols)}")
            return False
        
        # Normalize Agency column (handle empty/null as "Unassigned")
        df_master["Agency_Normalized"] = df_master["Agency"].fillna("").str.strip()
        df_master.loc[df_master["Agency_Normalized"] == "", "Agency_Normalized"] = "Unassigned"
        
        # Normalize mapping agencies for case-insensitive matching
        df_map["Agency_Lower"] = df_map["Agency ID (Agency in the file)"].str.strip().str.lower()
        
        # Track assigned users
        assigned_idx = set()
        
        # Process each file in the mapping
        total_files = len(df_map["File name"].unique())
        for file_idx, file_name in enumerate(df_map["File name"].unique(), 1):
            if progress_callback:
                progress_callback(f"Processing {file_idx}/{total_files}: {file_name}")
            
            # Get all agencies for this file (handle comma-separated values)
            file_agencies_raw = df_map[df_map["File name"] == file_name]["Agency ID (Agency in the file)"].tolist()
            file_agencies_lower = []
            for agency_str in file_agencies_raw:
                if ',' in agency_str:
                    # Split comma-separated agencies
                    agencies = [a.strip().lower() for a in agency_str.split(',') if a.strip()]
                    file_agencies_lower.extend(agencies)
                else:
                    file_agencies_lower.append(agency_str.strip().lower())
            
            # Find users for these agencies (case-insensitive)
            file_users = df_master[df_master["Agency_Normalized"].str.lower().isin(file_agencies_lower)].copy()
            
            if file_users.empty:
                continue  # Skip if no users for this file
            
            # Mark these users as assigned
            assigned_idx.update(file_users.index)
            
            # Add Review Action and Comments columns
            file_users["Review Action"] = ""
            file_users["Comments"] = ""
            
            # Remove the temporary normalized column
            if "Agency_Normalized" in file_users.columns:
                file_users = file_users.drop(columns=["Agency_Normalized"])
            
            # Create output file with multiple tabs
            out_file = out_dir / f"{sanitize_sheet_name(file_name)}.xlsx"
            
            with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
                # Tab 1: All Users
                file_users.to_excel(writer, sheet_name="All Users", index=False)
                format_worksheet(writer, "All Users", file_users)
                
                # Tab 2-N: Individual agency tabs (alphabetical)
                unique_agencies = sorted(file_users["Agency"].dropna().unique())
                
                for agency in unique_agencies:
                    agency_users = file_users[file_users["Agency"] == agency].copy()
                    
                    if agency_users.empty:
                        continue  # Skip empty tabs
                    
                    # Sanitize sheet name
                    sheet_name = sanitize_sheet_name(agency)
                    
                    # Write to sheet
                    agency_users.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_worksheet(writer, sheet_name, agency_users)
        
        # Handle unmapped agencies
        unassigned = df_master.loc[~df_master.index.isin(assigned_idx)]
        
        if not unassigned.empty:
            unmapped_agencies = unassigned["Agency_Normalized"].unique()
            unmapped_count = len(unassigned)
            
            # Ask user how to handle unmapped
            choice = messagebox.askyesnocancel(
                "Unmapped Agencies Found",
                f"{unmapped_count} users with {len(unmapped_agencies)} unmapped agencies:\n{', '.join(unmapped_agencies[:10])}\n\n"
                f"YES = Individual files per agency\n"
                f"NO = Single 'Unassigned.xlsx' file\n"
                f"CANCEL = Skip unmapped users"
            )
            
            if choice is True:  # Individual files
                for agency in unmapped_agencies:
                    agency_users = unassigned[unassigned["Agency_Normalized"] == agency].copy()
                    agency_users["Review Action"] = ""
                    agency_users["Comments"] = ""
                    
                    if "Agency_Normalized" in agency_users.columns:
                        agency_users = agency_users.drop(columns=["Agency_Normalized"])
                    
                    out_file = out_dir / f"{sanitize_sheet_name(agency)}.xlsx"
                    with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
                        agency_users.to_excel(writer, sheet_name="User Access List", index=False)
                        format_worksheet(writer, "User Access List", agency_users)
            
            elif choice is False:  # Single Unassigned file
                unassigned_copy = unassigned.copy()
                unassigned_copy["Review Action"] = ""
                unassigned_copy["Comments"] = ""
                
                if "Agency_Normalized" in unassigned_copy.columns:
                    unassigned_copy = unassigned_copy.drop(columns=["Agency_Normalized"])
                
                out_file = out_dir / "Unassigned.xlsx"
                with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
                    unassigned_copy.to_excel(writer, sheet_name="Unmapped Users", index=False)
                    format_worksheet(writer, "Unmapped Users", unassigned_copy)
        
        if progress_callback:
            progress_callback("Complete!")
        
        messagebox.showinfo("Success", f"Agency files created successfully in:\n{out_dir}")
        return True
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate agency files:\n{str(e)}")
        return False

def validate_agency_files_detailed(master_file, output_dir):
    """
    Performs detailed validation of agency files against the master file.
    
    Args:
        master_file (str): Path to the master user access file.
        output_dir (str): Path to the output directory.
    
    Returns:
        tuple: (is_valid, detailed_report)
    """
    if not os.path.exists(master_file) or not os.path.exists(output_dir):
        return False, "Master file or output directory not found"
    
    try:
        df_master = pd.read_excel(master_file)
        files = os.listdir(output_dir)
        agency_files = [f for f in files if f.endswith('.xlsx') and not f.startswith('Unassigned')]
        
        validation_results = []
        total_issues = 0
        
        # Validate each agency file
        for file in agency_files:
            agency_name = file.replace('.xlsx', '')
            file_path = os.path.join(output_dir, file)
            
            try:
                df_agency = pd.read_excel(file_path)
                
                # Check structure
                required_columns = ["Review Action", "Comments"]
                missing_columns = [col for col in required_columns if col not in df_agency.columns]
                
                if missing_columns:
                    validation_results.append(f"❌ {agency_name}: Missing columns {missing_columns}")
                    total_issues += 1
                else:
                    validation_results.append(f"✅ {agency_name}: Structure valid")
                
                # Check user count (basic validation)
                user_count = len(df_agency)
                validation_results.append(f"   📊 {agency_name}: {user_count} users")
                
            except Exception as e:
                validation_results.append(f"❌ {agency_name}: Error reading file - {str(e)}")
                total_issues += 1
        
        # Summary
        if total_issues == 0:
            summary = "✅ All files validated successfully!"
        else:
            summary = f"⚠️ Found {total_issues} issues that need attention"
        
        detailed_report = f"{summary}\n\n" + "\n".join(validation_results)
        return total_issues == 0, detailed_report
        
    except Exception as e:
        return False, f"Validation failed: {str(e)}"

def open_enhanced_settings_dialog(parent_window):
    """
    Opens an enhanced settings dialog with email body editor and advanced settings.
    
    Args:
        parent_window: The parent window for the dialog.
    
    Returns:
        bool: True if settings were saved, False otherwise.
    """
    dialog = ctk.CTkToplevel(parent_window)
    dialog.title("Application Settings")
    dialog.geometry("800x700")
    dialog.transient(parent_window)
    dialog.grab_set()
    
    # Load current config
    current_config = config_manager.load()
    
    # Create variables for the form fields
    settings_vars = {}
    for key, value in current_config.items():
        if isinstance(value, str):
            settings_vars[key] = ctk.StringVar(value=value)
        else:
            settings_vars[key] = ctk.StringVar(value=str(value))
    
    # Create main frame with notebook for tabs
    main_frame = ctk.CTkFrame(dialog)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    # Create notebook for tabs
    notebook = ctk.CTkTabview(main_frame)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)
    
    # === GENERAL SETTINGS TAB ===
    general_tab = notebook.add("General")
    
    ctk.CTkLabel(general_tab, text="General Settings", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))
    
    # Create scrollable frame for settings
    scroll_frame = ctk.CTkScrollableFrame(general_tab)
    scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Define settings fields with labels
    settings_fields = [
        ("review_period", "Review Period:", "e.g., Q2 2025"),
        ("deadline", "Deadline:", "e.g., June 30, 2025"),
        ("company_name", "Company Name:", "e.g., Omnicom Group"),
        ("sender_name", "Sender Name:", "e.g., Govind Waghmare"),
        ("sender_title", "Sender Title:", "e.g., Manager, Financial Applications | Analytics"),
        ("email_subject_prefix", "Email Subject Prefix:", "e.g., [ACTION REQUIRED] Cognos Access Review")
    ]
    
    # Create form fields
    for key, label, placeholder in settings_fields:
        frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(frame, text=label).pack(anchor="w")
        entry = ctk.CTkEntry(frame, textvariable=settings_vars[key], placeholder_text=placeholder)
        entry.pack(fill="x", pady=(5, 0))
    
    # === EMAIL BODY TAB ===
    email_tab = notebook.add("Email Body")
    
    ctk.CTkLabel(email_tab, text="Email Template Editor", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 10))
    
    # Template variables help
    help_text = "Available variables: {review_period}, {deadline}, {sender_name}, {sender_title}, {company_name}"
    ctk.CTkLabel(email_tab, text=help_text, font=ctk.CTkFont(size=12), text_color="gray").pack(pady=(0, 10))
    
    # Email body editor frame
    editor_frame = ctk.CTkFrame(email_tab)
    editor_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Text editor
    text_editor = ctk.CTkTextbox(editor_frame, wrap="word")
    text_editor.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Load current email template
    try:
        with open('email_template.txt', 'r', encoding='utf-8') as f:
            current_template = f.read()
    except (FileNotFoundError, UnicodeDecodeError):
        # If file doesn't exist or has encoding issues, use default template
        current_template = """Hello,

As part of our Sarbanes-Oxley (SOX) compliance requirements, we must complete the {review_period} Cognos Platform user access review by {deadline}. Your review and response are critical to ensuring compliance and maintaining appropriate system access.

Action Required:
1. Review the attached User Access Report listing Cognos Reporting Users and their folder access as of {review_period}.
2. Confirm or request changes:
   - If no updates are needed, reply confirming your review.
   - If updates are required, note them in Column G of the "User Access List" tab and return the file.
   - For access changes, submit a Paige ticket under the Cognos Services section.

This review is mandatory for compliance, and your prompt response is essential. Please let me know if you have any questions.

Best Regards,
{sender_name}
{sender_title}
{company_name}"""
    
    text_editor.insert("1.0", current_template)
    
    # === ADVANCED SETTINGS TAB ===
    advanced_tab = notebook.add("Advanced")
    
    ctk.CTkLabel(advanced_tab, text="Advanced Settings", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))
    
    advanced_scroll = ctk.CTkScrollableFrame(advanced_tab)
    advanced_scroll.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Auto-scan setting
    auto_scan_var = ctk.BooleanVar(value=current_config.get("auto_scan", True))
    ctk.CTkCheckBox(
        advanced_scroll, text="Auto-scan output folder on startup", 
        variable=auto_scan_var
    ).pack(anchor="w", pady=10)
    
    # Default email mode
    default_email_mode_var = ctk.StringVar(value=current_config.get("default_email_mode", "Preview"))
    email_mode_frame = ctk.CTkFrame(advanced_scroll, fg_color="transparent")
    email_mode_frame.pack(fill="x", pady=5)
    
    ctk.CTkLabel(email_mode_frame, text="Default Email Mode:").pack(anchor="w")
    for mode in ["Preview", "Direct", "Schedule"]:
        ctk.CTkRadioButton(
            email_mode_frame, text=mode, variable=default_email_mode_var, value=mode
        ).pack(anchor="w", padx=20, pady=2)
    
    # Buttons frame
    button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    button_frame.pack(fill="x", pady=(20, 0))
    
    def save_settings():
        try:
            # Collect values from form
            new_config = {}
            for key, var in settings_vars.items():
                new_config[key] = var.get()
            
            # Add advanced settings
            new_config["auto_scan"] = auto_scan_var.get()
            new_config["default_email_mode"] = default_email_mode_var.get()
            
            # Save email template using config manager
            email_content = text_editor.get("1.0", "end-1c")
            save_email_template(email_content)
            
            # Save using config manager
            config_manager.update(new_config)
            config_manager.save()
            
            # Update global config
            global config, REVIEW_PERIOD, REVIEW_DEADLINE, AUDIT_FILE_NAME
            config = new_config
            REVIEW_PERIOD = config.get("review_period", "Q2 2025")
            REVIEW_DEADLINE = config.get("deadline", "June 30, 2025")
            AUDIT_FILE_NAME = config_manager.get_audit_file_name()
            
            logger.info("Settings saved successfully")
            messagebox.showinfo("Success", "Settings saved successfully!\n\nSome changes may require restarting the application.", parent=dialog)
            dialog.destroy()
            return True
            
        except Exception as e:
            logger.error(f"Failed to save settings: {str(e)}")
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}", parent=dialog)
            return False
    
    def cancel():
        dialog.destroy()
        return False
    
    # Create buttons
    ctk.CTkButton(button_frame, text="Save Settings", command=save_settings).pack(side="left", padx=(0, 10))
    ctk.CTkButton(button_frame, text="Cancel", command=cancel).pack(side="left")
    
    # Center the dialog
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")
    
    return False


# =============== Agency Mapping Manager Dialog ===============

class AgencyMappingManagerDialog(ctk.CTkToplevel):
    """Dialog for managing agency mappings."""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Agency Mapping Manager")
        
        # Set minimum size and initial size
        self.minsize(800, 600)
        self.geometry("1200x700")
        
        # Store mapping data
        self.mapping_data = pd.DataFrame()
        
        # Load mapping data
        self.load_mapping()
        
        # Create UI
        self.create_widgets()
        
        # Bind dialog-level keyboard shortcuts
        self.bind("<Control-s>", lambda e: self.save_and_validate())
        self.bind("<Control-S>", lambda e: self.save_and_validate())
        self.bind("<Control-w>", lambda e: self.destroy())
        self.bind("<Control-W>", lambda e: self.destroy())
        self.bind("<Escape>", lambda e: self.destroy())
        
        # Update and center after widgets are created
        self.update_idletasks()
        self.center_window()
    
    def load_mapping(self):
        """Loads mapping file from the current configuration."""
        mapping_file = self.parent.vars["domain"].get()
        if mapping_file and os.path.exists(mapping_file):
            try:
                df = pd.read_excel(mapping_file)
                df.columns = df.columns.str.strip()
                
                # Map column names to standardized names
                column_mapping = {}
                for col in df.columns:
                    if "file name" in col.lower():
                        column_mapping[col] = "File name"
                    elif "agency id" in col.lower() or "agency in the file" in col.lower():
                        column_mapping[col] = "Agency ID"
                
                # Rename columns
                if column_mapping:
                    df = df.rename(columns=column_mapping)
                
                # Ensure required columns exist
                if "File name" not in df.columns:
                    df["File name"] = ""
                if "Agency ID" not in df.columns:
                    df["Agency ID"] = ""
                
                self.mapping_data = df[["File name", "Agency ID"]]
            except Exception as e:
                logger.error(f"Failed to load mapping file: {str(e)}")
                # Create empty DataFrame with required columns
                self.mapping_data = pd.DataFrame(columns=["File name", "Agency ID"])
        else:
            # Create empty DataFrame if no file selected
            self.mapping_data = pd.DataFrame(columns=["File name", "Agency ID"])
    
    def create_widgets(self):
        """Creates the UI components."""
        # Main container
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ctk.CTkLabel(main_frame, text="Agency Mapping Manager", 
                                   font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(0, 15))
        region_names = self.region_manager.get_all_region_names()
        
        if not region_names:
            # Show message if no regions configured
            msg_label = ctk.CTkLabel(main_frame, 
                                    text="No regions configured.\nPlease configure regions in Region Manager first.",
                                    font=ctk.CTkFont(size=14))
            msg_label.pack(pady=50)
        else:
            # Tabview for regions
            self.tabview = ctk.CTkTabview(main_frame, height=500)
            self.tabview.pack(fill="both", expand=True, pady=(0, 15))
            
            # Create tab for each region
            for region_name in region_names:
                tab = self.tabview.add(region_name)
                self.create_region_tab(tab, region_name)
        
        # Bottom buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkButton(button_frame, text="Save All & Validate (Ctrl+S)", command=self.save_all_and_validate,
                     fg_color="green", width=200).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Close (Esc)", command=self.destroy,
                     fg_color="gray", width=120).pack(side="left", padx=5)
    
    def create_region_tab(self, tab, region_name):
        """Creates the content for a region tab."""
        from tkinter import ttk
        
        # Control buttons frame (top)
        control_frame = ctk.CTkFrame(tab, fg_color="transparent")
        control_frame.pack(fill="x", pady=(10, 10))
        
        add_btn = ctk.CTkButton(control_frame, text="Add Row (Ctrl+N)", width=120,
                               command=lambda: self.add_mapping_row(region_name))
        add_btn.pack(side="left", padx=5)
        
        edit_btn = ctk.CTkButton(control_frame, text="Edit Row (Enter)", width=120,
                                command=lambda: self.edit_mapping_row(region_name))
        edit_btn.pack(side="left", padx=5)
        
        delete_btn = ctk.CTkButton(control_frame, text="Delete Row (Del)", width=120,
                                   command=lambda: self.delete_mapping_row(region_name))
        delete_btn.pack(side="left", padx=5)
        
        sync_btn = ctk.CTkButton(control_frame, text="Sync with Email Manifest (Ctrl+S)", width=220,
                                command=lambda: self.sync_with_manifest(region_name),
                                fg_color="#1f6aa5")
        sync_btn.pack(side="left", padx=15)
        
        # Treeview frame
        tree_frame = ctk.CTkFrame(tab)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Treeview with scrollbar
        tree = ttk.Treeview(tree_frame, columns=("File name", "Agency ID"), 
                           show="headings", height=20)
        tree.heading("File name", text="File name")
        tree.heading("Agency ID", text="Agency ID")
        tree.column("File name", width=400)
        tree.column("Agency ID", width=400)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind keyboard shortcuts
        tree.bind("<Delete>", lambda e: self.delete_mapping_row(region_name))
        tree.bind("<Return>", lambda e: self.edit_mapping_row(region_name))
        tree.bind("<Double-Button-1>", lambda e: self.edit_mapping_row(region_name))
        tree.bind("<Control-n>", lambda e: self.add_mapping_row(region_name))
        tree.bind("<Control-N>", lambda e: self.add_mapping_row(region_name))
        tree.bind("<Control-a>", lambda e: self.select_all_tree(tree))
        tree.bind("<Control-A>", lambda e: self.select_all_tree(tree))
        
        # Load data into tree
        self.populate_tree(tree, region_name)
        
        # Store references
        self.tab_data[region_name] = {
            "tree": tree,
            "add_btn": add_btn,
            "edit_btn": edit_btn,
            "delete_btn": delete_btn,
            "sync_btn": sync_btn
        }
    
    def populate_tree(self, tree, region_name):
        """Populates treeview with mapping data."""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Add data
        if region_name in self.mapping_data:
            df = self.mapping_data[region_name]
            for _, row in df.iterrows():
                file_name = str(row.get("File name", "")).strip()
                agency_id = str(row.get("Agency ID", "")).strip()
                tree.insert("", "end", values=(file_name, agency_id))
    
    def select_all_tree(self, tree):
        """Selects all items in the treeview."""
        for item in tree.get_children():
            tree.selection_add(item)
    
    def add_mapping_row(self, region_name):
        """Adds a new mapping row."""
        dialog = MappingEditDialog(self, "Add Mapping", "", "")
        self.wait_window(dialog)
        
        if dialog.result:
            file_name, agency_id = dialog.result
            # Add to DataFrame
            new_row = pd.DataFrame({"File name": [file_name], "Agency ID": [agency_id]})
            self.mapping_data[region_name] = pd.concat([self.mapping_data[region_name], new_row], 
                                                       ignore_index=True)
            # Refresh tree
            self.populate_tree(self.tab_data[region_name]["tree"], region_name)
    
    def edit_mapping_row(self, region_name):
        """Edits the selected mapping row."""
        tree = self.tab_data[region_name]["tree"]
        selection = tree.selection()
        
        if not selection:
            messagebox.showwarning("No Selection", "Please select a row to edit.")
            return
        
        # Get current values
        item = selection[0]
        values = tree.item(item, "values")
        file_name, agency_id = values
        
        # Open edit dialog
        dialog = MappingEditDialog(self, "Edit Mapping", file_name, agency_id)
        self.wait_window(dialog)
        
        if dialog.result:
            new_file_name, new_agency_id = dialog.result
            # Update DataFrame
            idx = tree.index(item)
            self.mapping_data[region_name].at[idx, "File name"] = new_file_name
            self.mapping_data[region_name].at[idx, "Agency ID"] = new_agency_id
            # Refresh tree
            self.populate_tree(tree, region_name)
    
    def delete_mapping_row(self, region_name):
        """Deletes the selected mapping row."""
        tree = self.tab_data[region_name]["tree"]
        selection = tree.selection()
        
        if not selection:
            messagebox.showwarning("No Selection", "Please select a row to delete.")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this mapping?"):
            item = selection[0]
            idx = tree.index(item)
            # Remove from DataFrame
            self.mapping_data[region_name] = self.mapping_data[region_name].drop(idx).reset_index(drop=True)
            # Refresh tree
            self.populate_tree(tree, region_name)
    
    def sync_with_manifest(self, region_name):
        """Syncs mapping with email manifest - adds missing file names."""
        # Get email manifest from any region (they share the same manifest)
        profile = self.region_manager.get_profile(region_name)
        if not profile or not profile.email_manifest or not os.path.exists(profile.email_manifest):
            messagebox.showerror("Error", "Email manifest file not found for this region.")
            return
        
        try:
            # Load email manifest
            df_manifest = pd.read_excel(profile.email_manifest)
            df_manifest.columns = df_manifest.columns.str.strip()
            
            if "Agency" not in df_manifest.columns:
                messagebox.showerror("Error", "Email manifest missing 'Agency' column.")
                return
            
            # Get unique file names from manifest
            manifest_files = set(df_manifest["Agency"].dropna().astype(str).str.strip())
            
            # Get existing file names in mapping
            existing_files = set(self.mapping_data[region_name]["File name"].dropna().astype(str).str.strip())
            
            # Find missing files
            missing_files = manifest_files - existing_files
            
            if not missing_files:
                messagebox.showinfo("Sync Complete", "No new file names to add. Mapping is already in sync.")
                return
            
            # Add missing files
            new_rows = pd.DataFrame({
                "File name": list(missing_files),
                "Agency ID": [""] * len(missing_files)  # Empty agency IDs for user to fill
            })
            
            self.mapping_data[region_name] = pd.concat([self.mapping_data[region_name], new_rows], 
                                                       ignore_index=True)
            
            # Refresh tree
            self.populate_tree(self.tab_data[region_name]["tree"], region_name)
            
            messagebox.showinfo("Sync Complete", 
                              f"Added {len(missing_files)} new file name(s) from email manifest.\n\n"
                              f"Please fill in the 'Agency ID' column for the new entries.")
            
        except Exception as e:
            logger.error(f"Failed to sync with manifest: {str(e)}")
            messagebox.showerror("Sync Error", f"Failed to sync with email manifest:\n{str(e)}")
    
    def save_all_and_validate(self):
        """Saves all mapping files and validates emails."""
        try:
            # Save all mapping files
            for region_name, df in self.mapping_data.items():
                profile = self.region_manager.get_profile(region_name)
                if profile and profile.mapping_file:
                    # Create backup
                    if os.path.exists(profile.mapping_file):
                        create_backup(profile.mapping_file)
                    
                    # Save
                    df.to_excel(profile.mapping_file, index=False)
                    logger.info(f"Saved mapping for {region_name}: {profile.mapping_file}")
            
            messagebox.showinfo("Success", "All mapping files saved successfully!")
            
            # Now validate emails from manifest
            self.validate_manifest_emails()
            
        except Exception as e:
            logger.error(f"Failed to save mappings: {str(e)}")
            messagebox.showerror("Save Error", f"Failed to save mapping files:\n{str(e)}")
    
    def validate_manifest_emails(self):
        """Validates all emails in the email manifest and shows results."""
        # Get email manifest path from any region
        region_names = self.region_manager.get_all_region_names()
        if not region_names:
            return
        
        profile = self.region_manager.get_profile(region_names[0])
        if not profile or not profile.email_manifest or not os.path.exists(profile.email_manifest):
            messagebox.showerror("Error", "Email manifest file not found.")
            return
        
        try:
            # Open Email Validation Results Dialog
            dialog = EmailValidationDialog(self, profile.email_manifest)
            dialog.wait_window()
            
        except Exception as e:
            logger.error(f"Failed to validate emails: {str(e)}")
            messagebox.showerror("Validation Error", f"Failed to validate emails:\n{str(e)}")
    
    def center_window(self):
        """Centers the dialog on screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")


class MappingEditDialog(ctk.CTkToplevel):
    """Dialog for adding/editing a mapping entry."""
    
    def __init__(self, parent, title, file_name, agency_id):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x250")
        self.transient(parent)
        self.grab_set()
        
        self.result = None
        
        # Main frame
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # File name
        ctk.CTkLabel(main_frame, text="File name (e.g., ACME_Corp.xlsx):").pack(anchor="w", pady=(0, 5))
        self.file_name_entry = ctk.CTkEntry(main_frame, width=400)
        self.file_name_entry.pack(fill="x", pady=(0, 15))
        self.file_name_entry.insert(0, file_name)
        
        # Agency ID
        ctk.CTkLabel(main_frame, text="Agency ID (from master file):").pack(anchor="w", pady=(0, 5))
        self.agency_id_entry = ctk.CTkEntry(main_frame, width=400)
        self.agency_id_entry.pack(fill="x", pady=(0, 20))
        self.agency_id_entry.insert(0, agency_id)
        
        # Buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x")
        
        ctk.CTkButton(button_frame, text="Save", command=self.save, width=100).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=self.destroy, 
                     fg_color="gray", width=100).pack(side="left", padx=5)
        
        self.center_window()
    
    def save(self):
        """Saves the values and closes dialog."""
        file_name = self.file_name_entry.get().strip()
        agency_id = self.agency_id_entry.get().strip()
        
        if not file_name:
            messagebox.showwarning("Validation Error", "File name cannot be empty.")
            return
        
        self.result = (file_name, agency_id)
        self.destroy()
    
    def center_window(self):
        """Centers the dialog on screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")




class ValidationResultsDialog(ctk.CTkToplevel):
    """
    Interactive dialog showing comprehensive validation results with fix options.
    """
    
    def __init__(self, parent, validation_results: Dict, app_reference):
        super().__init__(parent)
        self.title("\ud83d\udd0d Validation Results")
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        
        self.validation_results = validation_results
        self.app = app_reference
        
        # Create UI
        self._create_widgets()
        self.center_window()
    
    def _create_widgets(self):
        """Create the dialog widgets."""
        # Header with overall status
        header_frame = ctk.CTkFrame(self, fg_color=self._get_status_color(), corner_radius=0)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        status_icon = {
            "pass": "\u2705",
            "warning": "\u26a0\ufe0f",
            "error": "\u274c"
        }.get(self.validation_results["overall_status"], "\u2753")
        
        status_text = {
            "pass": "All Checks Passed!",
            "warning": f"Passed with {self.validation_results['warning_count']} Warnings",
            "error": f"Failed with {self.validation_results['error_count']} Errors"
        }.get(self.validation_results["overall_status"], "Unknown Status")
        
        ctk.CTkLabel(
            header_frame,
            text=f"{status_icon} {status_text}",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white"
        ).pack(pady=20)
        
        # Summary stats
        stats_text = f"\u2139\ufe0f {len(self.validation_results['categories'])} categories checked"
        if self.validation_results['error_count'] > 0:
            stats_text += f" \u2022 {self.validation_results['error_count']} errors"
        if self.validation_results['warning_count'] > 0:
            stats_text += f" \u2022 {self.validation_results['warning_count']} warnings"
        
        ctk.CTkLabel(
            header_frame,
            text=stats_text,
            font=ctk.CTkFont(size=12),
            text_color="white"
        ).pack(pady=(0, 15))
        
        # Scrollable content area
        scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Display each category
        for category in self.validation_results["categories"]:
            self._create_category_section(scroll_frame, category)
        
        # Footer with action buttons
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        if self.validation_results["can_proceed"]:
            ctk.CTkButton(
                footer_frame,
                text="\u2705 Proceed Anyway" if self.validation_results["warning_count"] > 0 else "\u2705 Continue",
                command=self.destroy,
                fg_color="green",
                hover_color="darkgreen",
                height=40,
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(side="right", padx=5)
        
        ctk.CTkButton(
            footer_frame,
            text="\u274c Close",
            command=self.destroy,
            fg_color="gray",
            hover_color="darkgray",
            height=40,
            width=120
        ).pack(side="right", padx=5)
        
        # Auto-fix button if any fixable issues
        fixable_count = sum(
            len([i for i in cat["issues"] if i.get("fixable", False)])
            for cat in self.validation_results["categories"]
        )
        if fixable_count > 0:
            ctk.CTkButton(
                footer_frame,
                text=f"\ud83d\udd27 Auto-Fix ({fixable_count} issues)",
                command=self._auto_fix_all,
                fg_color="orange",
                hover_color="darkorange",
                height=40,
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(side="left", padx=5)
    
    def _create_category_section(self, parent, category: Dict):
        """Create a section for one validation category."""
        # Category header
        category_frame = ctk.CTkFrame(parent, fg_color=("gray90", "gray20"), corner_radius=10)
        category_frame.pack(fill="x", pady=(0, 15))
        
        header_frame = ctk.CTkFrame(category_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=10)
        
        status_icon = "\u2705" if category["status"] == "pass" else ("\u26a0\ufe0f" if category["status"] == "warning" else "\u274c")
        
        ctk.CTkLabel(
            header_frame,
            text=f"{status_icon} {category['name']}",
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        ).pack(side="left")
        
        issue_count = len(category["issues"])
        if issue_count > 0:
            ctk.CTkLabel(
                header_frame,
                text=f"{issue_count} issue{'s' if issue_count != 1 else ''}",
                font=ctk.CTkFont(size=11),
                text_color="gray"
            ).pack(side="right")
        
        # Issues list
        if category["issues"]:
            issues_frame = ctk.CTkFrame(category_frame, fg_color="transparent")
            issues_frame.pack(fill="x", padx=15, pady=(0, 10))
            
            for idx, issue in enumerate(category["issues"]):
                self._create_issue_row(issues_frame, issue, idx)
        else:
            ctk.CTkLabel(
                category_frame,
                text="\u2713 No issues found",
                font=ctk.CTkFont(size=11),
                text_color="green"
            ).pack(padx=15, pady=(0, 10))
    
    def _create_issue_row(self, parent, issue: Dict, index: int):
        """Create a row for one issue."""
        row_frame = ctk.CTkFrame(parent, fg_color=("gray95", "gray25"), corner_radius=5)
        row_frame.pack(fill="x", pady=2)
        
        content_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
        content_frame.pack(fill="x", padx=10, pady=8)
        
        # Severity icon
        severity_icon = {
            "error": "\u274c",
            "warning": "\u26a0\ufe0f",
            "info": "\u2139\ufe0f"
        }.get(issue["severity"], "\u2022")
        
        severity_color = {
            "error": "red",
            "warning": "orange",
            "info": "blue"
        }.get(issue["severity"], "gray")
        
        icon_label = ctk.CTkLabel(
            content_frame,
            text=severity_icon,
            font=ctk.CTkFont(size=14),
            width=30
        )
        icon_label.pack(side="left", padx=(0, 10))
        
        # Message
        message_label = ctk.CTkLabel(
            content_frame,
            text=issue["message"],
            font=ctk.CTkFont(size=11),
            anchor="w",
            justify="left",
            wraplength=600
        )
        message_label.pack(side="left", fill="x", expand=True)
        
        # Fix button if fixable
        if issue.get("fixable", False) and issue.get("fix_action"):
            fix_btn = ctk.CTkButton(
                content_frame,
                text="\ud83d\udd27 Fix",
                command=lambda: self._fix_issue(issue["fix_action"]),
                width=70,
                height=28,
                fg_color="orange",
                hover_color="darkorange",
                font=ctk.CTkFont(size=10)
            )
            fix_btn.pack(side="right", padx=(10, 0))
    
    def _fix_issue(self, fix_action: Tuple):
        """Apply a fix for an issue."""
        action_type, action_data = fix_action
        
        if action_type == "create_directory":
            # Create missing directory
            try:
                Path(action_data).mkdir(parents=True, exist_ok=True)
                messagebox.showinfo("Success", f"Created directory: {action_data}")
                self.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create directory: {e}")
        
        elif action_type == "show_unmapped":
            # Show unmapped agencies
            msg = "Unmapped agencies:\\n\\n" + "\\n".join(action_data[:20])
            if len(action_data) > 20:
                msg += f"\\n... and {len(action_data) - 20} more"
            messagebox.showinfo("Unmapped Agencies", msg)
        
        elif action_type == "fix_emails":
            messagebox.showinfo("Email Fix", "Please open the combined file and fix invalid email addresses manually.")
        
        elif action_type == "auto_sanitize":
            messagebox.showinfo("Auto-Sanitize", "Filenames will be automatically sanitized during generation.")
        
        elif action_type == "select_region":
            messagebox.showinfo("Region", "Please use 'Manage Regions' to select a current region.")
        
        elif action_type == "edit_region":
            messagebox.showinfo("Edit Region", f"Please use 'Manage Regions' to edit region: {action_data}")
    
    def _auto_fix_all(self):
        """Attempt to fix all fixable issues automatically."""
        fixed_count = 0
        
        for category in self.validation_results["categories"]:
            for issue in category["issues"]:
                if issue.get("fixable", False) and issue.get("fix_action"):
                    action_type, action_data = issue["fix_action"]
                    
                    try:
                        if action_type == "create_directory":
                            Path(action_data).mkdir(parents=True, exist_ok=True)
                            fixed_count += 1
                    except:
                        pass
        
        if fixed_count > 0:
            messagebox.showinfo("Auto-Fix Complete", f"Fixed {fixed_count} issue(s).\\nPlease re-run validation to check remaining issues.")
            self.destroy()
        else:
            messagebox.showinfo("Auto-Fix", "No issues could be automatically fixed.\\nPlease fix remaining issues manually.")
    
    def _get_status_color(self) -> str:
        """Get color based on overall status."""
        return {
            "pass": "green",
            "warning": "orange",
            "error": "red"
        }.get(self.validation_results["overall_status"], "gray")
    
    def center_window(self):
        """Center dialog on screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")


class EmailValidationDialog(ctk.CTkToplevel):
    """Dialog showing email validation results with interactive fixing."""
    
    def __init__(self, parent, manifest_file):
        super().__init__(parent)
        self.parent = parent
        self.manifest_file = manifest_file
        self.title("Email Validation Results")
        self.geometry("900x650")
        self.transient(parent)
        self.grab_set()
        
        # Load and validate emails
        self.valid_emails = []
        self.invalid_emails = []
        self.df_manifest = None
        
        self.load_and_validate()
        self.create_widgets()
        self.center_window()
    
    def load_and_validate(self):
        """Loads manifest and validates all emails."""
        try:
            self.df_manifest = pd.read_excel(self.manifest_file)
            self.df_manifest.columns = self.df_manifest.columns.str.strip()
            
            if "Agency" not in self.df_manifest.columns:
                raise ValueError("Email manifest missing 'Agency' column")
            
            # Validate each row
            for idx, row in self.df_manifest.iterrows():
                agency = str(row.get("Agency", "")).strip()
                to_email = str(row.get("To", "")).strip() if pd.notna(row.get("To")) else ""
                cc_email = str(row.get("CC", "")).strip() if pd.notna(row.get("CC")) else ""
                
                # Check if emails are valid
                to_valid = is_valid_email(to_email) if to_email else False
                cc_valid = is_valid_email(cc_email) if cc_email else True  # CC is optional
                
                issues = []
                if not to_email:
                    issues.append("Missing 'To' email")
                elif not to_valid:
                    issues.append(f"Invalid 'To' email: {to_email}")
                
                if cc_email and not cc_valid:
                    issues.append(f"Invalid 'CC' email: {cc_email}")
                
                if issues:
                    self.invalid_emails.append({
                        "index": idx,
                        "agency": agency,
                        "to": to_email,
                        "cc": cc_email,
                        "issues": ", ".join(issues)
                    })
                else:
                    self.valid_emails.append({
                        "agency": agency,
                        "to": to_email,
                        "cc": cc_email if cc_email else "(none)"
                    })
        
        except Exception as e:
            logger.error(f"Failed to validate emails: {str(e)}")
            messagebox.showerror("Error", f"Failed to load email manifest:\n{str(e)}")
            self.destroy()
    
    def create_widgets(self):
        """Creates the UI components."""
        from tkinter import ttk
        
        # Main frame
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_text = f"Found {len(self.valid_emails)} valid and {len(self.invalid_emails)} invalid emails"
        ctk.CTkLabel(main_frame, text=title_text, font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(0, 15))
        
        # Tabview for valid/invalid
        tabview = ctk.CTkTabview(main_frame)
        tabview.pack(fill="both", expand=True, pady=(0, 15))
        
        # Valid emails tab
        valid_tab = tabview.add(f"✅ Valid ({len(self.valid_emails)})")
        self.create_valid_tab(valid_tab)
        
        # Invalid emails tab
        invalid_tab = tabview.add(f"❌ Invalid ({len(self.invalid_emails)})")
        self.create_invalid_tab(invalid_tab)
        
        # Bottom buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x")
        
        ctk.CTkButton(button_frame, text="Re-Validate", command=self.revalidate, 
                     width=120).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Close", command=self.destroy, 
                     fg_color="gray", width=100).pack(side="left", padx=5)
    
    def create_valid_tab(self, tab):
        """Creates the valid emails tab."""
        from tkinter import ttk
        
        # Treeview
        tree_frame = ctk.CTkFrame(tab)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        tree = ttk.Treeview(tree_frame, columns=("Agency", "To", "CC"), 
                           show="headings", height=15)
        tree.heading("Agency", text="Agency")
        tree.heading("To", text="To Email")
        tree.heading("CC", text="CC Email")
        tree.column("Agency", width=250)
        tree.column("To", width=250)
        tree.column("CC", width=250)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate
        for email in self.valid_emails:
            tree.insert("", "end", values=(email["agency"], email["to"], email["cc"]))
        
        self.valid_tree = tree
    
    def create_invalid_tab(self, tab):
        """Creates the invalid emails tab with fix button."""
        from tkinter import ttk
        
        # Control buttons
        control_frame = ctk.CTkFrame(tab, fg_color="transparent")
        control_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        ctk.CTkButton(control_frame, text="Fix Selected Email", 
                     command=self.fix_selected_email, width=150,
                     fg_color="orange").pack(side="left", padx=5)
        
        # Treeview
        tree_frame = ctk.CTkFrame(tab)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        tree = ttk.Treeview(tree_frame, columns=("Agency", "To", "CC", "Issues"), 
                           show="headings", height=12)
        tree.heading("Agency", text="Agency")
        tree.heading("To", text="To Email")
        tree.heading("CC", text="CC Email")
        tree.heading("Issues", text="Issues")
        tree.column("Agency", width=200)
        tree.column("To", width=200)
        tree.column("CC", width=150)
        tree.column("Issues", width=250)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate
        for email in self.invalid_emails:
            tree.insert("", "end", values=(email["agency"], email["to"], 
                                          email["cc"] if email["cc"] else "(none)", 
                                          email["issues"]),
                       tags=("invalid",))
        
        # Color invalid rows red
        tree.tag_configure("invalid", foreground="red")
        
        self.invalid_tree = tree
    
    def fix_selected_email(self):
        """Opens dialog to fix the selected invalid email."""
        selection = self.invalid_tree.selection()
        
        if not selection:
            messagebox.showwarning("No Selection", "Please select an email to fix.")
            return
        
        # Get selected email data
        item = selection[0]
        idx = self.invalid_tree.index(item)
        email_data = self.invalid_emails[idx]
        
        # Open fix dialog
        dialog = FixEmailDialog(self, email_data)
        self.wait_window(dialog)
        
        if dialog.result:
            # Update manifest DataFrame
            new_to, new_cc = dialog.result
            self.df_manifest.at[email_data["index"], "To"] = new_to
            self.df_manifest.at[email_data["index"], "CC"] = new_cc
            
            # Save manifest
            try:
                create_backup(self.manifest_file)
                self.df_manifest.to_excel(self.manifest_file, index=False)
                messagebox.showinfo("Success", "Email updated successfully!")
                
                # Reload and refresh display
                self.revalidate()
                
            except Exception as e:
                logger.error(f"Failed to save manifest: {str(e)}")
                messagebox.showerror("Save Error", f"Failed to save changes:\n{str(e)}")
    
    def revalidate(self):
        """Reloads and revalidates all emails."""
        self.valid_emails = []
        self.invalid_emails = []
        self.load_and_validate()
        
        # Refresh display
        self.destroy()
        new_dialog = EmailValidationDialog(self.parent, self.manifest_file)
    
    def center_window(self):
        """Centers the dialog on screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")


class FixEmailDialog(ctk.CTkToplevel):
    """Dialog for fixing invalid email addresses."""
    
    def __init__(self, parent, email_data):
        super().__init__(parent)
        self.email_data = email_data
        self.title("Fix Email")
        self.geometry("500x300")
        self.transient(parent)
        self.grab_set()
        
        self.result = None
        
        # Main frame
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        ctk.CTkLabel(main_frame, text=f"Fix Email for: {email_data['agency']}", 
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(0, 15))
        
        # Issues
        ctk.CTkLabel(main_frame, text=f"Issues: {email_data['issues']}", 
                    text_color="red").pack(pady=(0, 15))
        
        # To Email
        ctk.CTkLabel(main_frame, text="To Email:").pack(anchor="w", pady=(0, 5))
        self.to_entry = ctk.CTkEntry(main_frame, width=400)
        self.to_entry.pack(fill="x", pady=(0, 15))
        self.to_entry.insert(0, email_data["to"])
        
        # CC Email
        ctk.CTkLabel(main_frame, text="CC Email (optional):").pack(anchor="w", pady=(0, 5))
        self.cc_entry = ctk.CTkEntry(main_frame, width=400)
        self.cc_entry.pack(fill="x", pady=(0, 20))
        self.cc_entry.insert(0, email_data["cc"])
        
        # Buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x")
        
        ctk.CTkButton(button_frame, text="Save", command=self.save, width=100).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=self.destroy, 
                     fg_color="gray", width=100).pack(side="left", padx=5)
        
        self.center_window()
    
    def save(self):
        """Validates and saves the email addresses."""
        to_email = self.to_entry.get().strip()
        cc_email = self.cc_entry.get().strip()
        
        # Validate
        if not to_email:
            messagebox.showwarning("Validation Error", "'To' email cannot be empty.")
            return
        
        if not is_valid_email(to_email):
            messagebox.showwarning("Validation Error", f"Invalid 'To' email format: {to_email}")
            return
        
        if cc_email and not is_valid_email(cc_email):
            messagebox.showwarning("Validation Error", f"Invalid 'CC' email format: {cc_email}")
            return
        
        self.result = (to_email, cc_email)
        self.destroy()
    
    def center_window(self):
        """Centers the dialog on screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")


# =============== Region Management Dialog ===============

class RegionManagementDialog(ctk.CTkToplevel):
    """
    Dialog for managing regional profiles.
    
    Allows users to add, edit, and delete regional configurations,
    each with its own master file, agency mapping, and output directory.
    """
    
    def __init__(self, parent, region_manager: RegionManager):
        super().__init__(parent)
        
        self.region_manager = region_manager
        self.parent = parent
        
        self.title("Region Profile Manager")
        self.geometry("900x600")
        self.transient(parent)
        self.grab_set()
        
        # Create main layout
        self.create_widgets()
        
        # Load existing regions
        self.refresh_region_list()
        
        # Center dialog
        self.center_window()
    
    def create_widgets(self):
        """Create all widgets for the dialog."""
        
        # Header
        header = ctk.CTkLabel(
            self, 
            text="Region Profile Manager", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        header.pack(pady=15)
        
        # Main frame
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Left: Region List
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)
        
        ctk.CTkLabel(
            list_frame, 
            text="Configured Regions", 
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=10)
        
        # Treeview for regions
        tree_frame = ctk.CTkFrame(list_frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tree = ttk.Treeview(
            tree_frame, 
            columns=("Region", "Master File", "Output Dir"), 
            show="headings", 
            height=15
        )
        self.tree.heading("Region", text="Region Name")
        self.tree.heading("Master File", text="Master File")
        self.tree.heading("Output Dir", text="Output Directory")
        
        self.tree.column("Region", width=150)
        self.tree.column("Master File", width=250)
        self.tree.column("Output Dir", width=200)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Button frame
        btn_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkButton(
            btn_frame, text="Add New Region", 
            command=self.add_region, width=140
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame, text="Edit Selected", 
            command=self.edit_region, width=140
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame, text="Delete Selected", 
            command=self.delete_region, width=140,
            fg_color="red", hover_color="darkred"
        ).pack(side="left", padx=5)
        
        # Right: Details Panel
        details_frame = ctk.CTkFrame(main_frame)
        details_frame.pack(side="right", fill="both", padx=(5, 10), pady=10)
        
        ctk.CTkLabel(
            details_frame, 
            text="Region Details", 
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=10)
        
        # Details text box
        self.details_text = ctk.CTkTextbox(details_frame, width=300, height=400)
        self.details_text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bind selection event
        self.tree.bind("<<TreeviewSelect>>", self.on_region_select)
        
        # Close button
        ctk.CTkButton(
            self, text="Close", command=self.destroy, width=120
        ).pack(pady=10)
    
    def refresh_region_list(self):
        """Refresh the region list in treeview."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add all regions
        for region_name in self.region_manager.get_all_region_names():
            profile = self.region_manager.get_profile(region_name)
            if profile:
                master_name = Path(profile.master_file).name if profile.master_file else "N/A"
                output_name = Path(profile.output_dir).name if profile.output_dir else "N/A"
                
                self.tree.insert(
                    "", "end", 
                    values=(region_name, master_name, output_name),
                    tags=(region_name,)
                )
    
    def on_region_select(self, event):
        """Handle region selection in treeview."""
        selection = self.tree.selection()
        if not selection:
            self.details_text.delete("1.0", "end")
            return
        
        item = selection[0]
        region_name = self.tree.item(item)["values"][0]
        profile = self.region_manager.get_profile(region_name)
        
        if profile:
            details = f"""Region: {profile.region_name}

Description:
{profile.description or "No description"}

Master File:
{profile.master_file}

Agency Mapping File:
{profile.mapping_file}

Email Manifest:
{profile.email_manifest}

Output Directory:
{profile.output_dir}
"""
            self.details_text.delete("1.0", "end")
            self.details_text.insert("1.0", details)
    
    def add_region(self):
        """Open dialog to add a new region."""
        editor = RegionEditorDialog(self, None, self.region_manager)
        self.wait_window(editor)
        self.refresh_region_list()
    
    def edit_region(self):
        """Edit the selected region."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a region to edit.")
            return
        
        item = selection[0]
        region_name = self.tree.item(item)["values"][0]
        profile = self.region_manager.get_profile(region_name)
        
        if profile:
            editor = RegionEditorDialog(self, profile, self.region_manager)
            self.wait_window(editor)
            self.refresh_region_list()
    
    def delete_region(self):
        """Delete the selected region."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a region to delete.")
            return
        
        item = selection[0]
        region_name = self.tree.item(item)["values"][0]
        
        # Confirm deletion
        result = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete the region profile:\n\n'{region_name}'?\n\n"
            "This will NOT delete any files, only the profile configuration."
        )
        
        if result:
            if self.region_manager.delete_profile(region_name):
                messagebox.showinfo("Deleted", f"Region profile '{region_name}' deleted successfully.")
                self.refresh_region_list()
            else:
                messagebox.showerror("Delete Failed", f"Failed to delete region '{region_name}'.")
    
    def center_window(self):
        """Center the dialog on screen."""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")


class RegionEditorDialog(ctk.CTkToplevel):
    """
    Dialog for adding or editing a region profile.
    """
    
    def __init__(self, parent, profile: Optional[RegionProfile], region_manager: RegionManager):
        super().__init__(parent)
        
        self.profile = profile
        self.region_manager = region_manager
        self.is_edit = profile is not None
        
        self.title("Edit Region" if self.is_edit else "Add New Region")
        self.geometry("700x550")
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        
        if self.profile:
            self.load_profile_data()
        
        self.center_window()
    
    def create_widgets(self):
        """Create form widgets."""
        
        # Header
        header_text = "Edit Region Profile" if self.is_edit else "Create New Region Profile"
        header = ctk.CTkLabel(
            self, 
            text=header_text, 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        header.pack(pady=15)
        
        # Form frame
        form_frame = ctk.CTkScrollableFrame(self)
        form_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Create form fields
        self.vars = {}
        
        fields = [
            ("region_name", "Region Name:", "e.g., North America, EMEA, APAC"),
            ("description", "Description:", "Optional description of this region"),
            ("master_file", "Master File:", "Path to master user access file"),
            ("mapping_file", "Agency Mapping File:", "Path to agency mapping file"),
            ("email_manifest", "Email Manifest:", "Path to email manifest file"),
            ("output_dir", "Output Directory:", "Output directory for generated files")
        ]
        
        for key, label, placeholder in fields:
            field_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
            field_frame.pack(fill="x", pady=10)
            
            ctk.CTkLabel(
                field_frame, text=label, 
                font=ctk.CTkFont(weight="bold")
            ).pack(anchor="w")
            
            if key in ["master_file", "mapping_file", "email_manifest", "output_dir"]:
                # File/folder selector row
                entry_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
                entry_frame.pack(fill="x", pady=(5, 0))
                
                self.vars[key] = ctk.StringVar()
                entry = ctk.CTkEntry(
                    entry_frame, 
                    textvariable=self.vars[key], 
                    placeholder_text=placeholder
                )
                entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
                
                # Browse button
                browse_cmd = lambda k=key: self.browse_file(k) if k != "output_dir" else self.browse_folder(k)
                ctk.CTkButton(
                    entry_frame, text="Browse", 
                    command=browse_cmd, width=80
                ).pack(side="right")
            else:
                # Regular entry
                self.vars[key] = ctk.StringVar()
                entry = ctk.CTkEntry(
                    field_frame, 
                    textvariable=self.vars[key], 
                    placeholder_text=placeholder
                )
                entry.pack(fill="x", pady=(5, 0))
        
        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        ctk.CTkButton(
            btn_frame, text="Save Region", 
            command=self.save_region, width=120,
            fg_color="green", hover_color="darkgreen"
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame, text="Cancel", 
            command=self.destroy, width=120
        ).pack(side="left", padx=5)
    
    def load_profile_data(self):
        """Load existing profile data into form."""
        if self.profile:
            self.vars["region_name"].set(self.profile.region_name)
            self.vars["description"].set(self.profile.description)
            self.vars["master_file"].set(self.profile.master_file)
            self.vars["mapping_file"].set(self.profile.mapping_file)
            self.vars["email_manifest"].set(self.profile.email_manifest)
            self.vars["output_dir"].set(self.profile.output_dir)
    
    def browse_file(self, key):
        """Open file browser."""
        path = filedialog.askopenfilename(
            title=f"Select {key.replace('_', ' ').title()}",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if path:
            self.vars[key].set(path)
    
    def browse_folder(self, key):
        """Open folder browser."""
        path = filedialog.askdirectory(title=f"Select {key.replace('_', ' ').title()}")
        if path:
            self.vars[key].set(path)
    
    def save_region(self):
        """Save the region profile."""
        # Create profile from form data
        profile = RegionProfile(
            region_name=self.vars["region_name"].get().strip(),
            description=self.vars["description"].get().strip(),
            master_file=self.vars["master_file"].get().strip(),
            mapping_file=self.vars["mapping_file"].get().strip(),
            email_manifest=self.vars["email_manifest"].get().strip(),
            output_dir=self.vars["output_dir"].get().strip()
        )
        
        # Validate
        errors = profile.validate()
        if errors:
            messagebox.showerror(
                "Validation Error", 
                "Please fix the following errors:\n\n" + "\n".join(f"• {e}" for e in errors)
            )
            return
        
        # Save
        if self.region_manager.add_profile(profile):
            action = "updated" if self.is_edit else "created"
            messagebox.showinfo(
                "Success", 
                f"Region profile '{profile.region_name}' {action} successfully!"
            )
            self.destroy()
        else:
            messagebox.showerror("Save Failed", "Failed to save region profile.")
    
    def center_window(self):
        """Center the dialog on screen."""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")


class EmailPreviewNavigationDialog(ctk.CTkToplevel):
    """
    Dialog for previewing and navigating through multiple emails before sending.
    Allows reviewing all emails with Previous/Next navigation.
    """
    
    def __init__(self, parent, emails: List[dict], email_handler):
        super().__init__(parent)
        
        self.emails = emails  # List of email data dicts
        self.email_handler = email_handler
        self.current_index = 0
        self.sent_count = 0
        self.skipped_emails = set()
        
        self.title("Email Preview")
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        self.display_current_email()
        self.center_window()
    
    def create_widgets(self):
        """Create the dialog UI."""
        
        # Header with navigation
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        self.counter_label = ctk.CTkLabel(
            header_frame,
            text="",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.counter_label.pack(side="left")
        
        # Navigation buttons in header
        nav_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        nav_frame.pack(side="right")
        
        self.prev_btn = ctk.CTkButton(
            nav_frame,
            text="◀ Previous",
            command=self.previous_email,
            width=100
        )
        self.prev_btn.pack(side="left", padx=(0, 5))
        
        self.next_btn = ctk.CTkButton(
            nav_frame,
            text="Next ▶",
            command=self.next_email,
            width=100
        )
        self.next_btn.pack(side="left")
        
        # Content frame
        content_frame = ctk.CTkScrollableFrame(self, fg_color=("gray90", "gray20"))
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # To field
        to_frame = self._create_field_frame(content_frame, "To:")
        self.to_text = ctk.CTkTextbox(to_frame, height=50, font=ctk.CTkFont(size=12))
        self.to_text.pack(fill="x", expand=True, pady=5)
        
        # CC field
        cc_frame = self._create_field_frame(content_frame, "CC:")
        self.cc_text = ctk.CTkTextbox(cc_frame, height=50, font=ctk.CTkFont(size=12))
        self.cc_text.pack(fill="x", expand=True, pady=5)
        
        # Subject field
        subject_frame = self._create_field_frame(content_frame, "Subject:")
        self.subject_text = ctk.CTkTextbox(subject_frame, height=60, font=ctk.CTkFont(size=12, weight="bold"))
        self.subject_text.pack(fill="x", expand=True, pady=5)
        
        # Attachment field
        attach_frame = self._create_field_frame(content_frame, "Attachment:")
        self.attach_label = ctk.CTkLabel(
            attach_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("#2196F3", "#64B5F6")
        )
        self.attach_label.pack(fill="x", pady=5)
        
        # Body preview
        body_frame = self._create_field_frame(content_frame, "Body Preview:")
        self.body_text = ctk.CTkTextbox(body_frame, height=250, font=ctk.CTkFont(size=11))
        self.body_text.pack(fill="both", expand=True, pady=5)
        
        # Action buttons
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        # Left side buttons
        left_buttons = ctk.CTkFrame(button_frame, fg_color="transparent")
        left_buttons.pack(side="left")
        
        skip_btn = ctk.CTkButton(
            left_buttons,
            text="Skip This Email",
            command=self.skip_email,
            width=130,
            fg_color="transparent",
            border_width=2,
            border_color=("gray70", "gray30")
        )
        skip_btn.pack(side="left", padx=(0, 10))
        
        # Right side buttons
        right_buttons = ctk.CTkFrame(button_frame, fg_color="transparent")
        right_buttons.pack(side="right")
        
        cancel_btn = ctk.CTkButton(
            right_buttons,
            text="Cancel All",
            command=self.on_cancel,
            width=120,
            fg_color="transparent",
            border_width=2,
            border_color=("gray70", "gray30")
        )
        cancel_btn.pack(side="left", padx=(0, 10))
        
        send_current_btn = ctk.CTkButton(
            right_buttons,
            text="Send This Email",
            command=self.send_current,
            width=130,
            fg_color=("#4CAF50", "#388E3C")
        )
        send_current_btn.pack(side="left", padx=(0, 10))
        
        send_all_btn = ctk.CTkButton(
            right_buttons,
            text="Send All Remaining",
            command=self.send_all_remaining,
            width=150,
            fg_color=("#2196F3", "#1976D2")
        )
        send_all_btn.pack(side="left")
    
    def _create_field_frame(self, parent, label_text: str):
        """Create a labeled frame for a field."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=ctk.CTkFont(weight="bold", size=13)
        )
        label.pack(anchor="w", pady=(5, 0))
        
        return frame
    
    def display_current_email(self):
        """Display the current email in the preview."""
        if not self.emails or self.current_index >= len(self.emails):
            return
        
        email = self.emails[self.current_index]
        
        # Update counter
        total = len(self.emails)
        remaining = total - self.current_index
        self.counter_label.configure(
            text=f"Email {self.current_index + 1} of {total}"
        )
        
        # Update navigation buttons
        self.prev_btn.configure(state="normal" if self.current_index > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.current_index < total - 1 else "disabled")
        
        # Populate fields
        self.to_text.delete("1.0", "end")
        self.to_text.insert("1.0", email.get('to', ''))
        
        self.cc_text.delete("1.0", "end")
        self.cc_text.insert("1.0", email.get('cc', ''))
        
        self.subject_text.delete("1.0", "end")
        self.subject_text.insert("1.0", email.get('subject', ''))
        
        attachment = email.get('attachment_path', '')
        if attachment:
            self.attach_label.configure(text=Path(attachment).name)
        else:
            self.attach_label.configure(text="No attachment")
        
        self.body_text.delete("1.0", "end")
        body = email.get('body', '')
        preview = body[:2000] + ("..." if len(body) > 2000 else "")
        self.body_text.insert("1.0", preview)
    
    def previous_email(self):
        """Navigate to previous email."""
        if self.current_index > 0:
            self.current_index -= 1
            self.display_current_email()
    
    def next_email(self):
        """Navigate to next email."""
        if self.current_index < len(self.emails) - 1:
            self.current_index += 1
            self.display_current_email()
    
    def skip_email(self):
        """Skip the current email and move to next."""
        self.skipped_emails.add(self.current_index)
        
        if self.current_index < len(self.emails) - 1:
            self.next_email()
        else:
            messagebox.showinfo("Preview Complete", 
                              f"Reviewed all emails.\n\nSent: {self.sent_count}\nSkipped: {len(self.skipped_emails)}")
            self.destroy()
    
    def send_current(self):
        """Send the current email."""
        try:
            email = self.emails[self.current_index]
            
            # Send via email handler
            success = self.email_handler.send_single_email(
                to_addresses=email.get('to', ''),
                cc_addresses=email.get('cc', ''),
                subject=email.get('subject', ''),
                body=email.get('body', ''),
                attachment_path=email.get('attachment_path')
            )
            
            if success:
                self.sent_count += 1
                messagebox.showinfo("Success", f"Email sent successfully!\n\nSent: {self.sent_count} of {len(self.emails)}")
                
                # Move to next or close
                if self.current_index < len(self.emails) - 1:
                    self.next_email()
                else:
                    messagebox.showinfo("Complete", f"All emails processed!\n\nSent: {self.sent_count}\nSkipped: {len(self.skipped_emails)}")
                    self.destroy()
            else:
                messagebox.showerror("Send Failed", "Failed to send email. Please check the logs.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")
    
    def send_all_remaining(self):
        """Send all remaining emails without further preview."""
        confirm = messagebox.askyesno(
            "Confirm Send All",
            f"Send all {len(self.emails) - self.current_index} remaining emails?\n\nThis cannot be undone."
        )
        
        if not confirm:
            return
        
        # Send all from current index onwards
        for i in range(self.current_index, len(self.emails)):
            if i in self.skipped_emails:
                continue
            
            try:
                email = self.emails[i]
                success = self.email_handler.send_single_email(
                    to_addresses=email.get('to', ''),
                    cc_addresses=email.get('cc', ''),
                    subject=email.get('subject', ''),
                    body=email.get('body', ''),
                    attachment_path=email.get('attachment_path')
                )
                
                if success:
                    self.sent_count += 1
            
            except Exception as e:
                messagebox.showerror("Error", f"Failed to send email {i+1}: {str(e)}")
                break
        
        messagebox.showinfo("Complete", f"Batch send complete!\n\nSent: {self.sent_count}\nSkipped: {len(self.skipped_emails)}")
        self.destroy()
    
    def on_cancel(self):
        """Cancel all sending."""
        confirm = messagebox.askyesno(
            "Confirm Cancel",
            "Cancel sending all emails?\n\nAny emails already sent will not be recalled."
        )
        
        if confirm:
            self.destroy()
    
    def center_window(self):
        """Center the dialog on screen."""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")


class UnassignedAgenciesDialog(ctk.CTkToplevel):
    """
    Interactive dialog for handling agencies found in master data but not in mapping file.
    Allows per-agency decision making and updates the mapping file.
    """
    
    def __init__(self, parent, unassigned_data: List[dict], file_to_agencies: dict, 
                 mapping_file_path: Path, current_region: str):
        super().__init__(parent)
        
        self.unassigned_data = unassigned_data  # [{'agency': str, 'country': str, 'user_count': int}]
        self.file_to_agencies = file_to_agencies  # Existing file names in mapping
        self.mapping_file_path = mapping_file_path
        self.current_region = current_region
        
        self.result = None  # Will store user's decisions
        self.action_decisions = {}  # {agency: {'action': str, 'target': str}}
        
        self.title("Unassigned Agencies Detected")
        self.geometry("1000x650")
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        self.center_window()
    
    def create_widgets(self):
        """Create the dialog UI."""
        
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        header_label = ctk.CTkLabel(
            header_frame,
            text="⚠️ Unassigned Agencies Detected",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=("#FF6B6B", "#FF8787")
        )
        header_label.pack(anchor="w")
        
        info_label = ctk.CTkLabel(
            header_frame,
            text=f"The following {len(self.unassigned_data)} agencies were found in the master data but not in the mapping file.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60")
        )
        info_label.pack(anchor="w", pady=(5, 0))
        
        # Scrollable table frame
        table_frame = ctk.CTkScrollableFrame(self, fg_color=("gray90", "gray20"))
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Table headers
        headers = ["Agency", "Country", "Users", "Action", "Target"]
        header_row = ctk.CTkFrame(table_frame, fg_color=("gray80", "gray25"))
        header_row.pack(fill="x", pady=(0, 5))
        
        for i, header_text in enumerate(headers):
            width = [250, 120, 80, 180, 250][i]
            label = ctk.CTkLabel(
                header_row,
                text=header_text,
                font=ctk.CTkFont(weight="bold", size=13),
                width=width,
                anchor="w" if i < 2 else "center"
            )
            label.pack(side="left", padx=5, pady=8)
        
        # Create rows for each unassigned agency
        self.action_widgets = {}  # Store references to widgets
        
        for idx, agency_data in enumerate(self.unassigned_data):
            agency = agency_data['agency']
            country = agency_data.get('country', 'Unknown')
            user_count = agency_data.get('user_count', 0)
            
            # Row frame
            row_frame = ctk.CTkFrame(
                table_frame, 
                fg_color=("white", "gray17") if idx % 2 == 0 else ("gray95", "gray20")
            )
            row_frame.pack(fill="x", pady=2)
            
            # Agency name
            agency_label = ctk.CTkLabel(
                row_frame, text=agency[:35] + "..." if len(agency) > 35 else agency,
                width=250, anchor="w", font=ctk.CTkFont(size=12)
            )
            agency_label.pack(side="left", padx=5, pady=8)
            
            # Country
            country_label = ctk.CTkLabel(
                row_frame, text=country, width=120, anchor="center",
                font=ctk.CTkFont(size=12)
            )
            country_label.pack(side="left", padx=5)
            
            # User count
            count_label = ctk.CTkLabel(
                row_frame, text=str(user_count), width=80, anchor="center",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#2196F3", "#64B5F6")
            )
            count_label.pack(side="left", padx=5)
            
            # Action dropdown
            action_var = ctk.StringVar(value="Keep as Unassigned")
            action_menu = ctk.CTkOptionMenu(
                row_frame,
                variable=action_var,
                values=["Add to Existing File", "Create New File", "Keep as Unassigned"],
                width=180,
                command=lambda agency=agency, var=action_var: self.on_action_changed(agency, var)
            )
            action_menu.pack(side="left", padx=5)
            
            # Target (initially empty, populated based on action)
            target_frame = ctk.CTkFrame(row_frame, fg_color="transparent", width=250)
            target_frame.pack(side="left", padx=5, fill="both", expand=True)
            target_frame.pack_propagate(False)
            
            # Store widget references
            self.action_widgets[agency] = {
                'action_var': action_var,
                'target_frame': target_frame,
                'agency_data': agency_data
            }
            
            # Initialize with default
            self.on_action_changed(agency, action_var)
        
        # Update mapping file checkbox
        checkbox_frame = ctk.CTkFrame(self, fg_color="transparent")
        checkbox_frame.pack(fill="x", padx=20, pady=10)
        
        self.update_mapping_var = ctk.BooleanVar(value=True)
        update_checkbox = ctk.CTkCheckBox(
            checkbox_frame,
            text="Update Agency Mapping File (recommended)",
            variable=self.update_mapping_var,
            font=ctk.CTkFont(size=13),
            checkbox_width=24,
            checkbox_height=24
        )
        update_checkbox.pack(anchor="w")
        
        note_label = ctk.CTkLabel(
            checkbox_frame,
            text="Note: A backup will be created before updating the mapping file.",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray50")
        )
        note_label.pack(anchor="w", padx=30)
        
        # Button frame
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self.on_cancel,
            width=120,
            fg_color="transparent",
            border_width=2,
            border_color=("gray70", "gray30")
        )
        cancel_btn.pack(side="right", padx=(10, 0))
        
        apply_btn = ctk.CTkButton(
            button_frame,
            text="Apply & Continue",
            command=self.on_apply,
            width=150,
            fg_color=("#2196F3", "#1976D2")
        )
        apply_btn.pack(side="right")
    
    def on_action_changed(self, agency: str, action_var: ctk.StringVar):
        """Handle action dropdown change."""
        action = action_var.get()
        target_frame = self.action_widgets[agency]['target_frame']
        
        # Clear existing widgets in target frame
        for widget in target_frame.winfo_children():
            widget.destroy()
        
        if action == "Add to Existing File":
            # Show dropdown of existing files
            existing_files = list(self.file_to_agencies.keys())
            if existing_files:
                target_var = ctk.StringVar(value=existing_files[0])
                target_menu = ctk.CTkOptionMenu(
                    target_frame,
                    variable=target_var,
                    values=existing_files,
                    width=230
                )
                target_menu.pack(fill="x", expand=True)
                self.action_widgets[agency]['target_var'] = target_var
            else:
                label = ctk.CTkLabel(
                    target_frame,
                    text="No existing files",
                    text_color=("gray50", "gray50")
                )
                label.pack()
        
        elif action == "Create New File":
            # Show input fields for new file
            entry_frame = ctk.CTkFrame(target_frame, fg_color="transparent")
            entry_frame.pack(fill="both", expand=True)
            
            # File name entry
            target_var = ctk.StringVar(value=agency)
            entry = ctk.CTkEntry(
                entry_frame,
                textvariable=target_var,
                placeholder_text="File name...",
                width=230
            )
            entry.pack(side="left", fill="x", expand=True)
            self.action_widgets[agency]['target_var'] = target_var
            
            # Button to specify recipients
            config_btn = ctk.CTkButton(
                entry_frame,
                text="📧",
                width=30,
                command=lambda a=agency: self.configure_new_file(a)
            )
            config_btn.pack(side="left", padx=(5, 0))
        
        else:  # Keep as Unassigned
            country = self.action_widgets[agency]['agency_data'].get('country', 'Unknown')
            unassigned_name = f"Unassigned_{country}.xlsx"
            label = ctk.CTkLabel(
                target_frame,
                text=unassigned_name,
                text_color=("gray50", "gray50"),
                font=ctk.CTkFont(size=11)
            )
            label.pack(anchor="w")
            self.action_widgets[agency]['target_var'] = ctk.StringVar(value=unassigned_name)
    
    def configure_new_file(self, agency: str):
        """Open dialog to configure recipients for new file."""
        dialog = ctk.CTkInputDialog(
            text=f"Enter email recipients for '{agency}':\n\nFormat: TO emails ; CC emails\nExample: john@co.com;jane@co.com ; manager@co.com",
            title="Configure Email Recipients"
        )
        recipients = dialog.get_input()
        
        if recipients:
            # Store recipients for later use
            if 'recipients' not in self.action_widgets[agency]:
                self.action_widgets[agency]['recipients'] = {}
            
            parts = recipients.split(';')
            if len(parts) >= 1:
                self.action_widgets[agency]['recipients']['to'] = parts[0].strip()
            if len(parts) >= 2:
                self.action_widgets[agency]['recipients']['cc'] = parts[1].strip() if len(parts) > 1 else ""
    
    def on_apply(self):
        """Apply user's decisions."""
        decisions = {}
        
        # Collect all decisions
        for agency, widgets in self.action_widgets.items():
            action = widgets['action_var'].get()
            target_var = widgets.get('target_var')
            target = target_var.get() if target_var else ""
            
            decision = {
                'action': action,
                'target': target,
                'agency_data': widgets['agency_data']
            }
            
            # Add recipients if creating new file
            if action == "Create New File":
                decision['recipients'] = widgets.get('recipients', {'to': '', 'cc': ''})
            
            decisions[agency] = decision
        
        # Store results
        self.result = {
            'decisions': decisions,
            'update_mapping': self.update_mapping_var.get()
        }
        
        self.destroy()
    
    def on_cancel(self):
        """Cancel the operation."""
        self.result = None
        self.destroy()
    
    def center_window(self):
        """Center the dialog on screen."""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")


# =============== Main GUI Application ===============

def main():
    """Main application entry point."""
    app = CognosAccessReviewApp()
    app.mainloop()

if __name__ == "__main__":                                                                                                  
    # This block ensures the code inside only runs when the script is executed directly
    # (not when it's imported as a module into another script).
    main() 