"""
Cognos Access Review Tool - Enterprise Edition

SOX compliance user access review automation tool.
Manages the full lifecycle: file generation, email distribution, response tracking, and reporting.

Requires: pandas, customtkinter, pywin32 (Windows), tkcalendar, pytz, openpyxl, xlsxwriter
"""

import os
import re
import json
import shutil
import logging
import threading
import tkinter as tk
import tkinter.ttk as ttk
from pathlib import Path
from enum import Enum
from typing import Optional, Tuple, List, Dict, Set, Callable, TypedDict
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from collections import defaultdict

import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

try:
    import pytz
    HAS_PYTZ = True
except ImportError:
    HAS_PYTZ = False

try:
    import win32com.client as win32
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    import hashlib
    HAS_HASHLIB = True
except ImportError:
    HAS_HASHLIB = False


# ============================================================================
# COLOR SCHEME SYSTEM
# ============================================================================

PRIMARY_BLUE_LIGHT = "#007A9B"
PRIMARY_BLUE_DARK = "#00B4E1"
PRIMARY_BLUE_HOVER = "#005575"
PRIMARY_BLUE_ACTIVE = "#00475C"

ACCENT_TEAL_LIGHT = "#17A2B8"
ACCENT_TEAL_DARK = "#20C9A6"
ACCENT_TEAL_HOVER = "#138496"

STATUS_COMPLETE = "#28A745"
STATUS_PENDING = "#FFC107"
STATUS_OVERDUE = "#DC3545"
STATUS_IN_PROGRESS = "#007A9B"

STATUS_COMPLETE_BG = "#D4EDDA"
STATUS_COMPLETE_TEXT = "#155724"
STATUS_PENDING_BG = "#FFF3CD"
STATUS_PENDING_TEXT = "#856404"
STATUS_OVERDUE_BG = "#F8D7DA"
STATUS_OVERDUE_TEXT = "#721C24"
STATUS_INFO_BG = "#D1ECF1"
STATUS_INFO_TEXT = "#0C5460"

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

CURRENT_THEME = "light"
CURRENT_COLORS = LIGHT_COLORS.copy()

FONT_SIZES = {
    "logo": 40, "section_title": 18, "header": 16,
    "normal": 13, "label": 12, "small": 11, "help": 12,
}

SPACING = {"tiny": 4, "small": 8, "normal": 16, "large": 24, "xl": 32}

DIMENSIONS = {
    "button_height": 44, "sidebar_width": 220, "sidebar_collapsed": 60,
    "header_height": 64, "footer_height": 64, "min_button_width": 120, "icon_size": 24,
}

ICONS = {
    "complete": "✓", "pending": "⏳", "overdue": "⚠", "in_progress": "•",
    "error": "✕", "info": "ℹ", "setup": "📋", "process": "🔄",
    "reporting": "📊", "tools": "🔧", "help": "❓", "about": "ℹ",
}


def set_theme(theme_name: str) -> None:
    global CURRENT_THEME, CURRENT_COLORS
    if theme_name.lower() not in ["light", "dark"]:
        raise ValueError("Theme must be 'light' or 'dark'")
    CURRENT_THEME = theme_name.lower()
    CURRENT_COLORS = LIGHT_COLORS.copy() if CURRENT_THEME == "light" else DARK_COLORS.copy()


def get_color(color_name: str, theme: str = None) -> str:
    if theme:
        color_dict = LIGHT_COLORS if theme.lower() == "light" else DARK_COLORS
    else:
        color_dict = CURRENT_COLORS
    return color_dict.get(color_name, "#000000")


def get_all_colors(theme: str = None) -> dict:
    if theme:
        return LIGHT_COLORS.copy() if theme.lower() == "light" else DARK_COLORS.copy()
    return CURRENT_COLORS.copy()


# ============================================================================
# CONSTANTS & DATA MODELS
# ============================================================================

class ColumnNames:
    AGENCY = "Agency"
    DOMAIN = "Domain"
    USER_NAME = "User Name"
    EMAIL = "Email"
    EMAIL_ADDRESS = "EmailAddress"
    USERNAME = "UserName"
    COUNTRY = "Country"
    PROFILE_NAME = "ProfileName"
    REGION = "Region"
    FOLDER = "Folder"
    SUBFOLDER = "SubFolder"

    SOURCE_FILE_NAME = "source_file_name"
    FILE_NAME = "File name"
    AGENCY_ID = "agency_id"
    RECIPIENTS_TO = "recipients_to"
    RECIPIENTS_CC = "recipients_cc"

    SENT_DATE = "Sent Email Date"
    RESPONSE_DATE = "Response Received Date"
    STATUS = "Status"
    COMMENTS = "Comments"
    TO = "To"
    CC = "CC"

    REVIEW_ACTION = "Review Action"
    REVIEW_COMMENTS = "Comments"


class SheetNames:
    ALL_USERS = "All Users"
    USER_ACCESS_LIST = "User Access List"
    USER_ACCESS_SUMMARY = "User Access Summary"
    UNASSIGNED = "Unassigned"


class FileNames:
    UNASSIGNED_FILE = "Unassigned.xlsx"
    AUDIT_LOG_TEMPLATE = "Audit_CognosAccessReview_{period}.xlsx"
    BACKUP_DIR = "backups"
    LOG_FILE = "cognos_review.log"
    CONFIG_FILE = "config.json"
    EMAIL_TEMPLATE = "email_template.txt"
    REGION_PROFILES = "region_profiles.json"


class EmailMode(Enum):
    PREVIEW = "Preview"
    DIRECT = "Direct"
    SCHEDULE = "Schedule"


class AuditStatus(Enum):
    NOT_SENT = "Not Sent"
    SENT = "Sent"
    RESPONDED = "Responded"
    OVERDUE = "Overdue"


class ValidationSeverity(Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


# ============================================================================
# TYPED DICTS
# ============================================================================

class AppConfig(TypedDict, total=False):
    review_period: str
    deadline: str
    company_name: str
    sender_name: str
    sender_title: str
    email_subject_prefix: str
    auto_scan: bool
    default_email_mode: str
    current_region: Optional[str]
    email_font_family: str
    email_font_size: str
    email_font_color: str
    email_format: str
    email_send_delay: float


class RegionProfileDict(TypedDict):
    region_name: str
    master_file: str
    mapping_file: str
    output_dir: str
    email_manifest: str
    description: str


class EmailChangeRecord(TypedDict):
    agency: str
    field: str
    old_value: str
    new_value: str
    timestamp: str


class ValidationIssue(TypedDict):
    severity: str
    message: str
    details: Optional[str]


class ValidationResults(TypedDict):
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
    to: str
    cc: str


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class AgencyMapping:
    file_name: str
    agency_id: str

    def __post_init__(self):
        self.file_name = self.file_name.strip()
        self.agency_id = self.agency_id.strip()

    @property
    def excel_file_name(self) -> str:
        return f"{self.file_name}.xlsx"


@dataclass
class CombinedMapping:
    source_file_name: str
    agency_id: str
    recipients_to: str
    recipients_cc: str
    source_tab: str = ""

    def __post_init__(self):
        self.source_file_name = self.source_file_name.strip()
        self.agency_id = self.agency_id.strip()
        self.recipients_to = self.recipients_to.strip() if self.recipients_to else ""
        self.recipients_cc = self.recipients_cc.strip() if self.recipients_cc else ""
        self.source_tab = self.source_tab.strip()

    @property
    def excel_file_name(self) -> str:
        return f"{self.source_file_name}.xlsx"

    def to_agency_mapping(self) -> AgencyMapping:
        return AgencyMapping(file_name=self.source_file_name, agency_id=self.agency_id)

    def get_email_recipients(self) -> EmailRecipients:
        return {"to": self.recipients_to, "cc": self.recipients_cc}


@dataclass
class AuditLogEntry:
    agency: str
    sent_date: Optional[datetime] = None
    response_date: Optional[datetime] = None
    to: str = ""
    cc: str = ""
    status: str = AuditStatus.NOT_SENT.value
    comments: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {
            ColumnNames.AGENCY: self.agency,
            ColumnNames.SENT_DATE: self.sent_date.strftime("%Y-%m-%d %H:%M") if self.sent_date else "",
            ColumnNames.RESPONSE_DATE: self.response_date.strftime("%Y-%m-%d %H:%M") if self.response_date else "",
            ColumnNames.TO: self.to,
            ColumnNames.CC: self.cc,
            ColumnNames.STATUS: self.status,
            ColumnNames.COMMENTS: self.comments,
        }


@dataclass
class FileGenerationRequest:
    master_file: Path
    agency_map_file: Path
    output_dir: Path
    progress_callback: Optional[Callable] = None

    def validate(self) -> List[str]:
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
    combined_file: Path
    output_dir: Path
    file_names: List[str]
    universal_attachment: Optional[Path] = None
    mode: EmailMode = EmailMode.PREVIEW

    def validate(self) -> List[str]:
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
    total_agencies: int
    sent_count: int
    responded_count: int
    not_sent_count: int
    overdue_count: int
    completion_percentage: float
    days_left: int

    @classmethod
    def calculate(cls, audit_df, deadline_str: str, total_agencies: int) -> 'DashboardMetrics':
        if audit_df.empty:
            sent = responded = 0
        else:
            sent = (audit_df[ColumnNames.STATUS] == AuditStatus.SENT.value).sum()
            responded = (audit_df[ColumnNames.STATUS] == AuditStatus.RESPONDED.value).sum()
        not_sent = total_agencies - sent - responded
        completion = (responded / total_agencies * 100) if total_agencies > 0 else 0.0
        try:
            deadline = datetime.strptime(deadline_str, "%B %d, %Y")
            days_left = (deadline - datetime.now()).days
        except (ValueError, TypeError):
            days_left = -1
        overdue = 0
        if not audit_df.empty and ColumnNames.SENT_DATE in audit_df.columns:
            for _, row in audit_df.iterrows():
                if row[ColumnNames.STATUS] == AuditStatus.SENT.value:
                    sent_date_str = row.get(ColumnNames.SENT_DATE, "")
                    if sent_date_str:
                        try:
                            sent_date = datetime.strptime(str(sent_date_str), "%Y-%m-%d %H:%M")
                            if (datetime.now() - sent_date).days > 7:
                                overdue += 1
                        except (ValueError, TypeError):
                            pass
        return cls(
            total_agencies=total_agencies, sent_count=sent, responded_count=responded,
            not_sent_count=not_sent, overdue_count=overdue,
            completion_percentage=completion, days_left=days_left,
        )


@dataclass
class ExcelFormatting:
    freeze_header: bool = True
    bold_header: bool = True
    header_bg_color: str = "#D9E1F2"
    auto_width: bool = True
    max_column_width: int = 50
    border_header: bool = True


@dataclass
class GenerationReport:
    """Post-generation verification report."""
    total_master_users: int = 0
    total_generated_users: int = 0
    total_files_created: int = 0
    unmapped_users: int = 0
    file_details: List[Dict] = field(default_factory=list)
    discrepancies: List[str] = field(default_factory=list)

    @property
    def is_clean(self) -> bool:
        return len(self.discrepancies) == 0 and self.total_master_users == (self.total_generated_users + self.unmapped_users)


# ============================================================================
# EXCEPTIONS
# ============================================================================

class CognosReviewError(Exception):
    pass

class ValidationError(CognosReviewError):
    pass

class FileProcessingError(CognosReviewError):
    pass

class EmailError(CognosReviewError):
    pass

class ConfigurationError(CognosReviewError):
    pass


# ============================================================================
# LOGGING SETUP
# ============================================================================

def _get_log_path() -> Path:
    """Get a writable path for the log file, falling back gracefully."""
    candidates = [
        Path(os.path.dirname(os.path.abspath(__file__))) / FileNames.LOG_FILE,
        Path.home() / ".cognos_review" / FileNames.LOG_FILE,
        Path(os.environ.get("TEMP", os.environ.get("TMP", "."))) / FileNames.LOG_FILE,
    ]
    for path in candidates:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.touch(exist_ok=True)
            return path
        except (PermissionError, OSError):
            continue
    return Path(FileNames.LOG_FILE)


def setup_logging() -> logging.Logger:
    log_formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    _logger = logging.getLogger("cognos_review")
    _logger.setLevel(logging.INFO)
    if not _logger.handlers:
        try:
            log_path = _get_log_path()
            file_handler = RotatingFileHandler(
                str(log_path), maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8'
            )
            file_handler.setFormatter(log_formatter)
            _logger.addHandler(file_handler)
        except (PermissionError, OSError):
            pass  # Console-only if all file paths fail

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(log_formatter)
        _logger.addHandler(console_handler)
    return _logger

logger = setup_logging()


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\']', '-', str(name))
    name = name.strip().strip('.')
    if len(name) > 255:
        name = name[:255]
    return name if name else "File"


def sanitize_sheet_name(name: str) -> str:
    name = re.sub(r'[\\/*?\[\]]', '', str(name))
    if len(name) > 31:
        name = name[:31]
    return name if name else "Sheet1"


def format_email_list(emails: str) -> List[str]:
    if not emails or not isinstance(emails, str):
        return []
    return [email.strip() for email in emails.split(';') if email.strip()]


def is_valid_email(email: str) -> bool:
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email.strip()))


def extract_emails_from_text(text: str) -> List[str]:
    if not text or not isinstance(text, str):
        return []
    email_pattern = r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'
    potential_emails = re.findall(email_pattern, text, re.IGNORECASE)
    valid_emails = []
    seen_emails = set()
    for email in potential_emails:
        email = email.strip().lower()
        if email not in seen_emails and is_valid_email(email):
            valid_emails.append(email)
            seen_emails.add(email)
    return valid_emails


def parse_email_forwarding_message(message: str) -> Dict[str, List[str]]:
    if not message or not isinstance(message, str):
        return {"to": [], "cc": []}
    message_lower = message.lower()
    all_emails = extract_emails_from_text(message)
    result = {"to": [], "cc": []}
    cc_patterns = [
        r'cc[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'copy[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'also\s+(?:include|cc|copy)[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'please\s+(?:cc|copy)[:\s]+([^.\n]*?)(?=\n|$|\.)',
    ]
    to_patterns = [
        r'(?:forward|send)[^@]*?to[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'please\s+(?:send|forward)[^@]*?(?:to|email)[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'send\s+(?:email\s+)?to[:\s]+([^.\n]*?)(?=\n|$|\.)',
        r'forward\s+(?:email\s+)?to[:\s]+([^.\n]*?)(?=\n|$|\.)',
    ]
    cc_emails = set()
    for pattern in cc_patterns:
        matches = re.findall(pattern, message_lower, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            cc_emails.update(extract_emails_from_text(match))
    to_emails = set()
    for pattern in to_patterns:
        matches = re.findall(pattern, message_lower, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            to_emails.update(extract_emails_from_text(match))
    remaining = set(all_emails) - cc_emails - to_emails
    if not to_emails and remaining:
        to_emails = remaining
    result["to"] = sorted(list(to_emails))
    result["cc"] = sorted(list(cc_emails))
    return result


def smart_email_verification(text: str) -> Dict:
    result = {
        "found_emails": [], "valid_emails": [], "invalid_emails": [],
        "parsing": {"to": [], "cc": []}, "suggestions": [], "confidence": "low",
    }
    if not text or not isinstance(text, str):
        return result
    found_emails = extract_emails_from_text(text)
    result["found_emails"] = found_emails
    for email in found_emails:
        if is_valid_email(email):
            result["valid_emails"].append(email)
        else:
            result["invalid_emails"].append(email)
    result["parsing"] = parse_email_forwarding_message(text)
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
        result["suggestions"].append("No email addresses detected in text")
    return result


def smart_agency_match(agency1: str, agency2: str) -> bool:
    if not agency1 or not agency2:
        return False
    norm1 = re.sub(r'\s+', ' ', str(agency1).strip().upper())
    norm2 = re.sub(r'\s+', ' ', str(agency2).strip().upper())
    return norm1 == norm2


# ============================================================================
# CONFIG MANAGER
# ============================================================================

class ConfigManager:
    DEFAULT_CONFIG: AppConfig = {
        "review_period": "Q2 2025",
        "deadline": "June 30, 2025",
        "company_name": "Omnicom Group",
        "sender_name": "Govind Waghmare",
        "sender_title": "Manager, Financial Applications | Analytics",
        "email_subject_prefix": "[ACTION REQUIRED] Cognos Access Review",
        "auto_scan": True,
        "default_email_mode": EmailMode.PREVIEW.value,
        "current_region": None,
        "email_font_family": "Calibri",
        "email_font_size": "11",
        "email_font_color": "#000000",
        "email_format": "html",
        "email_send_delay": 2.0,
    }

    REQUIRED_FIELDS = ["review_period", "deadline", "company_name", "sender_name", "sender_title"]

    def __init__(self, config_path: Optional[Path] = None):
        self.config_path = config_path or Path(FileNames.CONFIG_FILE)
        self._config: AppConfig = {}
        self._loaded = False

    def load(self) -> AppConfig:
        if self._loaded:
            return self._config
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                self._config = {**self.DEFAULT_CONFIG, **loaded_config}
                self.validate()
                logger.info(f"Configuration loaded from {self.config_path}")
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON in config file: {e}")
                raise ConfigurationError(f"Invalid JSON in config file: {e}")
            except ConfigurationError:
                raise
            except Exception as e:
                logger.error(f"Failed to load config: {e}")
                raise ConfigurationError(f"Failed to load config: {e}")
        else:
            logger.warning(f"Config file not found at {self.config_path}, using defaults")
            self._config = self.DEFAULT_CONFIG.copy()
        self._loaded = True
        return self._config

    def save(self) -> None:
        if not self._config:
            self._config = self.DEFAULT_CONFIG.copy()
        try:
            self.validate()
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2, ensure_ascii=False)
            logger.info(f"Configuration saved to {self.config_path}")
        except Exception as e:
            logger.error(f"Failed to save config: {e}")
            raise ConfigurationError(f"Failed to save config: {e}")

    def validate(self) -> None:
        errors = []
        for fld in self.REQUIRED_FIELDS:
            if fld not in self._config or not self._config[fld]:
                errors.append(f"Missing required field: {fld}")
        if "default_email_mode" in self._config:
            mode = self._config["default_email_mode"]
            valid_modes = [e.value for e in EmailMode]
            if mode not in valid_modes:
                errors.append(f"Invalid email mode '{mode}'. Must be one of: {valid_modes}")
        if errors:
            error_msg = "Configuration validation failed:\n" + "\n".join(f"  - {e}" for e in errors)
            raise ConfigurationError(error_msg)

    def get(self, key: str, default=None):
        if not self._loaded:
            self.load()
        return self._config.get(key, default)

    def update(self, updates: Dict) -> None:
        if not self._loaded:
            self.load()
        self._config.update(updates)
        self.validate()
        logger.info(f"Configuration updated with {len(updates)} changes")

    def get_all(self) -> AppConfig:
        if not self._loaded:
            self.load()
        return self._config.copy()

    def reset_to_defaults(self) -> None:
        self._config = self.DEFAULT_CONFIG.copy()
        self._loaded = True

    def get_audit_file_name(self, region_name: Optional[str] = None) -> str:
        if not self._loaded:
            self.load()
        period = self._config.get("review_period", "Q2_2025")
        period_safe = period.replace(' ', '_')
        if region_name:
            region_safe = region_name.replace(' ', '_')
            return f"Audit_CognosAccessReview_{period_safe}_{region_safe}.xlsx"
        return FileNames.AUDIT_LOG_TEMPLATE.format(period=period_safe)

    def set_current_region(self, region_name: Optional[str]) -> None:
        if not self._loaded:
            self.load()
        self._config["current_region"] = region_name
        self.save()

    def get_current_region(self) -> Optional[str]:
        if not self._loaded:
            self.load()
        return self._config.get("current_region")


def load_email_template(template_path: Optional[Path] = None) -> str:
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
            logger.warning(f"Failed to load email template: {e}, using default")
    return DEFAULT_TEMPLATE


def save_email_template(content: str, template_path: Optional[Path] = None) -> None:
    template_path = template_path or Path(FileNames.EMAIL_TEMPLATE)
    try:
        with open(template_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logger.info(f"Email template saved to {template_path}")
    except Exception as e:
        raise ConfigurationError(f"Failed to save email template: {e}")


_global_config_manager: Optional[ConfigManager] = None

def get_config_manager() -> ConfigManager:
    global _global_config_manager
    if _global_config_manager is None:
        _global_config_manager = ConfigManager()
    return _global_config_manager


# ============================================================================
# APP STATE - Consolidated global state
# ============================================================================

class AppState:
    """Singleton holding all application-wide state."""
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        self.config_manager = get_config_manager()
        self._config = self.config_manager.load()
        self._file_hashes: Dict[str, str] = {}

    @property
    def review_period(self) -> str:
        return self._config.get("review_period", "Q2 2025")

    @property
    def review_deadline(self) -> str:
        return self._config.get("deadline", "June 30, 2025")

    @property
    def audit_file_name(self) -> str:
        return self.config_manager.get_audit_file_name()

    def reload_config(self):
        self.config_manager._loaded = False
        self._config = self.config_manager.load()

    def track_file_hash(self, filepath: str) -> str:
        """Store hash of a file for change detection."""
        try:
            content = Path(filepath).read_bytes()
            file_hash = hashlib.md5(content).hexdigest()
            self._file_hashes[filepath] = file_hash
            return file_hash
        except Exception:
            return ""

    def has_file_changed(self, filepath: str) -> bool:
        """Check if a tracked file has been modified since last load."""
        if filepath not in self._file_hashes:
            return False
        try:
            content = Path(filepath).read_bytes()
            current_hash = hashlib.md5(content).hexdigest()
            return current_hash != self._file_hashes[filepath]
        except Exception:
            return True


# ============================================================================
# AUDIT LOGGER
# ============================================================================

class AuditLogger:
    def __init__(self, output_dir: Optional[Path] = None, audit_file_name: Optional[str] = None):
        self.output_dir = Path(output_dir) if output_dir else None
        self.audit_file_name = audit_file_name or "Audit_CognosAccessReview.xlsx"
        self.audit_file_path = (self.output_dir / self.audit_file_name) if self.output_dir else None
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")
        self._df: Optional[pd.DataFrame] = None

    def _get_columns(self) -> List[str]:
        return [
            ColumnNames.AGENCY, ColumnNames.SENT_DATE, ColumnNames.RESPONSE_DATE,
            ColumnNames.TO, ColumnNames.CC, ColumnNames.STATUS, ColumnNames.COMMENTS,
        ]

    def load(self) -> pd.DataFrame:
        if self._df is not None:
            return self._df
        if self.audit_file_path and self.audit_file_path.exists():
            try:
                self._df = pd.read_excel(self.audit_file_path)
                for col in self._get_columns():
                    if col not in self._df.columns:
                        self._df[col] = ""
                self._logger.info(f"Loaded audit log: {self.audit_file_path}")
            except Exception as e:
                self._logger.error(f"Failed to load audit log: {e}")
                self._df = pd.DataFrame(columns=self._get_columns())
        else:
            self._df = pd.DataFrame(columns=self._get_columns())
        return self._df

    def save(self) -> bool:
        if self._df is None:
            return False
        try:
            if self.audit_file_path and self.audit_file_path.exists():
                self.create_backup()
            if self.output_dir:
                self.output_dir.mkdir(parents=True, exist_ok=True)
            if self.audit_file_path:
                self._df.to_excel(self.audit_file_path, index=False)
                self._logger.info(f"Saved audit log: {self.audit_file_path}")
            return True
        except Exception as e:
            self._logger.error(f"Failed to save audit log: {e}")
            return False

    def create_backup(self) -> Optional[Path]:
        if not self.audit_file_path or not self.audit_file_path.exists():
            return None
        try:
            backup_dir = self.output_dir / FileNames.BACKUP_DIR
            backup_dir.mkdir(exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = backup_dir / f"{timestamp}_{self.audit_file_name}"
            shutil.copy2(self.audit_file_path, backup_path)
            return backup_path
        except Exception as e:
            self._logger.error(f"Failed to create backup: {e}")
            return None

    def initialize_log(self, agencies: List[str], preserve_existing: bool = True) -> None:
        df = self.load()
        if preserve_existing and not df.empty:
            existing = set(df[ColumnNames.AGENCY].str.upper())
            new_agencies = [a for a in agencies if a.upper() not in existing]
            if new_agencies:
                new_entries = pd.DataFrame({
                    ColumnNames.AGENCY: new_agencies, ColumnNames.SENT_DATE: "",
                    ColumnNames.RESPONSE_DATE: "", ColumnNames.TO: "", ColumnNames.CC: "",
                    ColumnNames.STATUS: AuditStatus.NOT_SENT.value, ColumnNames.COMMENTS: "",
                })
                self._df = pd.concat([df, new_entries], ignore_index=True)
        else:
            self._df = pd.DataFrame({
                ColumnNames.AGENCY: agencies, ColumnNames.SENT_DATE: "",
                ColumnNames.RESPONSE_DATE: "", ColumnNames.TO: "", ColumnNames.CC: "",
                ColumnNames.STATUS: AuditStatus.NOT_SENT.value, ColumnNames.COMMENTS: "",
            })

    def mark_sent(self, agency: str, to: str = "", cc: str = "", comments: str = "") -> None:
        df = self.load()
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        if mask.any():
            idx = df[mask].index[0]
            df.loc[idx, ColumnNames.SENT_DATE] = now_str
            df.loc[idx, ColumnNames.TO] = to
            df.loc[idx, ColumnNames.CC] = cc
            df.loc[idx, ColumnNames.STATUS] = AuditStatus.SENT.value
            if comments:
                df.loc[idx, ColumnNames.COMMENTS] = comments
        else:
            new_row = pd.DataFrame([{
                ColumnNames.AGENCY: agency, ColumnNames.SENT_DATE: now_str,
                ColumnNames.RESPONSE_DATE: "", ColumnNames.TO: to, ColumnNames.CC: cc,
                ColumnNames.STATUS: AuditStatus.SENT.value, ColumnNames.COMMENTS: comments,
            }])
            self._df = pd.concat([df, new_row], ignore_index=True)

    def mark_responded(self, agency: str, comments: str = "") -> None:
        df = self.load()
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        if mask.any():
            idx = df[mask].index[0]
            df.loc[idx, ColumnNames.RESPONSE_DATE] = datetime.now().strftime("%Y-%m-%d %H:%M")
            df.loc[idx, ColumnNames.STATUS] = AuditStatus.RESPONDED.value
            if comments:
                existing = str(df.loc[idx, ColumnNames.COMMENTS])
                df.loc[idx, ColumnNames.COMMENTS] = f"{existing}; {comments}" if existing and existing != "nan" else comments

    def reset_agency(self, agency: str) -> None:
        """Reset an agency status to allow re-sending."""
        df = self.load()
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        if mask.any():
            idx = df[mask].index[0]
            df.loc[idx, ColumnNames.SENT_DATE] = ""
            df.loc[idx, ColumnNames.RESPONSE_DATE] = ""
            df.loc[idx, ColumnNames.STATUS] = AuditStatus.NOT_SENT.value
            df.loc[idx, ColumnNames.COMMENTS] = f"Reset on {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            self._logger.info(f"Reset agency status: {agency}")

    def get_status(self, agency: str) -> Optional[str]:
        df = self.load()
        mask = df[ColumnNames.AGENCY].str.upper() == agency.upper()
        if mask.any():
            return df.loc[mask.idxmax(), ColumnNames.STATUS]
        return None

    def get_metrics(self, deadline_str: str, total_agencies: int) -> DashboardMetrics:
        return DashboardMetrics.calculate(self.load(), deadline_str, total_agencies)

    def get_summary_stats(self) -> Dict[str, int]:
        df = self.load()
        if df.empty:
            return {"total": 0, "sent": 0, "responded": 0, "not_sent": 0, "overdue": 0}
        return {
            "total": len(df),
            "sent": (df[ColumnNames.STATUS] == AuditStatus.SENT.value).sum(),
            "responded": (df[ColumnNames.STATUS] == AuditStatus.RESPONDED.value).sum(),
            "not_sent": (df[ColumnNames.STATUS] == AuditStatus.NOT_SENT.value).sum(),
            "overdue": (df[ColumnNames.STATUS] == AuditStatus.OVERDUE.value).sum(),
        }


# ============================================================================
# REGION MANAGER
# ============================================================================

@dataclass
class RegionProfile:
    region_name: str
    master_file: str = ""
    mapping_file: str = ""
    output_dir: str = ""
    email_manifest: str = ""
    description: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'RegionProfile':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


class RegionManager:
    def __init__(self, config_file: Optional[Path] = None):
        self.config_file = config_file or Path(FileNames.REGION_PROFILES)
        self._profiles: Dict[str, RegionProfile] = {}
        self._current_profile: Optional[str] = None
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")
        self.load_profiles()

    def load_profiles(self) -> None:
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._profiles = {
                    name: RegionProfile.from_dict(profile_data)
                    for name, profile_data in data.get("profiles", {}).items()
                }
                self._current_profile = data.get("current_profile")
            except Exception as e:
                self._logger.error(f"Failed to load region profiles: {e}")

    def save_profiles(self) -> None:
        try:
            data = {
                "profiles": {name: profile.to_dict() for name, profile in self._profiles.items()},
                "current_profile": self._current_profile,
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self._logger.error(f"Failed to save region profiles: {e}")

    def add_profile(self, profile: RegionProfile) -> bool:
        if profile.region_name in self._profiles:
            return False
        self._profiles[profile.region_name] = profile
        self.save_profiles()
        return True

    def delete_profile(self, region_name: str) -> bool:
        if region_name not in self._profiles:
            return False
        del self._profiles[region_name]
        if self._current_profile == region_name:
            self._current_profile = None
        self.save_profiles()
        return True

    def get_profile(self, region_name: str) -> Optional[RegionProfile]:
        return self._profiles.get(region_name)

    def get_all_region_names(self) -> List[str]:
        return sorted(self._profiles.keys())

    def switch_region(self, region_name: str) -> Optional[RegionProfile]:
        profile = self._profiles.get(region_name)
        if profile:
            self._current_profile = region_name
            self.save_profiles()
        return profile

    def get_current_profile(self) -> Optional[RegionProfile]:
        if self._current_profile:
            return self._profiles.get(self._current_profile)
        return None

    def update_profile(self, region_name: str, profile: RegionProfile) -> bool:
        if region_name not in self._profiles:
            return False
        if region_name != profile.region_name:
            del self._profiles[region_name]
        self._profiles[profile.region_name] = profile
        if self._current_profile == region_name:
            self._current_profile = profile.region_name
        self.save_profiles()
        return True


# ============================================================================
# FILE VALIDATOR
# ============================================================================

class FileValidator:
    def __init__(self):
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")

    def validate_file_exists(self, file_path, file_type: str = "file") -> Tuple[bool, str]:
        if not file_path:
            return False, f"Please select a {file_type}"
        path = Path(file_path) if isinstance(file_path, str) else file_path
        if not path.exists():
            return False, f"{file_type} not found: {path}"
        return True, ""

    def validate_excel_file(self, file_path) -> Tuple[bool, str]:
        try:
            pd.read_excel(file_path, nrows=1)
            return True, ""
        except Exception as e:
            return False, f"Invalid Excel file: {str(e)}"

    def validate_master_file(self, file_path: Path) -> List[ValidationIssue]:
        issues = []
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            if ColumnNames.AGENCY not in df.columns:
                issues.append({"severity": ValidationSeverity.ERROR.value, "message": "Master file missing 'Agency' column", "details": f"Columns: {', '.join(df.columns)}"})
                return issues
            has_username = any(col in df.columns for col in [ColumnNames.USERNAME, ColumnNames.USER_NAME])
            if not has_username:
                issues.append({"severity": ValidationSeverity.ERROR.value, "message": "Master file missing 'UserName' or 'User Name' column", "details": f"Columns: {', '.join(df.columns)}"})
            if df.empty:
                issues.append({"severity": ValidationSeverity.WARNING.value, "message": "Master file is empty", "details": None})
            null_count = df[ColumnNames.AGENCY].isna().sum()
            if null_count > 0:
                issues.append({"severity": ValidationSeverity.INFO.value, "message": f"{null_count} users have no agency assigned", "details": None})
        except Exception as e:
            issues.append({"severity": ValidationSeverity.ERROR.value, "message": "Failed to read master file", "details": str(e)})
        return issues

    def validate_combined_file(self, file_path: Path, file_names: Set[str] = None) -> List[ValidationIssue]:
        issues = []
        try:
            combined_loader = CombinedFileLoader(file_path)
            combined_data = combined_loader.load_all_data()
            if not combined_data:
                issues.append({"severity": ValidationSeverity.ERROR.value, "message": "No data found in combined file", "details": None})
                return issues
            for idx, mapping in enumerate(combined_data):
                to_emails = mapping.recipients_to
                if to_emails:
                    for email in format_email_list(to_emails):
                        if not is_valid_email(email):
                            issues.append({"severity": ValidationSeverity.WARNING.value, "message": f"Invalid To email in row {idx+1}", "details": email})
                else:
                    if mapping.source_file_name:
                        issues.append({"severity": ValidationSeverity.WARNING.value, "message": f"No To addresses for '{mapping.source_file_name}'", "details": None})
        except Exception as e:
            issues.append({"severity": ValidationSeverity.ERROR.value, "message": f"Failed to validate combined file: {e}", "details": None})
        return issues

    def validate_all_files(self, master_file, combined_file, output_dir) -> Tuple[bool, str]:
        master_file = Path(master_file) if isinstance(master_file, str) else master_file
        combined_file = Path(combined_file) if combined_file and isinstance(combined_file, str) else combined_file
        output_dir = Path(output_dir) if isinstance(output_dir, str) else output_dir
        results = []
        for fp, ft in [(master_file, "Master File"), (combined_file, "Combined Mapping File")]:
            ok, msg = self.validate_file_exists(fp, ft)
            results.append(f"{'✅' if ok else '❌'} {msg or f'{ft} found'}")
        if not output_dir or str(output_dir).strip() == "":
            results.append("❌ Please select an Output Folder")
        elif not output_dir.exists():
            results.append(f"❌ Output folder not found: {output_dir}")
        else:
            results.append("✅ Output folder is valid")
        for fp, ft in [(master_file, "Master File"), (combined_file, "Combined File")]:
            if fp and fp.exists():
                ok, msg = self.validate_excel_file(fp)
                results.append(f"{'✅' if ok else '❌'} {ft}: {msg or 'readable'}")
        has_errors = any("❌" in r for r in results)
        status = "\n".join(results)
        if has_errors:
            status += "\n\nPlease fix errors before proceeding."
        else:
            status += "\n\nAll files valid and ready!"
        return not has_errors, status


# ============================================================================
# COMBINED FILE LOADER
# ============================================================================

class CombinedFileLoader:
    def __init__(self, file_path: Path):
        self.file_path = Path(file_path)
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")
        self._combined_data: List[CombinedMapping] = []
        self._loaded = False

    def load_all_tabs(self, selected_tabs: Optional[List[str]] = None) -> List[CombinedMapping]:
        if self._loaded and not selected_tabs:
            return self._combined_data
        try:
            excel_file = pd.ExcelFile(self.file_path)
            all_tabs = excel_file.sheet_names
            tabs_to_process = selected_tabs if selected_tabs else all_tabs
            combined_mappings = []
            for tab_name in tabs_to_process:
                if tab_name not in all_tabs:
                    self._logger.warning(f"Tab '{tab_name}' not found")
                    continue
                df = pd.read_excel(self.file_path, sheet_name=tab_name)
                df.columns = df.columns.str.strip()
                column_mapping = {}
                for col in df.columns:
                    cl = col.lower()
                    if "source_file_name" in cl or "file name" in cl:
                        column_mapping[col] = ColumnNames.SOURCE_FILE_NAME
                    elif "agency_id" in cl or "agency id" in cl:
                        column_mapping[col] = ColumnNames.AGENCY_ID
                    elif "recipients_to" in cl or cl in ["to", "recipients to"]:
                        column_mapping[col] = ColumnNames.RECIPIENTS_TO
                    elif "recipients_cc" in cl or cl in ["cc", "recipients cc"]:
                        column_mapping[col] = ColumnNames.RECIPIENTS_CC
                if column_mapping:
                    df = df.rename(columns=column_mapping)
                required = [ColumnNames.SOURCE_FILE_NAME, ColumnNames.AGENCY_ID]
                if not all(c in df.columns for c in required):
                    self._logger.error(f"Required columns missing in tab '{tab_name}'")
                    continue
                for opt_col in [ColumnNames.RECIPIENTS_TO, ColumnNames.RECIPIENTS_CC]:
                    if opt_col not in df.columns:
                        df[opt_col] = ""
                for _, row in df.iterrows():
                    src = str(row[ColumnNames.SOURCE_FILE_NAME]).strip() if pd.notna(row[ColumnNames.SOURCE_FILE_NAME]) else ""
                    aid = str(row[ColumnNames.AGENCY_ID]).strip() if pd.notna(row[ColumnNames.AGENCY_ID]) else ""
                    if not src or not aid:
                        continue
                    rto = str(row[ColumnNames.RECIPIENTS_TO]) if pd.notna(row[ColumnNames.RECIPIENTS_TO]) else ""
                    rcc = str(row[ColumnNames.RECIPIENTS_CC]) if pd.notna(row[ColumnNames.RECIPIENTS_CC]) else ""
                    if ',' in aid:
                        for single in [a.strip() for a in aid.split(',') if a.strip()]:
                            combined_mappings.append(CombinedMapping(src, single, rto, rcc, tab_name))
                    else:
                        combined_mappings.append(CombinedMapping(src, aid, rto, rcc, tab_name))
            if not selected_tabs:
                self._combined_data = combined_mappings
                self._loaded = True
            return combined_mappings
        except Exception as e:
            raise FileProcessingError(f"Failed to load combined file: {e}")

    def load_all_data(self, selected_tabs: Optional[List[str]] = None) -> List[CombinedMapping]:
        return self.load_all_tabs(selected_tabs)

    def get_agency_mappings(self, selected_tabs: Optional[List[str]] = None) -> List[AgencyMapping]:
        return [m.to_agency_mapping() for m in self.load_all_tabs(selected_tabs)]

    def get_email_mappings(self, selected_tabs: Optional[List[str]] = None) -> Dict[str, EmailRecipients]:
        mappings = {}
        for m in self.load_all_tabs(selected_tabs):
            if m.source_file_name not in mappings:
                mappings[m.source_file_name] = m.get_email_recipients()
        return mappings

    def get_available_tabs(self) -> List[str]:
        try:
            return pd.ExcelFile(self.file_path).sheet_names
        except Exception:
            return []


# ============================================================================
# FILE PROCESSOR
# ============================================================================

class FileProcessor:
    def __init__(self, formatting: Optional[ExcelFormatting] = None):
        self.formatting = formatting or ExcelFormatting()
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")

    def format_worksheet(self, writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
        try:
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            if self.formatting.freeze_header:
                worksheet.freeze_panes(1, 0)
            if self.formatting.bold_header:
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': self.formatting.header_bg_color,
                    'border': 1 if self.formatting.border_header else 0,
                })
                for col_idx, col_name in enumerate(df.columns):
                    worksheet.write(0, col_idx, col_name, header_format)
            if self.formatting.auto_width:
                for col_idx, col_name in enumerate(df.columns):
                    max_len = max(len(str(col_name)), *(len(str(v)) for v in df[col_name].astype(str)))
                    worksheet.set_column(col_idx, col_idx, min(max_len + 2, self.formatting.max_column_width))
        except Exception as e:
            self._logger.warning(f"Failed to format worksheet {sheet_name}: {e}")

    def load_agency_mappings(self, mapping_file: Path, selected_tabs: Optional[List[str]] = None) -> List[AgencyMapping]:
        try:
            return CombinedFileLoader(mapping_file).get_agency_mappings(selected_tabs)
        except Exception as e:
            raise FileProcessingError(f"Failed to load agency mappings: {e}")

    def generate_agency_files(
        self, master_file: Path, combined_map_file: Path, output_dir: Path,
        progress_callback: Optional[Callable] = None, handle_unmapped: str = "prompt",
        selected_tabs: Optional[List[str]] = None, unmapped_callback: Optional[Callable] = None,
    ) -> GenerationReport:
        report = GenerationReport()
        try:
            master_file = Path(master_file)
            combined_map_file = Path(combined_map_file)
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)

            if progress_callback:
                progress_callback(0.1, "Loading master file...")

            df_master = pd.read_excel(master_file)
            df_master.columns = df_master.columns.str.strip()
            report.total_master_users = len(df_master)

            agency_col = None
            for col in df_master.columns:
                if col.lower() == ColumnNames.AGENCY.lower():
                    agency_col = col
                    break
            if not agency_col:
                raise FileProcessingError(f"Master file must have '{ColumnNames.AGENCY}' column")

            if progress_callback:
                progress_callback(0.2, "Loading agency mappings...")

            mappings = self.load_agency_mappings(combined_map_file, selected_tabs)
            file_to_agencies: Dict[str, List[str]] = {}
            for m in mappings:
                file_to_agencies.setdefault(m.file_name, []).append(m.agency_id)

            assigned_indices = set()
            total_files = len(file_to_agencies)

            for file_idx, (file_name, agencies) in enumerate(file_to_agencies.items()):
                if progress_callback:
                    progress_callback(0.3 + (file_idx / max(total_files, 1)) * 0.5, f"Generating {file_name}...")
                user_count = self._generate_single_file(df_master, agency_col, file_name, agencies, output_dir, assigned_indices)
                report.total_files_created += 1
                report.file_details.append({"file_name": file_name, "agencies": agencies, "user_count": user_count})
                report.total_generated_users += user_count

            if progress_callback:
                progress_callback(0.85, "Processing unmapped users...")

            unassigned_indices = set(df_master.index) - assigned_indices
            report.unmapped_users = len(unassigned_indices)

            if unassigned_indices:
                df_unassigned = df_master.loc[list(unassigned_indices)].copy()
                if handle_unmapped == "prompt" and unmapped_callback:
                    unmapped_agencies = df_unassigned[agency_col].fillna("[No Agency]").unique()
                    unassigned_data = []
                    for agency in unmapped_agencies:
                        adf = df_unassigned[df_unassigned[agency_col].fillna("[No Agency]") == agency]
                        country = adf[ColumnNames.COUNTRY].iloc[0] if ColumnNames.COUNTRY in adf.columns else "Unknown"
                        unassigned_data.append({'agency': str(agency), 'country': country, 'user_count': len(adf)})
                    dialog_result = unmapped_callback(unassigned_data, file_to_agencies, combined_map_file)
                    if dialog_result and dialog_result.get('decisions'):
                        for agency, decision in dialog_result['decisions'].items():
                            action = decision['action']
                            target = decision['target']
                            if action == "Add to Existing File" and target in file_to_agencies:
                                file_to_agencies[target].append(agency)
                                self._generate_single_file(df_master, agency_col, target, file_to_agencies[target], output_dir, assigned_indices)
                            elif action == "Create New File":
                                self._generate_single_file(df_master, agency_col, target, [agency], output_dir, assigned_indices)
                            elif action == "Keep as Unassigned":
                                country = decision.get('agency_data', {}).get('country', 'Unknown')
                                self._create_unassigned_file(df_unassigned[df_unassigned[agency_col].fillna("[No Agency]") == agency], output_dir, f"Unassigned_{country}")
                    else:
                        handle_unmapped = "single"
                if handle_unmapped == "individual":
                    self._create_individual_unmapped_files(df_unassigned, agency_col, output_dir)
                elif handle_unmapped == "single":
                    self._create_unassigned_file(df_unassigned, output_dir)

            # Verify counts
            if report.total_master_users != (report.total_generated_users + report.unmapped_users):
                report.discrepancies.append(
                    f"User count mismatch: Master has {report.total_master_users}, "
                    f"generated {report.total_generated_users}, unmapped {report.unmapped_users}"
                )

            if progress_callback:
                progress_callback(1.0, "Complete!")

            self._logger.info(
                f"Generation complete: {report.total_files_created} files, "
                f"{report.total_generated_users} users, {report.unmapped_users} unmapped"
            )
            return report

        except Exception as e:
            self._logger.error(f"File generation failed: {e}")
            raise FileProcessingError(f"File generation failed: {e}")

    def _generate_single_file(self, df_master, agency_col, file_name, agencies, output_dir, assigned_indices) -> int:
        agency_alternatives = "|".join([f"^{re.escape(ag.strip())}$" for ag in agencies])
        agency_pattern = f"(?i)({agency_alternatives})"
        matched = df_master[df_master[agency_col].astype(str).str.match(agency_pattern, na=False)]
        if matched.empty:
            self._logger.warning(f"No users found for {file_name}")
            return 0
        assigned_indices.update(matched.index)
        safe_name = sanitize_filename(file_name)
        output_file = output_dir / (safe_name if safe_name.lower().endswith('.xlsx') else f"{safe_name}.xlsx")
        df_output = matched.copy()
        df_output[ColumnNames.REVIEW_ACTION] = ""
        df_output[ColumnNames.COMMENTS] = ""
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            primary_sheet = SheetNames.USER_ACCESS_LIST
            df_output.to_excel(writer, sheet_name=primary_sheet, index=False)
            self.format_worksheet(writer, primary_sheet, df_output)
            self._create_pivot_summary(writer, df_output, SheetNames.USER_ACCESS_SUMMARY)
            for agency in sorted(agencies, key=str.upper):
                agency_match = df_output[df_output[agency_col].astype(str).str.upper() == agency.upper()]
                if not agency_match.empty:
                    sname = sanitize_sheet_name(agency)
                    agency_match.to_excel(writer, sheet_name=sname, index=False)
                    self.format_worksheet(writer, sname, agency_match)
        self._logger.info(f"Created {output_file} ({len(matched)} users, {len(agencies)} agencies)")
        return len(matched)

    def _create_pivot_summary(self, writer, df, sheet_name):
        try:
            user_col = None
            for col in [ColumnNames.USERNAME, ColumnNames.USER_NAME, "UserName", "User Name"]:
                if col in df.columns:
                    user_col = col
                    break
            if not user_col or ColumnNames.FOLDER not in df.columns or ColumnNames.SUBFOLDER not in df.columns:
                return
            df_pivot = df.drop(columns=[c for c in [ColumnNames.REVIEW_ACTION, ColumnNames.REVIEW_COMMENTS, ColumnNames.COMMENTS] if c in df.columns])
            pivot = pd.pivot_table(df_pivot, index=user_col, columns=[ColumnNames.FOLDER, ColumnNames.SUBFOLDER], aggfunc='size', fill_value=0)
            pivot_display = pivot.reset_index()
            if isinstance(pivot_display.columns, pd.MultiIndex):
                pivot_display.columns = [col[0] if col[1] == '' else f"{col[0]} - {col[1]}" if isinstance(col, tuple) else str(col) for col in pivot_display.columns]
            for col in pivot_display.columns:
                if col != user_col:
                    pivot_display[col] = pivot_display[col].apply(lambda x: 'X' if x > 0 else '')
            pivot_display.to_excel(writer, sheet_name=sheet_name, index=False)
            wb = writer.book
            ws = writer.sheets[sheet_name]
            hfmt = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center', 'text_wrap': True})
            for i, val in enumerate(pivot_display.columns.values):
                ws.write(0, i, str(val), hfmt)
            for i, col in enumerate(pivot_display.columns):
                mx = max(pivot_display[col].astype(str).apply(len).max(), len(str(col)))
                ws.set_column(i, i, min(mx + 2, 50))
            ws.freeze_panes(1, 1)
        except Exception as e:
            self._logger.warning(f"Failed to create pivot summary: {e}")

    def _create_unassigned_file(self, df_unassigned, output_dir, name="Unassigned"):
        output_file = output_dir / f"{name}.xlsx"
        df_out = df_unassigned.copy()
        df_out[ColumnNames.REVIEW_ACTION] = ""
        df_out[ColumnNames.COMMENTS] = ""
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            df_out.to_excel(writer, sheet_name=SheetNames.ALL_USERS, index=False)
            self.format_worksheet(writer, SheetNames.ALL_USERS, df_out)

    def _create_individual_unmapped_files(self, df_unassigned, agency_col, output_dir):
        for agency in df_unassigned[agency_col].fillna("Unassigned").unique():
            astr = str(agency).strip() or "Unassigned"
            adf = df_unassigned[df_unassigned[agency_col].fillna("Unassigned") == agency].copy()
            adf[ColumnNames.REVIEW_ACTION] = ""
            adf[ColumnNames.COMMENTS] = ""
            out = output_dir / f"{sanitize_filename(astr)}.xlsx"
            with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                adf.to_excel(writer, sheet_name=SheetNames.ALL_USERS, index=False)
                self.format_worksheet(writer, SheetNames.ALL_USERS, adf)

    def update_combined_mapping_file(self, mapping_file_path, updates, tab_name):
        try:
            backup_dir = Path(mapping_file_path).parent / "backups"
            backup_dir.mkdir(exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(mapping_file_path, backup_dir / f"{Path(mapping_file_path).stem}_backup_{ts}{Path(mapping_file_path).suffix}")
            xl = pd.ExcelFile(mapping_file_path)
            sheets = {s: pd.read_excel(xl, sheet_name=s) for s in xl.sheet_names}
            if tab_name not in sheets:
                return False
            df = sheets[tab_name].copy()
            for agency, sfn in updates.get('add_to_existing', []):
                mask = df[ColumnNames.SOURCE_FILE_NAME] == sfn
                if mask.any():
                    idx = df[mask].index[0]
                    cur = str(df.loc[idx, ColumnNames.AGENCY_ID])
                    df.loc[idx, ColumnNames.AGENCY_ID] = f"{cur}, {agency}" if cur and cur != 'nan' else agency
            for agency, sfn, to_emails, cc_emails in updates.get('create_new', []):
                df = pd.concat([df, pd.DataFrame([{
                    ColumnNames.SOURCE_FILE_NAME: sfn, ColumnNames.AGENCY_ID: agency,
                    ColumnNames.RECIPIENTS_TO: to_emails, ColumnNames.RECIPIENTS_CC: cc_emails,
                }])], ignore_index=True)
            sheets[tab_name] = df
            with pd.ExcelWriter(mapping_file_path, engine='openpyxl') as writer:
                for sn, sdf in sheets.items():
                    sdf.to_excel(writer, sheet_name=sn, index=False)
            return True
        except Exception as e:
            self._logger.error(f"Failed to update mapping file: {e}")
            return False


# ============================================================================
# EMAIL HANDLER
# ============================================================================

class EmailHandler:
    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")
        self._outlook = None
        if HAS_WIN32:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

    def _get_outlook(self, force_new: bool = False):
        if not HAS_WIN32:
            raise EmailError("pywin32 not installed. Email features require Windows with Outlook.")
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        if force_new:
            self._outlook = None
        if self._outlook is None:
            try:
                try:
                    self._outlook = win32.GetActiveObject("Outlook.Application")
                except Exception:
                    self._outlook = win32.Dispatch("Outlook.Application")
                namespace = self._outlook.GetNamespace("MAPI")
                _ = namespace.GetDefaultFolder(6)
            except Exception as e:
                self._outlook = None
                raise EmailError(f"Failed to connect to Outlook: {e}")
        return self._outlook

    def reset_outlook_connection(self):
        self._outlook = None

    def test_connection(self) -> Tuple[bool, str]:
        try:
            outlook = self._get_outlook()
            mail = outlook.CreateItem(0)
            mail.Subject = "Test Email - Cognos Access Review Tool"
            mail.Body = "This is a test email to verify Outlook connectivity.\nYou can close this without sending."
            mail.To = "test@example.com"
            mail.Display()
            return True, "Outlook connection successful! Test email displayed."
        except Exception as e:
            return False, f"Outlook connection failed: {str(e)}"

    def load_email_manifest(self, combined_file: Path, selected_tabs: Optional[List[str]] = None) -> Dict[str, EmailRecipients]:
        try:
            return CombinedFileLoader(combined_file).get_email_mappings(selected_tabs)
        except Exception as e:
            raise EmailError(f"Failed to load email manifest: {e}")

    def create_email(self, to, cc, subject, body, attachments=None, html_body=None, retry_on_fail=True):
        try:
            if HAS_WIN32:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            outlook = self._get_outlook()
            mail = outlook.CreateItem(0)
            mail.To = to
            mail.CC = cc
            mail.Subject = subject
            if html_body:
                mail.HTMLBody = html_body
            else:
                mail.Body = body
            if attachments:
                for ap in attachments:
                    if ap and Path(ap).exists():
                        mail.Attachments.Add(str(ap))
            return mail
        except Exception as e:
            if retry_on_fail:
                self.reset_outlook_connection()
                return self.create_email(to, cc, subject, body, attachments, html_body, False)
            raise EmailError(f"Failed to create email: {e}")

    def _build_html_body(self, plain_text: str, config: dict) -> str:
        font_family = config.get("email_font_family", "Calibri")
        font_size = config.get("email_font_size", "11")
        font_color = config.get("email_font_color", "#000000")
        escaped = plain_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        paragraphs = escaped.split('\n')
        html_paragraphs = ''.join(f'<p style="margin:0 0 8px 0;">{p if p.strip() else "&nbsp;"}</p>' for p in paragraphs)
        return f"""<html><body style="font-family: {font_family}, sans-serif; font-size: {font_size}pt; color: {font_color};">{html_paragraphs}</body></html>"""

    def send_single_email(self, to_addresses, cc_addresses, subject, body, attachment_path=None, html_body=None) -> Tuple[bool, str]:
        """Returns (success, error_message_or_empty)."""
        try:
            attachments = [attachment_path] if attachment_path else []
            mail = self.create_email(to_addresses, cc_addresses, subject, body, attachments, html_body)
            mail.Send()

            # Verify it appeared in Sent Items (basic confirmation)
            confirmation = self._verify_sent(subject)
            if confirmation:
                self._logger.info(f"Confirmed sent: {subject}")
            else:
                self._logger.warning(f"Send called but could not confirm in Sent Items: {subject}")

            return True, ""
        except Exception as e:
            self._logger.error(f"Failed to send email: {e}")
            return False, str(e)

    def _verify_sent(self, subject: str, timeout_seconds: int = 5) -> bool:
        """Basic check if email appeared in Sent Items."""
        try:
            import time
            time.sleep(min(timeout_seconds, 3))
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            sent_folder = namespace.GetDefaultFolder(5)
            messages = sent_folder.Items
            messages.Sort("[SentOn]", True)
            for msg in messages:
                try:
                    if msg.Subject == subject:
                        return True
                except Exception:
                    continue
                break  # Only check most recent
            return False
        except Exception:
            return False

    def send_emails(
        self, combined_file, output_dir, file_names, mode=EmailMode.PREVIEW,
        universal_attachment=None, progress_callback=None, scheduled_time=None,
        selected_tabs=None, send_delay=None,
    ) -> int:
        try:
            recipients = self.load_email_manifest(Path(combined_file), selected_tabs)
            template = load_email_template()
            config = self.config.get_all()
            review_period = config.get("review_period", "Q2 2025")
            delay = send_delay or config.get("email_send_delay", 2.0)
            use_html = config.get("email_format", "html") == "html"

            if mode == EmailMode.DIRECT:
                self.setup_compliance_folders(review_period)

            processed = 0
            total = len(file_names)

            for idx, file_name in enumerate(file_names):
                if progress_callback:
                    progress_callback((idx / max(total, 1)) * 100, f"Processing {file_name} ({idx+1}/{total})...")

                if file_name not in recipients:
                    self._logger.warning(f"No recipients for: {file_name}")
                    continue
                recipient = recipients[file_name]
                if not recipient["to"]:
                    self._logger.warning(f"No To addresses for: {file_name}")
                    continue

                subject = f"{config.get('email_subject_prefix', '')} {review_period} - {file_name}".strip()
                agency_file = Path(output_dir) / f"{file_name}.xlsx"
                user_count = 0
                agency_list = ""

                if agency_file.exists():
                    try:
                        df_check = pd.read_excel(agency_file, sheet_name=0)
                        user_count = len(df_check)
                        xl = pd.ExcelFile(agency_file)
                        agency_sheets = [s for s in xl.sheet_names if s not in [SheetNames.USER_ACCESS_LIST, SheetNames.ALL_USERS, SheetNames.USER_ACCESS_SUMMARY]]
                        agency_list = ", ".join(agency_sheets[:5])
                        if len(agency_sheets) > 5:
                            agency_list += f" and {len(agency_sheets) - 5} more"
                    except Exception:
                        pass

                recipient_name = ""
                try:
                    recipient_name = recipient["to"].split(';')[0].strip().split('@')[0].replace('.', ' ').title()
                except Exception:
                    pass

                template_vars = {
                    'review_period': review_period, 'deadline': config.get("deadline", "TBD"),
                    'sender_name': config.get("sender_name", ""), 'sender_title': config.get("sender_title", ""),
                    'company_name': config.get("company_name", ""),
                    'QUARTER': review_period, 'DEADLINE': config.get("deadline", "TBD"),
                    'SOURCE_FILE': file_name, 'AGENCY_LIST': agency_list or "your assigned agencies",
                    'USER_COUNT': str(user_count), 'RECIPIENT_NAME': recipient_name or "there",
                }
                body = template
                for key, value in template_vars.items():
                    body = body.replace(f'{{{key}}}', str(value))
                    body = body.replace(f'{{{key.lower()}}}', str(value))

                attachments = []
                if agency_file.exists():
                    attachments.append(agency_file)
                else:
                    self._logger.warning(f"Agency file not found: {agency_file}")
                    continue
                if universal_attachment:
                    attachments.append(Path(universal_attachment))

                html_body = self._build_html_body(body, config) if use_html else None
                mail = self.create_email(recipient["to"], recipient["cc"], subject, body, attachments, html_body)

                if mode == EmailMode.PREVIEW:
                    mail.Display()
                elif mode == EmailMode.DIRECT:
                    mail.Send()
                    self._logger.info(f"Sent: {file_name}")
                    if delay > 0 and idx < total - 1:
                        import time
                        time.sleep(delay)
                elif mode == EmailMode.SCHEDULE:
                    if scheduled_time:
                        mail.DeferredDeliveryTime = scheduled_time
                        mail.Send()
                    else:
                        mail.Save()

                processed += 1

            if progress_callback:
                progress_callback(100, f"Complete! {processed}/{total} processed.")
            return processed
        except Exception as e:
            raise EmailError(f"Email sending failed: {e}")

    def prepare_email_batch(self, combined_file, output_dir, file_names, universal_attachment=None, selected_tabs=None) -> List[dict]:
        try:
            recipients = self.load_email_manifest(Path(combined_file), selected_tabs)
            template = load_email_template()
            config = self.config.get_all()
            review_period = config.get("review_period", "Q2 2025")
            use_html = config.get("email_format", "html") == "html"
            emails = []
            for file_name in file_names:
                if file_name not in recipients or not recipients[file_name]["to"]:
                    continue
                r = recipients[file_name]
                subject = f"{config.get('email_subject_prefix', '')} {review_period} - {file_name}".strip()
                agency_file = Path(output_dir) / f"{file_name}.xlsx"
                user_count = 0
                try:
                    if agency_file.exists():
                        user_count = len(pd.read_excel(agency_file, sheet_name=0))
                except Exception:
                    pass
                recipient_name = ""
                try:
                    recipient_name = r["to"].split(';')[0].strip().split('@')[0].replace('.', ' ').title()
                except Exception:
                    pass
                tvars = {
                    'review_period': review_period, 'deadline': config.get("deadline", "TBD"),
                    'sender_name': config.get("sender_name", ""), 'sender_title': config.get("sender_title", ""),
                    'company_name': config.get("company_name", ""),
                    'SOURCE_FILE': file_name, 'AGENCY_LIST': "", 'USER_COUNT': str(user_count),
                    'RECIPIENT_NAME': recipient_name or "there",
                }
                body = template
                for k, v in tvars.items():
                    body = body.replace(f'{{{k}}}', str(v)).replace(f'{{{k.lower()}}}', str(v))
                html_body = self._build_html_body(body, config) if use_html else None
                emails.append({
                    'to': r["to"], 'cc': r["cc"], 'subject': subject, 'body': body,
                    'html_body': html_body,
                    'attachment_path': str(agency_file) if agency_file.exists() else None,
                    'file_name': file_name,
                })
            return emails
        except Exception as e:
            raise EmailError(f"Failed to prepare email batch: {e}")

    def scan_inbox(self, folder_name="Compliance", subject_keywords=None) -> List[Dict]:
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            try:
                target = inbox.Folders[folder_name]
            except Exception:
                target = inbox
            messages = target.Items
            messages.Sort("[ReceivedTime]", True)
            responses = []
            keywords = subject_keywords or ["cognos", "access review"]
            for message in messages:
                try:
                    subj = str(message.Subject).lower()
                    if any(kw.lower() in subj for kw in keywords):
                        responses.append({
                            "subject": message.Subject, "sender": message.SenderName,
                            "sender_email": message.SenderEmailAddress,
                            "received": message.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
                            "body_preview": str(message.Body)[:200],
                        })
                except Exception:
                    continue
            return responses
        except Exception as e:
            raise EmailError(f"Inbox scan failed: {e}")

    def create_folder(self, folder_name, parent_folder="Inbox") -> bool:
        try:
            outlook = self._get_outlook()
            namespace = outlook.GetNamespace("MAPI")
            if parent_folder.lower() == "inbox":
                parent = namespace.GetDefaultFolder(6)
            elif parent_folder.lower() == "sent items":
                parent = namespace.GetDefaultFolder(5)
            else:
                parent = namespace.GetDefaultFolder(6).Folders[parent_folder]
            try:
                parent.Folders[folder_name]
                return True
            except Exception:
                pass
            parent.Folders.Add(folder_name)
            return True
        except Exception:
            return False

    def setup_compliance_folders(self, review_period=None) -> Dict[str, bool]:
        if not review_period:
            review_period = self.config.get("review_period", "Q2 2025")
        results = {}
        for name in [f"Compliance {review_period} - Sent", f"Compliance {review_period} - Replies"]:
            results[name] = self.create_folder(name)
        return results


# ============================================================================
# REPORT GENERATOR
# ============================================================================

class ReportGenerator:
    def __init__(self, output_dir: Optional[Path] = None):
        self.output_dir = Path(output_dir) if output_dir else Path("reports")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._logger = logging.getLogger(f"cognos_review.{self.__class__.__name__}")

    def generate_compliance_report(self, audit_df, metrics, report_type="summary", output_format="xlsx", review_period="Q4 FY25") -> str:
        if report_type not in ["summary", "detailed", "exceptions"]:
            raise ValueError(f"Invalid report_type: {report_type}")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"SOX_Compliance_Report_{report_type.title()}_{review_period.replace(' ', '_')}_{ts}"
        if output_format == "xlsx":
            return self._generate_excel_report(audit_df, metrics, report_type, self.output_dir / f"{filename}.xlsx", review_period)
        return self._generate_html_report_file(audit_df, metrics, report_type, self.output_dir / f"{filename}.html", review_period)

    def _generate_excel_report(self, audit_df, metrics, report_type, output_path, review_period) -> str:
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            wb = writer.book
            title_fmt = wb.add_format({'bold': True, 'font_size': 16})
            header_fmt = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
            label_fmt = wb.add_format({'bold': True, 'align': 'right'})
            ws = wb.add_worksheet("Executive Summary")
            ws.write(0, 0, f"SOX Compliance Report - {review_period}", title_fmt)
            ws.write(2, 0, "Generated:", label_fmt)
            ws.write(2, 1, datetime.now().strftime("%Y-%m-%d %H:%M"))
            for i, (lbl, val) in enumerate([
                ("Total Agencies", metrics.total_agencies), ("Emails Sent", metrics.sent_count),
                ("Responses Received", metrics.responded_count), ("Not Sent", metrics.not_sent_count),
                ("Overdue", metrics.overdue_count), ("Completion Rate", f"{metrics.completion_percentage:.1f}%"),
                ("Days Until Deadline", metrics.days_left),
            ]):
                ws.write(4 + i, 0, lbl + ":", label_fmt)
                ws.write(4 + i, 1, str(val) if isinstance(val, str) else val)
            ws.set_column(0, 0, 25)
            ws.set_column(1, 1, 20)
            if report_type in ["detailed", "summary"] and not audit_df.empty:
                status_df = audit_df.groupby('Status').size().reset_index(name='Count')
                status_df.to_excel(writer, sheet_name="Status Distribution", index=False)
            if report_type == "detailed" and not audit_df.empty:
                audit_df.to_excel(writer, sheet_name="Full Audit Log", index=False)
            elif report_type == "exceptions" and not audit_df.empty:
                exc = audit_df[(audit_df['Status'] == 'Not Sent') | (audit_df['Status'] == 'Overdue')]
                exc.to_excel(writer, sheet_name="Action Required", index=False)
        return str(output_path)

    def _generate_html_report_file(self, audit_df, metrics, report_type, output_path, review_period) -> str:
        html = f"""<!DOCTYPE html><html><head><title>SOX Report - {review_period}</title>
<style>body{{font-family:Arial,sans-serif;margin:40px}}h1{{color:#2c3e50}}table{{border-collapse:collapse;width:100%;margin-top:20px}}th{{background:#4472C4;color:white;padding:12px}}td{{border:1px solid #ddd;padding:10px}}tr:nth-child(even){{background:#f2f2f2}}.m{{margin:10px 0}}.ml{{font-weight:bold;display:inline-block;width:200px}}</style></head><body>
<h1>SOX Compliance Report - {review_period}</h1>
<p><b>Generated:</b> {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
<h2>Key Metrics</h2>"""
        for lbl, val in [("Total Agencies", metrics.total_agencies), ("Sent", metrics.sent_count), ("Responded", metrics.responded_count), ("Not Sent", metrics.not_sent_count), ("Overdue", metrics.overdue_count), ("Completion", f"{metrics.completion_percentage:.1f}%"), ("Days Left", metrics.days_left)]:
            html += f'<div class="m"><span class="ml">{lbl}:</span>{val}</div>'
        if report_type == "detailed" and not audit_df.empty:
            html += f"<h2>Full Audit Log</h2>{audit_df.to_html(index=False)}"
        elif report_type == "exceptions" and not audit_df.empty:
            exc = audit_df[(audit_df['Status'] == 'Not Sent') | (audit_df['Status'] == 'Overdue')]
            html += f"<h2>Action Required</h2>{exc.to_html(index=False)}"
        html += "</body></html>"
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        return str(output_path)


# ============================================================================
# EMAIL MANIFEST MANAGER
# ============================================================================

@dataclass
class EmailChange:
    agency: str
    field: str
    old_value: str
    new_value: str

    def __str__(self):
        return f"{self.agency} | {self.field}: '{self.old_value}' -> '{self.new_value}'"


class EmailManifestManager:
    EMAIL_PATTERN = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

    def __init__(self, manifest_file: str):
        self.manifest_file = Path(manifest_file)
        self.df = None
        self._hash = None
        self._load()

    def _load(self):
        try:
            self.df = pd.read_excel(self.manifest_file)
            self.df.columns = self.df.columns.str.strip()
            self._hash = self._calculate_hash()
        except Exception as e:
            logger.error(f"Failed to load email manifest: {e}")
            raise

    def _calculate_hash(self) -> str:
        if self.df is None:
            return ""
        return hashlib.md5(self.df.to_json().encode()).hexdigest()

    def validate_emails(self) -> Dict[str, List[str]]:
        invalid = {"To": [], "CC": []}
        if self.df is None:
            return invalid
        for _, row in self.df.iterrows():
            agency = row.get("Agency", "Unknown")
            for field_name in ["To", "CC"]:
                raw = str(row.get(field_name, ""))
                if pd.notna(raw) and raw and raw != "nan":
                    for email in [e.strip() for e in raw.split(";") if e.strip()]:
                        if not re.match(self.EMAIL_PATTERN, email.strip()):
                            invalid[field_name].append(f"{agency}: {email}")
        return invalid

    def detect_changes(self, old_file: str) -> List[EmailChange]:
        changes = []
        try:
            old_df = pd.read_excel(old_file)
            old_df.columns = old_df.columns.str.strip()
        except Exception:
            return changes
        if self.df is None:
            return changes
        old_lookup = {}
        for _, row in old_df.iterrows():
            agency = str(row.get("Agency", "")).strip()
            if agency:
                old_lookup[agency] = {"To": str(row.get("To", "")), "CC": str(row.get("CC", ""))}
        for _, row in self.df.iterrows():
            agency = str(row.get("Agency", "")).strip()
            if not agency:
                continue
            for fld in ["To", "CC"]:
                new_val = str(row.get(fld, ""))
                old_val = old_lookup.get(agency, {}).get(fld, "[NEW]")
                if new_val != old_val:
                    changes.append(EmailChange(agency, fld, old_val, new_val))
        return changes

    def has_file_changed(self) -> bool:
        try:
            current = hashlib.md5(pd.read_excel(self.manifest_file).to_json().encode()).hexdigest()
            return current != self._hash
        except Exception:
            return True

    def save(self):
        self.df.to_excel(self.manifest_file, index=False)
        self._hash = self._calculate_hash()

    def get_agencies(self) -> List[str]:
        if self.df is None:
            return []
        return self.df["Agency"].dropna().str.strip().tolist() if "Agency" in self.df.columns else []

    def update_email(self, agency: str, field: str, new_value: str) -> bool:
        if self.df is None or field not in ["To", "CC"]:
            return False
        mask = self.df["Agency"].str.strip() == agency
        if mask.any():
            self.df.loc[mask, field] = new_value
            return True
        return False

    def create_backup(self, backup_dir="backups") -> Optional[str]:
        try:
            bp = Path(backup_dir)
            bp.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = bp / f"{ts}_{self.manifest_file.name}"
            self.df.to_excel(dest, index=False)
            return str(dest)
        except Exception:
            return None


# ============================================================================
# INITIALIZE APP STATE
# ============================================================================

app_state = AppState()


# ============================================================================
# SMART EMAIL VERIFIER DIALOG
# ============================================================================

class SmartEmailVerifierDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Smart Email Verifier")
        self.geometry("700x550")
        self.transient(parent)
        self.grab_set()
        self._create_widgets()
        self._center()

    def _create_widgets(self):
        main = ctk.CTkFrame(self)
        main.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main, text="Smart Email Verifier", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 10))
        ctk.CTkLabel(main, text="Paste text containing email addresses or forwarding instructions:", font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.text_input = ctk.CTkTextbox(main, height=150)
        self.text_input.pack(fill="x", pady=(5, 10))
        ctk.CTkButton(main, text="Analyze", command=self._analyze, height=36).pack(pady=(0, 10))
        self.result_text = ctk.CTkTextbox(main, height=200, state="disabled")
        self.result_text.pack(fill="both", expand=True)

    def _analyze(self):
        text = self.text_input.get("1.0", "end-1c")
        if not text.strip():
            return
        result = smart_email_verification(text)
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        out = f"Confidence: {result['confidence'].upper()}\n\n"
        out += f"Found Emails ({len(result['found_emails'])}):\n"
        for e in result['found_emails']:
            out += f"  {e}\n"
        out += f"\nValid: {len(result['valid_emails'])} | Invalid: {len(result['invalid_emails'])}\n"
        if result['parsing']['to']:
            out += f"\nTO: {'; '.join(result['parsing']['to'])}\n"
        if result['parsing']['cc']:
            out += f"CC: {'; '.join(result['parsing']['cc'])}\n"
        for s in result['suggestions']:
            out += f"\n> {s}"
        self.result_text.insert("1.0", out)
        self.result_text.configure(state="disabled")

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


# ============================================================================
# MAIN APPLICATION GUI
# ============================================================================

class CognosAccessReviewApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("Cognos Access Review Tool - Enterprise Edition")
        self.geometry("1100x800")

        try:
            self.state('zoomed')
        except Exception:
            self.attributes('-zoomed', True) if os.name != 'nt' else None

        ctk.set_appearance_mode("System")

        self.file_validator = FileValidator()
        self.file_processor = FileProcessor()
        self.email_handler = EmailHandler(app_state.config_manager)
        self.audit_logger = None

        self.agencies = []
        self.audit_df = pd.DataFrame()
        self.agency_checkboxes = {}
        self.agency_folders = {}

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_content()
        self._build_sections()
        self.switch_section("setup")

    # ── Sidebar ─────────────────────────────────────────────────────────

    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=DIMENSIONS["sidebar_width"], corner_radius=0, fg_color=get_color("sidebar"))
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_rowconfigure(10, weight=1)

        ctk.CTkLabel(
            self.sidebar, text="OMNICOM\nCognos Review",
            font=ctk.CTkFont(size=18, weight="bold"), text_color=get_color("primary"),
        ).grid(row=0, column=0, pady=20, sticky="ew", padx=10)

        ctk.CTkFrame(self.sidebar, height=2, fg_color=get_color("border")).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

        self.sidebar_buttons = {}
        self.current_section = None
        sections = [
            ("setup", f"{ICONS['setup']} Setup"),
            ("process", f"{ICONS['process']} Process"),
            ("reporting", f"{ICONS['reporting']} Reporting"),
            ("tools", f"{ICONS['tools']} Tools"),
            ("help", f"{ICONS['help']} Help"),
        ]
        for idx, (sid, label) in enumerate(sections):
            btn = ctk.CTkButton(
                self.sidebar, text=label, font=ctk.CTkFont(size=14, weight="bold"),
                fg_color="transparent", text_color=get_color("text"), hover_color=get_color("secondary"),
                height=45, command=lambda s=sid: self.switch_section(s), corner_radius=10, anchor="w",
            )
            btn.grid(row=idx + 2, column=0, sticky="ew", padx=10, pady=3)
            self.sidebar_buttons[sid] = btn

        ctk.CTkFrame(self.sidebar, height=2, fg_color=get_color("border")).grid(row=8, column=0, sticky="ew", padx=10, pady=10)

        footer = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        footer.grid(row=9, column=0, sticky="ew", padx=10, pady=10)

        ctk.CTkLabel(footer, text="Theme:", font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(0, 5))
        self.theme_var = ctk.StringVar(value="System")
        ctk.CTkOptionMenu(footer, variable=self.theme_var, values=["Light", "Dark", "System"], command=self.change_theme, height=32, fg_color=get_color("primary")).pack(fill="x", pady=(0, 10))
        ctk.CTkButton(footer, text="⚙ Settings", command=self.settings, height=36, fg_color=get_color("primary"), hover_color=get_color("primary_hover")).pack(fill="x", pady=(0, 5))
        ctk.CTkButton(footer, text="About", command=self.show_about, height=36, fg_color=get_color("secondary"), hover_color=get_color("secondary_hover")).pack(fill="x")

    # ── Main Content Frame ──────────────────────────────────────────────

    def _build_main_content(self):
        self.main_content = ctk.CTkFrame(self, corner_radius=0, fg_color=get_color("background"))
        self.main_content.grid(row=0, column=1, sticky="nsew")
        self.main_content.grid_columnconfigure(0, weight=1)
        self.main_content.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.main_content, fg_color=get_color("surface"), corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        ctk.CTkLabel(header, text="Cognos Access Review", font=ctk.CTkFont(size=24, weight="bold"), text_color=get_color("primary")).pack(side="left", padx=20, pady=15)

        self.status_label = ctk.CTkLabel(header, text="Ready", font=ctk.CTkFont(size=12), text_color=get_color("text_secondary"))
        self.status_label.pack(side="right", padx=20)

        self.section_container = ctk.CTkFrame(self.main_content, fg_color="transparent")
        self.section_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        self.section_container.grid_columnconfigure(0, weight=1)
        self.section_container.grid_rowconfigure(0, weight=1)

        pbar_frame = ctk.CTkFrame(self.main_content, fg_color=get_color("surface"), height=40, corner_radius=0)
        pbar_frame.grid(row=2, column=0, sticky="ew")
        pbar_frame.grid_propagate(False)
        self.progress_bar = ctk.CTkProgressBar(pbar_frame, height=6)
        self.progress_bar.pack(fill="x", padx=20, pady=(8, 0))
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(pbar_frame, text="", font=ctk.CTkFont(size=11), text_color=get_color("text_secondary"))
        self.progress_label.pack(padx=20, anchor="w")

    # ── Section Frames ──────────────────────────────────────────────────

    def _build_sections(self):
        self.sections = {}
        self._build_setup_section()
        self._build_process_section()
        self._build_reporting_section()
        self._build_tools_section()
        self._build_help_section()

    def switch_section(self, section_id):
        for sid, btn in self.sidebar_buttons.items():
            if sid == section_id:
                btn.configure(fg_color=get_color("sidebar_active"), text_color=get_color("sidebar_active_text"))
            else:
                btn.configure(fg_color="transparent", text_color=get_color("text"))
        for frame in self.sections.values():
            frame.grid_forget()
        if section_id in self.sections:
            self.sections[section_id].grid(row=0, column=0, sticky="nsew")
        self.current_section = section_id

    # ── SETUP SECTION ───────────────────────────────────────────────────

    def _build_setup_section(self):
        frame = ctk.CTkScrollableFrame(self.section_container, fg_color="transparent")
        self.sections["setup"] = frame

        ctk.CTkLabel(frame, text="File Configuration", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w", pady=(0, 15))

        self.vars = {
            "master": ctk.StringVar(), "combined": ctk.StringVar(),
            "output": ctk.StringVar(), "attach": ctk.StringVar(),
        }
        self.file_status_labels = {}

        file_fields = [
            ("master", "Master User Access File", "Excel files", "*.xlsx"),
            ("combined", "Combined Agency & Email Mapping", "Excel files", "*.xlsx"),
        ]
        for key, label, ftype, fext in file_fields:
            card = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
            card.pack(fill="x", pady=5)
            top = ctk.CTkFrame(card, fg_color="transparent")
            top.pack(fill="x", padx=15, pady=(10, 5))
            ctk.CTkLabel(top, text=label, font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
            status_lbl = ctk.CTkLabel(top, text="", font=ctk.CTkFont(size=12))
            status_lbl.pack(side="right")
            self.file_status_labels[key] = status_lbl
            bottom = ctk.CTkFrame(card, fg_color="transparent")
            bottom.pack(fill="x", padx=15, pady=(0, 10))
            entry = ctk.CTkEntry(bottom, textvariable=self.vars[key], height=36)
            entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkButton(bottom, text="Browse", width=80, height=36, command=lambda k=key, ft=ftype, fe=fext: self._browse_file(k, ft, fe)).pack(side="left")

        # Output directory
        card = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        card.pack(fill="x", pady=5)
        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(top, text="Output Directory", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        self.file_status_labels["output"] = ctk.CTkLabel(top, text="", font=ctk.CTkFont(size=12))
        self.file_status_labels["output"].pack(side="right")
        bottom = ctk.CTkFrame(card, fg_color="transparent")
        bottom.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkEntry(bottom, textvariable=self.vars["output"], height=36).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(bottom, text="Browse", width=80, height=36, command=self._browse_output_dir).pack(side="left")

        # Universal attachment
        card = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        card.pack(fill="x", pady=5)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=15, pady=10)
        ctk.CTkLabel(inner, text="Universal Attachment (optional)", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkEntry(inner, textvariable=self.vars["attach"], height=36, width=300).pack(side="left", padx=(15, 8))
        ctk.CTkButton(inner, text="Browse", width=80, height=36, command=lambda: self._browse_file("attach", "All files", "*.*")).pack(side="left")

        # Region/Tab selection
        ctk.CTkLabel(frame, text="Region / Tab Selection", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(20, 10))
        self.tab_frame = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        self.tab_frame.pack(fill="x", pady=5)
        self.tab_checkboxes = {}
        self.tab_vars = {}
        info = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        info.pack(fill="x", padx=15, pady=10)
        ctk.CTkLabel(info, text="Select region tabs to process (leave all unchecked for all regions):", font=ctk.CTkFont(size=12)).pack(anchor="w")
        self.tab_checkbox_frame = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        self.tab_checkbox_frame.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkButton(self.tab_frame, text="Load Tabs from File", height=32, width=160, command=self._load_tabs_from_file).pack(padx=15, pady=(0, 10), anchor="w")

        # Validate button
        ctk.CTkButton(frame, text="Validate All Files", height=44, font=ctk.CTkFont(size=14, weight="bold"), fg_color=get_color("primary"), hover_color=get_color("primary_hover"), command=self.validate_files).pack(fill="x", pady=(20, 10))

    def _browse_file(self, key, ftype, fext):
        path = filedialog.askopenfilename(filetypes=[(ftype, fext), ("All files", "*.*")])
        if path:
            self.vars[key].set(path)
            app_state.track_file_hash(path)
            self._validate_single_file(key, path)
            if key == "combined":
                self._load_tabs_from_file()

    def _browse_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.vars["output"].set(path)
            self.file_status_labels["output"].configure(text="✅", text_color=get_color("success"))
            self._auto_scan_output(path)

    def _validate_single_file(self, key, path):
        lbl = self.file_status_labels.get(key)
        if not lbl:
            return
        ok, msg = self.file_validator.validate_file_exists(path, key)
        if ok:
            ok2, msg2 = self.file_validator.validate_excel_file(path)
            if ok2:
                lbl.configure(text="✅ Valid", text_color=get_color("success"))
            else:
                lbl.configure(text=f"⚠ {msg2[:30]}", text_color=get_color("warning"))
        else:
            lbl.configure(text="❌ Not found", text_color=get_color("danger"))

    def _load_tabs_from_file(self):
        combined = self.vars["combined"].get()
        if not combined or not Path(combined).exists():
            return
        for w in self.tab_checkbox_frame.winfo_children():
            w.destroy()
        self.tab_vars.clear()
        try:
            tabs = CombinedFileLoader(Path(combined)).get_available_tabs()
            for tab in tabs:
                var = ctk.IntVar(value=0)
                self.tab_vars[tab] = var
                ctk.CTkCheckBox(self.tab_checkbox_frame, text=tab, variable=var).pack(side="left", padx=8, pady=5)
        except Exception as e:
            logger.error(f"Failed to load tabs: {e}")

    def _auto_scan_output(self, path):
        try:
            files = [f.replace(".xlsx", "") for f in os.listdir(path) if f.endswith(".xlsx") and not f.startswith("~$") and not f.startswith("Audit_") and not f.startswith("Unassigned")]
            if files:
                self.agencies = sorted(files)
                self._rebuild_agency_list()
                audit_path = os.path.join(path, app_state.audit_file_name)
                if os.path.exists(audit_path):
                    self.audit_df = pd.read_excel(audit_path)
                self.update_dashboard_metrics()
                self.status_label.configure(text=f"Loaded {len(self.agencies)} agencies")
        except Exception as e:
            logger.error(f"Auto-scan failed: {e}")

    # ── PROCESS SECTION ─────────────────────────────────────────────────

    def _build_process_section(self):
        frame = ctk.CTkScrollableFrame(self.section_container, fg_color="transparent")
        self.sections["process"] = frame

        ctk.CTkLabel(frame, text="Generate & Send", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w", pady=(0, 15))

        # Step 1: Generate
        step1 = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        step1.pack(fill="x", pady=5)
        ctk.CTkLabel(step1, text="Step 1: Generate Agency Files", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=15, pady=(10, 5))
        ctk.CTkLabel(step1, text="Creates per-agency Excel files from the master data and mapping.", font=ctk.CTkFont(size=12), text_color=get_color("text_secondary")).pack(anchor="w", padx=15)
        ctk.CTkButton(step1, text="Generate Files", height=40, fg_color=get_color("primary"), command=self.generate_files).pack(padx=15, pady=10, anchor="w")

        # Step 2: Agency selection and email mode
        step2 = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        step2.pack(fill="x", pady=5)
        ctk.CTkLabel(step2, text="Step 2: Select Agencies & Send", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=15, pady=(10, 5))

        mode_frame = ctk.CTkFrame(step2, fg_color="transparent")
        mode_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(mode_frame, text="Email Mode:", font=ctk.CTkFont(size=12)).pack(side="left")
        self.email_mode = ctk.StringVar(value="Preview")
        for mode in ["Preview", "Direct", "Schedule"]:
            ctk.CTkRadioButton(mode_frame, text=mode, variable=self.email_mode, value=mode).pack(side="left", padx=10)

        # Agency filter and list
        filter_frame = ctk.CTkFrame(step2, fg_color="transparent")
        filter_frame.pack(fill="x", padx=15, pady=5)
        self.agency_filter = ctk.StringVar()
        self.agency_filter.trace_add("write", lambda *_: self.filter_agencies())
        ctk.CTkEntry(filter_frame, textvariable=self.agency_filter, placeholder_text="Filter agencies...", height=36).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(filter_frame, text="Select All", width=90, height=36, command=self.select_all_agencies).pack(side="left", padx=(0, 5))
        ctk.CTkButton(filter_frame, text="Deselect All", width=100, height=36, command=self.deselect_all_agencies).pack(side="left")

        self.agency_scroll = ctk.CTkScrollableFrame(step2, height=200, fg_color=get_color("background"))
        self.agency_scroll.pack(fill="x", padx=15, pady=5)

        btn_frame = ctk.CTkFrame(step2, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=10)
        ctk.CTkButton(btn_frame, text="Send Emails", height=40, fg_color=get_color("primary"), command=self.send).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Preview Batch", height=40, fg_color=get_color("secondary"), command=self.preview_batch).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Add Folder", height=40, fg_color="transparent", border_width=2, border_color=get_color("border"), command=self.add_agency_folder).pack(side="left")

    # ── REPORTING SECTION ───────────────────────────────────────────────

    def _build_reporting_section(self):
        frame = ctk.CTkScrollableFrame(self.section_container, fg_color="transparent")
        self.sections["reporting"] = frame

        ctk.CTkLabel(frame, text="Dashboard & Reporting", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w", pady=(0, 15))

        # Quick stats
        stats_frame = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        stats_frame.pack(fill="x", pady=5)
        stats_inner = ctk.CTkFrame(stats_frame, fg_color="transparent")
        stats_inner.pack(fill="x", padx=15, pady=15)

        self.stat_labels = {}
        for col, (key, label, color) in enumerate([
            ("total", "Total", get_color("primary")), ("sent", "Sent", STATUS_IN_PROGRESS),
            ("responded", "Responded", STATUS_COMPLETE), ("pending", "Pending", STATUS_PENDING),
            ("overdue", "Overdue", STATUS_OVERDUE),
        ]):
            card = ctk.CTkFrame(stats_inner, fg_color="transparent")
            card.pack(side="left", expand=True, fill="x", padx=5)
            self.stat_labels[key] = ctk.CTkLabel(card, text="0", font=ctk.CTkFont(size=28, weight="bold"), text_color=color)
            self.stat_labels[key].pack()
            ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=12), text_color=get_color("text_secondary")).pack()

        # Completion and deadline
        meta_frame = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        meta_frame.pack(fill="x", pady=5)
        meta_inner = ctk.CTkFrame(meta_frame, fg_color="transparent")
        meta_inner.pack(fill="x", padx=15, pady=10)
        self.completion_label = ctk.CTkLabel(meta_inner, text="0.0%", font=ctk.CTkFont(size=20, weight="bold"))
        self.completion_label.pack(side="left", padx=(0, 20))
        ctk.CTkLabel(meta_inner, text="Completion", font=ctk.CTkFont(size=12), text_color=get_color("text_secondary")).pack(side="left", padx=(0, 40))
        self.days_left_label = ctk.CTkLabel(meta_inner, text="-", font=ctk.CTkFont(size=20, weight="bold"))
        self.days_left_label.pack(side="left", padx=(0, 20))
        ctk.CTkLabel(meta_inner, text="Days Left", font=ctk.CTkFont(size=12), text_color=get_color("text_secondary")).pack(side="left")

        # Actions
        action_frame = ctk.CTkFrame(frame, fg_color="transparent")
        action_frame.pack(fill="x", pady=10)
        ctk.CTkButton(action_frame, text="Mark Responded", height=40, fg_color=STATUS_COMPLETE, command=self.mark_responded).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action_frame, text="Re-send Selected", height=40, fg_color=STATUS_PENDING, command=self.resend_selected).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action_frame, text="Open Dashboard", height=40, fg_color=get_color("primary"), command=self.open_dashboard).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action_frame, text="Generate Report", height=40, fg_color=get_color("secondary"), command=self.generate_report).pack(side="left")

    # ── TOOLS SECTION ───────────────────────────────────────────────────

    def _build_tools_section(self):
        frame = ctk.CTkScrollableFrame(self.section_container, fg_color="transparent")
        self.sections["tools"] = frame

        ctk.CTkLabel(frame, text="Tools & Utilities", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w", pady=(0, 15))

        tools = [
            ("Smart Email Verifier", "Extract and validate emails from text", self._open_email_verifier),
            ("Scan Inbox for Replies", "Check Outlook for compliance responses", self.scan_inbox),
            ("Test Email Connection", "Verify Outlook is connected and working", self.test_email_connection),
            ("Open Output Folder", "Open the output directory in file explorer", self.open_output_folder),
            ("Email Template Editor", "Edit the email template and font settings", self.open_template_editor),
            ("Region Manager", "Manage regional profiles and settings", self.open_region_manager),
            ("Fix Unmapped Agencies", "Find and fix agencies missing from mapping", self.fix_errors),
        ]
        for title, desc, cmd in tools:
            card = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
            card.pack(fill="x", pady=3)
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=15, pady=10)
            text_frame = ctk.CTkFrame(inner, fg_color="transparent")
            text_frame.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(text_frame, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
            ctk.CTkLabel(text_frame, text=desc, font=ctk.CTkFont(size=11), text_color=get_color("text_secondary")).pack(anchor="w")
            ctk.CTkButton(inner, text="Open", width=80, height=36, command=cmd).pack(side="right")

    # ── HELP SECTION ────────────────────────────────────────────────────

    def _build_help_section(self):
        frame = ctk.CTkScrollableFrame(self.section_container, fg_color="transparent")
        self.sections["help"] = frame

        ctk.CTkLabel(frame, text="Help & Documentation", font=ctk.CTkFont(size=20, weight="bold")).pack(anchor="w", pady=(0, 15))

        card = ctk.CTkFrame(frame, fg_color=get_color("surface"), corner_radius=8)
        card.pack(fill="x", pady=5)
        help_text = """Quick Start Guide:

1. SETUP: Browse and select your Master File and Combined Mapping File
2. SETUP: Select your Output Directory
3. SETUP: Optionally select region tabs to process, then click Validate
4. PROCESS: Click "Generate Files" to create per-agency Excel files
5. PROCESS: Select agencies and choose email mode (Preview/Direct/Schedule)
6. PROCESS: Click "Send Emails" to distribute review files
7. REPORTING: Track responses, mark agencies as responded
8. REPORTING: Generate SOX compliance reports

File Format Requirements:
- Master File: Must have columns: EmailAddress, UserName, Agency, Country, Region, Folder, SubFolder
- Combined Mapping: Columns: source_file_name, agency_id, recipients_to, recipients_cc
  Each tab represents a region (AMER, APAC, EMEA, etc.)

Keyboard Shortcuts:
- Ctrl+S: Save settings
- Ctrl+R: Refresh dashboard"""
        ctk.CTkLabel(card, text=help_text, font=ctk.CTkFont(size=12), justify="left", wraplength=700).pack(padx=20, pady=20, anchor="w")

        ctk.CTkLabel(frame, text=f"Version 2.0 | Built {datetime.now().strftime('%Y')}", font=ctk.CTkFont(size=11), text_color=get_color("text_secondary")).pack(pady=10)

    # ── Core Actions ────────────────────────────────────────────────────

    def show_progress(self, show: bool):
        if show:
            self.progress_bar.set(0)
        else:
            self.progress_bar.set(0)
            self.progress_label.configure(text="")

    def update_progress(self, value: float, text: str = ""):
        self.progress_bar.set(value if value <= 1 else value / 100)
        self.progress_label.configure(text=text)

    def run_task_in_thread(self, target, args=()):
        thread = threading.Thread(target=target, args=args, daemon=True)
        thread.start()

    def _rebuild_agency_list(self):
        for w in self.agency_scroll.winfo_children():
            w.destroy()
        self.agency_checkboxes.clear()
        for agency in self.agencies:
            var = ctk.IntVar(value=0)
            status = ""
            if not self.audit_df.empty and ColumnNames.AGENCY in self.audit_df.columns:
                mask = self.audit_df[ColumnNames.AGENCY].str.upper() == agency.upper()
                if mask.any():
                    s = self.audit_df.loc[mask.idxmax(), ColumnNames.STATUS]
                    if s == AuditStatus.RESPONDED.value:
                        status = " ✓"
                    elif s == AuditStatus.SENT.value:
                        status = " →"
            cb = ctk.CTkCheckBox(self.agency_scroll, text=f"{agency}{status}", variable=var)
            cb.pack(anchor="w", padx=5, pady=2)
            self.agency_checkboxes[agency] = var

    def filter_agencies(self):
        term = self.agency_filter.get().lower()
        for w in self.agency_scroll.winfo_children():
            w.destroy()
        self.agency_checkboxes.clear()
        for agency in self.agencies:
            if term and term not in agency.lower():
                continue
            var = ctk.IntVar(value=0)
            ctk.CTkCheckBox(self.agency_scroll, text=agency, variable=var).pack(anchor="w", padx=5, pady=2)
            self.agency_checkboxes[agency] = var

    def get_selected_agencies(self) -> List[str]:
        return [a for a, v in self.agency_checkboxes.items() if v.get() == 1]

    def select_all_agencies(self):
        for v in self.agency_checkboxes.values():
            v.set(1)

    def deselect_all_agencies(self):
        for v in self.agency_checkboxes.values():
            v.set(0)

    def get_selected_tabs(self) -> Optional[List[str]]:
        selected = [tab for tab, var in self.tab_vars.items() if var.get() == 1]
        return selected if selected else None

    def update_dashboard_metrics(self):
        total = len(self.agencies)
        responded = 0
        sent = 0
        overdue = 0
        if not self.audit_df.empty and ColumnNames.STATUS in self.audit_df.columns:
            agency_set = set(a.upper() for a in self.agencies)
            for _, row in self.audit_df.iterrows():
                ag = str(row.get(ColumnNames.AGENCY, "")).upper()
                if ag not in agency_set:
                    continue
                st = row.get(ColumnNames.STATUS, "")
                if st == AuditStatus.RESPONDED.value:
                    responded += 1
                elif st == AuditStatus.SENT.value:
                    sent += 1
                    try:
                        sd = pd.to_datetime(row.get(ColumnNames.SENT_DATE, ""))
                        if (pd.Timestamp.now() - sd).days > 7:
                            overdue += 1
                    except Exception:
                        pass
        pending = total - responded - sent

        if hasattr(self, 'stat_labels'):
            self.stat_labels.get("total", ctk.CTkLabel(self)).configure(text=str(total))
            self.stat_labels.get("sent", ctk.CTkLabel(self)).configure(text=str(sent))
            self.stat_labels.get("responded", ctk.CTkLabel(self)).configure(text=str(responded))
            self.stat_labels.get("pending", ctk.CTkLabel(self)).configure(text=str(pending))
            self.stat_labels.get("overdue", ctk.CTkLabel(self)).configure(text=str(overdue))

        completion = (responded / total * 100) if total > 0 else 0
        self.completion_label.configure(text=f"{completion:.1f}%")
        try:
            deadline = pd.to_datetime(app_state.review_deadline)
            days_left = (deadline - pd.Timestamp.now()).days
            self.days_left_label.configure(text=str(days_left))
        except Exception:
            self.days_left_label.configure(text="-")

    def refresh(self):
        self._rebuild_agency_list()
        self.update_dashboard_metrics()

    # ── File Operations ─────────────────────────────────────────────────

    def validate_files(self):
        try:
            self.show_progress(True)
            self.update_progress(0.2, "Validating files...")
            master = self.vars["master"].get()
            combined = self.vars["combined"].get()
            output = self.vars["output"].get()
            ok, msg = self.file_validator.validate_all_files(master, combined, output)
            self.update_progress(1.0, "Validation complete")
            self.show_progress(False)

            # Check for file changes
            change_warnings = []
            for key in ["master", "combined"]:
                path = self.vars[key].get()
                if path and app_state.has_file_changed(path):
                    change_warnings.append(f"⚠ {key.title()} file has been modified since last load!")

            if change_warnings:
                msg = "\n".join(change_warnings) + "\n\n" + msg

            messagebox.showinfo("Validation Results", msg)
        except Exception as e:
            self.show_progress(False)
            messagebox.showerror("Error", str(e))

    def generate_files(self):
        master = self.vars["master"].get()
        combined = self.vars["combined"].get()
        output = self.vars["output"].get()
        if not all([master, combined, output]):
            messagebox.showwarning("Missing Files", "Please select all required files in Setup.")
            return
        self.run_task_in_thread(target=self._generate_files_thread)

    def _generate_files_thread(self):
        try:
            self.after(0, self.show_progress, True)
            self.after(0, self.update_progress, 0.1, "Starting generation...")

            def progress_cb(val, msg):
                self.after(0, self.update_progress, val, msg)

            selected_tabs = self.get_selected_tabs()
            report = self.file_processor.generate_agency_files(
                master_file=Path(self.vars["master"].get()),
                combined_map_file=Path(self.vars["combined"].get()),
                output_dir=Path(self.vars["output"].get()),
                progress_callback=progress_cb,
                selected_tabs=selected_tabs,
                handle_unmapped="single",
            )

            self.after(0, self._auto_scan_output, self.vars["output"].get())
            self.after(0, self.show_progress, False)

            summary = (
                f"Generation Complete!\n\n"
                f"Files Created: {report.total_files_created}\n"
                f"Total Users: {report.total_generated_users}\n"
                f"Unmapped Users: {report.unmapped_users}\n"
                f"Master Total: {report.total_master_users}\n"
            )
            if report.discrepancies:
                summary += "\nDiscrepancies:\n" + "\n".join(f"  - {d}" for d in report.discrepancies)
            else:
                summary += "\n✅ All users accounted for!"

            self.after(0, messagebox.showinfo, "Generation Report", summary)
        except Exception as e:
            self.after(0, self.show_progress, False)
            self.after(0, messagebox.showerror, "Error", str(e))

    # ── Email Operations ────────────────────────────────────────────────

    def send(self):
        selected = self.get_selected_agencies()
        if not selected:
            return messagebox.showwarning("Warning", "Please select one or more agencies.")
        combined = self.vars["combined"].get()
        ok, msg = self.file_validator.validate_file_exists(combined, "Combined Mapping File")
        if not ok:
            return messagebox.showerror("Error", msg)
        mode = self.email_mode.get()
        if mode == "Schedule":
            self.schedule_email_dialog(selected)
        elif mode == "Direct":
            confirm = messagebox.askyesno(
                "Confirm Direct Send",
                f"Send {len(selected)} emails immediately?\n\nThis cannot be undone.",
                icon="warning",
            )
            if confirm:
                self.run_task_in_thread(target=self._send_emails_thread, args=(selected, mode))
        else:
            self.run_task_in_thread(target=self._send_emails_thread, args=(selected, mode))

    def _send_emails_thread(self, selected, mode):
        try:
            self.after(0, self.show_progress, True)
            self.after(0, self.update_progress, 0.1, "Preparing emails...")

            def progress_cb(val, msg):
                self.after(0, self.update_progress, val / 100, msg)

            email_mode = EmailMode.PREVIEW if mode == "Preview" else EmailMode.DIRECT
            sent_count = self.email_handler.send_emails(
                combined_file=Path(self.vars["combined"].get()),
                output_dir=Path(self.vars["output"].get()),
                file_names=selected,
                mode=email_mode,
                universal_attachment=Path(self.vars["attach"].get()) if self.vars["attach"].get() else None,
                progress_callback=progress_cb,
                selected_tabs=self.get_selected_tabs(),
            )

            # Update audit log
            output_dir = Path(self.vars["output"].get())
            self.audit_logger = AuditLogger(output_dir=output_dir, audit_file_name=app_state.audit_file_name)
            self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
            try:
                loader = CombinedFileLoader(Path(self.vars["combined"].get()))
                data = loader.load_all_data()
                addr = {m.source_file_name.upper(): (m.recipients_to, m.recipients_cc) for m in data}
            except Exception:
                addr = {}
            for agency in selected:
                to, cc = addr.get(agency.upper(), ("", ""))
                self.audit_logger.mark_sent(agency, to=to, cc=cc, comments=f"Mode: {mode}")
            self.audit_logger.save()
            self.audit_df = self.audit_logger.load()
            self.after(0, self.refresh)
            self.after(0, self.show_progress, False)
            self.after(0, messagebox.showinfo, "Done", f"Processed {sent_count}/{len(selected)} emails.")
        except Exception as e:
            self.after(0, self.show_progress, False)
            self.after(0, messagebox.showerror, "Error", str(e))

    def preview_batch(self):
        selected = self.get_selected_agencies()
        if not selected:
            return messagebox.showwarning("Warning", "Select agencies first.")
        combined = self.vars["combined"].get()
        output = self.vars["output"].get()
        if not combined or not output:
            return messagebox.showwarning("Warning", "Configure files in Setup first.")
        try:
            emails = self.email_handler.prepare_email_batch(
                Path(combined), Path(output), selected, selected_tabs=self.get_selected_tabs(),
            )
            if emails:
                EmailPreviewDialog(self, emails, self.email_handler)
            else:
                messagebox.showinfo("Info", "No emails to preview.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def schedule_email_dialog(self, agencies):
        if not HAS_TKCALENDAR:
            messagebox.showerror("Missing Package", "tkcalendar is required for scheduling.")
            return
        win = ctk.CTkToplevel(self)
        win.title("Schedule Emails")
        win.geometry("400x280")
        win.transient(self)
        win.grab_set()
        ctk.CTkLabel(win, text=f"Schedule {len(agencies)} emails:", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(15, 5))
        ctk.CTkLabel(win, text="Date & Time:").pack()
        date_entry = DateEntry(win, date_pattern='yyyy-mm-dd', width=12)
        date_entry.pack(pady=5)
        time_entry = ctk.CTkEntry(win, placeholder_text="HH:MM (24hr)")
        time_entry.insert(0, "09:00")
        time_entry.pack(pady=5)

        def do_schedule():
            dt_str = f"{date_entry.get()} {time_entry.get()}"
            try:
                naive_dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                if naive_dt <= datetime.now():
                    messagebox.showerror("Invalid", "Time must be in the future.", parent=win)
                    return
                sent = self.email_handler.send_emails(
                    combined_file=Path(self.vars["combined"].get()),
                    output_dir=Path(self.vars["output"].get()),
                    file_names=agencies,
                    mode=EmailMode.SCHEDULE,
                    scheduled_time=naive_dt,
                    selected_tabs=self.get_selected_tabs(),
                )
                win.destroy()
                messagebox.showinfo("Scheduled", f"{sent} emails scheduled for {naive_dt.strftime('%Y-%m-%d %H:%M')}")
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=win)

        ctk.CTkButton(win, text="Schedule", command=do_schedule).pack(pady=15)

    # ── Reporting Actions ───────────────────────────────────────────────

    def mark_responded(self):
        selected = self.get_selected_agencies()
        if not selected:
            return messagebox.showwarning("Warning", "Select agencies to mark as responded.")
        output = self.vars["output"].get()
        if not output:
            return messagebox.showwarning("Warning", "Set output directory first.")
        comments = simpledialog.askstring("Comments", "Add response notes (optional):", parent=self) or ""
        self.audit_logger = AuditLogger(output_dir=Path(output), audit_file_name=app_state.audit_file_name)
        self.audit_logger.initialize_log(self.agencies, preserve_existing=True)
        for agency in selected:
            self.audit_logger.mark_responded(agency, comments=comments)
        self.audit_logger.save()
        self.audit_df = self.audit_logger.load()
        self.refresh()
        messagebox.showinfo("Done", f"Marked {len(selected)} agencies as responded.")

    def resend_selected(self):
        selected = self.get_selected_agencies()
        if not selected:
            return messagebox.showwarning("Warning", "Select agencies to re-send.")
        confirm = messagebox.askyesno("Confirm Re-send", f"Reset status and allow re-sending for {len(selected)} agencies?")
        if not confirm:
            return
        output = self.vars["output"].get()
        if not output:
            return
        self.audit_logger = AuditLogger(output_dir=Path(output), audit_file_name=app_state.audit_file_name)
        self.audit_logger.load()
        for agency in selected:
            self.audit_logger.reset_agency(agency)
        self.audit_logger.save()
        self.audit_df = self.audit_logger.load()
        self.refresh()
        messagebox.showinfo("Done", f"Reset {len(selected)} agencies. You can now re-send.")

    def open_dashboard(self):
        output = self.vars["output"].get()
        if not output or not self.agencies:
            return messagebox.showwarning("Warning", "Load agencies first.")
        dashboard = ctk.CTkToplevel(self)
        dashboard.title("Compliance Dashboard")
        dashboard.geometry("1000x600")
        dashboard.transient(self)

        style = ttk.Style()
        try:
            bg = get_color("background")
            fg = get_color("text")
            style.configure("Dashboard.Treeview", background=bg, foreground=fg, fieldbackground=bg, rowheight=28)
            style.configure("Dashboard.Treeview.Heading", background=get_color("primary"), foreground="white", font=('Arial', 11, 'bold'))
        except Exception:
            pass

        cols = ["Agency", "Status", "Sent Date", "Response Date", "To", "CC", "Comments"]
        tree = ttk.Treeview(dashboard, columns=cols, show="headings", style="Dashboard.Treeview")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=130)
        tree.column("Agency", width=200)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        if not self.audit_df.empty:
            for _, row in self.audit_df.iterrows():
                values = [str(row.get(c, "")) for c in cols]
                tree.insert("", "end", values=values)

        btn_frame = ctk.CTkFrame(dashboard, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(btn_frame, text="Refresh", command=lambda: self._refresh_dashboard(tree, cols)).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Export CSV", command=lambda: self._export_csv(dashboard)).pack(side="left", padx=5)

    def _refresh_dashboard(self, tree, cols):
        tree.delete(*tree.get_children())
        if not self.audit_df.empty:
            for _, row in self.audit_df.iterrows():
                tree.insert("", "end", values=[str(row.get(c, "")) for c in cols])

    def _export_csv(self, parent):
        if self.audit_df.empty:
            return messagebox.showinfo("Info", "No data to export.", parent=parent)
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], parent=parent)
        if path:
            self.audit_df.to_csv(path, index=False)
            messagebox.showinfo("Exported", f"Saved to {path}", parent=parent)

    def generate_report(self):
        if self.audit_df.empty:
            return messagebox.showwarning("Warning", "No audit data available.")
        try:
            config = app_state.config_manager.get_all()
            metrics = DashboardMetrics.calculate(self.audit_df, config.get("deadline", ""), len(self.agencies))
            generator = ReportGenerator()
            path = generator.generate_compliance_report(self.audit_df, metrics, "summary", "xlsx", config.get("review_period", ""))
            if messagebox.askyesno("Report Generated", f"Report saved to:\n{path}\n\nOpen now?"):
                try:
                    os.startfile(path)
                except AttributeError:
                    import subprocess
                    subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ── Tool Actions ────────────────────────────────────────────────────

    def _open_email_verifier(self):
        SmartEmailVerifierDialog(self)

    def scan_inbox(self):
        try:
            replies = self.email_handler.scan_inbox()
            if replies:
                msg = f"Found {len(replies)} replies:\n\n"
                for r in replies[:10]:
                    msg += f"From: {r['sender']} ({r['received']})\n  {r['subject']}\n\n"
                messagebox.showinfo("Inbox Scan", msg)
            else:
                messagebox.showinfo("Inbox Scan", "No matching replies found.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def test_email_connection(self):
        ok, msg = self.email_handler.test_connection()
        if ok:
            messagebox.showinfo("Connection Test", msg)
        else:
            messagebox.showerror("Connection Test", msg)

    def open_output_folder(self):
        output = self.vars["output"].get()
        if not output or not os.path.exists(output):
            return messagebox.showwarning("Warning", "Output directory not set or doesn't exist.")
        try:
            os.startfile(output)
        except AttributeError:
            import subprocess
            subprocess.Popen(['xdg-open', output])

    def open_template_editor(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Email Template Editor")
        dialog.geometry("800x600")
        dialog.transient(self)
        dialog.grab_set()

        main = ctk.CTkFrame(dialog)
        main.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(main, text="Email Template Editor", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 10))

        # Font settings
        font_frame = ctk.CTkFrame(main, fg_color=get_color("surface"), corner_radius=8)
        font_frame.pack(fill="x", pady=(0, 10))
        font_inner = ctk.CTkFrame(font_frame, fg_color="transparent")
        font_inner.pack(fill="x", padx=15, pady=10)

        config = app_state.config_manager.get_all()

        ctk.CTkLabel(font_inner, text="Font:").pack(side="left", padx=(0, 5))
        font_var = ctk.StringVar(value=config.get("email_font_family", "Calibri"))
        ctk.CTkComboBox(font_inner, variable=font_var, values=["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia", "Tahoma", "Segoe UI"], width=150).pack(side="left", padx=(0, 15))

        ctk.CTkLabel(font_inner, text="Size:").pack(side="left", padx=(0, 5))
        size_var = ctk.StringVar(value=config.get("email_font_size", "11"))
        ctk.CTkComboBox(font_inner, variable=size_var, values=["9", "10", "11", "12", "13", "14", "16"], width=80).pack(side="left", padx=(0, 15))

        ctk.CTkLabel(font_inner, text="Color:").pack(side="left", padx=(0, 5))
        color_var = ctk.StringVar(value=config.get("email_font_color", "#000000"))
        ctk.CTkEntry(font_inner, textvariable=color_var, width=100).pack(side="left", padx=(0, 15))

        format_var = ctk.StringVar(value=config.get("email_format", "html"))
        ctk.CTkCheckBox(font_inner, text="Send as HTML", variable=format_var, onvalue="html", offvalue="plain").pack(side="left")

        ctk.CTkLabel(main, text="Template (use {review_period}, {deadline}, {sender_name}, etc.):", font=ctk.CTkFont(size=12)).pack(anchor="w")
        text_box = ctk.CTkTextbox(main, height=350)
        text_box.pack(fill="both", expand=True, pady=(5, 10))
        text_box.insert("1.0", load_email_template())

        def save():
            save_email_template(text_box.get("1.0", "end-1c"))
            app_state.config_manager.update({
                "email_font_family": font_var.get(), "email_font_size": size_var.get(),
                "email_font_color": color_var.get(), "email_format": format_var.get(),
            })
            app_state.config_manager.save()
            messagebox.showinfo("Saved", "Template and font settings saved.", parent=dialog)
            dialog.destroy()

        ctk.CTkButton(main, text="Save", command=save, height=40).pack(side="left", padx=(0, 10))
        ctk.CTkButton(main, text="Cancel", command=dialog.destroy, height=40, fg_color="transparent", border_width=2, border_color=get_color("border")).pack(side="left")

    def open_region_manager(self):
        rm = RegionManager()
        dialog = ctk.CTkToplevel(self)
        dialog.title("Region Manager")
        dialog.geometry("600x400")
        dialog.transient(self)
        dialog.grab_set()

        main = ctk.CTkFrame(dialog)
        main.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main, text="Region Profiles", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 10))

        listbox = tk.Listbox(main, height=10)
        listbox.pack(fill="both", expand=True, pady=5)
        for name in rm.get_all_region_names():
            listbox.insert("end", name)

        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)

        def add_region():
            name = simpledialog.askstring("New Region", "Region name:", parent=dialog)
            if name:
                rm.add_profile(RegionProfile(region_name=name))
                listbox.insert("end", name)

        def delete_region():
            sel = listbox.curselection()
            if sel:
                name = listbox.get(sel[0])
                if messagebox.askyesno("Confirm", f"Delete region '{name}'?", parent=dialog):
                    rm.delete_profile(name)
                    listbox.delete(sel[0])

        ctk.CTkButton(btn_frame, text="Add", command=add_region, width=80).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Delete", command=delete_region, width=80, fg_color=get_color("danger")).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Close", command=dialog.destroy, width=80).pack(side="right", padx=5)

    def fix_errors(self):
        """Find and fix unmapped agencies between master and combined mapping files."""
        master_file = self.vars["master"].get()
        combined_file = self.vars["combined"].get()
        if not master_file or not combined_file:
            messagebox.showerror("Missing Files", "Please select Master File and Combined Mapping File in Setup.")
            return
        try:
            df_master = pd.read_excel(master_file)
            df_master.columns = df_master.columns.str.strip()

            # Load all agencies from the combined mapping file (agency_id column)
            loader = CombinedFileLoader(Path(combined_file))
            combined_data = loader.load_all_data()
            mapped_agencies = set()
            for m in combined_data:
                mapped_agencies.add(m.agency_id.strip().upper())

            # Compare with master file's Agency column
            master_agencies = set(df_master[ColumnNames.AGENCY].astype(str).str.strip().str.upper().unique())
            missing = sorted(master_agencies - mapped_agencies)

            if not missing:
                messagebox.showinfo("No Issues", "All agencies in the master file are mapped.")
                return

            msg = f"Found {len(missing)} unmapped agencies:\n\n"
            msg += "\n".join(f"  - {a}" for a in missing[:20])
            if len(missing) > 20:
                msg += f"\n  ... and {len(missing) - 20} more"
            msg += "\n\nThese agencies exist in the master file but have no mapping in the combined file."
            messagebox.showwarning("Unmapped Agencies", msg)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to check mappings: {str(e)}")

    def add_agency_folder(self):
        path = filedialog.askdirectory(title="Select Additional Agency Folder")
        if not path or not os.path.isdir(path):
            return
        try:
            files = [f.replace(".xlsx", "") for f in os.listdir(path) if f.endswith(".xlsx") and not f.startswith("~$") and not f.startswith("Unassigned")]
            if not files:
                messagebox.showinfo("No Files", f"No agency Excel files found in:\n{path}")
                return
            added = 0
            for f in files:
                if f not in self.agencies:
                    self.agencies.append(f)
                    self.agency_folders[f] = path
                    added += 1
            self.agencies = sorted(self.agencies)

            # Try to merge audit data
            audit_file = os.path.join(path, app_state.audit_file_name)
            if os.path.exists(audit_file):
                try:
                    new_df = pd.read_excel(audit_file)
                    if self.audit_df.empty:
                        self.audit_df = new_df
                    else:
                        self.audit_df = pd.concat([self.audit_df, new_df], ignore_index=True).drop_duplicates(subset=['Agency'], keep='first')
                except Exception:
                    pass

            self._rebuild_agency_list()
            self.update_dashboard_metrics()
            messagebox.showinfo("Added", f"Added {added} agencies from:\n{path}\nTotal: {len(self.agencies)}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ── Settings & Theme ────────────────────────────────────────────────

    def settings(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Settings")
        dialog.geometry("500x450")
        dialog.transient(self)
        dialog.grab_set()

        main = ctk.CTkFrame(dialog)
        main.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main, text="Application Settings", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 15))

        scroll = ctk.CTkScrollableFrame(main)
        scroll.pack(fill="both", expand=True)

        config = app_state.config_manager.get_all()
        entries = {}
        fields = [
            ("review_period", "Review Period"), ("deadline", "Deadline"),
            ("company_name", "Company Name"), ("sender_name", "Sender Name"),
            ("sender_title", "Sender Title"), ("email_subject_prefix", "Email Subject Prefix"),
            ("email_send_delay", "Send Delay (seconds)"),
        ]
        for key, label in fields:
            ctk.CTkLabel(scroll, text=label, font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(8, 2))
            var = ctk.StringVar(value=str(config.get(key, "")))
            ctk.CTkEntry(scroll, textvariable=var, height=36).pack(fill="x")
            entries[key] = var

        def save():
            try:
                updates = {k: v.get() for k, v in entries.items()}
                try:
                    updates["email_send_delay"] = float(updates["email_send_delay"])
                except ValueError:
                    updates["email_send_delay"] = 2.0
                app_state.config_manager.update(updates)
                app_state.config_manager.save()
                app_state.reload_config()
                messagebox.showinfo("Saved", "Settings saved.", parent=dialog)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dialog)

        btn_frame = ctk.CTkFrame(main, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 0))
        ctk.CTkButton(btn_frame, text="Save", command=save).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Cancel", command=dialog.destroy, fg_color="transparent", border_width=2, border_color=get_color("border")).pack(side="left")

    def change_theme(self, choice):
        ctk.set_appearance_mode(choice)

    def show_about(self):
        messagebox.showinfo(
            "About",
            "Cognos Access Review Tool\nEnterprise Edition v2.0\n\n"
            "SOX Compliance User Access Review Automation\n\n"
            f"Review Period: {app_state.review_period}\n"
            f"Deadline: {app_state.review_deadline}",
        )


# ============================================================================
# EMAIL PREVIEW DIALOG
# ============================================================================

class EmailPreviewDialog(ctk.CTkToplevel):
    def __init__(self, parent, emails: List[dict], email_handler: EmailHandler):
        super().__init__(parent)
        self.emails = emails
        self.email_handler = email_handler
        self.current_index = 0
        self.sent_count = 0
        self.skipped = set()

        self.title("Email Preview")
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self._build_ui()
        self._display_current()
        self._center()

    def _build_ui(self):
        main = ctk.CTkFrame(self)
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # Navigation
        nav = ctk.CTkFrame(main, fg_color="transparent")
        nav.pack(fill="x", pady=(0, 10))
        self.prev_btn = ctk.CTkButton(nav, text="← Previous", width=100, command=self._prev)
        self.prev_btn.pack(side="left")
        self.counter_label = ctk.CTkLabel(nav, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self.counter_label.pack(side="left", expand=True)
        self.next_btn = ctk.CTkButton(nav, text="Next →", width=100, command=self._next)
        self.next_btn.pack(side="right")

        # Fields
        for label_text, attr in [("To:", "to_text"), ("CC:", "cc_text"), ("Subject:", "subject_text")]:
            ctk.CTkLabel(main, text=label_text, font=ctk.CTkFont(weight="bold")).pack(anchor="w")
            tb = ctk.CTkTextbox(main, height=30)
            tb.pack(fill="x", pady=(0, 5))
            setattr(self, attr, tb)

        ctk.CTkLabel(main, text="Attachment:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.attach_label = ctk.CTkLabel(main, text="")
        self.attach_label.pack(anchor="w", pady=(0, 5))

        ctk.CTkLabel(main, text="Body:", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.body_text = ctk.CTkTextbox(main, height=250)
        self.body_text.pack(fill="both", expand=True, pady=(0, 10))

        # Buttons
        btns = ctk.CTkFrame(main, fg_color="transparent")
        btns.pack(fill="x")
        ctk.CTkButton(btns, text="Skip", command=self._skip, width=80, fg_color="transparent", border_width=2, border_color=get_color("border")).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btns, text="Cancel All", command=self._cancel, width=100, fg_color="transparent", border_width=2, border_color=get_color("border")).pack(side="left")
        ctk.CTkButton(btns, text="Send All Remaining", command=self._send_all, width=150, fg_color=get_color("primary")).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btns, text="Send This", command=self._send_current, width=100, fg_color=STATUS_COMPLETE).pack(side="right")

    def _display_current(self):
        if not self.emails or self.current_index >= len(self.emails):
            return
        e = self.emails[self.current_index]
        self.counter_label.configure(text=f"Email {self.current_index + 1} of {len(self.emails)}")
        self.prev_btn.configure(state="normal" if self.current_index > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.current_index < len(self.emails) - 1 else "disabled")
        for tb, key in [(self.to_text, 'to'), (self.cc_text, 'cc'), (self.subject_text, 'subject')]:
            tb.delete("1.0", "end")
            tb.insert("1.0", e.get(key, ''))
        self.attach_label.configure(text=Path(e.get('attachment_path', '')).name if e.get('attachment_path') else "None")
        self.body_text.delete("1.0", "end")
        self.body_text.insert("1.0", e.get('body', '')[:2000])

    def _prev(self):
        if self.current_index > 0:
            self.current_index -= 1
            self._display_current()

    def _next(self):
        if self.current_index < len(self.emails) - 1:
            self.current_index += 1
            self._display_current()

    def _skip(self):
        self.skipped.add(self.current_index)
        if self.current_index < len(self.emails) - 1:
            self._next()
        else:
            self._finish()

    def _send_current(self):
        e = self.emails[self.current_index]
        ok, err = self.email_handler.send_single_email(e['to'], e['cc'], e['subject'], e['body'], e.get('attachment_path'), e.get('html_body'))
        if ok:
            self.sent_count += 1
            if self.current_index < len(self.emails) - 1:
                self._next()
            else:
                self._finish()
        else:
            messagebox.showerror("Send Failed", err)

    def _send_all(self):
        if not messagebox.askyesno("Confirm", f"Send all {len(self.emails) - self.current_index} remaining emails?"):
            return
        for i in range(self.current_index, len(self.emails)):
            if i in self.skipped:
                continue
            e = self.emails[i]
            ok, _ = self.email_handler.send_single_email(e['to'], e['cc'], e['subject'], e['body'], e.get('attachment_path'), e.get('html_body'))
            if ok:
                self.sent_count += 1
        self._finish()

    def _cancel(self):
        if messagebox.askyesno("Confirm", "Cancel sending? Already sent emails won't be recalled."):
            self.destroy()

    def _finish(self):
        messagebox.showinfo("Complete", f"Sent: {self.sent_count} | Skipped: {len(self.skipped)}")
        self.destroy()

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


# ============================================================================
# UNASSIGNED AGENCIES DIALOG
# ============================================================================

class UnassignedAgenciesDialog(ctk.CTkToplevel):
    def __init__(self, parent, unassigned_data, file_to_agencies, mapping_file_path, current_region=""):
        super().__init__(parent)
        self.unassigned_data = unassigned_data
        self.file_to_agencies = file_to_agencies
        self.mapping_file_path = mapping_file_path
        self.result = None
        self.action_widgets = {}

        self.title("Unassigned Agencies")
        self.geometry("900x600")
        self.transient(parent)
        self.grab_set()
        self._build_ui()
        self._center()

    def _build_ui(self):
        ctk.CTkLabel(self, text=f"⚠ {len(self.unassigned_data)} Unassigned Agencies", font=ctk.CTkFont(size=18, weight="bold")).pack(padx=20, pady=(15, 5), anchor="w")

        scroll = ctk.CTkScrollableFrame(self, fg_color=get_color("surface"))
        scroll.pack(fill="both", expand=True, padx=20, pady=10)

        existing_files = list(self.file_to_agencies.keys())
        for idx, data in enumerate(self.unassigned_data):
            agency = data['agency']
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text=f"{agency} ({data.get('user_count', 0)} users)", width=300, anchor="w").pack(side="left", padx=5)
            action_var = ctk.StringVar(value="Keep as Unassigned")
            ctk.CTkOptionMenu(row, variable=action_var, values=["Add to Existing File", "Create New File", "Keep as Unassigned"], width=180).pack(side="left", padx=5)
            target_var = ctk.StringVar(value=existing_files[0] if existing_files else agency)
            ctk.CTkComboBox(row, variable=target_var, values=existing_files + [agency], width=200).pack(side="left", padx=5)
            self.action_widgets[agency] = {'action_var': action_var, 'target_var': target_var, 'agency_data': data}

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(btn_frame, text="Apply", command=self._apply, fg_color=get_color("primary")).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="Cancel", command=self._cancel, fg_color="transparent", border_width=2, border_color=get_color("border")).pack(side="right")

    def _apply(self):
        decisions = {}
        for agency, w in self.action_widgets.items():
            decisions[agency] = {
                'action': w['action_var'].get(), 'target': w['target_var'].get(),
                'agency_data': w['agency_data'], 'recipients': {'to': '', 'cc': ''},
            }
        self.result = {'decisions': decisions, 'update_mapping': True}
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    app = CognosAccessReviewApp()
    app.mainloop()

if __name__ == "__main__":
    main()
