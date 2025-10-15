import sys
import os
import shutil
import subprocess
import re
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote
import concurrent.futures
import threading

import shutil as _shutil
import requests
import urllib3
from bs4 import BeautifulSoup

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QPoint
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QWidget,
    QPushButton,
    QFileDialog,
    QListWidget,
    QListWidgetItem,
    QLabel,
    QHBoxLayout,
    QComboBox,
    QSplitter,
    QMessageBox,
    QMenu,
    QDialog,
    QLineEdit,
    QFormLayout,
    QDialogButtonBox,
    QProgressBar,
    QTextEdit,
    QFrame,
)
from PyQt6.QtPdfWidgets import QPdfView
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtGui import QWheelEvent


# ------------------ SETTINGS ------------------
def get_upload_dir():
    """Return the appropriate upload directory based on whether running in a PyInstaller bundle."""
    if getattr(sys, "frozen", False):  # Running in a PyInstaller bundle
        # Use a user-writable directory, e.g., ~/Documents/VCEViewer/uploaded_reports
        base_dir = Path.home() / "Documents" / "VCEViewer"
    else:
        # Use the project directory for development
        base_dir = Path.cwd()
    upload_dir = base_dir / "uploaded_reports"
    return upload_dir


UPLOAD_DIR = get_upload_dir()
CONVERTED_DIR = UPLOAD_DIR / "converted"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
CONVERTED_DIR.mkdir(exist_ok=True)
SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".docx"}  # include .doc too for older reports
WORD_EXTENSIONS = {".doc", ".docx"}

VCAA_BASE = "https://www.vcaa.vic.edu.au"
VCAA_SUBJECTS_PAGE = (
    VCAA_BASE
    + "/assessment/vce/examination-specifications-past-examinations-and-examination-reports/"
    + "examination-specifications-past-examinations-and-external-assessment-reports"
)
VCAA_NHT_SUBJECTS_PAGE = (
    VCAA_BASE
    + "/assessment/vce/examination-specifications-past-examinations-and-examination-reports/"
    + "nht-examination-specifications-past-examinations-and-examination-reports"
)
VCAA_SUBJECTS_PREFIX = (
    "/assessment/vce/examination-specifications-past-examinations-and-examination-reports/"
)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Always skip these (case-insensitive)
EXCLUDE_HINTS = [
    "sample",  # sample exams
    "formula",  # formula sheets/booklets
    "data book",  # data books (various spellings)
    "data-book",
    "databook",
    "assessment guide",  # not a report
    "transcript",  # video transcripts etc.
]

# We accept links whose VISIBLE TEXT contains "report"
REPORT_TOKEN = "report"

# ------------------ SUBJECT NORMALISATION (for filename parsing only) ------------------
SUBJECT_ALIASES = {
    "mathmethodscas": "MathMethodsCAS",
    "mathematicalmethods": "MathMethods",
    "mathmethods": "MathMethods",
    "methods": "MathMethods",
    "method": "MathMethods",
    "mmcas": "MathMethodsCAS",
    "maths1": "MathMethods",
    "mm": "MathMethods",
    "mmcas2": "MathMethodsCAS",
    "specialist": "SpecialistMaths",
    "sm": "SpecialistMaths",
    "chemistry": "Chemistry",
    "chem": "Chemistry",
}

SUBJECT_KEY_ALIASES = {
    "mathmethodscas": "mathmethodscas",
    "mathematicalmethods": "mathmethods",
    "mathmethods": "mathmethods",
    "methods": "mathmethods",
    "method": "mathmethods",
    "mmcas": "mathmethodscas",
    "mm": "mathmethods",
    "maths1": "mathmethods",
    "mmcas2": "mathmethodscas",
    "specialist": "specialistmaths",
    "specialistmathematics": "specialistmaths",
    "specialistmaths": "specialistmaths",
    "sm": "specialistmaths",
    "chemistry": "chemistry",
    "chem": "chemistry",
}

SUBJECT_DISPLAY_NAMES = {
    "mathmethods": "Mathematical Methods",
    "mathmethodscas": "Mathematical Methods (CAS)",
    "specialistmaths": "Specialist Mathematics",
    "chemistry": "Chemistry",
}


# ------------------ UTILS ------------------
def soffice_cmd():
    """
    Return the path to the LibreOffice CLI binary for headless mode, or None if not found.
    Checks Windows, macOS, and Linux paths, as well as system PATH.
    """
    possible_paths = []
    if os.name == "nt":  # Windows
        possible_paths.extend([
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"C:\LibreOffice\program\soffice.exe",  # Custom install
            r"D:\LibreOffice\program\soffice.exe",  # Custom drive
            "soffice.exe",  # System PATH
        ])
    else:  # macOS and Linux
        possible_paths.extend([
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
            "/usr/local/bin/soffice",  # Linux or macOS Homebrew
            "/opt/local/bin/soffice",  # Alternative
            "/usr/bin/soffice",  # Common Linux path
            "soffice",  # System PATH
            "libreoffice",  # Fallback name
        ])

    for path in possible_paths:
        if _shutil.which(path) or os.path.exists(path):
            return path
    return None  # Indicate LibreOffice is not available


# ------------------ PARSING ------------------
def parse_filename(file_path: Path, title_hint=None):
    """
    Parse (best-effort) subject, year and exam number from a file path's name.
    Returns: (subject, year, exam_number)
    subject is inferred from aliases when possible; otherwise title-cased basename.
    year = '20xx' or 'Unknown'. exam_number = 'exam1'/'exam2' or 'Unknown'.
    """
    raw_parts = [file_path.stem]
    if title_hint:
        raw_parts.append(str(title_hint))
    combined = " ".join(p for p in raw_parts if p).lower()

    # Year
    year = "Unknown"
    year_match = re.search(r"(20\d{2})", combined)
    if year_match:
        year = year_match.group(1)
    else:
        trailing_pair = re.search(r"\b(\d{2})\b", combined)
        if trailing_pair:
            y = int(trailing_pair.group(1))
            if 0 <= y <= 30:
                year = f"20{y:02d}"

    # Exam number
    exam_number = "Unknown"
    exam_patterns = [
        r"exam\s*(?:number\s*)?([12])",
        r"exam[-_]?([12])",
        r"examrep(?:ort)?[-_\s]*([12])",
        r"assess(?:ment)?rep(?:ort)?[-_\s]*([12])",
        r"externalassessmentreport[-_\s]*([12])",
        r"paper[-_\s]*([12])",
        r"report[-_\s]*([12])",
    ]
    for pattern in exam_patterns:
        ex_match = re.search(pattern, combined)
        if ex_match:
            exam_number = f"exam{ex_match.group(1)}"
            break

    if exam_number == "Unknown":
        word_match = re.search(
            r"(?:exam|paper)[-_\s]*(one|two|i{1,3}|iv|v)\b",
            combined,
        )
        if word_match:
            token = word_match.group(1).lower()
            token_map = {"one": "1", "two": "2", "i": "1", "ii": "2"}
            mapped = token_map.get(token)
            if mapped:
                exam_number = f"exam{mapped}"

    name = combined
    name = re.sub(
        r"[-_\s]?(assessrep|examreport|examrep|externalassessmentreport|report|exam)",
        " ",
        name,
    )
    name = re.sub(r"\s*\(\d+\)", "", name)
    if year != "Unknown":
        name = name.replace(year.lower(), " ")
    name = re.sub(r"\b20\d{2}\b", " ", name)
    name = re.sub(r"\bnorthern\s+hemisphere(?:\s+timetable)?\b", " ", name)
    name = re.sub(r"\bnht\b", " ", name)
    name = re.sub(r"[^a-z0-9\s]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()

    if exam_number == "Unknown":
        ex_match = re.search(r"(?:ex|exam)?[-_\s]?([12])\b", name)
        if ex_match:
            exam_number = f"exam{ex_match.group(1)}"
            name = re.sub(r"(?:ex|exam)?[-_\s]?[12]\b", "", name).strip()
        else:
            trailing_digit = re.search(r"(\d)$", name)
            if trailing_digit:
                exam_number = f"exam{trailing_digit.group(1)}"
                name = re.sub(r"\d$", "", name).strip()

    # Subject (very rough heuristic — we override this with the user-selected subject on VCAA downloads)
    tokens = [t for t in re.split(r"[^a-z0-9]+", combined) if t]
    joined = "".join(tokens)

    subject = "Unknown"
    for key, canonical in SUBJECT_ALIASES.items():
        if len(key) <= 3:
            if key in tokens or any(tok.startswith(key) for tok in tokens):
                subject = canonical
                break
        else:
            if key in tokens or (key in joined if joined else False):
                subject = canonical
                break

    if subject == "Unknown" and name:
        fallback = " ".join(t for t in tokens if t)
        subject = fallback.title() if fallback else "Unknown"

    return subject, year, exam_number


# ------------------ DOCX CONVERTER (sequential) ------------------
class DocxConverterThread(QThread):
    progress = pyqtSignal(str, int)  # doc_path, percent
    finished = pyqtSignal(str, str)  # doc_path, pdf_path
    error = pyqtSignal(str, str)  # doc_path, error_message

    def __init__(self, docx_path: str, output_dir: str):
        super().__init__()
        self.docx_path = docx_path
        self.output_dir = output_dir

    def run(self):
        soffice = soffice_cmd()
        if not soffice:
            error_msg = (
                "LibreOffice is not installed or not found in PATH. "
                "Please install LibreOffice from https://www.libreoffice.org/download/download/ "
                "to enable .doc/.docx conversion."
            )
            self.error.emit(self.docx_path, error_msg)
            return

        try:
            # Signal "started conversion"
            self.progress.emit(self.docx_path, 60)
            cmd = [
                soffice,
                "--headless",
                "--nologo",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                str(self.output_dir),  # Ensure string path
                str(self.docx_path),  # Ensure string path
            ]
            subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True,
                shell=(os.name == "nt")  # Use shell on Windows to handle paths with spaces
            )

            pdf_path = Path(self.output_dir) / (Path(self.docx_path).stem + ".pdf")
            if pdf_path.exists():
                self.progress.emit(self.docx_path, 100)
                self.finished.emit(self.docx_path, str(pdf_path))
            else:
                self.error.emit(self.docx_path, "Conversion failed: PDF not created.")
        except subprocess.CalledProcessError as e:
            error_msg = f"LibreOffice conversion failed: {e.stderr or str(e)}"
            self.error.emit(self.docx_path, error_msg)
        except Exception as e:
            self.error.emit(self.docx_path, str(e))


# ------------------ VCAA SCRAPER ------------------
class VCAASubjectScraperThread(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    @staticmethod
    def _normalise_subject_key(label: str) -> str:
        cleaned = re.sub(r"\(.*?nht.*?\)", "", label or "", flags=re.IGNORECASE)
        cleaned = re.sub(
            r"northern\s+hemisphere\s+timetable",
            " ",
            cleaned,
            flags=re.IGNORECASE,
        )
        cleaned = re.sub(r"northern\s+hemisphere", " ", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\bnht\b", " ", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(
            r"\b(examination|exam|report|reports|assessment|external|paper|papers|specification|specifications)\b",
            " ",
            cleaned,
            flags=re.IGNORECASE,
        )
        cleaned = re.sub(r"\b20\d{2}\b", " ", cleaned)
        cleaned = re.sub(r"[^A-Za-z0-9\s]+", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        alias_key = re.sub(r"[^a-z0-9]+", "", cleaned.lower())
        return SUBJECT_KEY_ALIASES.get(alias_key, alias_key)

    @staticmethod
    def _canonicalize(text: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", (text or "").lower())

    @staticmethod
    def _clean_subject_label(text: str) -> str:
        if not text:
            return ""
        spaced = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
        spaced = re.sub(r"\(.*?\)", " ", spaced)
        spaced = re.sub(
            r"northern\s+hemisphere\s+timetable|northern\s+hemisphere|\bnht\b",
            " ",
            spaced,
            flags=re.IGNORECASE,
        )
        spaced = re.sub(
            r"\b(examination|exam|report|reports|assessment|external|paper|papers|specification|specifications)\b",
            " ",
            spaced,
            flags=re.IGNORECASE,
        )
        spaced = re.sub(r"\b20\d{2}\b", " ", spaced)
        spaced = re.sub(r"[^A-Za-z0-9\s]+", " ", spaced)
        spaced = re.sub(r"\s+", " ", spaced).strip()
        return spaced.title()

    @staticmethod
    def _scrape_subject_page(page_url: str, headers: dict):
        resp = requests.get(page_url, headers=headers, timeout=30, verify=False)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        subjects = {}
        for link in soup.find_all("a", href=True):
            text = link.get_text(strip=True)
            href = link["href"].strip()
            if not href or not text:
                continue
            full = urljoin(VCAA_BASE, href)
            path_lower = full.lower()
            # Keep only VCE study pages under this subtree, skip VET and other hubs
            if "/vce-vet-" in path_lower:
                continue
            if VCAA_SUBJECTS_PREFIX not in path_lower:
                continue
            full_stripped = full.rstrip("/")
            if full_stripped in {
                page_url.rstrip("/"),
                VCAA_SUBJECTS_PAGE.rstrip("/"),
                VCAA_NHT_SUBJECTS_PAGE.rstrip("/"),
            }:
                continue
            key = VCAASubjectScraperThread._normalise_subject_key(text)
            subjects.setdefault(
                key,
                {
                    "label": text,
                    "urls": [],
                },
            )
            subjects[key]["urls"].append({"url": full, "is_nht": False})
        return subjects

    @classmethod
    def _match_subject_candidate(cls, candidate: str, subjects: dict):
        if not candidate:
            return None, None

        alias_key = cls._normalise_subject_key(candidate)
        alias_key = SUBJECT_KEY_ALIASES.get(alias_key, alias_key)
        if alias_key in subjects:
            label = subjects[alias_key].get("label")
            if not label:
                label = SUBJECT_DISPLAY_NAMES.get(alias_key) or candidate
            return alias_key, label

        cleaned = cls._clean_subject_label(candidate)
        if cleaned and cleaned != candidate:
            alias_key = cls._normalise_subject_key(cleaned)
            alias_key = SUBJECT_KEY_ALIASES.get(alias_key, alias_key)
            if alias_key in subjects:
                label = subjects[alias_key].get("label")
                if not label:
                    label = SUBJECT_DISPLAY_NAMES.get(alias_key) or cleaned
                return alias_key, label

        variants = set()
        candidate_lower = (candidate or "").lower()
        if candidate_lower:
            variants.add(candidate_lower)
        cleaned_lower = (cleaned or "").lower()
        if cleaned_lower:
            variants.add(cleaned_lower)

        best = (None, None, 0)
        for key, info in subjects.items():
            label = info.get("label") or SUBJECT_DISPLAY_NAMES.get(key) or key
            label_lower = (label or "").lower()
            if not label_lower:
                continue
            for variant in variants:
                if not variant:
                    continue
                if label_lower in variant or variant in label_lower:
                    score = len(label_lower)
                    if score > best[2]:
                        best = (key, info.get("label") or label, score)
        if best[0]:
            return best[0], best[1]
        return None, None

    @classmethod
    def _extract_subject_from_nht_link(cls, text: str, url: str, subjects: dict):
        candidates = []
        parsed = urlparse(url)
        filename = unquote(parsed.path.split("/")[-1]) if parsed.path else ""
        if filename:
            stem = Path(filename).stem
            if stem:
                candidates.append(stem)
            file_subject, _, _ = parse_filename(Path(filename), title_hint=text)
            if file_subject and file_subject != "Unknown":
                candidates.append(file_subject)
        if text:
            candidates.append(text)
        if parsed.path:
            for segment in parsed.path.split("/"):
                if segment:
                    candidates.append(segment)

        for candidate in candidates:
            key, label = cls._match_subject_candidate(candidate, subjects)
            if key:
                return key, label

        return None, None

    @classmethod
    def _collect_nht_documents(cls, headers: dict, subjects: dict):
        resp = requests.get(
            VCAA_NHT_SUBJECTS_PAGE, headers=headers, timeout=30, verify=False
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        collected = {}
        for link in soup.find_all("a", href=True):
            text = link.get_text(strip=True)
            href = link["href"].strip()
            if not href:
                continue
            text_lower = (text or "").lower()
            href_lower = href.lower()
            if "nht" not in text_lower and "northern hemisphere" not in text_lower and "nht" not in href_lower:
                continue
            if REPORT_TOKEN not in text_lower and REPORT_TOKEN not in href_lower:
                continue
            full = urljoin(VCAA_BASE, href)
            parsed = urlparse(full)
            path_lower = (parsed.path or "").lower()
            if not path_lower.endswith((".pdf", ".doc", ".docx")):
                continue
            key, label = cls._extract_subject_from_nht_link(text, full, subjects)
            if not key:
                continue
            entry = collected.setdefault(
                key,
                {
                    "label": label
                    or SUBJECT_DISPLAY_NAMES.get(key)
                    or "",
                    "urls": [],
                },
            )
            if label:
                entry["label"] = label
            entry["urls"].append({"url": full, "is_nht": True, "title": text})
        return collected

    def run(self):
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            subjects = self._scrape_subject_page(VCAA_SUBJECTS_PAGE, headers)
            normalised_subjects = {}
            for key, info in subjects.items():
                label = info.get("label") or key
                alias_key = self._normalise_subject_key(label)
                canonical_key = SUBJECT_KEY_ALIASES.get(alias_key, alias_key) if alias_key else key
                canonical_key = SUBJECT_KEY_ALIASES.get(canonical_key, canonical_key)
                preferred_label = SUBJECT_DISPLAY_NAMES.get(canonical_key) or label
                entry = normalised_subjects.setdefault(
                    canonical_key,
                    {
                        "label": preferred_label,
                        "urls": [],
                    },
                )
                if SUBJECT_DISPLAY_NAMES.get(canonical_key):
                    entry["label"] = SUBJECT_DISPLAY_NAMES[canonical_key]
                elif info.get("label") and not entry.get("label"):
                    entry["label"] = info["label"]
                entry["urls"].extend(info.get("urls", []))
            subjects = normalised_subjects
            try:
                nht_subjects = self._collect_nht_documents(headers, subjects)
            except Exception:
                nht_subjects = {}

            for key, info in nht_subjects.items():
                canonical_key = SUBJECT_KEY_ALIASES.get(key, key)
                preferred_label = SUBJECT_DISPLAY_NAMES.get(canonical_key) or info.get("label")
                if not preferred_label:
                    preferred_label = "NHT Subject"
                entry = subjects.setdefault(
                    canonical_key,
                    {
                        "label": preferred_label,
                        "urls": [],
                    },
                )
                if SUBJECT_DISPLAY_NAMES.get(canonical_key):
                    entry["label"] = SUBJECT_DISPLAY_NAMES[canonical_key]
                elif not entry.get("label"):
                    entry["label"] = preferred_label
                entry["urls"].extend(info.get("urls", []))

            for key, info in subjects.items():
                preferred_label = SUBJECT_DISPLAY_NAMES.get(key)
                if preferred_label:
                    info["label"] = preferred_label
                elif not info.get("label"):
                    info["label"] = self._clean_subject_label(key)
                seen_urls = {}
                unique_urls = []
                for entry in info.get("urls", []):
                    if isinstance(entry, dict):
                        url = entry.get("url")
                        is_nht = bool(entry.get("is_nht"))
                        title = entry.get("title")
                    else:
                        url = entry
                        is_nht = False
                        title = None
                    if not url:
                        continue
                    if url in seen_urls:
                        if is_nht:
                            seen_urls[url]["is_nht"] = True
                        if title and not seen_urls[url].get("title"):
                            seen_urls[url]["title"] = title
                        continue
                    stored = {"url": url, "is_nht": is_nht}
                    if title:
                        stored["title"] = title
                    seen_urls[url] = stored
                    unique_urls.append(stored)
                info["urls"] = unique_urls

            if not subjects:
                self.error.emit("No subjects found on the VCAA index page.")
            else:
                self.finished.emit(subjects)
        except Exception as e:
            self.error.emit(str(e))


class VCAADownloadThread(QThread):
    progress = pyqtSignal(
        str, int, int
    )  # message, current, total (only counts matched links)
    file_done = pyqtSignal(str)  # saved file path
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, subject_name, subject_urls):
        super().__init__()
        self.subject_name = subject_name
        self.subject_urls = (
            subject_urls if isinstance(subject_urls, (list, tuple)) else [subject_urls]
        )

    @staticmethod
    def _should_skip(link_href: str, link_text: str) -> bool:
        h = (link_href or "").lower()
        t = (link_text or "").lower()
        if any(k in h or k in t for k in EXCLUDE_HINTS):
            return True
        # We ONLY want reports: visible text must include 'report'
        if REPORT_TOKEN not in t and REPORT_TOKEN not in h:
            return True
        # Keep only PDF/DOC/DOCX
        if not (h.endswith(".pdf") or h.endswith(".docx") or h.endswith(".doc")):
            return True
        return False

    def run(self):
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            links = []

            for source in self.subject_urls:
                if isinstance(source, dict):
                    page_url = source.get("url")
                    source_is_nht = bool(source.get("is_nht"))
                    source_title = source.get("title")
                else:
                    page_url = source
                    source_is_nht = False
                    source_title = None
                if not page_url:
                    continue

                parsed = urlparse(page_url)
                path_lower = (parsed.path or "").lower()
                if path_lower.endswith((".pdf", ".doc", ".docx")):
                    links.append(
                        {
                            "url": page_url,
                            "is_nht": source_is_nht
                            or "nht" in path_lower
                            or "northern-hemisphere" in path_lower
                            or "northernhemisphere" in path_lower,
                            "title": source_title,
                        }
                    )
                    continue

                resp = requests.get(
                    page_url, headers=headers, timeout=30, verify=False
                )
                resp.raise_for_status()
                soup = BeautifulSoup(resp.text, "html.parser")

                for a in soup.find_all("a", href=True):
                    href = a["href"].strip()
                    text = a.get_text(strip=True)
                    if not href:
                        continue
                    if self._should_skip(href, text):
                        continue
                    file_url = urljoin(VCAA_BASE, href)
                    combined_lower = href.lower()
                    text_lower = (text or "").lower()
                    links.append(
                        {
                            "url": file_url,
                            "is_nht": source_is_nht
                            or "nht" in combined_lower
                            or "northern hemisphere" in text_lower
                            or "northern-hemisphere" in combined_lower
                            or "northernhemisphere" in combined_lower,
                            "title": text,
                        }
                    )

            seen = {}
            unique_links = []
            for link in links:
                url = link.get("url") if isinstance(link, dict) else link
                if not url:
                    continue
                is_nht = bool(link.get("is_nht")) if isinstance(link, dict) else False
                title = link.get("title") if isinstance(link, dict) else None
                if url in seen:
                    if is_nht:
                        seen[url]["is_nht"] = True
                    if title and not seen[url].get("title"):
                        seen[url]["title"] = title
                    continue
                stored = {"url": url, "is_nht": is_nht}
                if title:
                    stored["title"] = title
                seen[url] = stored
                unique_links.append(stored)

            total = len(unique_links)
            if total == 0:
                self.finished.emit(
                    f"No examination reports found for {self.subject_name}."
                )
                return

            subject_folder = UPLOAD_DIR / self._safe_subject_folder(self.subject_name)
            subject_folder.mkdir(parents=True, exist_ok=True)

            completed = 0
            lock = threading.Lock()

            def download_one(link_info):
                nonlocal completed
                file_url = link_info.get("url") if isinstance(link_info, dict) else link_info
                is_nht_doc = bool(link_info.get("is_nht")) if isinstance(link_info, dict) else False
                title_hint = link_info.get("title") if isinstance(link_info, dict) else None
                if title_hint:
                    lower_title = title_hint.lower()
                    if "nht" in lower_title or "northern hemisphere" in lower_title:
                        is_nht_doc = True
                filename = file_url.split("/")[-1]
                try:
                    r = requests.get(
                        file_url, headers=headers, timeout=120, verify=False
                    )
                    r.raise_for_status()

                    # Use original name temporarily to parse year/exam number
                    temp_path = subject_folder / filename
                    with open(temp_path, "wb") as f:
                        f.write(r.content)

                    # Infer year/exam from filename or link text; fix subject to dialog selection
                    _, year, exam_number = parse_filename(temp_path, title_hint=title_hint)

                    # Build final filename
                    ext = temp_path.suffix.lower()
                    parts = [self._clean_filename(self.subject_name)]
                    if year and year != "Unknown":
                        parts.append(year)
                    if exam_number and exam_number != "Unknown":
                        parts.append(exam_number)
                    if is_nht_doc or "nht" in filename.lower():
                        parts.append("NHT")
                    final_stem = "_".join(parts) if parts else temp_path.stem
                    final_path = subject_folder / f"{final_stem}{ext}"

                    # Avoid overwrites
                    counter = 2
                    while final_path.exists():
                        final_path = subject_folder / f"{final_stem}_{counter}{ext}"
                        counter += 1

                    temp_path.rename(final_path)
                    self.file_done.emit(str(final_path))

                    with lock:
                        completed += 1
                        self.progress.emit(
                            f"Finished downloading {filename}", completed, total
                        )
                except Exception as e:
                    self.error.emit(f"Failed to download {filename}: {str(e)}")

            self.progress.emit("Starting concurrent downloads...", 0, total)

            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(download_one, link) for link in unique_links]
                concurrent.futures.wait(futures)

            self.finished.emit(f"All reports for {self.subject_name} downloaded.")
        except Exception as e:
            self.error.emit(str(e))

    @staticmethod
    def _clean_filename(name: str) -> str:
        cleaned = re.sub(r"[\\/:*?\"<>|]", "_", name).strip()
        cleaned = re.sub(r"\s+", "_", cleaned)
        return cleaned

    @staticmethod
    def _safe_subject_folder(name: str) -> str:
        s = VCAADownloadThread._clean_filename(name)
        return s or "Subject"


# ------------------ EDIT DIALOG ------------------
class EditPropertiesDialog(QDialog):
    def __init__(self, subject, year, exam_number, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Exam Properties")
        self.subject_input = QLineEdit(subject)
        self.year_input = QLineEdit(year)
        self.exam_input = QLineEdit(exam_number)
        form_layout = QFormLayout()
        form_layout.addRow("Subject:", self.subject_input)
        form_layout.addRow("Year:", self.year_input)
        form_layout.addRow("Exam Number:", self.exam_input)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout = QVBoxLayout()
        layout.addLayout(form_layout)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_values(self):
        return (
            self.subject_input.text().strip(),
            self.year_input.text().strip(),
            self.exam_input.text().strip().lower(),
        )


# ------------------ CUSTOM PDF VIEW ------------------
class NoZoomPdfView(QPdfView):
    def __init__(self, parent=None):
        super().__init__(parent)

    def wheelEvent(self, event: QWheelEvent):
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            event.ignore()
            return
        if not event.pixelDelta().isNull() or event.angleDelta().y() != 0:
            super().wheelEvent(event)
        else:
            event.ignore()


# ------------------ MAIN APP ------------------
def open_reports_folder():
    """Open the reports folder in the system file explorer."""
    folder_path = str(UPLOAD_DIR.resolve())
    if os.name == "nt":  # Windows
        subprocess.run(["explorer", folder_path], check=False)
    elif sys.platform == "darwin":  # macOS
        subprocess.run(["open", folder_path], check=False)
    else:  # Linux
        subprocess.run(["xdg-open", folder_path], check=False)


class VCEViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VCE Exam Report Viewer")
        self.resize(1280, 860)
        self.threads = []

        # Track conversion queue/active
        self._conversion_queue = []  # list[str path]
        self._conversion_active = False
        self._queued_set = set()  # to avoid duplicate enqueues
        self._progress_by_path = {}  # str(path) -> int progress (0..100)

        # Buttons
        self.upload_btn = QPushButton("Upload Reports")
        self.upload_btn.clicked.connect(self.upload_files)
        self.open_folder_btn = QPushButton("Open Reports Folder")
        self.open_folder_btn.clicked.connect(open_reports_folder)

        # Filters
        self.subject_filter = QComboBox()
        self.subject_filter.addItem("All Subjects")
        self.subject_filter.currentIndexChanged.connect(self.populate_file_list)
        self.year_filter = QComboBox()
        self.year_filter.addItem("All Years")
        self.year_filter.currentIndexChanged.connect(self.populate_file_list)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Filter by Subject:"))
        filter_layout.addWidget(self.subject_filter)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(QLabel("Filter by Year:"))
        filter_layout.addWidget(self.year_filter)

        left_layout = QVBoxLayout()
        left_layout.addWidget(self.upload_btn)
        left_layout.addWidget(self.open_folder_btn)
        left_layout.addLayout(filter_layout)

        # File list
        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self.open_file)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_context_menu)
        left_layout.addWidget(self.file_list)
        left_widget = QWidget()
        left_widget.setLayout(left_layout)

        # PDF view with zoom controls
        self.pdf_doc = QPdfDocument(self)
        self.pdf_view = NoZoomPdfView(self)
        self.pdf_view.setDocument(self.pdf_doc)
        self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        self.pdf_view.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.pdf_view.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.current_zoom = self.pdf_view.zoomFactor()  # Track current zoom factor

        # Zoom buttons
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setFixedSize(30, 30)
        self.zoom_in_btn.clicked.connect(self.zoom_in)
        self.zoom_out_btn = QPushButton("-")
        self.zoom_out_btn.setFixedSize(30, 30)
        self.zoom_out_btn.clicked.connect(self.zoom_out)

        zoom_layout = QHBoxLayout()
        zoom_layout.addWidget(QLabel("Zoom:"))
        zoom_layout.addWidget(self.zoom_in_btn)
        zoom_layout.addWidget(self.zoom_out_btn)
        zoom_layout.addStretch()

        right_layout = QVBoxLayout()
        right_layout.addLayout(zoom_layout)
        right_layout.addWidget(self.pdf_view)
        right_widget = QWidget()
        right_widget.setLayout(right_layout)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([460, 820])

        container = QWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Menu bar for VCAA Downloader
        menubar = self.menuBar()
        vcaa_menu = menubar.addMenu("VCAA")
        download_action = vcaa_menu.addAction("Download Reports")
        download_action.triggered.connect(self.open_vcaa_download_dialog)

        self.files = []  # list of dict entries
        self.current_pdf_path = None

        self.load_files()

    @staticmethod
    def _is_nht_path(path: Path) -> bool:
        stem = path.stem.lower()
        normalized = stem.replace("_", " ").replace("-", " ")
        return "nht" in stem or "northern hemisphere" in normalized

    def zoom_in(self):
        """Increase the zoom factor by 10%."""
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.current_zoom = min(self.current_zoom * 1.1, 4.0)  # Max zoom 400%
        self.pdf_view.setZoomFactor(self.current_zoom)

    def zoom_out(self):
        """Decrease the zoom factor by 10%."""
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.current_zoom = max(self.current_zoom / 1.1, 0.25)  # Min zoom 25%
        self.pdf_view.setZoomFactor(self.current_zoom)

    # ---------- FILES ----------
    def load_files(self):
        self.files.clear()
        subjects = set()
        years = set()

        # Walk subject folders, skipping CONVERTED_DIR
        for subject_folder in sorted(
                p for p in UPLOAD_DIR.iterdir() if p.is_dir() and p != CONVERTED_DIR
        ):
            for file in sorted(subject_folder.glob("*")):
                if file.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                subject, year, exam_number = parse_filename(file)
                folder_subject = subject_folder.name or subject
                subject = folder_subject

                # Compute pdf path: for .pdf, itself; for Word, converted path in CONVERTED_DIR
                if file.suffix.lower() in WORD_EXTENSIONS:
                    conv_pdf = CONVERTED_DIR / (file.stem + ".pdf")
                    pdf_path = conv_pdf if conv_pdf.exists() else None
                    # progress: keep last known, else 0
                    prog = self._progress_by_path.get(str(file), 0)
                    # if not converted and not already queued, enqueue
                    if pdf_path is None and str(file) not in self._queued_set:
                        self.enqueue_conversion(str(file))
                else:
                    pdf_path = file
                    prog = 0  # PDFs don't show a bar

                entry = {
                    "id": str(file),  # stable id by path
                    "path": file,
                    "subject": subject,
                    "year": year,
                    "exam_number": exam_number,
                    "pdf_path": pdf_path,
                    "progress": prog,
                    "is_nht": self._is_nht_path(file),
                }
                self.files.append(entry)
                subjects.add(subject)
                years.add(year)

        # Sort files by subject (asc, case-insensitive), then year (descending, with "Unknown" last), then exam_number
        def get_sort_key(entry):
            year = entry["year"]
            year_val = int(year) if year.isdigit() else 0  # Treat "Unknown" as 0
            return (
                entry["subject"].lower(),
                -year_val,
                entry["exam_number"],
            )

        self.files.sort(key=get_sort_key)

        self.update_filters(subjects, years)
        self.populate_file_list()
        # kick the queue (if any)
        self._start_next_conversion_if_idle()

    def update_filters(self, subjects, years):
        self.subject_filter.blockSignals(True)
        self.year_filter.blockSignals(True)
        self.subject_filter.clear()
        self.subject_filter.addItem("All Subjects")
        for s in sorted(subjects):
            self.subject_filter.addItem(s)
        self.year_filter.clear()
        self.year_filter.addItem("All Years")
        for y in sorted(years):
            self.year_filter.addItem(y)
        self.subject_filter.blockSignals(False)
        self.year_filter.blockSignals(False)

    def _matches_filters(self, entry):
        subj = self.subject_filter.currentText()
        year = self.year_filter.currentText()
        return (subj == "All Subjects" or entry["subject"] == subj) and (
                year == "All Years" or entry["year"] == year
        )

    def populate_file_list(self):
        # Save current scroll position
        scroll_bar = self.file_list.verticalScrollBar()
        scroll_position = scroll_bar.value() if scroll_bar else 0

        # Rebuild visible list based on filters
        self.file_list.clear()
        filtered = [e for e in self.files if self._matches_filters(e)]

        # Sort filtered entries (though self.files is already sorted, but in case)
        def get_group_key(e):
            y = int(e["year"]) if e["year"].isdigit() else 0
            return e["subject"].lower(), -y, e["exam_number"]

        filtered.sort(key=get_group_key)

        subj_filter = self.subject_filter.currentText()
        year_filter = self.year_filter.currentText()

        add_subject_headers = subj_filter == "All Subjects"
        add_year_subheaders = year_filter == "All Years"

        current_subject = None

        for entry in filtered:
            subj = entry["subject"]
            year = entry["year"]

            if add_subject_headers and subj != current_subject:
                header_item = QListWidgetItem(f"--- {subj} ---")
                header_item.setFlags(Qt.ItemFlag.NoItemFlags)
                font = header_item.font()
                font.setBold(True)
                header_item.setFont(font)
                self.file_list.addItem(header_item)
                current_subject = subj

            exam_number = entry.get("exam_number", "Unknown") or "Unknown"
            exam_display = ""
            if exam_number != "Unknown":
                lower_exam = exam_number.lower()
                if lower_exam.startswith("exam"):
                    suffix = exam_number[4:].strip()
                    exam_display = f"Exam {suffix}" if suffix else "Exam"
                else:
                    exam_display = exam_number.title()
            if entry.get("is_nht"):
                exam_display = (
                    f"{exam_display} (NHT)" if exam_display else "NHT Report"
                )

            label_parts = [entry["subject"]]
            if year != "Unknown":
                label_parts.append(year)
            if exam_display:
                label_parts.append(exam_display)
            else:
                label_parts.append("Report")
            label_text = " - ".join(label_parts)

            row_widget = QWidget()
            vbox = QVBoxLayout(row_widget)
            vbox.setContentsMargins(6, 6, 6, 6)

            lbl = QLabel(label_text)
            lbl.setStyleSheet("font-weight: 500;")
            vbox.addWidget(lbl)

            # Show a progress bar ONLY for Word files not yet converted
            show_bar = (entry["path"].suffix.lower() in WORD_EXTENSIONS) and (
                    entry["pdf_path"] is None
            )
            if show_bar:
                bar = QProgressBar()
                bar.setRange(0, 100)
                bar.setValue(int(entry.get("progress", 0)))
                vbox.addWidget(bar)

            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, entry["id"])
            self.file_list.addItem(item)
            self.file_list.setItemWidget(item, row_widget)
            item.setSizeHint(row_widget.sizeHint())

        # Restore scroll position
        if scroll_bar:
            scroll_bar.setValue(scroll_position)

    def _update_progress_ui(self, path_str: str, value: int):
        # Update stored progress and the visible bar if the row is currently displayed
        self._progress_by_path[path_str] = value
        # also update in files
        for entry in self.files:
            if entry["id"] == path_str:
                entry["progress"] = value
                break
        # Update visible row if present
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == path_str:
                widget = self.file_list.itemWidget(item)
                if widget:
                    bars = widget.findChildren(QProgressBar)
                    if bars:
                        bars[0].setValue(int(value))
                break

    def open_file(self, item):
        entry_id = item.data(Qt.ItemDataRole.UserRole)
        entry = next((f for f in self.files if f["id"] == entry_id), None)
        if not entry:
            return
        if entry["pdf_path"] and Path(entry["pdf_path"]).exists():
            self.pdf_doc.load(str(entry["pdf_path"]))
            self.current_pdf_path = str(entry["pdf_path"])
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
            self.current_zoom = self.pdf_view.zoomFactor()
        else:
            QMessageBox.information(
                self, "Not ready", "This item isn't a PDF yet (conversion pending)."
            )

    # ---------- UPLOAD ----------
    def upload_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Exam Reports",
            str(Path.home()),
            "Reports (*.pdf *.doc *.docx)",
        )
        for p in paths:
            original_name = Path(p).name
            subject, _, _ = parse_filename(Path(p))
            subj_folder_name = subject if subject != "Unknown" else "Misc"
            subject_folder = UPLOAD_DIR / subj_folder_name
            subject_folder.mkdir(exist_ok=True)
            dest = subject_folder / original_name
            if not dest.exists():
                shutil.copy(p, dest)
        self.load_files()

    # ---------- CONVERSION QUEUE ----------
    def enqueue_conversion(self, doc_path_str: str):
        # only Word files
        if Path(doc_path_str).suffix.lower() not in WORD_EXTENSIONS:
            return
        if doc_path_str in self._queued_set:
            return
        self._queued_set.add(doc_path_str)
        self._conversion_queue.append(doc_path_str)
        # ensure an initial 0% progress bar is visible for this item
        self._progress_by_path.setdefault(doc_path_str, 0)

    def _start_next_conversion_if_idle(self):
        if self._conversion_active:
            return
        if not self._conversion_queue:
            return
        self._conversion_active = True
        doc_path = self._conversion_queue.pop(0)
        # kick UI to 0 if not already set
        self._update_progress_ui(doc_path, self._progress_by_path.get(doc_path, 0))
        # Start converter
        thread = DocxConverterThread(doc_path, str(CONVERTED_DIR))
        thread.progress.connect(self._on_conv_progress)
        thread.finished.connect(self._on_conv_finished)
        thread.error.connect(self._on_conv_error)
        self.threads.append(thread)
        thread.start()

    def _on_conv_progress(self, doc_path, value):
        self._update_progress_ui(doc_path, value)

    def _on_conv_finished(self, doc_path, pdf_path):
        # Update entries: set pdf_path, progress 100, remove from queued set
        self._update_progress_ui(doc_path, 100)
        self._queued_set.discard(doc_path)
        # Reflect converted PDF in file model
        for f in self.files:
            if f["id"] == doc_path:
                f["pdf_path"] = Path(pdf_path)
                break
        # Reload files to update sorting
        self.load_files()
        # Allow next conversion
        self._conversion_active = False
        self._start_next_conversion_if_idle()

    def _on_conv_error(self, doc_path, msg):
        self._queued_set.discard(doc_path)
        QMessageBox.warning(
            self,
            "Conversion Failed",
            f"Failed to convert {Path(doc_path).name}:\n{msg}",
            QMessageBox.StandardButton.Ok,
        )
        # Reset progress to 0
        self._update_progress_ui(doc_path, 0)
        self._conversion_active = False
        self._start_next_conversion_if_idle()

    # ---------- CONTEXT ----------
    def show_context_menu(self, point: QPoint):
        item = self.file_list.itemAt(point)
        if not item:
            return
        menu = QMenu()
        edit_action = menu.addAction("Edit Properties")
        delete_action = menu.addAction("Delete Report")
        action = menu.exec(self.file_list.mapToGlobal(point))
        if action == edit_action:
            self.edit_properties(item)
        elif action == delete_action:
            self.delete_report(item)

    def edit_properties(self, item):
        entry_id = item.data(Qt.ItemDataRole.UserRole)
        entry = next((f for f in self.files if f["id"] == entry_id), None)
        if not entry:
            return
        dialog = EditPropertiesDialog(
            entry["subject"], entry["year"], entry["exam_number"], self
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_subject, new_year, new_exam = dialog.get_values()
            new_folder = UPLOAD_DIR / (new_subject if new_subject else "Misc")
            new_folder.mkdir(exist_ok=True)
            ext = entry["path"].suffix
            new_path = new_folder / f"{new_subject}_{new_year}_{new_exam}{ext}"
            try:
                # If currently open, close before rename
                if entry["pdf_path"] and self.current_pdf_path == str(
                        entry["pdf_path"]
                ):
                    self.pdf_doc.close()
                    self.current_pdf_path = None

                # Update queue bookkeeping if needed
                was_queued = entry["id"] in self._queued_set
                if was_queued:
                    # remove old path from queue list if present
                    try:
                        self._conversion_queue.remove(entry["id"])
                    except ValueError:
                        pass
                    self._queued_set.discard(entry["id"])

                # Rename source file
                old_path = entry["path"]
                entry["path"].rename(new_path)

                # If Word with converted PDF, rename converted too
                if ext.lower() in WORD_EXTENSIONS:
                    old_converted = CONVERTED_DIR / (old_path.stem + ".pdf")
                    new_converted = CONVERTED_DIR / (new_path.stem + ".pdf")
                    if old_converted.exists():
                        old_converted.rename(new_converted)
                        entry["pdf_path"] = new_converted
                    else:
                        entry["pdf_path"] = None
                else:
                    entry["pdf_path"] = new_path  # it is the PDF

                # Update entry fields
                entry["id"] = str(new_path)
                entry["subject"] = new_subject
                entry["year"] = new_year
                entry["exam_number"] = new_exam
                entry["path"] = new_path

                # carry progress state to new key
                prog = self._progress_by_path.pop(
                    str(old_path), entry.get("progress", 0)
                )
                self._progress_by_path[str(new_path)] = prog

                # re-enqueue under new path if needed
                if was_queued:
                    self.enqueue_conversion(str(new_path))

                # Reload files to update sorting
                self.load_files()

            except Exception as e:
                QMessageBox.warning(self, "Rename Failed", str(e))

    def delete_report(self, item):
        entry_id = item.data(Qt.ItemDataRole.UserRole)
        entry = next((f for f in self.files if f["id"] == entry_id), None)
        if not entry:
            return
        if entry["pdf_path"] and self.current_pdf_path == str(entry["pdf_path"]):
            self.pdf_doc.close()
            self.current_pdf_path = None
        reply = QMessageBox.question(
            self,
            "Delete Report",
            f"Delete {entry['path'].name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            try:
                # If queued, remove from queue
                if entry["id"] in self._queued_set:
                    try:
                        self._conversion_queue.remove(entry["id"])
                    except ValueError:
                        pass
                    self._queued_set.discard(entry["id"])
                # Delete source file
                if entry["path"].exists():
                    entry["path"].unlink()
                # Delete converted PDF if it exists and is distinct
                if (
                        entry["pdf_path"]
                        and Path(entry["pdf_path"]).exists()
                        and Path(entry["pdf_path"]) != entry["path"]
                ):
                    Path(entry["pdf_path"]).unlink()
            except Exception as e:
                QMessageBox.warning(self, "Delete Failed", str(e))
            self.load_files()

    # ---------- VCAA DOWNLOAD ----------
    def open_vcaa_download_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("VCAA Downloader")
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Select Subject:"))
        combo = QComboBox()
        combo.setEditable(False)
        layout.addWidget(combo)

        progress_bar = QProgressBar()
        progress_bar.setRange(0, 1)
        progress_bar.setValue(0)
        layout.addWidget(progress_bar)

        progress_label = QLabel()
        layout.addWidget(progress_label)

        log_box = QTextEdit()
        log_box.setReadOnly(True)
        log_box.setMinimumHeight(160)
        layout.addWidget(log_box)

        download_btn = QPushButton("Download Reports")
        layout.addWidget(download_btn)
        dialog.setLayout(layout)

        # Load subjects asynchronously
        def on_subjects_loaded(subjects: dict):
            combo.clear()
            sorted_subjects = sorted(subjects.values(), key=lambda v: v["label"].lower())
            for info in sorted_subjects:
                combo.addItem(info["label"], info["urls"])
            progress_label.setText(f"Loaded {len(sorted_subjects)} subjects from VCAA")

        def on_subject_error(msg):
            QMessageBox.warning(dialog, "Error", msg)

        scraper_thread = VCAASubjectScraperThread()
        scraper_thread.finished.connect(on_subjects_loaded)
        scraper_thread.error.connect(on_subject_error)
        scraper_thread.start()
        self.threads.append(scraper_thread)

        def download_selected():
            subject_name = combo.currentText()
            subject_urls = combo.currentData()
            if not subject_urls:
                QMessageBox.warning(dialog, "Error", "No subject URLs found.")
                return
            download_thread = VCAADownloadThread(subject_name, subject_urls)

            def update_progress(msg, current, total):
                progress_bar.setMaximum(total)
                progress_bar.setValue(current)
                progress_label.setText(f"{current}/{total}")
                log_box.append(msg)

            def on_finished(msg):
                log_box.append(msg)
                QMessageBox.information(dialog, "Done", msg)
                # Reload once at the end to update sorting
                self.load_files()

            download_thread.progress.connect(update_progress)
            download_thread.file_done.connect(
                lambda _: self.load_files()
            )  # Update after each file
            download_thread.finished.connect(on_finished)
            download_thread.error.connect(
                lambda msg: QMessageBox.warning(dialog, "Download Error", msg)
            )
            download_thread.start()
            self.threads.append(download_thread)

        download_btn.clicked.connect(download_selected)
        dialog.exec()


# ------------------ MAIN ------------------
def main():
    app = QApplication(sys.argv)
    window = VCEViewer()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
