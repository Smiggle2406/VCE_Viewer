import sys
import os
import shutil
import subprocess
import re
import tempfile
from pathlib import Path
from urllib.parse import urljoin
import concurrent.futures
import threading

import shutil as _shutil
import requests
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
    if getattr(sys, "frozen", False):
        base_dir = Path.home() / "Documents" / "VCEViewer"
    else:
        base_dir = Path.cwd()
    upload_dir = base_dir / "uploaded_reports"
    return upload_dir


UPLOAD_DIR = get_upload_dir()
CONVERTED_DIR = UPLOAD_DIR / "converted"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
CONVERTED_DIR.mkdir(exist_ok=True)
SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".docx"}
WORD_EXTENSIONS = {".doc", ".docx"}

VCAA_BASE = "https://www.vcaa.vic.edu.au"
VCAA_SUBJECTS_PAGE = (
        VCAA_BASE
        + "/assessment/vce/examination-specifications-past-examinations-and-examination-reports/"
        + "examination-specifications-past-examinations-and-external-assessment-reports"
)

EXCLUDE_HINTS = [
    "sample",
    "formula",
    "data book",
    "data-book",
    "databook",
    "assessment guide",
    "transcript",
]

REPORT_TOKEN = "report"

# ------------------ SUBJECT NORMALISATION ------------------
SUBJECT_ALIASES = {
    "mathmethodscas": "MathMethodsCAS",
    "mathematicalmethods": "MathMethods",
    "mathmethods": "MathMethods",
    "mmcas": "MathMethodsCAS",
    "maths1": "MathMethods",
    "mm": "MathMethods",
    "mmcas2": "MathMethodsCAS",
    "specialist": "SpecialistMaths",
    "sm": "SpecialistMaths",
    "chemistry": "Chemistry",
    "chem": "Chemistry",
}


# ------------------ UTILS ------------------
def soffice_cmd():
    """
    Return the path to the LibreOffice CLI binary for headless mode, or None if not found.
    Checks Windows, macOS, and Linux paths, as well as system PATH.
    """
    possible_paths = []
    if os.name == "nt":  # Windows
        possible_paths.extend(
            [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                r"C:\LibreOffice\program\soffice.exe",
                r"D:\LibreOffice\program\soffice.exe",
                "soffice.exe",
            ]
        )
    else:  # macOS and Linux
        possible_paths.extend(
            [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS standard
                "/usr/local/bin/soffice",  # macOS Homebrew or Linux
                "/opt/local/bin/soffice",  # macOS alternative
                "/usr/bin/soffice",  # Common Linux
                "soffice",  # System PATH
                "libreoffice",  # Fallback
            ]
        )

    for path in possible_paths:
        if _shutil.which(path) or os.path.exists(path):
            return path
    return None


# ------------------ PARSING ------------------
def parse_filename(file_path: Path):
    """
    Parse (best-effort) subject, year and exam number from a file path's name.
    Returns: (subject, year, exam_number)
    """
    name = file_path.stem.lower()
    name = re.sub(
        r"[-_\s]?(assessrep|examreport|examrep|externalassessmentreport|report|exam)",
        "",
        name,
    )
    name = re.sub(r"\s*\(\d+\)", "", name)
    name = name.strip()

    year = "Unknown"
    year_match = re.search(r"(20\d{2})", name)
    if year_match:
        year = year_match.group(1)
        name = name.replace(year, "").strip("-_ ")
    else:
        trailing_digit = re.search(r"(\d{2})$", name)
        if trailing_digit:
            y = int(trailing_digit.group(1))
            if y <= 30:
                year = f"20{y:02d}"
                name = re.sub(r"\d{2}$", "", name).strip("-_ ")

    exam_number = "Unknown"
    ex_match = re.search(r"(?:ex|exam)?[-_]?([12])\b", name)
    if ex_match:
        exam_number = f"exam{ex_match.group(1)}"
        name = re.sub(r"(?:ex|exam)?[-_]?[12]\b", "", name).strip("-_ ")
    else:
        trailing_digit = re.search(r"(\d)$", name)
        if trailing_digit:
            exam_number = f"exam{trailing_digit.group(1)}"
            name = re.sub(r"\d$", "", name).strip("-_ ")

    subject = "Unknown"
    for key in SUBJECT_ALIASES:
        if key in name:
            subject = SUBJECT_ALIASES[key]
            break
    if subject == "Unknown" and name:
        subject = name.title()

    return subject, year, exam_number


# ------------------ DOCX CONVERTER ------------------
class DocxConverterThread(QThread):
    progress = pyqtSignal(str, int)
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str, str)

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
            self.progress.emit(self.docx_path, 60)
            cmd = [
                soffice,
                "--headless",
                "--nologo",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                str(self.output_dir),
                str(self.docx_path),
            ]
            subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True,
                shell=(os.name == "nt"),  # Shell only for Windows
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

    def run(self):
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            resp = requests.get(
                VCAA_SUBJECTS_PAGE, headers=headers, timeout=30, verify=False
            )
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")
            subjects = {}
            for link in soup.find_all("a", href=True):
                text = link.get_text(strip=True)
                href = link["href"].strip()
                if not href:
                    continue
                full = urljoin(VCAA_BASE, href)
                path_lower = full.lower()
                if "/vce-vet-" in path_lower:
                    continue
                if (
                        "/assessment/vce/examination-specifications-past-examinations-and-examination-reports/"
                        in path_lower
                ):
                    if full.rstrip("/") != VCAA_SUBJECTS_PAGE.rstrip("/") and text:
                        subjects[text] = full
            if not subjects:
                self.error.emit("No subjects found on the VCAA index page.")
            else:
                self.finished.emit(subjects)
        except Exception as e:
            self.error.emit(str(e))


class VCAADownloadThread(QThread):
    progress = pyqtSignal(str, int, int)
    file_done = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, subject_name, subject_url):
        super().__init__()
        self.subject_name = subject_name
        self.subject_url = subject_url

    @staticmethod
    def _should_skip(link_href: str, link_text: str) -> bool:
        h = (link_href or "").lower()
        t = (link_text or "").lower()
        if any(k in h or k in t for k in EXCLUDE_HINTS):
            return True
        if REPORT_TOKEN not in t:
            return True
        if not (h.endswith(".pdf") or h.endswith(".docx") or h.endswith(".doc")):
            return True
        return False

    def run(self):
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            resp = requests.get(
                self.subject_url, headers=headers, timeout=30, verify=False
            )
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")

            links = []
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                text = a.get_text(strip=True)
                if not href:
                    continue
                if self._should_skip(href, text):
                    continue
                links.append(urljoin(VCAA_BASE, href))

            total = len(links)
            if total == 0:
                self.finished.emit(
                    f"No examination reports found for {self.subject_name}."
                )
                return

            subject_folder = UPLOAD_DIR / self._safe_subject_folder(self.subject_name)
            subject_folder.mkdir(parents=True, exist_ok=True)

            completed = 0
            lock = threading.Lock()

            def download_one(file_url):
                nonlocal completed
                filename = file_url.split("/")[-1]
                try:
                    r = requests.get(
                        file_url, headers=headers, timeout=120, verify=False
                    )
                    r.raise_for_status()

                    temp_path = subject_folder / filename
                    with open(temp_path, "wb") as f:
                        f.write(r.content)

                    _, year, exam_number = parse_filename(temp_path)
                    ext = temp_path.suffix.lower()
                    parts = [self._clean_filename(self.subject_name)]
                    if year and year != "Unknown":
                        parts.append(year)
                    if exam_number and exam_number != "Unknown":
                        parts.append(exam_number)
                    final_stem = "_".join(parts) if parts else temp_path.stem
                    final_path = subject_folder / f"{final_stem}{ext}"

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
                futures = [executor.submit(download_one, url) for url in links]
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
        # Allow vertical scrolling, ignore for zooming
        if event.angleDelta().y() != 0:
            super().wheelEvent(event)  # Process scrolling
        else:
            event.ignore()


# ------------------ MAIN APP ------------------
class VCEViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VCE Exam Report Viewer")
        self.resize(1280, 860)
        self.threads = []

        # Track conversion queue/active
        self._conversion_queue = []
        self._conversion_active = False
        self._queued_set = set()
        self._progress_by_path = {}

        # Buttons
        self.upload_btn = QPushButton("Upload Reports")
        self.upload_btn.clicked.connect(self.upload_files)
        self.open_folder_btn = QPushButton("Open Reports Folder")
        self.open_folder_btn.clicked.connect(self.open_reports_folder)

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
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.current_zoom = 1.0
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)

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

        self.files = []
        self.current_pdf_path = None
        self.temp_pdf_path = None

        self.load_files()

    def zoom_in(self):
        """Increase the zoom factor by 10%."""
        self.current_zoom = min(self.current_zoom * 1.1, 4.0)
        self.pdf_view.setZoomFactor(self.current_zoom)

    def zoom_out(self):
        """Decrease the zoom factor by 10%."""
        self.current_zoom = max(self.current_zoom / 1.1, 0.25)
        self.pdf_view.setZoomFactor(self.current_zoom)

    def open_reports_folder(self):
        """Open the reports folder in the system file explorer."""
        folder_path = str(UPLOAD_DIR.resolve())
        if os.name == "nt":
            subprocess.run(["explorer", folder_path], check=False)
        elif sys.platform == "darwin":
            subprocess.run(["open", folder_path], check=False)
        else:
            subprocess.run(["xdg-open", folder_path], check=False)

    def load_files(self):
        self.files.clear()
        subjects = set()
        years = set()

        for subject_folder in sorted(
                p for p in UPLOAD_DIR.iterdir() if p.is_dir() and p != CONVERTED_DIR
        ):
            for file in sorted(subject_folder.glob("*")):
                if file.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                subject, year, exam_number = parse_filename(file)
                folder_subject = subject_folder.name or subject
                subject = folder_subject

                if file.suffix.lower() in WORD_EXTENSIONS:
                    conv_pdf = CONVERTED_DIR / (file.stem + ".pdf")
                    pdf_path = conv_pdf if conv_pdf.exists() else None
                    prog = self._progress_by_path.get(str(file), 0)
                    if pdf_path is None and str(file) not in self._queued_set:
                        self.enqueue_conversion(str(file))
                else:
                    pdf_path = file
                    prog = 0

                entry = {
                    "id": str(file),
                    "path": file,
                    "subject": subject,
                    "year": year,
                    "exam_number": exam_number,
                    "pdf_path": pdf_path,
                    "progress": prog,
                }
                self.files.append(entry)
                subjects.add(subject)
                years.add(year)

        def get_sort_key(entry):
            year = entry["year"]
            year_val = int(year) if year.isdigit() else 0
            return (
                entry["subject"].lower(),
                -year_val,
                entry["exam_number"],
            )

        self.files.sort(key=get_sort_key)

        self.update_filters(subjects, years)
        self.populate_file_list()
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
        scroll_bar = self.file_list.verticalScrollBar()
        scroll_position = scroll_bar.value() if scroll_bar else 0

        self.file_list.clear()
        filtered = [e for e in self.files if self._matches_filters(e)]

        def get_group_key(e):
            y = int(e["year"]) if e["year"].isdigit() else 0
            return e["subject"].lower(), -y, e["exam_number"]

        filtered.sort(key=get_group_key)

        subj_filter = self.subject_filter.currentText()
        year_filter = self.year_filter.currentText()

        add_subject_headers = subj_filter == "All Subjects"
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

            label_text = (
                f"{entry['subject']}_{entry['year']}_{entry['exam_number']}.pdf"
            )

            row_widget = QWidget()
            vbox = QVBoxLayout(row_widget)
            vbox.setContentsMargins(6, 6, 6, 6)

            lbl = QLabel(label_text)
            lbl.setStyleSheet("font-weight: 500;")
            vbox.addWidget(lbl)

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

        if scroll_bar:
            scroll_bar.setValue(scroll_position)

    def _update_progress_ui(self, path_str: str, value: int):
        self._progress_by_path[path_str] = value
        for entry in self.files:
            if entry["id"] == path_str:
                entry["progress"] = value
                break
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
            if self.temp_pdf_path and Path(self.temp_pdf_path).exists():
                try:
                    Path(self.temp_pdf_path).unlink()
                except:
                    pass
            self.temp_pdf_path = None
            self.pdf_doc.load(str(entry["pdf_path"]))
            self.current_pdf_path = str(entry["pdf_path"])
            self.pdf_view.setZoomFactor(self.current_zoom)
        else:
            QMessageBox.information(
                self, "Not ready", "This item isn't a PDF yet (conversion pending)."
            )

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

    def enqueue_conversion(self, doc_path_str: str):
        if Path(doc_path_str).suffix.lower() not in WORD_EXTENSIONS:
            return
        if doc_path_str in self._queued_set:
            return
        self._queued_set.add(doc_path_str)
        self._conversion_queue.append(doc_path_str)
        self._progress_by_path.setdefault(doc_path_str, 0)

    def _start_next_conversion_if_idle(self):
        if self._conversion_active:
            return
        if not self._conversion_queue:
            return
        self._conversion_active = True
        doc_path = self._conversion_queue.pop(0)
        self._update_progress_ui(doc_path, self._progress_by_path.get(doc_path, 0))
        thread = DocxConverterThread(doc_path, str(CONVERTED_DIR))
        thread.progress.connect(self._on_conv_progress)
        thread.finished.connect(self._on_conv_finished)
        thread.error.connect(self._on_conv_error)
        self.threads.append(thread)
        thread.start()

    def _on_conv_progress(self, doc_path, value):
        self._update_progress_ui(doc_path, value)

    def _on_conv_finished(self, doc_path, pdf_path):
        self._update_progress_ui(doc_path, 100)
        self._queued_set.discard(doc_path)
        for f in self.files:
            if f["id"] == doc_path:
                f["pdf_path"] = Path(pdf_path)
                break
        self.load_files()
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
        self._update_progress_ui(doc_path, 0)
        self._conversion_active = False
        self._start_next_conversion_if_idle()

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
                temp_file = None
                if entry["pdf_path"] and self.current_pdf_path == str(
                        entry["pdf_path"]
                ):
                    temp_file = (
                            Path(tempfile.gettempdir())
                            / f"vce_temp_{Path(entry['pdf_path']).name}"
                    )
                    shutil.copy(entry["pdf_path"], temp_file)
                    self.pdf_doc.close()
                    self.pdf_doc.load(str(temp_file))
                    self.current_pdf_path = str(temp_file)
                    self.temp_pdf_path = str(temp_file)
                    self.pdf_view.setZoomFactor(self.current_zoom)

                was_queued = entry["id"] in self._queued_set
                if was_queued:
                    try:
                        self._conversion_queue.remove(entry["id"])
                    except ValueError:
                        pass
                    self._queued_set.discard(entry["id"])

                old_path = entry["path"]
                entry["path"].rename(new_path)

                if ext.lower() in WORD_EXTENSIONS:
                    old_converted = CONVERTED_DIR / (old_path.stem + ".pdf")
                    new_converted = CONVERTED_DIR / (new_path.stem + ".pdf")
                    if old_converted.exists():
                        old_converted.rename(new_converted)
                        entry["pdf_path"] = new_converted
                    else:
                        entry["pdf_path"] = None
                else:
                    entry["pdf_path"] = new_path

                entry["id"] = str(new_path)
                entry["subject"] = new_subject
                entry["year"] = new_year
                entry["exam_number"] = new_exam
                entry["path"] = new_path

                prog = self._progress_by_path.pop(
                    str(old_path), entry.get("progress", 0)
                )
                self._progress_by_path[str(new_path)] = prog

                if was_queued:
                    self.enqueue_conversion(str(new_path))

                if temp_file and entry["pdf_path"] and Path(entry["pdf_path"]).exists():
                    self.pdf_doc.close()
                    self.pdf_doc.load(str(entry["pdf_path"]))
                    self.current_pdf_path = str(entry["pdf_path"])
                    self.temp_pdf_path = None
                    self.pdf_view.setZoomFactor(self.current_zoom)
                    try:
                        temp_file.unlink()
                    except:
                        pass

                self.load_files()

            except Exception as e:
                if temp_file and temp_file.exists():
                    self.pdf_doc.close()
                    self.pdf_doc.load(str(entry["pdf_path"]))
                    self.current_pdf_path = str(entry["pdf_path"])
                    self.temp_pdf_path = None
                    self.pdf_view.setZoomFactor(self.current_zoom)
                    try:
                        temp_file.unlink()
                    except:
                        pass
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
                if entry["id"] in self._queued_set:
                    try:
                        self._conversion_queue.remove(entry["id"])
                    except ValueError:
                        pass
                    self._queued_set.discard(entry["id"])
                if entry["path"].exists():
                    entry["path"].unlink()
                if (
                        entry["pdf_path"]
                        and Path(entry["pdf_path"]).exists()
                        and Path(entry["pdf_path"]) != entry["path"]
                ):
                    Path(entry["pdf_path"]).unlink()
            except Exception as e:
                QMessageBox.warning(self, "Delete Failed", str(e))
            self.load_files()

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

        def on_subjects_loaded(subjects: dict):
            combo.clear()
            for s in sorted(subjects.keys()):
                combo.addItem(s, subjects[s])
            progress_label.setText(f"Loaded {len(subjects)} subjects from VCAA")

        def on_subject_error(msg):
            QMessageBox.warning(dialog, "Error", msg)

        scraper_thread = VCAASubjectScraperThread()
        scraper_thread.finished.connect(on_subjects_loaded)
        scraper_thread.error.connect(on_subject_error)
        scraper_thread.start()
        self.threads.append(scraper_thread)

        def download_selected():
            subject_name = combo.currentText()
            subject_url = combo.currentData()
            if not subject_url:
                QMessageBox.warning(dialog, "Error", "No subject URL found.")
                return
            download_thread = VCAADownloadThread(subject_name, subject_url)

            def update_progress(msg, current, total):
                progress_bar.setMaximum(total)
                progress_bar.setValue(current)
                progress_label.setText(f"{current}/{total}")
                log_box.append(msg)

            def on_finished(msg):
                log_box.append(msg)
                QMessageBox.information(dialog, "Done", msg)
                self.load_files()

            download_thread.progress.connect(update_progress)
            download_thread.file_done.connect(lambda _: self.load_files())
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
