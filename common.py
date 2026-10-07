"""Helpers shared by the converter (app.py) and the exam module (exam.py)."""
import shutil
import subprocess
import zipfile
from pathlib import Path

MAX_UNZIPPED_MB = 300          # zip-bomb guard for the .docx container


def pandoc_available() -> bool:
    return shutil.which("pandoc") is not None


_HELP_CACHE: dict = {}


def pandoc_supports(flag: str) -> bool:
    """True if the installed pandoc lists `flag` in --help (older versions lack some)."""
    if "help" not in _HELP_CACHE:
        try:
            _HELP_CACHE["help"] = subprocess.run(
                ["pandoc", "--help"], capture_output=True, text=True, timeout=15).stdout
        except Exception:
            _HELP_CACHE["help"] = ""
    return flag in _HELP_CACHE["help"]


def validate_docx(path: Path) -> None:
    """Raise ValueError if `path` is not a sane .docx container."""
    if not zipfile.is_zipfile(path):
        raise ValueError("That file is not a valid .docx (not a ZIP container). "
                         "Legacy .doc files must be re-saved as .docx first.")
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "word/document.xml" not in names:
            raise ValueError("This ZIP is not a Word document (word/document.xml missing).")
        if len(names) > 5000:
            raise ValueError("The document contains an unreasonable number of parts.")
        if sum(i.file_size for i in z.infolist()) > MAX_UNZIPPED_MB * 1024 * 1024:
            raise ValueError("The document expands to an unreasonable size.")
