"""Word, Excel and PowerPoint -> PDF with LibreOffice, tuned for Arabic files.

LibreOffice runs headless in its own throw-away profile for every file, so
conversions can run side by side and none can leave a broken profile behind.
Word files first go through `docx_fixes.prepare_docx`, which removes known
differences between how Word and LibreOffice read a file. `fonts.conf`
(installed by the Dockerfile) maps the Microsoft fonts Office files ask for to
free fonts with the same Latin metrics and a close Arabic design.
"""
from __future__ import annotations

import json
import os
import shutil
import signal
import subprocess
import tempfile
import zipfile
from pathlib import Path

from pypdf import PdfReader

from .docx_fixes import prepare_docx

MAX_FILE_MB = max(1, int(os.environ.get("MAX_UPLOAD_MB", "20")))
TIMEOUT_SECONDS = max(30, int(os.environ.get("OFFICE_TO_PDF_TIMEOUT_SECONDS", "180")))
MAX_ZIP_PARTS = 5_000
MAX_UNCOMPRESSED_BYTES = 300 * 1024 * 1024

# extension -> (part every such file has, LibreOffice PDF export filter, name)
FORMATS = {
    ".docx": ("word/document.xml", "writer_pdf_Export", "Word"),
    ".xlsx": ("xl/workbook.xml", "calc_pdf_Export", "Excel"),
    ".pptx": ("ppt/presentation.xml", "impress_pdf_Export", "PowerPoint"),
}
SUPPORTED_SUFFIXES = frozenset(FORMATS)
LEGACY = {".doc": "DOCX", ".xls": "XLSX", ".ppt": "PPTX"}

# Excel page layout: the file's own print settings, or each sheet on one page.
SHEET_LAYOUTS = ("print", "fit")


class OfficeToPdfError(ValueError):
    """Expected rejection; the message is safe to show to the user."""


def libreoffice_path() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def _check(source: Path, suffix: str) -> None:
    if suffix in LEGACY:
        raise OfficeToPdfError(
            f"صيغة {suffix.lstrip('.').upper()} القديمة غير مدعومة. "
            f"احفظ الملف بصيغة {LEGACY[suffix]} من Office أولًا."
        )
    if suffix not in FORMATS:
        raise OfficeToPdfError("الصيغ المدعومة هي DOCX وXLSX وPPTX.")
    required, _, kind = FORMATS[suffix]
    if source.stat().st_size == 0:
        raise OfficeToPdfError("الملف المرفوع فارغ.")
    if source.stat().st_size > MAX_FILE_MB * 1024 * 1024:
        raise OfficeToPdfError(f"حجم الملف أكبر من الحد المسموح ({MAX_FILE_MB} ميغابايت).")
    try:
        with zipfile.ZipFile(source) as archive:
            members = archive.infolist()
            names = {member.filename.lower() for member in members}
            if "encryptedpackage" in names or "encryptioninfo" in names:
                raise OfficeToPdfError(
                    f"ملفات {kind} المحمية بكلمة مرور غير مدعومة. أزل الحماية ثم أعد الإرسال."
                )
            if required not in names:
                raise OfficeToPdfError(f"محتوى الملف لا يطابق صيغة {kind}.")
            if len(members) > MAX_ZIP_PARTS or sum(m.file_size for m in members) > MAX_UNCOMPRESSED_BYTES:
                raise OfficeToPdfError(f"حجم ملف {kind} بعد فك الضغط يتجاوز الحد المسموح.")
    except zipfile.BadZipFile as error:
        raise OfficeToPdfError(
            f"الملف ليس ملف {kind} صالحًا. إذا كان بصيغة Office القديمة "
            "فاحفظه بالصيغة الحديثة أولًا."
        ) from error


def _filter(suffix: str, sheet_layout: str) -> str:
    name = FORMATS[suffix][1]
    if suffix == ".xlsx" and sheet_layout == "fit":
        # LibreOffice 7.4+ takes export options as JSON after the filter name.
        options = {"SinglePageSheets": {"type": "boolean", "value": "true"}}
        return f"pdf:{name}:{json.dumps(options)}"
    return f"pdf:{name}"


def _run_libreoffice(binary: str, source: Path, outdir: Path, export_filter: str) -> None:
    profile = outdir / "profile"
    command = [
        binary,
        "--headless",
        "--norestore",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        f"-env:UserInstallation={profile.as_uri()}",
        "--convert-to",
        export_filter,
        "--outdir",
        str(outdir),
        str(source),
    ]
    # The container's file system is read-only outside /tmp, so LibreOffice
    # gets the work directory as its home as well as its profile.
    environment = {**os.environ, "HOME": str(outdir)}
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=environment,
        start_new_session=True,  # its own process group, so a hang can be killed whole
    )
    try:
        _, stderr = process.communicate(timeout=TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired as error:
        os.killpg(process.pid, signal.SIGKILL)
        process.wait()
        raise OfficeToPdfError(
            "استغرق تحويل الملف وقتًا أطول من المسموح. جرّب ملفًا أصغر."
        ) from error
    if process.returncode != 0:
        raise RuntimeError(
            f"LibreOffice exited with {process.returncode}: "
            f"{stderr.decode('utf-8', 'replace')[-2000:]}"
        )


def office_to_pdf(
    source: Path,
    destination: Path,
    suffix: str | None = None,
    sheet_layout: str = "print",
) -> dict:
    """Convert one DOCX, XLSX or PPTX to PDF. Returns the page count.

    `suffix` is the uploaded file's extension when `source` was saved under
    another name. `sheet_layout` applies to Excel only: "print" follows the
    workbook's own print settings, "fit" puts each sheet on a single page.
    """
    source, destination = Path(source), Path(destination)
    suffix = (suffix or source.suffix).lower()
    _check(source, suffix)
    if sheet_layout not in SHEET_LAYOUTS:
        raise OfficeToPdfError("خيار تخطيط صفحات Excel غير معروف.")
    binary = libreoffice_path()
    if binary is None:
        raise OfficeToPdfError("خدمة التحويل إلى PDF غير مهيأة على هذا الخادم.")

    with tempfile.TemporaryDirectory(prefix="office_to_pdf_") as work:
        work = Path(work)
        prepared = work / f"document{suffix}"
        if suffix == ".docx":
            try:
                prepare_docx(source, prepared)
            except Exception:
                shutil.copyfile(source, prepared)  # an unusual file is converted as it is
        else:
            shutil.copyfile(source, prepared)
        _run_libreoffice(binary, prepared, work, _filter(suffix, sheet_layout))
        produced = work / "document.pdf"
        if not produced.is_file() or produced.stat().st_size == 0:
            raise OfficeToPdfError("تعذر تحويل الملف. تأكد من سلامته ثم حاول مجددًا.")
        shutil.move(str(produced), destination)

    try:
        pages = len(PdfReader(str(destination)).pages)
    except Exception:
        pages = 0
    return {"pages": pages}


def word_to_pdf(source: Path, destination: Path) -> dict:
    """DOCX only; kept for callers that convert Word files."""
    return office_to_pdf(source, destination, ".docx")


__all__ = [
    "OfficeToPdfError",
    "SHEET_LAYOUTS",
    "SUPPORTED_SUFFIXES",
    "libreoffice_path",
    "office_to_pdf",
    "prepare_docx",
    "word_to_pdf",
]
