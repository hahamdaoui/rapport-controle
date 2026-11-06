"""Conversion DOCX -> PDF (LibreOffice en mode headless, repli docx2pdf) et comptage de pages."""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path


class ConversionError(RuntimeError):
    """La conversion en PDF a échoué (message en français)."""


_WINDOWS_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    r"C:\Program Files\LibreOffice 7\program\soffice.exe",
]


def find_soffice() -> str | None:
    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).exists():
        return env
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for cand in _WINDOWS_CANDIDATES:
        if Path(cand).exists():
            return cand
    return None


def convert_with_libreoffice(src: Path, outdir: Path, target: str = "pdf", timeout: int = 180) -> Path:
    """Convertit ``src`` vers ``target`` (pdf/xlsx) dans ``outdir``."""
    soffice = find_soffice()
    if not soffice:
        raise ConversionError("LibreOffice est introuvable.")
    outdir.mkdir(parents=True, exist_ok=True)
    # Profil dédié : évite les conflits si LibreOffice est déjà ouvert
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile:
        cmd = [soffice, f"-env:UserInstallation={Path(profile).resolve().as_uri()}", "--headless",
               "--convert-to", target, "--outdir", str(outdir), str(src)]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise ConversionError("La conversion LibreOffice a dépassé le délai autorisé.") from exc
        except OSError as exc:
            raise ConversionError(f"Impossible de lancer LibreOffice : {exc}") from exc
    out = outdir / (src.stem + "." + target)
    if not out.exists():
        raise ConversionError("LibreOffice n'a pas produit de fichier : " + (proc.stderr or proc.stdout or "").strip()[:200])
    return out


def docx_to_pdf(docx_path: str | Path, pdf_path: str | Path) -> Path:
    """Convertit un DOCX en PDF. Essaie LibreOffice puis Word (docx2pdf)."""
    docx_path, pdf_path = Path(docx_path), Path(pdf_path)
    errors = []
    if find_soffice():
        try:
            with tempfile.TemporaryDirectory(prefix="pdfout_") as tmp:
                out = convert_with_libreoffice(docx_path, Path(tmp), "pdf")
                shutil.move(str(out), str(pdf_path))
            return pdf_path
        except ConversionError as exc:
            errors.append(str(exc))
    try:  # Microsoft Word installé (Windows)
        from docx2pdf import convert  # type: ignore
        convert(str(docx_path), str(pdf_path))
        if pdf_path.exists():
            return pdf_path
    except Exception as exc:  # noqa: BLE001
        errors.append(f"docx2pdf : {exc}")
    raise ConversionError(
        "Impossible de générer le PDF. Installez LibreOffice (gratuit : libreoffice.org) puis relancez l'application. "
        + (" Détail : " + " | ".join(errors) if errors else "")
    )


def count_pdf_pages(pdf_path: str | Path) -> int:
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(pdf_path)).pages)
    except Exception:  # noqa: BLE001 - repli sans dépendance
        data = Path(pdf_path).read_bytes()
        return len(re.findall(rb"/Type\s*/Page(?!s)", data))
