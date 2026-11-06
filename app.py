"""Application Flask : analyse des résultats d'examens et génération du rapport A4 (DOCX + PDF).

Lancement :  python app.py   (ou start.bat sous Windows)  ->  http://127.0.0.1:5000
"""
from __future__ import annotations

import logging
import re
import shutil
import threading
import time
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException, RequestEntityTooLarge

import config
from services import data_analyzer, excel_processor, pdf_converter, report_builder
from services.data_analyzer import AnalysisError
from services.excel_processor import ExcelReadError

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = config.MAX_CONTENT_LENGTH
log = logging.getLogger("rapport")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

TOKEN_RE = re.compile(r"^[a-f0-9]{32}$")
META_KEYS = ["school", "academy", "direction", "teacher", "subject", "class", "semester", "exam", "year"]
MAX_OBS_CHARS = 1200
_generate_lock = threading.Lock()  # une génération à la fois (LibreOffice + Matplotlib)
_last_cleanup = 0.0


class UserError(Exception):
    """Erreur affichable telle quelle à l'utilisateur."""

    def __init__(self, message: str, status: int = 400):
        super().__init__(message)
        self.message, self.status = message, status


# ---------------------------------------------------------------------------
# Nettoyage automatique des fichiers temporaires
# ---------------------------------------------------------------------------
def cleanup_old_files(force: bool = False) -> None:
    global _last_cleanup
    now = time.time()
    if not force and now - _last_cleanup < 600:
        return
    _last_cleanup = now
    rules = [(config.UPLOAD_DIR, config.UPLOAD_TTL_HOURS), (config.OUTPUT_DIR, config.OUTPUT_TTL_HOURS),
             (config.CHART_DIR, config.OUTPUT_TTL_HOURS)]
    for folder, hours in rules:
        if not folder.exists():
            continue
        for item in folder.iterdir():
            if item.name.startswith("."):
                continue
            try:
                if now - item.stat().st_mtime > hours * 3600:
                    shutil.rmtree(item, ignore_errors=True) if item.is_dir() else item.unlink(missing_ok=True)
            except OSError:
                pass


@app.before_request
def _before() -> None:
    for d in (config.UPLOAD_DIR, config.OUTPUT_DIR, config.CHART_DIR):
        d.mkdir(parents=True, exist_ok=True)
    cleanup_old_files()


# ---------------------------------------------------------------------------
# Utilitaires
# ---------------------------------------------------------------------------
def _upload_path(token: str) -> Path:
    if not isinstance(token, str) or not TOKEN_RE.match(token):
        raise UserError("Fichier introuvable. Veuillez importer à nouveau votre fichier Excel.", 404)
    for ext in (".xlsx", ".xls"):
        p = config.UPLOAD_DIR / f"{token}{ext}"
        if p.exists():
            return p
    raise UserError("Le fichier importé a expiré ou a été supprimé. Veuillez l'importer à nouveau.", 404)


def _sniff_extension(head: bytes) -> str | None:
    if head.startswith(b"PK\x03\x04"):
        return ".xlsx"
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return ".xls"
    return None


def _clean_mapping(raw: Any) -> dict | None:
    if not isinstance(raw, dict):
        return None
    out: dict[str, Any] = {}
    try:
        if raw.get("sheet"):
            out["sheet"] = str(raw["sheet"])
        if raw.get("name_cols"):
            out["name_cols"] = [int(c) for c in raw["name_cols"]][:2]
        if raw.get("grade_col") not in (None, ""):
            out["grade_col"] = int(raw["grade_col"])
        if "absence_col" in raw:
            out["absence_col"] = int(raw["absence_col"]) if raw["absence_col"] not in (None, "") else ""
        if raw.get("first_row"):
            out["first_row"] = max(1, int(raw["first_row"]))
    except (TypeError, ValueError) as exc:
        raise UserError("Correspondance de colonnes invalide.") from exc
    return out or None


def _clean_scale(raw: Any) -> float | None:
    if raw in (None, ""):
        return None
    try:
        value = float(str(raw).replace(",", "."))
    except ValueError as exc:
        raise UserError("Le barème doit être un nombre (par exemple 20).") from exc
    if value not in (5, 10, 20, 40, 50, 100):
        raise UserError("Barème non pris en charge. Valeurs possibles : 5, 10, 20, 40, 50 ou 100.")
    return value


def _process(token: str, mapping: dict | None, scale: float | None, filename: str | None = None) -> dict[str, Any]:
    result = excel_processor.process_excel(_upload_path(token), filename=filename, mapping=mapping, scale=scale)
    stats = None
    if result["has_valid"]:
        stats = data_analyzer.analyze(result["students"])
    result["stats"] = stats
    result["token"] = token
    return result


def _parse_iso_date(value: Any, label: str) -> date:
    if not value:
        raise UserError(f"Veuillez renseigner la {label}.")
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except ValueError as exc:
        raise UserError(f"La {label} n'est pas une date valide.") from exc


@app.errorhandler(UserError)
def _handle_user_error(err: UserError):
    return jsonify(ok=False, error=err.message), err.status


@app.errorhandler(RequestEntityTooLarge)
def _too_large(_err):
    return jsonify(ok=False, error=f"Le fichier est trop volumineux (maximum {config.MAX_UPLOAD_MB} Mo)."), 413


@app.errorhandler(HTTPException)
def _http_error(err: HTTPException):
    if request.path.startswith("/api/") or request.path.startswith("/download/"):
        return jsonify(ok=False, error="Ressource introuvable ou requête invalide."), err.code or 400
    return err


@app.errorhandler(Exception)
def _unexpected(err: Exception):
    log.exception("Erreur inattendue")
    return jsonify(ok=False, error="Une erreur inattendue est survenue. Vérifiez votre fichier et réessayez."), 500


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
@app.get("/")
def index():
    return render_template("index.html", max_mb=config.MAX_UPLOAD_MB, max_obs=MAX_OBS_CHARS)


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
@app.post("/api/upload")
def api_upload():
    file = request.files.get("file")
    if file is None or not file.filename:
        raise UserError("Aucun fichier reçu. Sélectionnez un fichier Excel (.xlsx ou .xls).")
    original = Path(file.filename).name
    if Path(original).suffix.lower() not in config.ALLOWED_EXTENSIONS:
        raise UserError("Format non valide. Seuls les fichiers Excel (.xlsx, .xls) sont acceptés.")
    head = file.stream.read(8)
    file.stream.seek(0)
    ext = _sniff_extension(head)
    if ext is None:
        raise UserError("Ce fichier n'est pas un classeur Excel valide (contenu non reconnu).")
    token = uuid.uuid4().hex
    dest = config.UPLOAD_DIR / f"{token}{ext}"
    file.save(dest)
    try:
        payload = _process(token, None, None, filename=original)
    except (ExcelReadError, AnalysisError) as exc:
        dest.unlink(missing_ok=True)
        raise UserError(str(exc)) from exc
    except Exception:
        dest.unlink(missing_ok=True)
        raise
    return jsonify(ok=True, **payload)


@app.post("/api/analyze")
def api_analyze():
    body = request.get_json(silent=True) or {}
    try:
        payload = _process(body.get("token", ""), _clean_mapping(body.get("mapping")), _clean_scale(body.get("scale")),
                           filename=str(body.get("filename") or "")[:200] or None)
    except (ExcelReadError, AnalysisError) as exc:
        raise UserError(str(exc)) from exc
    return jsonify(ok=True, **payload)


@app.post("/api/suggest")
def api_suggest():
    body = request.get_json(silent=True) or {}
    try:
        payload = _process(body.get("token", ""), _clean_mapping(body.get("mapping")), _clean_scale(body.get("scale")))
    except (ExcelReadError, AnalysisError) as exc:
        raise UserError(str(exc)) from exc
    if not payload["stats"]:
        raise UserError("Aucune note valide : impossible de proposer une observation.")
    return jsonify(ok=True, text=data_analyzer.suggest_observation(payload["stats"]))


@app.post("/api/generate")
def api_generate():
    body = request.get_json(silent=True) or {}
    exam = _parse_iso_date(body.get("exam_date"), "date de l'examen")
    corr = _parse_iso_date(body.get("correction_date"), "date de correction")
    if corr < exam:
        raise UserError("La date de correction ne peut pas être antérieure à la date de l'examen.")
    obs = str(body.get("observations") or "").strip()
    if len(obs) > MAX_OBS_CHARS:
        raise UserError(f"Les observations sont trop longues (maximum {MAX_OBS_CHARS} caractères) pour tenir sur une page.")

    try:
        payload = _process(body.get("token", ""), _clean_mapping(body.get("mapping")), _clean_scale(body.get("scale")))
    except (ExcelReadError, AnalysisError) as exc:
        raise UserError(str(exc)) from exc
    if not payload["stats"]:
        raise UserError("Aucune note valide n'a été trouvée : le rapport ne peut pas être généré.")

    meta = dict(payload["metadata"])
    user_meta = body.get("metadata") if isinstance(body.get("metadata"), dict) else {}
    for key in META_KEYS:
        if key in user_meta:  # l'utilisateur peut corriger ou compléter (jamais d'invention côté serveur)
            meta[key] = re.sub(r"\s+", " ", str(user_meta[key] or "")).strip()[:120]

    ctx = {"metadata": meta, "stats": payload["stats"], "exam_date": exam.isoformat(),
           "correction_date": corr.isoformat(), "observations": obs}
    job = uuid.uuid4().hex
    with _generate_lock:
        try:
            result = report_builder.build_report(job, ctx, allow_multipage=bool(body.get("allow_multipage")))
        except Exception:
            report_builder.cleanup_job_files(job)
            raise

    (config.OUTPUT_DIR / job / "name.txt").write_text(result["basename"], encoding="utf-8")
    return jsonify(
        ok=True, job=job, stats=payload["stats"], metadata=meta, pages=result["pages"],
        fits_one_page=result["fits_one_page"], layout=result["preset"], warnings=result["warnings"],
        can_offer_two_pages=result["can_offer_two_pages"],
        docx_url=f"/download/{job}/docx", pdf_url=f"/download/{job}/pdf" if result["pdf"] else None,
        chart1_url=f"/static/charts/{job}_c1.png", chart2_url=f"/static/charts/{job}_c2.png",
        counts=payload["counts"],
    )


@app.delete("/api/file/<token>")
def api_delete_file(token: str):
    try:
        _upload_path(token).unlink(missing_ok=True)
    except UserError:
        pass
    return jsonify(ok=True)


@app.get("/api/status")
def api_status():
    """Indique si la conversion PDF est disponible (affiché dans l'interface)."""
    return jsonify(ok=True, pdf_available=pdf_converter.find_soffice() is not None)


@app.get("/download/<job>/<fmt>")
def download(job: str, fmt: str):
    if not TOKEN_RE.match(job) or fmt not in ("docx", "pdf"):
        raise UserError("Fichier introuvable.", 404)
    path = config.OUTPUT_DIR / job / f"rapport.{fmt}"
    if not path.exists():
        raise UserError("Ce rapport n'est plus disponible. Veuillez le générer à nouveau.", 404)
    name_file = config.OUTPUT_DIR / job / "name.txt"
    base = name_file.read_text(encoding="utf-8").strip() if name_file.exists() else "Rapport_controle"
    mime = ("application/vnd.openxmlformats-officedocument.wordprocessingml.document" if fmt == "docx"
            else "application/pdf")
    return send_file(path, as_attachment=True, download_name=f"{base}.{fmt}", mimetype=mime)


if __name__ == "__main__":
    for d in (config.UPLOAD_DIR, config.OUTPUT_DIR, config.CHART_DIR):
        d.mkdir(parents=True, exist_ok=True)
    cleanup_old_files(force=True)
    print(f"\n  Application prête : http://{config.HOST}:{config.PORT}\n  (Ctrl+C pour arrêter)\n")
    app.run(host=config.HOST, port=config.PORT, debug=False, threaded=True)
