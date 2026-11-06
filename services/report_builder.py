"""Orchestration : graphiques -> DOCX -> PDF -> vérification « une seule page A4 ».

Les mises en page (presets) sont essayées de la plus confortable à la plus
compacte. Pour chacune, le PDF est généré puis son nombre de pages vérifié ;
dès qu'un preset tient sur UNE page, il est retenu. Si aucun preset lisible ne
convient, un avertissement clair est renvoyé (et une mise en page sur 2 pages,
lisible, peut être demandée).
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any

import config
from services import chart_generator, docx_generator, pdf_converter

PAGE_H = 29.7

# De la plus confortable à la plus compacte (police minimale : 7,5 pt)
PRESETS: list[dict[str, Any]] = [
    {"name": "Standard", "layout": "stack", "title_pt": 18, "info_pt": 12, "body_pt": 9, "row_cm": 0.46,
     "chart1": (14.0, 5.0), "chart2": (12.0, 3.9), "support_cols": 1, "gap_pt": 4,
     "margin_top": 1.1, "margin_bottom": 1.0, "margin_side": 1.4},
    {"name": "Compact vertical", "layout": "stack", "title_pt": 18, "info_pt": 11.5, "body_pt": 9, "row_cm": 0.43,
     "chart1": (13.0, 4.4), "chart2": (11.0, 3.5), "support_cols": 2, "gap_pt": 3,
     "margin_top": 1.0, "margin_bottom": 0.9, "margin_side": 1.4},
    {"name": "Compact en colonnes", "layout": "grid", "title_pt": 17, "info_pt": 11, "body_pt": 8.5, "row_cm": 0.42,
     "chart1": (9.0, 4.7), "chart2": (9.0, 4.0), "support_cols": 2, "gap_pt": 3,
     "margin_top": 1.0, "margin_bottom": 0.9, "margin_side": 1.3},
    {"name": "Très compact", "layout": "grid", "title_pt": 16, "info_pt": 10, "body_pt": 8, "row_cm": 0.38,
     "chart1": (9.0, 4.0), "chart2": (9.0, 3.6), "support_cols": 2, "gap_pt": 2,
     "margin_top": 0.9, "margin_bottom": 0.8, "margin_side": 1.2},
    {"name": "Compact maximal lisible", "layout": "grid", "title_pt": 16, "info_pt": 10, "body_pt": 7.5, "row_cm": 0.35,
     "chart1": (8.8, 3.6), "chart2": (8.8, 3.3), "support_cols": 3, "gap_pt": 2,
     "margin_top": 0.8, "margin_bottom": 0.7, "margin_side": 1.2},
]

# Mise en page de secours, confortable, sur 2 pages maximum
TWO_PAGE_PRESET = dict(PRESETS[1], name="Deux pages (lisible)", body_pt=9.5, row_cm=0.5, info_pt=12,
                       chart1=(14.0, 5.5), chart2=(12.0, 4.5), support_cols=1)


def estimate_height_cm(p: dict[str, Any], n_support: int, obs_len: int) -> float:
    """Estimation grossière de la hauteur du contenu (sert uniquement à éviter des essais inutiles)."""
    rh = p["row_cm"]
    head = 1.0 + (0.8 if p["title_pt"] >= 17 else 0) + 3 * p["info_pt"] * 0.0423 + 0.15
    mastery = 0.6 + 7 * rh
    exam = 3 * rh
    obs = 0 if not obs_len else 0.55 + (obs_len / (95 if p["layout"] == "stack" else 48)) * (p["body_pt"] * 0.0423)
    cols = p["support_cols"]
    rows = -(-n_support // cols) + 1 if n_support else 1.5
    support = 0.7 + rows * rh
    if p["layout"] == "stack":
        body = p["chart1"][1] + mastery + exam + obs + p["chart2"][1] + 0.6
    else:
        body = max(p["chart1"][1], mastery) + max(p["chart2"][1], exam + obs + 0.2) + 0.3
    margins = p["margin_top"] + p["margin_bottom"]
    return margins + head + body + support


def _safe(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "", s or "")


def download_basename(meta: dict[str, str]) -> str:
    parts = ["Rapport_controle"]
    if _safe(meta.get("exam", "")):
        parts.append("N" + _safe(meta["exam"]))
    if _safe(meta.get("class", "")):
        parts.append(_safe(meta["class"]))
    return "_".join(parts)


def build_report(job_id: str, ctx: dict[str, Any], out_dir: Path | None = None,
                 allow_multipage: bool = False) -> dict[str, Any]:
    """Génère DOCX + PDF. Retourne chemins, nombre de pages, preset et avertissements."""
    out_dir = Path(out_dir or config.OUTPUT_DIR / job_id)
    out_dir.mkdir(parents=True, exist_ok=True)
    config.CHART_DIR.mkdir(parents=True, exist_ok=True)
    stats = ctx["stats"]
    obs_len = len((ctx.get("observations") or "").strip())
    n_support = stats["support_count"]
    have_pdf = pdf_converter.find_soffice() is not None
    warnings: list[str] = []

    if allow_multipage:
        candidates = [TWO_PAGE_PRESET]
    else:
        candidates = [p for p in PRESETS if estimate_height_cm(p, n_support, obs_len) <= PAGE_H + 3.5] or [PRESETS[-1]]

    result: dict[str, Any] | None = None
    last_attempt: dict[str, Any] | None = None
    for preset in candidates:
        charts = {
            "chart1": chart_generator.distribution_chart(stats, config.CHART_DIR / f"{job_id}_c1.png", *preset["chart1"]),
            "chart2": chart_generator.global_chart(stats, config.CHART_DIR / f"{job_id}_c2.png", *preset["chart2"]),
        }
        docx_path = docx_generator.generate_docx(ctx, preset, charts, out_dir / "rapport.docx")
        attempt = {"docx": docx_path, "pdf": None, "pages": None, "preset": preset["name"], "charts": charts}
        if have_pdf:
            try:
                pdf_path = pdf_converter.docx_to_pdf(docx_path, out_dir / "rapport.pdf")
                attempt["pdf"] = pdf_path
                attempt["pages"] = pdf_converter.count_pdf_pages(pdf_path)
            except pdf_converter.ConversionError as exc:
                warnings.append(str(exc))
        last_attempt = attempt
        if attempt["pages"] == 1 or allow_multipage:
            result = attempt
            break
        if attempt["pdf"] is None:  # impossible de vérifier : on ne boucle pas inutilement
            result = attempt
            break

    if result is None:
        result = last_attempt  # le plus compact lisible, mais > 1 page
    assert result is not None

    pages = result["pages"]
    fits = pages == 1
    if not have_pdf or result["pdf"] is None:
        warnings.append("Le PDF n'a pas pu être généré : installez LibreOffice (libreoffice.org) puis relancez. "
                        "Le nombre de pages n'a pas pu être vérifié.")
    elif not fits and not allow_multipage:
        warnings.append(
            f"Le contenu ne tient pas lisiblement sur une seule page A4 (il occupe {pages} pages avec la mise en page "
            "la plus compacte). Choisissez la mise en page « 2 pages (lisible) » plutôt qu'un document illisible."
        )
    elif allow_multipage and not fits:
        warnings.append(f"Rapport généré sur {pages} pages (mise en page lisible demandée).")

    base = download_basename(ctx["metadata"])
    return {"docx": result["docx"], "pdf": result["pdf"], "pages": pages, "fits_one_page": fits,
            "preset": result["preset"], "warnings": warnings, "charts": result["charts"], "basename": base,
            "can_offer_two_pages": (not fits) and (not allow_multipage) and result["pdf"] is not None}


def cleanup_job_files(job_id: str) -> None:
    shutil.rmtree(config.OUTPUT_DIR / job_id, ignore_errors=True)
    for suffix in ("_c1.png", "_c2.png"):
        (config.CHART_DIR / f"{job_id}{suffix}").unlink(missing_ok=True)
