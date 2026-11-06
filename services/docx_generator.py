"""Génération du rapport DOCX (python-docx) – format A4 portrait.

La mise en page est pilotée par un dictionnaire de « preset » (tailles de
police, hauteurs de graphiques, hauteur de lignes, nombre de colonnes de la
liste de soutien, disposition « stack » ou « grid ») afin de pouvoir adapter
automatiquement le rapport pour qu'il tienne sur UNE page (voir report_builder).
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import _Cell

import config
from services.data_analyzer import fmt, fmt_grade, fmt_pct

FONT = "Calibri"
CS_FONT = "Arial"  # police de secours pour l'arabe (script complexe)
PLACEHOLDER = "…………………"

_ARABIC = "\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF"
_AR_SEGMENT = re.compile(rf"[{_ARABIC}](?:[{_ARABIC}\s\-\.\d]*[{_ARABIC}])?")


def rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr.lstrip("#").upper())


# ---------------------------------------------------------------------------
# Texte (gestion des segments arabes : sens RTL + police script complexe)
# ---------------------------------------------------------------------------
def _style_run(run, size: float, bold: bool, color: str | None, italic: bool, arabic: bool) -> None:
    f = run.font
    f.name = FONT
    rfonts = run._r.get_or_add_rPr().get_or_add_rFonts()
    rfonts.set(qn("w:cs"), CS_FONT)
    f.size = Pt(size)
    sz = run._r.rPr.find(qn("w:sz"))
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(round(size * 2))))
    sz.addnext(szcs)
    f.bold = bold
    f.italic = italic
    if bold:
        f.cs_bold = True
    if color:
        f.color.rgb = rgb(color)
    if arabic:
        f.rtl = True


def add_text(par, text: str, size: float = 9, bold: bool = False, color: str | None = None,
             italic: bool = False) -> None:
    """Ajoute du texte au paragraphe en séparant les segments arabes."""
    pos = 0
    for m in _AR_SEGMENT.finditer(text):
        if m.start() > pos:
            _style_run(par.add_run(text[pos:m.start()]), size, bold, color, italic, False)
        _style_run(par.add_run(m.group(0)), size, bold, color, italic, True)
        pos = m.end()
    if pos < len(text):
        _style_run(par.add_run(text[pos:]), size, bold, color, italic, False)


# ---------------------------------------------------------------------------
# Paragraphes / cellules / tableaux
# ---------------------------------------------------------------------------
def fmt_par(par, align=None, before: float = 0, after: float = 0, line_pt: float | None = None) -> None:
    pf = par.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    if line_pt:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_pt)
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.keep_together = True
    if align is not None:
        par.alignment = align


def new_par(container, **kw):
    """Nouveau paragraphe ; réutilise le paragraphe vide terminal d'une cellule."""
    par = None
    if isinstance(container, _Cell):
        last = container._tc[-1] if len(container._tc) else None
        if last is not None and last.tag == qn("w:p") and not last.findall(qn("w:r")):
            from docx.text.paragraph import Paragraph
            par = Paragraph(last, container)
    if par is None:
        par = container.add_paragraph()
    fmt_par(par, **kw)
    return par


def shrink_trailing(container) -> None:
    """Réduit au minimum le dernier paragraphe vide (suit toujours un tableau)."""
    if isinstance(container, _Cell):
        last = container._tc[-1]
        if last.tag == qn("w:p") and not last.findall(qn("w:r")):
            from docx.text.paragraph import Paragraph
            p = Paragraph(last, container)
            fmt_par(p, line_pt=1)
            r = p.add_run("")
            _style_run(r, 1, False, None, False, False)


def _tc_props(cell, fill: str | None = None, borders: dict[str, tuple[float, str] | None] | None = None,
              valign: str = "center") -> None:
    """Propriétés de cellule dans l'ordre du schéma : tcW, tcBorders, shd, vAlign."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("w:tcBorders", "w:shd", "w:vAlign"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    if borders is not None:
        b = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            spec = borders.get(side)
            e = OxmlElement(f"w:{side}")
            if spec is None:
                e.set(qn("w:val"), "nil")
            else:
                e.set(qn("w:val"), "single")
                e.set(qn("w:sz"), str(int(spec[0] * 8)))
                e.set(qn("w:space"), "0")
                e.set(qn("w:color"), spec[1])
            b.append(e)
        tcPr.append(b)
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcPr.append(shd)
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), valign)
    tcPr.append(va)


def _tbl_props(table, width_cm: float, borders: tuple[float, str] | None, cell_mar=(15, 15, 60, 60),
               align: str = "center") -> None:
    """Reconstruit tblPr dans l'ordre du schéma : tblW, jc, tblBorders, tblLayout, tblCellMar, tblLook."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for child in list(tblPr):
        tblPr.remove(child)
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"), str(int(width_cm * 567)))
    w.set(qn("w:type"), "dxa")
    tblPr.append(w)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), align)
    tblPr.append(jc)
    tb = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        if borders is None:
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(int(borders[0] * 8)))
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), borders[1])
        tb.append(e)
    tblPr.append(tb)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    mar = OxmlElement("w:tblCellMar")
    for side, val in zip(("top", "left", "bottom", "right"), (cell_mar[0], cell_mar[2], cell_mar[1], cell_mar[3])):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)
    look = OxmlElement("w:tblLook")
    look.set(qn("w:val"), "0000")
    tblPr.append(look)


def _set_widths(table, widths_cm: list[float]) -> None:
    grid = table._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)


def _row_height(row, cm: float, exact: bool = False) -> None:
    trPr = row._tr.get_or_add_trPr()
    cs = OxmlElement("w:cantSplit")
    trPr.append(cs)
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(cm * 567)))
    h.set(qn("w:hRule"), "exact" if exact else "atLeast")
    trPr.append(h)


def make_table(container, rows: int, widths_cm: list[float], borders=(0.5, config.GRID_GREY), align="center",
               cell_mar=(15, 15, 60, 60)):
    table = container.add_table(rows, len(widths_cm))
    _tbl_props(table, sum(widths_cm), borders, cell_mar, align)
    _set_widths(table, widths_cm)
    return table


def put(cell, text: str, size: float, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.CENTER, italic=False,
        before: float = 0, after: float = 0) -> None:
    par = new_par(cell, align=align, before=before, after=after)
    add_text(par, text, size, bold, color, italic)


# ---------------------------------------------------------------------------
# Sections du rapport
# ---------------------------------------------------------------------------
def _fmt_date(value: str | date | None) -> str:
    if not value:
        return PLACEHOLDER
    if isinstance(value, (date, datetime)):
        return value.strftime("%d/%m/%Y")
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(value)


def section_title(container, text: str, p: dict, before: float = 3, after: float = 2, align=WD_ALIGN_PARAGRAPH.LEFT):
    par = new_par(container, align=align, before=before, after=after)
    par.paragraph_format.keep_with_next = True
    add_text(par, text, p["body_pt"] + 1.5, True, config.BLUE)


def build_header(doc, meta: dict[str, str], p: dict, width: float) -> None:
    exam = meta.get("exam") or "….."
    sem = meta.get("semester") or "….."
    cls = meta.get("class") or "……."
    part1 = f"Rapport de contrôle N° : {exam}  |  Semestre : {sem}"
    part2 = f"Classe : {cls}"
    size = p["title_pt"]
    avail_pt = (width - 0.9) * 28.35
    est = len(part1 + "  |  " + part2) * 0.5 * size
    t = make_table(doc, 1, [width], borders=(1.5, config.BLUE), cell_mar=(50, 50, 100, 100))
    cell = t.rows[0].cells[0]
    _tc_props(cell, fill=config.LIGHT_BLUE)
    par = new_par(cell, align=WD_ALIGN_PARAGRAPH.CENTER)
    if est <= avail_pt:
        add_text(par, f"{part1}  |  {part2}", size, True, config.BLUE)
    else:  # deux lignes équilibrées plutôt qu'une réduction excessive de la police
        add_text(par, part1, size, True, config.BLUE)
        par.add_run().add_break()
        add_text(par, part2, size, True, config.BLUE)
    _row_height(t.rows[0], 0.9)

    # Informations administratives (12 pt par défaut) – deux colonnes
    items = [("Établissement", meta.get("school")), ("Professeur", meta.get("teacher")),
             ("AREF", meta.get("academy")), ("Matière", meta.get("subject")),
             ("Direction provinciale", meta.get("direction")), ("Année scolaire", meta.get("year"))]
    t = make_table(doc, 3, [width * 0.5, width * 0.5], borders=None, cell_mar=(10, 10, 60, 60), align="center")
    for i, (label, value) in enumerate(items):
        cell = t.rows[i // 2].cells[i % 2]
        _tc_props(cell, valign="center")
        par = new_par(cell, align=WD_ALIGN_PARAGRAPH.LEFT)
        add_text(par, f"{label} : ", p["info_pt"], True, config.NAVY)
        add_text(par, (value or "").strip() or PLACEHOLDER, p["info_pt"], False, "000000")
    for r in t.rows:
        _row_height(r, p["info_pt"] * 0.0353 * 1.2)


def _header_cell(cell, text: str, size: float) -> None:
    _tc_props(cell, fill=config.NAVY)
    put(cell, text, size, True, "FFFFFF")


def build_mastery_table(container, stats: dict[str, Any], p: dict, width: float) -> None:
    section_title(container, "Tableau des niveaux de maîtrise des apprentissages", p)
    fs, rh = p["body_pt"], p["row_cm"]
    narrow = width < 12  # tableau dans une demi-page (mise en page en colonnes)
    ratios = (0.25, 0.20, 0.15, 0.20, 0.20) if narrow else (0.27, 0.24, 0.15, 0.16, 0.18)
    w = [width * x for x in ratios]
    t = make_table(container, 7, w, cell_mar=(15, 15, 40, 40) if narrow else (15, 15, 60, 60))
    for c, h in enumerate(["Niveau de maîtrise", "Intervalle de notes", "Nombre d'élèves", "Pourcentage", "Moyenne du niveau"]):
        _header_cell(t.rows[0].cells[c], h, fs - (1.0 if narrow else 0.5))
    for i, lv in enumerate(stats["levels"], start=1):
        fill = config.LIGHT_GREY if i % 2 == 0 else "FFFFFF"
        vals = [lv["label"], lv["interval"], str(lv["count"]), fmt_pct(lv["percent"]), fmt(lv["mean"])]
        for c, v in enumerate(vals):
            cell = t.rows[i].cells[c]
            _tc_props(cell, fill=fill)
            put(cell, v, fs, bold=(c == 0), color=lv["color"] if c == 0 else None)
    tot = ["Total / moyenne générale", "0 – 20", str(stats["n"]), "100 %", fmt(stats["mean"])]
    for c, v in enumerate(tot):
        cell = t.rows[6].cells[c]
        _tc_props(cell, fill=config.LIGHT_BLUE)
        put(cell, v, fs, True, config.NAVY)
    for r in t.rows:
        _row_height(r, rh)
    shrink_trailing(container)


def build_exam_table(container, exam_date: str, correction_date: str, p: dict, width: float) -> None:
    fs, rh = p["body_pt"], p["row_cm"]
    t = make_table(container, 3, [width * 0.5, width * 0.5])
    _header_cell(t.rows[0].cells[0], "Informations", fs)
    _header_cell(t.rows[0].cells[1], "Dates", fs)
    for i, (lab, val) in enumerate([("Date de l'examen", exam_date), ("Date de correction", correction_date)], start=1):
        _tc_props(t.rows[i].cells[0], fill=config.LIGHT_GREY)
        put(t.rows[i].cells[0], lab, fs, True, config.NAVY)
        _tc_props(t.rows[i].cells[1])
        put(t.rows[i].cells[1], _fmt_date(val), fs)
    for r in t.rows:
        _row_height(r, rh)
    shrink_trailing(container)


def build_observations(container, text: str, p: dict, width: float) -> None:
    t = make_table(container, 1, [width], borders=(0.5, config.GRID_GREY), cell_mar=(30, 30, 80, 80))
    cell = t.rows[0].cells[0]
    _tc_props(cell, fill=config.LIGHT_GREY, valign="top")
    par = new_par(cell, align=WD_ALIGN_PARAGRAPH.LEFT, after=1)
    add_text(par, "Observations et analyse des résultats", p["body_pt"], True, config.BLUE)
    for chunk in [c.strip() for c in text.splitlines() if c.strip()]:
        par = new_par(cell, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        add_text(par, chunk, p["body_pt"] - 0.5)
    shrink_trailing(container)


def build_support_table(container, stats: dict[str, Any], p: dict, width: float) -> None:
    section_title(container, "Liste des élèves nécessitant un soutien scolaire", p, before=4)
    support = stats["support"]
    fs, rh = p["body_pt"], p["row_cm"]
    if not support:
        t = make_table(container, 1, [width], borders=(0.5, config.GRID_GREY))
        _tc_props(t.rows[0].cells[0], fill="E8F5EC")
        put(t.rows[0].cells[0], "Aucun élève ne nécessite un soutien scolaire selon les résultats de ce contrôle.",
            fs, True, "1E8449")
        _row_height(t.rows[0], rh + 0.1)
        shrink_trailing(container)
        return

    groups = max(1, min(int(p.get("support_cols", 1)), 3))
    per = -(-len(support) // groups)  # ceil
    gap = 0.25
    unit = (width - gap * (groups - 1))
    base = [0.08, 0.50, 0.14, 0.28]  # N°, nom, note, niveau
    if groups == 1:
        base = [0.07, 0.52, 0.13, 0.28]
    widths: list[float] = []
    for g in range(groups):
        widths += [unit / groups * b for b in base]
        if g < groups - 1:
            widths.append(gap)
    t = make_table(container, per + 1, widths, borders=None)
    border = {s: (0.5, config.GRID_GREY) for s in ("top", "left", "bottom", "right")}
    heads = ["N°", "Nom et prénom de l'élève", "Note /20", "Niveau de maîtrise"]
    col = 0
    for g in range(groups):
        for h in heads:
            cell = t.rows[0].cells[col]
            _tc_props(cell, fill=config.NAVY, borders=border)
            put(cell, h, fs - 0.5, True, "FFFFFF")
            col += 1
        if g < groups - 1:
            _tc_props(t.rows[0].cells[col], borders={})
            col += 1
    for g in range(groups):
        c0 = g * (len(heads) + 1)
        for r in range(per):
            idx = g * per + r
            fill = config.LIGHT_GREY if r % 2 else "FFFFFF"
            cells = [t.rows[r + 1].cells[c0 + k] for k in range(4)]
            for cell in cells:
                _tc_props(cell, fill=fill, borders=border)
            if idx < len(support):
                s = support[idx]
                put(cells[0], str(idx + 1), fs)
                put(cells[1], s["name"], fs)
                put(cells[2], fmt_grade(s["grade"]), fs, True)
                put(cells[3], s["level"], fs, False, "C0392B")
            else:
                for cell in cells:
                    put(cell, "", fs)
        if g < groups - 1:
            for r in range(per + 1):
                _tc_props(t.rows[r].cells[c0 + 4], borders={})
    for r in t.rows:
        _row_height(r, rh)
    shrink_trailing(container)


def add_picture_par(container, path: Path, w_cm: float, h_cm: float, align=WD_ALIGN_PARAGRAPH.CENTER, before=2):
    par = new_par(container, align=align, before=before, after=0)
    par.add_run().add_picture(str(path), width=Cm(w_cm), height=Cm(h_cm))


# ---------------------------------------------------------------------------
# Assemblage
# ---------------------------------------------------------------------------
def generate_docx(ctx: dict[str, Any], preset: dict[str, Any], charts: dict[str, Path], out_path: str | Path) -> Path:
    """Construit le DOCX. ``ctx`` : metadata, stats, exam_date, correction_date, observations."""
    stats, meta = ctx["stats"], ctx["metadata"]
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    side = preset.get("margin_side", 1.4)
    sec.left_margin = sec.right_margin = Cm(side)
    sec.top_margin = Cm(preset.get("margin_top", 1.1))
    sec.bottom_margin = Cm(preset.get("margin_bottom", 1.0))
    sec.header_distance = sec.footer_distance = Cm(0.3)
    width = 21.0 - 2 * side

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(preset["body_pt"])
    normal.element.rPr.rFonts.set(qn("w:cs"), CS_FONT)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")  # attribut requis par le schéma OOXML
    doc.core_properties.title = "Rapport de contrôle"
    doc.core_properties.author = meta.get("teacher") or ""
    doc.core_properties.language = "fr-FR"

    obs = (ctx.get("observations") or "").strip()
    build_header(doc, meta, preset, width)

    def spacer(pt: float = 3):
        par = doc.add_paragraph()
        fmt_par(par, line_pt=pt)
        r = par.add_run("")
        _style_run(r, 1, False, None, False, False)

    spacer(preset.get("gap_pt", 3))

    if preset["layout"] == "stack":
        c1w, c1h = preset["chart1"]
        add_picture_par(doc, charts["chart1"], c1w, c1h, before=0)
        build_mastery_table(doc, stats, preset, width)
        spacer(preset.get("gap_pt", 3))
        build_exam_table(doc, ctx["exam_date"], ctx["correction_date"], preset, width * 0.55)
        if obs:
            spacer(preset.get("gap_pt", 3))
            build_observations(doc, obs, preset, width)
        c2w, c2h = preset["chart2"]
        add_picture_par(doc, charts["chart2"], c2w, c2h, before=preset.get("gap_pt", 3))
    else:  # grid : graphiques à côté des tableaux (ordre de lecture conservé)
        gap = 0.3
        half = (width - gap) / 2
        outer = make_table(doc, 2, [half, gap, half], borders=None, cell_mar=(0, 0, 0, 0))
        for r in outer.rows:
            for c in r.cells:
                _tc_props(c, valign="center")
        c1w, c1h = preset["chart1"]
        add_picture_par(outer.rows[0].cells[0], charts["chart1"], min(c1w, half), c1h, before=0)
        build_mastery_table(outer.rows[0].cells[2], stats, preset, half)
        left = outer.rows[1].cells[0]
        _tc_props(left, valign="center")
        build_exam_table(left, ctx["exam_date"], ctx["correction_date"], preset, half)
        if obs:
            new_par(left, line_pt=preset.get("gap_pt", 3))
            build_observations(left, obs, preset, half)
        c2w, c2h = preset["chart2"]
        add_picture_par(outer.rows[1].cells[2], charts["chart2"], min(c2w, half), c2h, before=0)
        for r in outer.rows:
            shrink_trailing(r.cells[0])
            shrink_trailing(r.cells[2])
        spacer(1)

    build_support_table(doc, stats, preset, width)
    # Dernier paragraphe obligatoire après un tableau : réduit à 1 pt pour ne jamais créer une 2e page.
    last = doc.add_paragraph()
    fmt_par(last, line_pt=1)
    _style_run(last.add_run(""), 1, False, None, False, False)

    out_path = Path(out_path)
    doc.save(out_path)
    return out_path
