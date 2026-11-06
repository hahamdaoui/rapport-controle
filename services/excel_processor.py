"""Lecture et validation des fichiers Excel de résultats d'examens.

Fonctionnalités principales
---------------------------
* lecture de TOUTES les feuilles (.xlsx via OpenPyXL, .xls via xlrd ou, à défaut,
  conversion LibreOffice) ;
* détection automatique des colonnes « nom » et « note » (en-têtes arabes ou
  français, y compris en-têtes sur plusieurs lignes, cellules fusionnées…) ;
* extraction des métadonnées (établissement, AREF, direction, professeur,
  matière, classe, semestre, année scolaire, numéro de contrôle) ;
* détection du barème ;
* validation ligne par ligne (note manquante, absent, invalide, hors barème,
  doublon) – jamais de note inventée, une note vide n'est jamais comptée 0 ;
* correspondance manuelle des colonnes (``mapping``) quand la détection échoue.

Le résultat est un dictionnaire JSON-sérialisable.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Normalisation de texte (arabe + français)
# --------------------------------------------------------------------------
_ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_ARABIC_LETTER_MAP = str.maketrans({"ى": "ي", "ة": "ه", "ؤ": "و", "ئ": "ي", "ک": "ك", "ی": "ي"})


def norm(text: Any) -> str:
    """Forme canonique d'un libellé : sans accents, voyelles, tatweel,
    ponctuation ni espaces ; hamzas d'alif unifiées ; minuscules."""
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text)).translate(_ARABIC_DIGITS)
    s = s.replace("\u0640", "")  # tatweel
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.translate(_ARABIC_LETTER_MAP).lower()
    return re.sub(r"[^\w]|_", "", s)


def clean_text(value: Any) -> str:
    """Nettoie un texte affichable : espaces multiples, tatweel, ':' final."""
    if value is None:
        return ""
    s = str(value).replace("\u0640", "").replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.rstrip(":").strip()


# --------------------------------------------------------------------------
# Libellés des métadonnées
# --------------------------------------------------------------------------
_LABELS: dict[str, list[str]] = {
    "academy": ["أكاديمية", "الأكاديمية", "أكاديمية جهوية", "AREF", "Académie", "Académie régionale"],
    "direction": ["م.الإقليمية", "المديرية الإقليمية", "المديرية", "مديرية", "Direction provinciale",
                  "Direction", "Province", "Délégation", "DP"],
    "school": ["مؤسسة", "المؤسسة", "Établissement", "Etablissement", "Ecole", "École"],
    "level": ["المستوى", "Niveau"],
    "class": ["القسم", "الفوج", "Classe", "Class"],
    "teacher": ["الأستاذ", "الاستاذ", "الأستاذ(ة)", "أستاذ", "Professeur", "Prof", "Enseignant", "Enseignante"],
    "subject": ["المادة", "Matière", "Matiere", "Discipline"],
    "semester": ["الدورة", "Semestre", "Session"],
    "exam": ["نقط", "الفرض", "Contrôle", "Controle", "Devoir", "Examen"],
    "year": ["السنة الدراسية", "Année scolaire", "Annee scolaire", "Année", "Annee"],
}
_INLINE_ONLY: dict[str, list[str]] = {"direction": ["إقليم", "Province"]}

_LABEL_NORM: dict[str, set[str]] = {k: {norm(x) for x in v} for k, v in _LABELS.items()}
_ALL_LABELS: set[str] = set().union(*_LABEL_NORM.values())
_INLINE_NORM: dict[str, set[str]] = {k: {norm(x) for x in v} for k, v in _INLINE_ONLY.items()}

_ORDINALS = {
    "1": ["الاول", "اولي", "premier", "premiere", "1er", "first"],
    "2": ["الثاني", "ثانيه", "deuxieme", "second", "seconde", "2eme", "2nd"],
    "3": ["الثالث", "ثالثه", "troisieme", "3eme"],
    "4": ["الرابع", "رابعه", "quatrieme", "4eme"],
    "5": ["الخامس", "خامسه", "cinquieme", "5eme"],
}


def ordinal_to_number(text: str) -> str:
    """« الدورة الأولى » -> « 1 », « الفرض الثاني » -> « 2 ». Texte brut si inconnu."""
    n = norm(text)
    m = re.search(r"\d+", n)
    if m:
        return m.group(0)
    for num, words in _ORDINALS.items():
        if any(norm(w) in n for w in words):
            return num
    return clean_text(text)


# --------------------------------------------------------------------------
# Lecture des classeurs
# --------------------------------------------------------------------------
class ExcelReadError(Exception):
    """Erreur de lecture du fichier (message en français destiné à l'utilisateur)."""


def _pad(matrix: list[list[Any]]) -> list[list[Any]]:
    width = max((len(r) for r in matrix), default=0)
    return [list(r) + [None] * (width - len(r)) for r in matrix]


def _trim(matrix: list[list[Any]]) -> list[list[Any]]:
    """Supprime les lignes vides finales."""
    while matrix and all(v is None or str(v).strip() == "" for v in matrix[-1]):
        matrix.pop()
    return matrix


def _read_xlsx(path: Path) -> dict[str, list[list[Any]]]:
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as exc:  # fichier corrompu, mot de passe…
        raise ExcelReadError(
            "Impossible de lire ce fichier Excel. Il est peut-être corrompu ou protégé par un mot de passe."
        ) from exc
    out = {}
    for ws in wb.worksheets:
        out[ws.title] = _pad(_trim([list(r) for r in ws.iter_rows(values_only=True)]))
    wb.close()
    return out


def _read_xls(path: Path) -> dict[str, list[list[Any]]]:
    try:
        import pandas as pd
        sheets = pd.read_excel(path, sheet_name=None, header=None, engine="xlrd", dtype=object)
    except ImportError:
        return _read_xls_via_libreoffice(path)
    except Exception as exc:
        raise ExcelReadError("Impossible de lire ce fichier .xls. Enregistrez-le au format .xlsx puis réessayez.") from exc
    out = {}
    for name, df in sheets.items():
        rows = [[None if (v is None or v != v) else v for v in row] for row in df.values.tolist()]
        out[str(name)] = _pad(_trim(rows))
    return out


def _read_xls_via_libreoffice(path: Path) -> dict[str, list[list[Any]]]:
    """Repli quand xlrd n'est pas installé : conversion .xls -> .xlsx par LibreOffice."""
    import tempfile
    from services.pdf_converter import convert_with_libreoffice, ConversionError
    with tempfile.TemporaryDirectory() as tmp:
        try:
            converted = convert_with_libreoffice(path, Path(tmp), target="xlsx")
        except ConversionError as exc:
            raise ExcelReadError(
                "Le format .xls nécessite le module « xlrd » ou LibreOffice. "
                "Installez xlrd (pip install xlrd) ou enregistrez le fichier au format .xlsx."
            ) from exc
        return _read_xlsx(converted)


def load_workbook_matrices(path: str | Path) -> dict[str, list[list[Any]]]:
    path = Path(path)
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(path)
    if ext == ".xls":
        return _read_xls(path)
    raise ExcelReadError("Format non pris en charge. Utilisez un fichier .xlsx ou .xls.")


# --------------------------------------------------------------------------
# Détection des métadonnées
# --------------------------------------------------------------------------
def _is_empty(v: Any) -> bool:
    return v is None or str(v).strip() == ""


def _match_field(cell_norm: str, table: dict[str, set[str]]) -> str | None:
    for field, labels in table.items():
        if cell_norm in labels:
            return field
    return None


def extract_metadata(matrix: list[list[Any]], limit_rows: int | None = None) -> dict[str, str]:
    """Cherche des paires « libellé -> valeur » dans les lignes d'en-tête."""
    found: dict[str, str] = {}
    rows = matrix if limit_rows is None else matrix[:limit_rows]
    for row in rows:
        for c, cell in enumerate(row):
            if _is_empty(cell) or isinstance(cell, (int, float)):
                continue
            text = str(cell)
            cn = norm(text)
            field = _match_field(cn, _LABEL_NORM)
            if field and field not in found:
                for v in row[c + 1:]:
                    if _is_empty(v):
                        continue
                    if norm(v) in _ALL_LABELS:
                        break  # la valeur est absente : on tombe sur le libellé suivant
                    found[field] = clean_text(v)
                    break
                continue
            # « Libellé : valeur » dans une même cellule
            if ":" in text or "：" in text:
                left, right = re.split(r"[:：]", text, maxsplit=1)
                f2 = _match_field(norm(left), _LABEL_NORM) or _match_field(norm(left), _INLINE_NORM)
                if f2 and f2 not in found and clean_text(right):
                    found[f2] = clean_text(right)
    # Nettoyages spécifiques
    if "direction" in found:
        found["direction"] = re.sub(r"^\s*(إقليم|اقليم|province)\s*[:：]?\s*", "", found["direction"], flags=re.I).strip()
    if "semester" in found:
        found["semester"] = ordinal_to_number(found["semester"])
    if "exam" in found:
        found["exam"] = ordinal_to_number(found["exam"])
    if "year" not in found:
        for row in rows:
            for cell in row:
                m = re.search(r"\b(20\d{2})\s*[/\-]\s*(20\d{2})\b", str(cell)) if not _is_empty(cell) else None
                if m:
                    found["year"] = f"{m.group(1)}/{m.group(2)}"
                    break
            if "year" in found:
                break
    return found


# --------------------------------------------------------------------------
# Détection des colonnes
# --------------------------------------------------------------------------
_NAME_LATIN = re.compile(
    r"^(nom|prenom|prenoms|nometprenom|nomprenom|prenomnom|nomcomplet|eleve|eleves|etudiant|"
    r"nomdeleleve|nomdeletudiant|name|fullname|student|studentname)$"
)
_GRADE_EXACT = {norm(x) for x in ["النقطة", "نقطة", "النقط", "Note", "Notes", "Moyenne", "Score", "Mark", "Marks",
                                   "Grade", "العلامة", "Note/20", "Note /20", "Note sur 20", "Points"]}
_GRADE_CONTAINS = [norm(x) for x in ["نقطة", "note", "moyenne", "grade", "score", "علامة", "معدل"]]
_GRADE_WEAK = [norm(x) for x in ["الفرض", "controle", "devoir", "examen", "ds", "cc"]]
_ABSENCE = [norm(x) for x in ["غياب", "تغيب", "absence", "absent", "abs", "مبرر", "justif"]]
_SUMMARY_NAMES = {norm(x) for x in ["Moyenne", "Moyenne de la classe", "Moyenne générale", "المعدل", "المعدل العام",
                                     "معدل القسم", "Effectif", "Total", "المجموع"]}
_HEADER_NAME_WORDS = {norm(x) for x in ["إسم التلميذ", "اسم التلميذ", "الاسم", "Nom", "Prénom", "Nom et prénom", "Élève"]}


def _is_name_header(text: Any) -> bool:
    n = norm(text)
    if not n:
        return False
    if "رقم" in n or "تاريخ" in n:
        return False
    if "اسم" in n and "ابو" not in n:
        return True
    if n in ("التلميذ", "التلميذه", "التلاميذ"):
        return True
    return bool(_NAME_LATIN.match(n))


def _grade_header_score(text: Any) -> int:
    n = norm(text)
    if not n:
        return 0
    if any(a in n for a in _ABSENCE):
        return -1
    if n in _GRADE_EXACT:
        return 3
    if any(k in n for k in _GRADE_CONTAINS):
        return 2
    if any(k == n or (len(k) > 2 and k in n) for k in _GRADE_WEAK):
        return 1
    return 0


def _find_header(matrix: list[list[Any]], max_scan: int = 80) -> tuple[int | None, list[int]]:
    """Retourne (ligne d'en-tête du nom, colonnes de noms)."""
    for r, row in enumerate(matrix[:max_scan]):
        cols = [c for c, v in enumerate(row) if not _is_empty(v) and not isinstance(v, (int, float)) and _is_name_header(v)]
        if cols:
            # « Nom » + « Prénom » dans deux colonnes distinctes -> on les fusionne
            ns = {c: norm(row[c]) for c in cols}
            if len(cols) >= 2 and {"nom", "prenom"} <= set(ns.values()):
                nom = next(c for c, n in ns.items() if n == "nom")
                pre = next(c for c, n in ns.items() if n == "prenom")
                return r, [nom, pre]
            return r, [cols[0]]
    return None, []


def _header_block_end(matrix: list[list[Any]], hr: int, name_cols: list[int]) -> int:
    """Dernière ligne appartenant à l'en-tête (en-têtes sur 2 lignes, fusionnés…)."""
    end = hr
    for r in range(hr + 1, min(hr + 3, len(matrix))):
        row = matrix[r]
        if all(_is_empty(v) for v in row):
            break
        name_cell = row[name_cols[0]] if name_cols and name_cols[0] < len(row) else None
        has_grade_hdr = any(_grade_header_score(v) != 0 for v in row if isinstance(v, str))
        if _is_empty(name_cell) or norm(name_cell) in _HEADER_NAME_WORDS:
            if has_grade_hdr or _is_empty(name_cell):
                end = r
                continue
        break
    return end


def _numeric_like(v: Any) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        return bool(re.fullmatch(r"\s*\d+([.,]\d+)?\s*", v.translate(_ARABIC_DIGITS)))
    return False


def _detect_grade_column(matrix, hr, hend, name_cols):
    """Retourne (colonne, colonne_absence, ambigu)."""
    ncols = len(matrix[0]) if matrix else 0
    scores: dict[int, int] = {}
    absence_col = None
    for c in range(ncols):
        if c in name_cols:
            continue
        col_scores = [_grade_header_score(matrix[r][c]) for r in range(max(0, hr - 1), hend + 1) if isinstance(matrix[r][c], str)]
        if any(s < 0 for s in col_scores):
            if absence_col is None:
                absence_col = c
            continue
        best = max(col_scores, default=0)
        if best > 0:
            scores[c] = best
    if not scores:
        return None, absence_col, True
    top = max(scores.values())
    tops = [c for c, s in scores.items() if s == top]
    return tops[0], absence_col, len(tops) > 1


def _guess_grade_column_by_content(matrix, hend, name_cols):
    best, best_n = None, 0
    for c in range(len(matrix[0]) if matrix else 0):
        if c in name_cols:
            continue
        n = sum(1 for row in matrix[hend + 1:] if _numeric_like(row[c]))
        if n > best_n:
            best, best_n = c, n
    return best


def _guess_name_column_by_content(matrix, start=0):
    best, best_n = None, 0
    for c in range(len(matrix[0]) if matrix else 0):
        n = sum(1 for row in matrix[start:] if isinstance(row[c], str) and re.search(r"[^\W\d_]{2,}", row[c]) and len(row[c]) > 3)
        if n > best_n:
            best, best_n = c, n
    return best


# --------------------------------------------------------------------------
# Notes
# --------------------------------------------------------------------------
_ABSENT_TOKENS = {norm(x) for x in ["abs", "absent", "absente", "ab", "غ", "غائب", "غائبة", "غياب"]}
_MISSING_TOKENS = {"", "-", "--", "—", "na", "nan", "n/a", "null", "none", "؟", "?"}


def parse_grade(value: Any) -> tuple[str, float | None, str]:
    """Interprète une cellule de note.

    Retourne (statut, valeur, message) avec statut ∈ {numeric, missing, absent, invalid}.
    """
    if value is None:
        return "missing", None, ""
    if isinstance(value, bool):
        return "invalid", None, "Valeur booléenne au lieu d'une note"
    if isinstance(value, (int, float)):
        if value != value:
            return "missing", None, ""
        if value < 0:
            return "invalid", None, "Note négative"
        return "numeric", float(value), ""
    text = str(value).strip()
    if text.lower() in _MISSING_TOKENS:
        return "missing", None, ""
    if norm(text) in _ABSENT_TOKENS:
        return "absent", None, ""
    t = text.translate(_ARABIC_DIGITS).replace("،", ".").replace(",", ".").replace(" ", "").replace("\xa0", "")
    t = re.sub(r"/\d+(\.\d+)?$", "", t)
    if re.fullmatch(r"\d+(\.\d+)?", t):
        return "numeric", float(t), ""
    return "invalid", None, f"Valeur non numérique : « {text[:30]} »"


_STANDARD_SCALES = [5, 10, 20, 40, 50, 100]


def detect_scale(matrix, hend, grade_col, explicit_user: float | None = None) -> tuple[float, str]:
    """(barème, source) – source ∈ explicit | inferred | default | user."""
    if explicit_user:
        return float(explicit_user), "user"
    pattern = re.compile(r"(?:/|\bsur\b|على|\bout of\b)\s*(\d{1,3})(?!\d)", re.I)
    for r in range(0, min(hend + 1, len(matrix))):
        for c, v in enumerate(matrix[r]):
            if isinstance(v, str) and (grade_col is None or abs(c - grade_col) <= 1 or r >= hend - 1):
                m = pattern.search(v.translate(_ARABIC_DIGITS))
                if m and int(m.group(1)) in _STANDARD_SCALES and not re.search(r"\d{4}", v):
                    return float(m.group(1)), "explicit"
    values = []
    if grade_col is not None:
        for row in matrix[hend + 1:]:
            st, val, _ = parse_grade(row[grade_col])
            if st == "numeric":
                values.append(val)
    if values:
        over = sum(1 for x in values if x > 20)
        if over / len(values) >= 0.3:
            m = max(values)
            for s in (40, 50, 100):
                if m <= s:
                    return float(s), "inferred"
    return 20.0, "default"


# --------------------------------------------------------------------------
# Analyse d'une feuille
# --------------------------------------------------------------------------
def _scan_sheet(name: str, matrix: list[list[Any]], mapping: dict | None, scale_user: float | None) -> dict[str, Any]:
    mapping = mapping or {}
    hr, name_cols = _find_header(matrix)
    auto_ambiguous = False
    notes: list[str] = []

    if mapping.get("name_cols"):
        name_cols = [int(c) for c in mapping["name_cols"]]
        if hr is None:
            hr = 0
        # on cherche tout de même un en-tête pour estimer le début des données
    elif hr is None:
        guessed = _guess_name_column_by_content(matrix)
        if guessed is not None:
            name_cols, hr, auto_ambiguous = [guessed], -1, True
            notes.append("Colonne des noms devinée d'après le contenu.")
    hend = _header_block_end(matrix, hr, name_cols) if hr is not None and hr >= 0 and name_cols else (hr if hr is not None else -1)

    grade_col, absence_col, ambiguous = None, None, False
    if matrix and name_cols:
        grade_col, absence_col, ambiguous = _detect_grade_column(matrix, max(hr, 0), max(hend, 0), name_cols)
        if grade_col is None:
            guessed = _guess_grade_column_by_content(matrix, hend, name_cols)
            if guessed is not None:
                grade_col, ambiguous = guessed, True
                notes.append("Colonne des notes devinée d'après le contenu.")
    if mapping.get("grade_col") is not None:
        grade_col, ambiguous = int(mapping["grade_col"]), False
    if mapping.get("absence_col") is not None:
        absence_col = int(mapping["absence_col"]) if mapping["absence_col"] != "" else None
    first_row = hend + 1 if hend is not None else 0
    if mapping.get("first_row"):
        first_row = max(0, int(mapping["first_row"]) - 1)

    scale, scale_source = detect_scale(matrix, first_row - 1, grade_col, scale_user)

    students: list[dict[str, Any]] = []
    if name_cols and grade_col is not None:
        for r in range(first_row, len(matrix)):
            row = matrix[r]
            raw_name = " ".join(clean_text(row[c]) for c in name_cols if c < len(row) and not _is_empty(row[c])).strip()
            raw_grade = row[grade_col] if grade_col < len(row) else None
            absent_flag = absence_col is not None and absence_col < len(row) and not _is_empty(row[absence_col])
            if not raw_name and _is_empty(raw_grade):
                continue  # ligne vide
            rec = {"sheet": name, "row": r + 1, "name": raw_name,
                   "raw_grade": "" if raw_grade is None else str(raw_grade).strip(),
                   "grade": None, "grade20": None, "status": "valid", "message": ""}
            if raw_name and norm(raw_name) in _SUMMARY_NAMES:
                continue  # ligne de synthèse (moyenne, total…)
            if not raw_name:
                rec.update(status="invalid", message="Nom de l'élève manquant")
            elif _numeric_like(raw_name):
                rec.update(status="invalid", message="Nom invalide (valeur numérique)")
            else:
                st, val, msg = parse_grade(raw_grade)
                if st == "missing" and absent_flag:
                    st = "absent"
                if st == "numeric":
                    if val > scale:
                        rec.update(status="invalid", message=f"Note supérieure au barème (/{_fmt_scale(scale)})", grade=val)
                    else:
                        rec["grade"] = val
                        rec["grade20"] = val if scale == 20 else round(val * 20.0 / scale, 6)
                elif st == "missing":
                    rec.update(status="missing", message="Note manquante")
                elif st == "absent":
                    rec.update(status="absent", message="Élève absent")
                else:
                    rec.update(status="invalid", message=msg)
            students.append(rec)

    # Doublons : même nom ET même note => doublon exact ignoré
    seen: dict[tuple[str, float], int] = {}
    names_to_grades: dict[str, set[float]] = {}
    for s in students:
        if s["status"] != "valid":
            continue
        key = (norm(s["name"]), s["grade"])
        if key in seen:
            s.update(status="duplicate", message=f"Doublon exact de la ligne {seen[key]} (ignoré)")
            s["grade20"] = None
        else:
            seen[key] = s["row"]
            names_to_grades.setdefault(norm(s["name"]), set()).add(s["grade"])
    same_name_diff_grade = sorted({s["name"] for s in students
                                   if s["status"] == "valid" and len(names_to_grades.get(norm(s["name"]), ())) > 1})

    n_valid = sum(1 for s in students if s["status"] == "valid")
    score = (1000 if (hr is not None and hr >= 0) else 0) + (500 if grade_col is not None else 0) + len(students) + 2 * n_valid
    return {"sheet": name, "hr": hr, "hend": hend, "name_cols": name_cols, "grade_col": grade_col,
            "absence_col": absence_col, "first_row": first_row, "scale": scale, "scale_source": scale_source,
            "students": students, "ambiguous": bool(ambiguous or auto_ambiguous), "notes": notes,
            "same_name_diff_grade": same_name_diff_grade, "score": score}


def _fmt_scale(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else str(v)


def _columns_info(matrix: list[list[Any]], scan: dict[str, Any]) -> list[dict[str, Any]]:
    hr = max(scan["hr"] if scan["hr"] is not None else 0, 0)
    hend = max(scan["hend"], hr)
    info = []
    for c in range(len(matrix[0]) if matrix else 0):
        header_parts = []
        for r in range(max(0, hr - 1), hend + 1):
            v = matrix[r][c]
            if not _is_empty(v) and clean_text(v) not in header_parts:
                header_parts.append(clean_text(v))
        samples = []
        for row in matrix[scan["first_row"]:]:
            if not _is_empty(row[c]):
                samples.append(clean_text(row[c])[:25])
            if len(samples) >= 3:
                break
        if header_parts or samples:
            info.append({"index": c, "letter": get_column_letter(c + 1), "header": " / ".join(header_parts)[:60],
                         "samples": samples})
    return info


# --------------------------------------------------------------------------
# API publique
# --------------------------------------------------------------------------
def process_excel(path: str | Path, filename: str | None = None, mapping: dict | None = None,
                  scale: float | None = None) -> dict[str, Any]:
    """Analyse complète d'un fichier Excel (voir le module pour le format du résultat)."""
    path = Path(path)
    matrices = load_workbook_matrices(path)
    if not matrices:
        raise ExcelReadError("Le fichier Excel ne contient aucune feuille.")

    requested = (mapping or {}).get("sheet")
    if requested and requested in matrices:
        scans = {requested: _scan_sheet(requested, matrices[requested], mapping, scale)}
        for n, m in matrices.items():
            if n != requested:
                scans[n] = _scan_sheet(n, m, None, scale)
        best = requested
    else:
        scans = {n: _scan_sheet(n, m, None, scale) for n, m in matrices.items()}
        best = max(scans, key=lambda n: scans[n]["score"])
    scan = scans[best]
    matrix = matrices[best]

    # Métadonnées : feuille retenue d'abord, puis les autres feuilles pour les champs manquants
    head_rows = max(scan["hend"], 0) + 1 if scan["hr"] is not None else 40
    metadata = extract_metadata(matrix, limit_rows=max(head_rows, 40))
    for n, m in matrices.items():
        if n == best:
            continue
        for k, v in extract_metadata(m, limit_rows=40).items():
            metadata.setdefault(k, v)
    meta_full = {k: metadata.get(k, "") for k in
                 ["school", "academy", "direction", "teacher", "subject", "level", "class", "semester", "exam", "year"]}

    students = scan["students"]
    counts = {"rows": len(students)}
    for st in ("valid", "absent", "missing", "invalid", "duplicate"):
        counts[st] = sum(1 for s in students if s["status"] == st)

    issues: list[dict[str, str]] = []

    def add(level, msg):
        issues.append({"level": level, "message": msg})

    for n in scan["notes"]:
        add("warning", n)
    if scan["hr"] is None and not scan["name_cols"]:
        add("error", "Colonne des noms introuvable : choisissez les colonnes manuellement ci-dessous.")
    if scan["grade_col"] is None:
        add("error", "Colonne des notes introuvable : choisissez-la manuellement ci-dessous.")
    elif scan["ambiguous"]:
        add("warning", "Plusieurs colonnes de notes possibles : vérifiez la colonne sélectionnée.")
    if not students and scan["grade_col"] is not None:
        add("error", "Aucune ligne d'élève n'a été trouvée sous l'en-tête.")
    if students and counts["valid"] == 0:
        if counts["missing"] == len(students):
            add("error", "Aucune note n'a été saisie dans ce fichier : la colonne des notes est vide pour les "
                         f"{counts['missing']} élèves. Renseignez les notes dans Excel (ou choisissez une autre "
                         "colonne) puis importez à nouveau le fichier.")
        else:
            add("error", "Aucune note valide n'a été trouvée dans la colonne sélectionnée.")
    if counts["missing"] and counts["valid"]:
        add("warning", f"{counts['missing']} élève(s) sans note : ils sont exclus des statistiques (non comptés comme 0).")
    if counts["absent"]:
        add("info", f"{counts['absent']} élève(s) absent(s) : exclus des statistiques.")
    if counts["invalid"]:
        add("warning", f"{counts['invalid']} enregistrement(s) invalide(s) exclu(s) (voir le détail ci-dessous).")
    if counts["duplicate"]:
        add("warning", f"{counts['duplicate']} doublon(s) exact(s) ignoré(s) pour éviter un double comptage.")
    if scan["same_name_diff_grade"]:
        add("warning", "Même nom avec des notes différentes (conservés tous les deux) : "
                       + ", ".join(scan["same_name_diff_grade"][:5]))
    if scan["scale_source"] == "inferred":
        add("warning", f"Barème détecté automatiquement : /{_fmt_scale(scan['scale'])}. Les notes sont ramenées à "
                       "/20 pour la classification.")
    elif scan["scale"] != 20:
        add("info", f"Barème : /{_fmt_scale(scan['scale'])}. Les notes sont ramenées à /20 pour la classification.")
    for key, label in (("school", "l'établissement"), ("teacher", "le professeur"), ("class", "la classe")):
        if not meta_full[key]:
            add("info", f"Information non trouvée dans le fichier : {label}. Vous pouvez la saisir ci-dessous.")

    needs_mapping = scan["grade_col"] is None or not scan["name_cols"] or scan["ambiguous"] or not students

    return {
        "filename": filename or path.name,
        "sheet": best,
        "sheets": list(matrices.keys()),
        "mapping": {"sheet": best, "name_cols": scan["name_cols"], "grade_col": scan["grade_col"],
                    "absence_col": scan["absence_col"], "first_row": scan["first_row"] + 1},
        "columns": _columns_info(matrix, scan),
        "metadata": meta_full,
        "scale": scan["scale"],
        "scale_source": scan["scale_source"],
        "students": students,
        "counts": counts,
        "issues": issues,
        "needs_mapping": bool(needs_mapping),
        "has_valid": counts["valid"] > 0,
    }
