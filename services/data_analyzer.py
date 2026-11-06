"""Calculs statistiques et classification des élèves par niveau de maîtrise.

Toutes les notes manipulées ici sont ramenées sur 20. Seuls les élèves ayant
une note valide (statut « valid ») sont comptés : une note vide n'est jamais 0.
"""
from __future__ import annotations

from typing import Any, Iterable

import numpy as np

import config

LEVELS = config.MASTERY_LEVELS


class AnalysisError(ValueError):
    """Données insuffisantes ou incohérentes pour produire des statistiques."""


def classify(grade: float) -> str:
    """Clé du niveau de maîtrise (intervalles mutuellement exclusifs).

    0 ≤ n < 10 | 10 ≤ n < 12 | 12 ≤ n < 14 | 14 ≤ n < 16 | 16 ≤ n ≤ 20
    """
    if grade is None or grade != grade or grade < 0 or grade > 20:
        raise AnalysisError(f"Note hors intervalle [0 ; 20] : {grade!r}")
    for key, _label, low, high, *_ in LEVELS[:-1]:
        if low <= grade < high:
            return key
    return LEVELS[-1][0]  # 16 ≤ n ≤ 20


def level_label(key: str) -> str:
    return next(lv[1] for lv in LEVELS if lv[0] == key)


def classify_global(grade: float) -> str:
    """Classification en trois groupes : Maîtrisé (≥ 12) / En cours (10–<12) / Non maîtrisé (< 10)."""
    if grade < config.SUPPORT_THRESHOLD:
        return "non_maitrise"
    if grade < config.MASTERED_THRESHOLD:
        return "en_cours"
    return "maitrise"


def _round(x: float | None, nd: int = 2) -> float | None:
    return None if x is None else round(float(x), nd)


def valid_students(students: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return [s for s in students if s.get("status") == "valid" and s.get("grade20") is not None]


def analyze(students: Iterable[dict[str, Any]]) -> dict[str, Any]:
    """Calcule toutes les statistiques nécessaires au rapport et à l'aperçu."""
    valid = valid_students(students)
    n = len(valid)
    if n == 0:
        raise AnalysisError("Aucun élève avec une note valide : impossible de calculer des statistiques.")

    grades = np.array([float(s["grade20"]) for s in valid], dtype=float)
    keys = [classify(g) for g in grades]

    levels = []
    for key, label, _lo, _hi, interval, bar_label, color in LEVELS:
        mask = np.array([k == key for k in keys])
        count = int(mask.sum())
        levels.append({
            "key": key, "label": label, "interval": interval, "bar_label": bar_label, "color": color,
            "count": count,
            "percent": _round(100.0 * count / n, 1),
            "mean": _round(grades[mask].mean(), 2) if count else None,  # uniquement les élèves du niveau
        })

    g3 = {"maitrise": 0, "en_cours": 0, "non_maitrise": 0}
    for g in grades:
        g3[classify_global(g)] += 1

    support = sorted(
        [{"name": s["name"], "grade": float(s["grade20"]), "level": level_label(classify(s["grade20"]))}
         for s in valid if s["grade20"] < config.SUPPORT_THRESHOLD],
        key=lambda x: (x["grade"], x["name"]),
    )

    stats = {
        "n": n,
        "mean": _round(grades.mean(), 2),
        "median": _round(np.median(grades), 2),
        "std": _round(grades.std(), 2),
        "max": _round(grades.max(), 2),
        "min": _round(grades.min(), 2),
        "pass_count": int((grades >= 10).sum()),
        "pass_rate": _round(100.0 * (grades >= 10).sum() / n, 1),
        "levels": levels,
        "global": [{"key": k, "label": lbl, "color": col, "count": g3[k], "percent": _round(100.0 * g3[k] / n, 1)}
                   for k, lbl, col in config.GLOBAL_LEVELS],
        "support": support,
        "support_count": len(support),
    }
    check_consistency(stats)
    return stats


def check_consistency(stats: dict[str, Any]) -> None:
    """Garde-fous : les totaux doivent toujours égaler le nombre d'élèves valides."""
    n = stats["n"]
    if sum(lv["count"] for lv in stats["levels"]) != n:
        raise AnalysisError("Incohérence : la somme des cinq niveaux ne correspond pas à l'effectif.")
    if sum(g["count"] for g in stats["global"]) != n:
        raise AnalysisError("Incohérence : la somme des trois groupes ne correspond pas à l'effectif.")
    if stats["support_count"] != stats["levels"][0]["count"]:
        raise AnalysisError("Incohérence : élèves en soutien ≠ élèves « Non maîtrisé ».")


# --------------------------------------------------------------------------
# Formatage à la française
# --------------------------------------------------------------------------
def fmt(value: float | None, nd: int = 2, dash: str = "—") -> str:
    """12.5 -> « 12,50 » ; None -> « — »."""
    if value is None:
        return dash
    return f"{value:.{nd}f}".replace(".", ",")


def fmt_grade(value: float) -> str:
    """Note affichée : 12 -> « 12 », 12.5 -> « 12,5 », 12.25 -> « 12,25 »."""
    s = f"{value:.2f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def fmt_pct(value: float | None) -> str:
    return "—" if value is None else f"{value:.1f}".replace(".", ",") + " %"


# --------------------------------------------------------------------------
# Observation suggérée
# --------------------------------------------------------------------------
def suggest_observation(stats: dict[str, Any]) -> str:
    """Texte d'observation modifiable, fondé uniquement sur les statistiques."""
    n, mean = stats["n"], stats["mean"]
    if mean >= 14:
        appr = "très satisfaisants"
    elif mean >= 12:
        appr = "satisfaisants"
    elif mean >= 10:
        appr = "moyens"
    else:
        appr = "insuffisants"
    lv = {l["key"]: l for l in stats["levels"]}
    parts = [
        f"Les résultats de ce contrôle sont {appr} : la moyenne de la classe est de {fmt(mean)}/20 "
        f"({n} élèves évalués, note la plus haute {fmt_grade(stats['max'])}, la plus basse {fmt_grade(stats['min'])}). "
        f"{stats['pass_count']} élève(s) ({fmt_pct(stats['pass_rate'])}) ont obtenu au moins 10/20."
    ]
    top = lv["bonne"]["count"] + lv["tres_bonne"]["count"]
    if top:
        parts.append(f"Points forts : {top} élève(s) atteignent une bonne ou très bonne maîtrise des apprentissages.")
    if stats["support_count"]:
        parts.append(
            f"Difficultés : {stats['support_count']} élève(s) ({fmt_pct(lv['non_maitrise']['percent'])}) n'ont pas "
            "atteint le seuil de maîtrise (note inférieure à 10/20), ce qui révèle des lacunes à combler."
        )
        parts.append("Actions proposées : séances de soutien ciblées pour ces élèves, exercices de remédiation "
                     "sur les prérequis non acquis et suivi individualisé jusqu'au prochain contrôle.")
    else:
        parts.append("Aucun élève n'est en dessous de 10/20 : poursuivre la consolidation et proposer des activités "
                     "d'approfondissement.")
    return " ".join(parts)
