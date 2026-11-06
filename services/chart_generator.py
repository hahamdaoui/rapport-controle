"""Graphiques Matplotlib du rapport (générés dynamiquement à partir des statistiques)."""
from __future__ import annotations

from pathlib import Path
from typing import Any

import matplotlib

matplotlib.use("Agg")  # pas d'interface graphique (serveur / Windows)
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import MaxNLocator  # noqa: E402

NAVY = "#1F3A5F"
GRID = "#D5DDE8"
CM = 1 / 2.54


def _title_size(title: str, width_cm: float, fs: float) -> float:
    """Taille de titre adaptée à la largeur du graphique (évite tout rognage)."""
    return max(6.0, min(fs + 2, width_cm * 28.35 * 0.90 / (len(title) * 0.56)))


def _style(ax, fs: float) -> None:
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color("#8896A8")
    ax.tick_params(colors=NAVY, labelsize=fs)
    ax.set_axisbelow(True)


def distribution_chart(stats: dict[str, Any], path: str | Path, width_cm: float = 14, height_cm: float = 5.0,
                       dpi: int = 220) -> Path:
    """Graphique 1 – barres horizontales : répartition par tranches de notes."""
    fs = max(6.5, min(9.0, height_cm * 1.6))
    levels = stats["levels"]
    labels = [lv["bar_label"] for lv in levels]
    counts = [lv["count"] for lv in levels]
    colors = [lv["color"] for lv in levels]

    fig, ax = plt.subplots(figsize=(width_cm * CM, height_cm * CM), dpi=dpi)
    bars = ax.barh(labels, counts, color=colors, height=0.62, edgecolor="white")
    ax.invert_yaxis()  # « Moins de 10 » en haut
    xmax = max(max(counts), 1)
    ax.set_xlim(0, xmax * 1.18 + 0.4)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    ax.grid(axis="x", color=GRID, linewidth=0.6)
    for b, c in zip(bars, counts):
        ax.text(b.get_width() + xmax * 0.015 + 0.05, b.get_y() + b.get_height() / 2, str(c),
                va="center", ha="left", fontsize=fs + 1, fontweight="bold", color=NAVY)
    ax.set_xlabel("Nombre d'élèves", fontsize=fs, color=NAVY)
    ax.set_ylabel("Tranches de notes", fontsize=fs, color=NAVY)
    title = "Répartition des élèves selon les tranches de notes"
    tfs = _title_size(title, width_cm, fs)
    fig.text(0.5, 0.985, title, ha="center", va="top", fontsize=tfs, fontweight="bold", color="#1F5FA8")
    _style(ax, fs)
    fig.tight_layout(pad=0.5, rect=(0, 0, 1, 1 - (tfs * 1.9) / (height_cm * 28.35)))
    path = Path(path)
    fig.savefig(path, dpi=dpi, facecolor="white")
    plt.close(fig)
    return path


def global_chart(stats: dict[str, Any], path: str | Path, width_cm: float = 12, height_cm: float = 4.0,
                 dpi: int = 220) -> Path:
    """Graphique 2 – barres verticales : Maîtrisé / En cours de maîtrise / Non maîtrisé."""
    fs = max(6.5, min(9.0, height_cm * 1.7))
    groups = stats["global"]
    labels = [g["label"] for g in groups]
    counts = [g["count"] for g in groups]
    colors = [g["color"] for g in groups]

    fig, ax = plt.subplots(figsize=(width_cm * CM, height_cm * CM), dpi=dpi)
    bars = ax.bar(labels, counts, color=colors, width=0.55, edgecolor="white")
    ymax = max(max(counts), 1)
    ax.set_ylim(0, ymax * 1.22 + 0.4)
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.grid(axis="y", color=GRID, linewidth=0.6)
    for b, c in zip(bars, counts):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + ymax * 0.02 + 0.05, str(c), ha="center", va="bottom",
                fontsize=fs + 1, fontweight="bold", color=NAVY)
    ax.set_xlabel("Niveaux de maîtrise", fontsize=fs, color=NAVY)
    ax.set_ylabel("Nombre d'élèves", fontsize=fs, color=NAVY)
    title = "Vue globale des niveaux de maîtrise"
    tfs = _title_size(title, width_cm, fs)
    fig.text(0.5, 0.985, title, ha="center", va="top", fontsize=tfs, fontweight="bold", color="#1F5FA8")
    _style(ax, fs)
    fig.tight_layout(pad=0.5, rect=(0, 0, 1, 1 - (tfs * 1.9) / (height_cm * 28.35)))
    path = Path(path)
    fig.savefig(path, dpi=dpi, facecolor="white")
    plt.close(fig)
    return path
