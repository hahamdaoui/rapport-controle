"""Configuration centrale de l'application de rapports de contrôle."""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
CHART_DIR = BASE_DIR / "static" / "charts"

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MAX_UPLOAD_MB = 10
MAX_CONTENT_LENGTH = MAX_UPLOAD_MB * 1024 * 1024

# Nettoyage automatique des fichiers temporaires (en heures)
UPLOAD_TTL_HOURS = 6
OUTPUT_TTL_HOURS = 24

HOST = "127.0.0.1"
PORT = 5000

# --- Niveaux de maîtrise (barème /20) -------------------------------------
# (clé, libellé, borne basse incluse, borne haute EXCLUE, intervalle affiché,
#  libellé de tranche pour le graphique, couleur)
# Le dernier niveau est borné à 20 inclus (voir data_analyzer.classify).
MASTERY_LEVELS = [
    ("non_maitrise", "Non maîtrisé", 0, 10, "0 ≤ Note < 10", "Moins de 10", "#C0392B"),
    ("minimale", "Maîtrise minimale", 10, 12, "10 ≤ Note < 12", "De 10 à moins de 12", "#E67E22"),
    ("moyenne", "Maîtrise moyenne", 12, 14, "12 ≤ Note < 14", "De 12 à moins de 14", "#D4AC0D"),
    ("bonne", "Bonne maîtrise", 14, 16, "14 ≤ Note < 16", "De 14 à moins de 16", "#2E86C1"),
    ("tres_bonne", "Très bonne maîtrise", 16, 20, "16 ≤ Note ≤ 20", "De 16 à 20", "#1E8449"),
]

SUPPORT_THRESHOLD = 10  # élèves strictement en dessous => soutien
MASTERED_THRESHOLD = 12  # note >= 12 => "Maîtrisé"

GLOBAL_LEVELS = [
    ("maitrise", "Maîtrisé", "#1E8449"),
    ("en_cours", "En cours de maîtrise", "#E67E22"),
    ("non_maitrise", "Non maîtrisé", "#C0392B"),
]

# Palette du rapport
NAVY = "1F3A5F"
BLUE = "1F5FA8"
LIGHT_BLUE = "DCEAF7"
LIGHT_GREY = "F2F5F9"
GRID_GREY = "B8C4D2"
