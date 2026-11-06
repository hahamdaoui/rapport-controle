# Rapport de contrôle – Analyse des résultats d'examens

Application web locale (Flask) pour les enseignants du collège : elle lit un fichier Excel de notes
(export **Massar** ou fichier libre, en arabe ou en français), calcule les statistiques, classe les élèves par
niveau de maîtrise et produit un **rapport A4 d'une seule page** en **DOCX** et **PDF**.

## 1. Installation (Windows)

1. **Python 3.11 ou supérieur** : <https://www.python.org/downloads/> (cochez *Add Python to PATH*).
2. **LibreOffice** (gratuit) : <https://www.libreoffice.org/download/> – nécessaire à la conversion PDF
   (et à la vérification « une seule page »). Si Microsoft Word est installé, `docx2pdf` sert de solution de repli.
3. Double-cliquez sur **`start.bat`**. Au premier lancement, il crée un environnement `.venv` et installe
   automatiquement les modules de `requirements.txt` (connexion Internet requise une seule fois).

Installation manuelle (équivalent) :

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## 2. Démarrer et ouvrir l'interface

`start.bat` lance le serveur et ouvre le navigateur sur **<http://127.0.0.1:5000>**.
Laissez la fenêtre noire ouverte pendant l'utilisation ; fermez-la pour arrêter l'application.
Les données ne quittent jamais votre ordinateur.

## 3. Utilisation pas à pas

1. **Importer** le fichier Excel (bouton, ou glisser-déposer). Le nom du fichier s'affiche ; « Remplacer » / « Supprimer » permettent de changer de fichier.
2. **Vérifier** : établissement, AREF, direction, professeur, matière, classe, semestre, n° du contrôle et année sont détectés automatiquement (modifiables, jamais inventés : un champ absent reste vide). L'aperçu montre les élèves, les notes, les anomalies (note manquante, absent, invalide, hors barème, doublon) et les statistiques.
   Si la détection est incorrecte, ouvrez **« Correspondance des colonnes »** pour choisir la feuille, la colonne des noms, celle des notes, la 1re ligne des élèves et le barème.
3. Saisir la **date de l'examen** et la **date de correction** (obligatoires, non extraites d'Excel).
4. Écrire les **observations** (facultatif) ou cliquer sur « Générer une observation suggérée » puis la modifier.
5. Cliquer sur **« Analyser et générer le rapport »**.
6. **Télécharger** le rapport DOCX et/ou PDF.

## 4. Règles de calcul

| Note /20 | Niveau |
|---|---|
| 0 ≤ n < 10 | Non maîtrisé |
| 10 ≤ n < 12 | Maîtrise minimale |
| 12 ≤ n < 14 | Maîtrise moyenne |
| 14 ≤ n < 16 | Bonne maîtrise |
| 16 ≤ n ≤ 20 | Très bonne maîtrise |

* Seules les notes valides sont comptées : une note vide n'est **jamais** comptée comme 0 ; les absents, notes invalides, notes hors barème et doublons exacts (même nom + même note) sont exclus et signalés.
* Pourcentages = effectif du niveau / nombre d'élèves valides. La moyenne d'un niveau ne porte que sur les élèves de ce niveau.
* Synthèse en 3 groupes : Maîtrisé (≥ 12), En cours de maîtrise (10 à < 12), Non maîtrisé (< 10).
* Soutien scolaire : note strictement inférieure à 10, triée par note croissante.
* Barème : détecté (« /10 », « sur 40 »…) ou à choisir manuellement ; les notes sont ramenées à /20.
* Des contrôles internes vérifient que la somme des 5 niveaux et celle des 3 groupes égalent l'effectif.

## 4 bis. Une seule page A4

Cinq mises en page sont essayées, de la plus confortable à la plus compacte (polices, hauteur des graphiques,
liste de soutien sur 1, 2 ou 3 colonnes, graphiques à côté des tableaux). Pour chacune, le PDF est généré et son
**nombre de pages vérifié** ; la première qui tient sur une page est retenue. Police minimale : 7,5 pt.
Si le contenu ne peut pas tenir lisiblement (ex. plus de ~50 élèves en soutien), un **avertissement** s'affiche et
la mise en page **« 2 pages (lisible) »** est proposée – jamais de rapport illisible ni multi-pages silencieux.
Les observations sont limitées à 1 200 caractères pour la même raison.

## 5. Structure du projet

```
app.py  config.py  requirements.txt  start.bat  README.md
services/  excel_processor.py  data_analyzer.py  chart_generator.py
           docx_generator.py  pdf_converter.py  report_builder.py
templates/ index.html  preview.html  result.html
static/    css/style.css  js/app.js  icons/sprite.svg  charts/
tests/     test_analyzer.py  test_excel_processor.py
sample_data/ exemple_notes_massar.xlsx  exemple_notes_francais.xlsx  exemple_rapport.docx  exemple_rapport.pdf
uploads/  outputs/        (fichiers temporaires, supprimés automatiquement)
```

## 6. Tests

```bat
.venv\Scripts\python -m unittest discover -s tests -v
```

Les fichiers d'exemple (noms et notes **fictifs**) se régénèrent avec `python sample_data/make_sample_excel.py`.

## 7. Sécurité et confidentialité

* Extension **et** contenu du fichier vérifiés (xlsx/xls), taille limitée à 10 Mo, nom de stockage aléatoire.
* Serveur écoutant uniquement sur `127.0.0.1` (non accessible depuis le réseau).
* Fichiers importés supprimés après 6 h, rapports après 24 h (voir `config.py`).

## 8. Dépannage

| Problème | Solution |
|---|---|
| « Aucune note n'a été saisie » | La colonne des notes du fichier est vide : saisissez les notes dans Excel (ou choisissez une autre colonne) et réimportez. |
| PDF indisponible | Installez LibreOffice, puis relancez `start.bat`. |
| Fichier .xls illisible | `pip install xlrd` (inclus dans requirements.txt) ou enregistrez en .xlsx. |
| Port 5000 occupé | Modifiez `PORT` dans `config.py`. |
| Colonnes mal détectées | Section « Correspondance des colonnes ». |

Note : les coins du titre sont droits (Word ne permet pas d'arrondir un cadre de tableau sans objets de dessin
fragiles) ; le titre utilise un cadre bleu clair avec bordure bleue.
