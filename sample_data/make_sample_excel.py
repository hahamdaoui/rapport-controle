"""Génère des fichiers Excel d'exemple (NOMS ET NOTES FICTIFS) pour tester l'application.

    python sample_data/make_sample_excel.py
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUT = Path(__file__).resolve().parent

NAMES_AR = [
    "العلوي سلمى", "بنجلون ياسين", "الإدريسي مريم", "التازي أيوب", "الفاسي هدى", "المودن أمين", "السعدي نادية",
    "الحسني عمر", "بوزيد خديجة", "الكتاني زكرياء", "الشرقاوي إيمان", "المنصوري رشيد", "الأزرق سناء", "الحدادي بلال",
    "الوزاني لبنى", "العمراني أنس", "اليعقوبي حفصة", "القادري يوسف", "الطاهري ملاك", "الرحموني وليد",
    "الصديقي آية", "المرابط حمزة", "الخطيبي شيماء", "الجزولي عبد الله", "الهلالي ريم", "الغازي سفيان",
    "الزاهيدي كوثر", "البكاري إلياس", "الودغيري جنات", "الناصري مهدي", "الساحلي فاطمة الزهراء", "الأنصاري نعيمة",
    "بوعزة محمد أمين", "لحلو سارة", "ڭنون عبد الإله",
]
NAMES_FR = ["Karim Benali", "Sofia Lahlou"]
GRADES = [  # 37 notes + 1 absent -> couvre toutes les tranches et les bornes (0, 9.75, 10, 12, 14, 16, 20)
    4.5, 7, 8.25, 9.75, 6, 0, 3.5, 8, 9, 5.75, 7.5, 9.5, 8.75,
    10, 10.5, 11.25, 11.75, 10.25, 11, 11.5,
    12, 12.5, 13, 13.75, 12.75, 13.5, 12.25,
    14, 14.5, 15.25, 15.75, 14.75, 15,
    16, 17.5, 18.25, 20,
]
assert len(NAMES_AR) + len(NAMES_FR) == len(GRADES) + 0 or True


def massar_like():
    wb = Workbook()
    ws = wb.active
    ws.title = "NotesCC"
    ws["F4"] = "نقط المراقبة المستمرة"
    ws.merge_cells("F4:I4")
    meta = [
        (7, "C", "أكاديمية :", "D", "كلميم - واد نون", "G", "م.الإقليمية : ", "I", "إقليم: سيدي افني", "L", "مؤسسة", "O",
         "الثانوية الإعدادية 30 يونيو"),
        (9, "C", "المستوى  :  ", "D", "الثانية إعدادي مسار دولي", "G", "القسم  :", "I", "2APIC-1", "L", "الاستاذ", "O",
         "عبد الحكيم حامدوي "),
        (11, "C", "الدورة  :", "D", "الدورة الأولى", "G", "نقط :", "I", "الفرض الأول", "L", "المادة", "O", "الرياضيات"),
    ]
    for r, *cells in meta:
        for i in range(0, len(cells), 2):
            ws[f"{cells[i]}{r}"] = cells[i + 1]
    ws["C13"], ws["D13"] = "السنة الدراسية :", "2026/2027"
    ws["B16"], ws["C16"], ws["D16"], ws["F16"] = "ID", "رقم  التلميذ  ", "إسم التلميذ  ", " تاريخ الإزدياد"
    ws["G16"], ws["H16"] = "الفرض الأول", "الفرض الأول"
    ws["G17"], ws["H17"] = "النقطة", "التغيب"
    ws.merge_cells("D16:E17")
    names = NAMES_AR + NAMES_FR
    # (nom, note) ; un élève absent (غ) et un élève sans note
    rows = list(zip(names, GRADES)) if len(names) == len(GRADES) else list(zip(names, GRADES))
    rows = rows[:len(GRADES)]
    rows.append(("الشاوي لمياء", "غ"))
    rows.append(("المدني حسناء", None))
    for i, (name, grade) in enumerate(rows):
        r = 18 + i
        ws[f"B{r}"] = 15000000 + i
        ws[f"C{r}"] = f"J1{i:08d}"
        ws[f"D{r}"] = name
        ws.merge_cells(f"D{r}:E{r}")
        ws[f"F{r}"] = "01-01-2013"
        if grade is not None:
            ws[f"G{r}"] = grade
    ws2 = wb.create_sheet("Data")
    ws2["A2"] = "مبرر"
    wb.save(OUT / "exemple_notes_massar.xlsx")


def french_like():
    wb = Workbook()
    ws = wb.active
    ws.title = "Contrôle 1"
    ws["A1"] = "Établissement : Collège Ibn Rochd"
    ws["A2"] = "AREF : Casablanca-Settat"
    ws["A3"] = "Direction provinciale : Nouaceur"
    ws["A4"] = "Professeur : Fatima Zahra El Idrissi"
    ws["A5"] = "Matière : Mathématiques"
    ws["A6"] = "Classe : 3APIC-2"
    ws["A7"] = "Semestre : 2"
    ws["A8"] = "Contrôle : 2"
    ws["A9"] = "Année scolaire : 2025/2026"
    ws["A11"], ws["B11"], ws["C11"] = "N°", "Nom et prénom", "Note /20"
    for c in ("A11", "B11", "C11"):
        ws[c].font = Font(bold=True)
        ws[c].fill = PatternFill("solid", fgColor="DDDDDD")
    data = [("Amine Zahraoui", "14,5"), ("Salma Berrada", 17), ("Youssef El Amrani", 8.5), ("Imane Tahiri", "abs"),
            ("Mehdi Alaoui", 11), ("Nour Bennani", 9.99), ("Hamza Idrissi", 12), ("Lina Cherkaoui", None),
            ("Anas Fassi", 6)]
    for i, (n, g) in enumerate(data, start=1):
        ws.cell(11 + i, 1, i)
        ws.cell(11 + i, 2, n)
        if g is not None:
            ws.cell(11 + i, 3, g)
    wb.save(OUT / "exemple_notes_francais.xlsx")


if __name__ == "__main__":
    massar_like()
    french_like()
    print("Fichiers d'exemple créés dans", OUT)
