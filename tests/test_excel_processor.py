"""Tests de la lecture / validation Excel (arabe, français, cas limites)."""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from services import excel_processor as xp  # noqa: E402
from services.pdf_converter import find_soffice  # noqa: E402


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def save(self, rows, name="t.xlsx", sheet="Feuil1", extra=None):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        for r in rows:
            ws.append(r)
        if extra:
            extra(wb)
        p = self.tmp / name
        wb.save(p)
        return p


class GradeParsingTests(unittest.TestCase):
    def test_parse(self):
        P = xp.parse_grade
        self.assertEqual(P(12.5), ("numeric", 12.5, ""))
        self.assertEqual(P("12,5")[:2], ("numeric", 12.5))
        self.assertEqual(P("١٤")[:2], ("numeric", 14.0))  # chiffres arabes
        self.assertEqual(P("15/20")[:2], ("numeric", 15.0))
        self.assertEqual(P(0)[:2], ("numeric", 0.0))
        self.assertEqual(P(None)[0], "missing")
        self.assertEqual(P("  ")[0], "missing")
        self.assertEqual(P("abs")[0], "absent")
        self.assertEqual(P("غ")[0], "absent")
        self.assertEqual(P("غائب")[0], "absent")
        self.assertEqual(P("abc")[0], "invalid")
        self.assertEqual(P(-3)[0], "invalid")
        self.assertEqual(P(True)[0], "invalid")

    def test_norm_and_ordinals(self):
        self.assertEqual(xp.norm("الأستاذ :"), xp.norm("الاستاذ"))
        self.assertEqual(xp.norm("Établissement :"), "etablissement")
        self.assertEqual(xp.ordinal_to_number("الدورة الأولى"), "1")
        self.assertEqual(xp.ordinal_to_number("الفرض الثاني"), "2")
        self.assertEqual(xp.ordinal_to_number("Semestre 2"), "2")


class FrenchSheetTests(Base):
    def rows(self):
        return [
            ["Établissement : Collège Test"], ["Classe : 3A"], ["Professeur : M. Dupont"], ["Matière : Maths"],
            ["Année scolaire : 2025/2026"], [],
            ["N°", "Nom et prénom", "Note /20"],
            [1, "Amine Zahraoui", "14,5"], [2, "Salma Berrada", 17], [3, "Youssef Amrani", 8.5],
            [4, "Imane Tahiri", "abs"], [5, "Lina Cherkaoui", None], [6, "Mehdi Alaoui", "xx"],
            [7, "Nour Bennani", 25], [8, "Anas Fassi", 9.99], [9, "Amine Zahraoui", "14,5"],
            [None, None, None],
        ]

    def test_columns_metadata_and_statuses(self):
        r = xp.process_excel(self.save(self.rows()))
        self.assertEqual(r["mapping"]["name_cols"], [1])
        self.assertEqual(r["mapping"]["grade_col"], 2)
        self.assertEqual(r["scale"], 20)
        self.assertEqual(r["metadata"]["school"], "Collège Test")
        self.assertEqual(r["metadata"]["class"], "3A")
        self.assertEqual(r["metadata"]["year"], "2025/2026")
        self.assertEqual(r["metadata"]["direction"], "")  # jamais inventé
        c = r["counts"]
        self.assertEqual((c["valid"], c["absent"], c["missing"], c["invalid"], c["duplicate"]), (4, 1, 1, 2, 1))
        by_name = {s["name"]: s for s in r["students"] if s["status"] == "valid"}
        self.assertEqual(by_name["Amine Zahraoui"]["grade"], 14.5)
        self.assertEqual(by_name["Anas Fassi"]["grade"], 9.99)
        invalid = [s for s in r["students"] if s["status"] == "invalid"]
        self.assertTrue(any("barème" in s["message"] for s in invalid))  # 25 > 20
        self.assertFalse(r["needs_mapping"])

    def test_missing_grade_is_not_zero(self):
        r = xp.process_excel(self.save([["Nom", "Note"], ["Aya Ben", 10], ["Omar Lo", None]]))
        self.assertEqual(r["counts"]["valid"], 1)
        self.assertEqual(r["counts"]["missing"], 1)

    def test_nom_prenom_two_columns(self):
        r = xp.process_excel(self.save([["Nom", "Prénom", "Note"], ["Alaoui", "Sara", 12], ["Idrissi", "Omar", 8]]))
        self.assertEqual(r["students"][0]["name"], "Alaoui Sara")

    def test_mapping_override_and_summary_rows_ignored(self):
        p = self.save([["Élève", "Contrôle 1", "Contrôle 2"], ["A B", 10, 15], ["C D", 12, 5], ["Moyenne", 11, 10]])
        r = xp.process_excel(p)
        self.assertTrue(r["needs_mapping"])  # deux colonnes de notes possibles
        r2 = xp.process_excel(p, mapping={"sheet": "Feuil1", "name_cols": [0], "grade_col": 2, "first_row": 2})
        self.assertEqual([s["grade"] for s in r2["students"]], [15.0, 5.0])
        self.assertEqual(r2["counts"]["rows"], 2)  # ligne « Moyenne » ignorée

    def test_scale_explicit_and_converted_to_20(self):
        r = xp.process_excel(self.save([["Nom", "Note /10"], ["A B", 8], ["C D", 4]]))
        self.assertEqual(r["scale"], 10)
        self.assertEqual([s["grade20"] for s in r["students"]], [16.0, 8.0])

    def test_scale_inferred_when_many_above_20(self):
        r = xp.process_excel(self.save([["Nom", "Note"], ["A B", 35], ["C D", 22], ["E F", 18], ["G H", 30]]))
        self.assertEqual(r["scale"], 40)
        self.assertEqual(r["scale_source"], "inferred")

    def test_same_name_different_grade_kept_with_warning(self):
        r = xp.process_excel(self.save([["Nom", "Note"], ["Ali Ben", 10], ["Ali Ben", 12]]))
        self.assertEqual(r["counts"]["valid"], 2)
        self.assertTrue(any("Même nom" in i["message"] for i in r["issues"]))

    def test_unknown_structure_requests_mapping(self):
        r = xp.process_excel(self.save([["x", "y"], ["Samir Alami", 12], ["Rim Tazi", 9]]))
        self.assertTrue(r["needs_mapping"])
        self.assertEqual(r["counts"]["valid"], 2)  # devinée d'après le contenu

    def test_multiple_sheets_best_chosen(self):
        def extra(wb):
            ws = wb.create_sheet("Notes")
            for row in [["Nom et prénom", "Note"], ["A B", 11], ["C D", 7]]:
                ws.append(row)
        r = xp.process_excel(self.save([["rien"]], extra=extra))
        self.assertEqual(r["sheet"], "Notes")
        self.assertEqual(r["sheets"], ["Feuil1", "Notes"])


class ArabicMassarTests(Base):
    def test_real_massar_layout_with_unicode_names(self):
        sample = ROOT / "sample_data" / "exemple_notes_massar.xlsx"
        r = xp.process_excel(sample)
        m = r["metadata"]
        self.assertEqual(m["class"], "2APIC-1")
        self.assertEqual(m["semester"], "1")
        self.assertEqual(m["exam"], "1")
        self.assertEqual(m["year"], "2026/2027")
        self.assertEqual(m["direction"], "سيدي افني")
        self.assertIn("الثانوية", m["school"])
        self.assertEqual(r["counts"]["valid"], 37)
        self.assertEqual(r["counts"]["absent"], 1)
        self.assertEqual(r["counts"]["missing"], 1)
        names = [s["name"] for s in r["students"]]
        self.assertIn("العلوي سلمى", names)  # encodage arabe intact
        self.assertIn("Karim Benali", names)

    def test_empty_grade_column_reports_clear_error(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["إسم التلميذ", "النقطة"])
        ws.append(["علي بنعمر", None])
        ws.append(["سلمى الإدريسي", None])
        p = self.tmp / "e.xlsx"
        wb.save(p)
        r = xp.process_excel(p)
        self.assertFalse(r["has_valid"])
        self.assertTrue(any(i["level"] == "error" and "Aucune note" in i["message"] for i in r["issues"]))


class ErrorTests(Base):
    def test_corrupted_file(self):
        p = self.tmp / "bad.xlsx"
        p.write_bytes(b"PK\x03\x04garbage")
        with self.assertRaises(xp.ExcelReadError):
            xp.process_excel(p)

    @unittest.skipUnless(find_soffice(), "LibreOffice requis pour créer un .xls de test")
    def test_xls_support(self):
        from services.pdf_converter import convert_with_libreoffice
        src = self.save([["Nom", "Note"], ["A B", 15], ["C D", 6]])
        xls = convert_with_libreoffice(src, self.tmp / "o", target="xls")
        r = xp.process_excel(xls)
        self.assertEqual(r["counts"]["valid"], 2)


if __name__ == "__main__":
    unittest.main()
