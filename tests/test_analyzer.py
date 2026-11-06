"""Tests des calculs statistiques (python -m unittest discover -s tests -v)."""
import random
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services import data_analyzer as da  # noqa: E402


def students(grades):
    return [{"name": f"Élève {i}", "status": "valid", "grade": g, "grade20": g} for i, g in enumerate(grades)]


class ClassifyTests(unittest.TestCase):
    def test_boundaries_are_mutually_exclusive(self):
        cases = {0: "non_maitrise", 9.99: "non_maitrise", 10: "minimale", 11.99: "minimale", 12: "moyenne",
                 13.99: "moyenne", 14: "bonne", 15.99: "bonne", 16: "tres_bonne", 20: "tres_bonne"}
        for grade, key in cases.items():
            self.assertEqual(da.classify(grade), key, grade)

    def test_out_of_range_rejected(self):
        for bad in (-0.1, 20.01, float("nan"), None):
            with self.assertRaises(da.AnalysisError):
                da.classify(bad)

    def test_global_groups(self):
        self.assertEqual(da.classify_global(9.99), "non_maitrise")
        self.assertEqual(da.classify_global(10), "en_cours")
        self.assertEqual(da.classify_global(11.99), "en_cours")
        self.assertEqual(da.classify_global(12), "maitrise")


class AnalyzeTests(unittest.TestCase):
    def test_known_distribution(self):
        g = [5, 9.5, 10, 11, 12, 13, 14, 15, 16, 20]
        s = da.analyze(students(g))
        self.assertEqual(s["n"], 10)
        self.assertEqual([l["count"] for l in s["levels"]], [2, 2, 2, 2, 2])
        self.assertEqual([l["percent"] for l in s["levels"]], [20.0] * 5)
        self.assertEqual(s["mean"], 12.55)
        self.assertEqual(s["max"], 20)
        self.assertEqual(s["min"], 5)
        self.assertEqual(s["levels"][0]["mean"], 7.25)  # (5 + 9.5) / 2 : uniquement le niveau concerné
        self.assertEqual(s["levels"][4]["mean"], 18.0)
        self.assertEqual([x["count"] for x in s["global"]], [6, 2, 2])  # maîtrisé, en cours, non maîtrisé
        self.assertEqual(s["support_count"], 2)

    def test_support_sorted_ascending_and_strictly_below_10(self):
        s = da.analyze(students([9.99, 10, 3, 0, 7.5, 12]))
        self.assertEqual([x["grade"] for x in s["support"]], [0, 3, 7.5, 9.99])
        self.assertTrue(all(x["level"] == "Non maîtrisé" for x in s["support"]))

    def test_empty_level_has_no_mean(self):
        s = da.analyze(students([5, 6]))
        self.assertIsNone(s["levels"][4]["mean"])
        self.assertEqual(da.fmt(None), "—")

    def test_only_valid_students_are_counted(self):
        data = students([8, 14]) + [
            {"name": "A", "status": "missing", "grade": None, "grade20": None},
            {"name": "B", "status": "absent", "grade": None, "grade20": None},
            {"name": "C", "status": "duplicate", "grade": 8, "grade20": None},
            {"name": "D", "status": "invalid", "grade": None, "grade20": None},
        ]
        s = da.analyze(data)
        self.assertEqual(s["n"], 2)
        self.assertEqual(s["mean"], 11.0)  # une note vide n'est jamais comptée comme 0

    def test_no_valid_students(self):
        with self.assertRaises(da.AnalysisError):
            da.analyze([{"name": "A", "status": "missing", "grade": None, "grade20": None}])

    def test_totals_always_equal_n_random(self):
        rnd = random.Random(42)
        for _ in range(200):
            n = rnd.randint(1, 80)
            g = [round(rnd.uniform(0, 20) * 4) / 4 for _ in range(n)]
            s = da.analyze(students(g))
            self.assertEqual(sum(l["count"] for l in s["levels"]), n)
            self.assertEqual(sum(x["count"] for x in s["global"]), n)
            self.assertEqual(s["support_count"], sum(1 for x in g if x < 10))
            self.assertAlmostEqual(sum(l["percent"] for l in s["levels"]), 100, delta=0.3)
            self.assertAlmostEqual(s["mean"], sum(g) / n, delta=0.0051)


class FormatTests(unittest.TestCase):
    def test_french_formatting(self):
        self.assertEqual(da.fmt(12.5), "12,50")
        self.assertEqual(da.fmt_grade(12), "12")
        self.assertEqual(da.fmt_grade(12.5), "12,5")
        self.assertEqual(da.fmt_grade(9.75), "9,75")
        self.assertEqual(da.fmt_pct(33.33), "33,3 %")

    def test_suggested_observation_uses_stats(self):
        s = da.analyze(students([5, 8, 14, 18]))
        text = da.suggest_observation(s)
        self.assertIn("11,25/20", text)
        self.assertIn("4 élèves", text)
        self.assertIn("soutien", text)
        s2 = da.analyze(students([12, 14]))
        self.assertIn("Aucun élève", da.suggest_observation(s2))


if __name__ == "__main__":
    unittest.main()
