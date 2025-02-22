from openpyxl import load_workbook
from grading import Grading
import unittest

wb = load_workbook("./data/grading_ranges.xlsx")
ws = wb.active
g = Grading(ws)


class TestGrading(unittest.TestCase):
    def test_equal(self):
        test_data = (("bio", 84, 2), ("bafs", 129.1, 3), ("chem", 154, 5))
        for datum in test_data:
            result = g.get_grade_by_score(datum[0], datum[1])
            want = datum[2]
            self.assertEqual(result, want)

    def test_unmatch(self):
        test_data = (("bio", 84, 3), ("bafs", 129.1, 2), ("chem", 154, 7))
        for datum in test_data:
            result = g.get_grade_by_score(datum[0], datum[1])
            want = datum[2]
            self.assertNotEqual(result, want)

    def test_out_of_range(self):
        with self.assertRaises(SystemExit) as cm:
            g.get_grade_by_score("bio", 221)
        self.assertEqual(cm.exception.code, 1)


if __name__ == "__main__":
    unittest.main()
