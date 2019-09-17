from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from typing import Optional, Union
from helper import bail


class Grading:
    def __init__(self, ws: Worksheet):
        self.ws = ws
        self.header_row = ws[1]
        self.subject_column = {
            cell.value: self.ws[get_column_letter(cell.column)]
            for cell in self.header_row
        }

    def get_grade_by_score(
        self, subject: str, score: Union[int, float]
    ) -> Optional[int]:
        score = int(round(score))

        if subject not in self.subject_column:
            bail(f"Subject ({subject}) not find")

        subject_column = self.subject_column[subject]
        for cell in subject_column[1:]:
            lower, upper = [int(n) for n in cell.value.split(", ")]

            isWithinRange = score >= lower and score <= upper
            if isWithinRange:
                return self.ws.cell(row=cell.row, column=1).value

        bail(f"Score: {score} of the subject {subject} is out of range")
        return None
