from openpyxl import worksheet
from openpyxl.utils import get_column_letter
from typing import Optional, Union
from helper import bail


class Grading:
    def __init__(self, ws: worksheet):
        self.ws = ws

    def getSubjectColumnLetter(self, subject: str) -> Optional[str]:
        """
            findSubjectCol returns the column letter of the subject,
            it returns None if the subject is not found
        """
        for cell in self.ws[1]:
            if cell.value == subject:
                return get_column_letter(cell.column)
        return None

    def findGradeBySubjectAndScore(
        self, subject: str, score: Union[int, float]
    ) -> Optional[int]:
        column_letter = self.getSubjectColumnLetter(subject)
        score = int(round(score))
        if column_letter is None:
            bail("Subject ({}) not find".format(subject))

        subject_column = self.ws[column_letter]
        for cell in subject_column:
            if cell.row != 1:
                grading_range = [int(n) for n in cell.value.split(", ")]
                lower_bound = grading_range[0]
                upper_bound = grading_range[1]

                isWithinRange = score >= lower_bound and score <= upper_bound
                if isWithinRange:
                    return self.ws.cell(row=cell.row, column=1).value

        bail("Score: {} of the subject {} is out of range".format(score, subject))
        return None
