import yaml
import argparse
from openpyxl import load_workbook
from grading import Grading
from scores import Scores
from statistics import mean
from typing import Union, List

parser = argparse.ArgumentParser(
    description="Tool for convert exam score to DSE predicted grade",
    formatter_class=argparse.ArgumentDefaultsHelpFormatter,
)

parser.add_argument(
    "config", nargs="?", metavar="CONFIG", help="config file", default="config.yaml"
)

args = parser.parse_args()
config = yaml.safe_load(open(args.config, "r", encoding="utf-8"))

wb = load_workbook(config["file"]["grading"])
ws = wb.active


if __name__ == "__main__":

    grading = Grading(ws)
    score_wb = load_workbook(config["file"]["scores"])
    student_wb = load_workbook(config["file"]["students"])
    student_sheet = student_wb.active
    classcode_column = student_sheet["A"]
    scores = Scores(score_wb, grading)

    ELECTIVES = ["m2", "x2", "x1"]
    for i, subject in enumerate(ELECTIVES):
        col_index = 7 - i
        student_sheet.insert_cols(col_index)
        for classcode_cell in classcode_column:
            content: Union[int, str, None] = None
            if classcode_cell.row == 1:
                content = subject + " grade"
            else:
                subject_code = student_sheet.cell(
                    row=classcode_cell.row, column=col_index - 1
                ).value
                if subject_code is None:
                    pass
                elif subject_code == "cscb":
                    grades: List[int] = []

                    for part in ["csc", "csb"]:
                        part_grade = scores.getGradeByClasscodeAndSubject(
                            classcode_cell.value, part
                        )
                        if part_grade is not None:
                            grades.append(part_grade)

                    content = int(round(mean(grades)))
                else:
                    content = scores.getGradeByClasscodeAndSubject(
                        classcode_cell.value, subject_code
                    )

            student_sheet.cell(row=classcode_cell.row, column=col_index).value = content

    CORE_SUBJECTS = ["chi", "eng", "math", "ls"]
    for i, subject in enumerate(CORE_SUBJECTS):
        col_index = 4 + i
        student_sheet.insert_cols(col_index, amount=1)
        for classcode_cell in classcode_column:
            content = None
            if classcode_cell.row == 1:
                content = subject
            else:
                content = scores.getGradeByClasscodeAndSubject(
                    classcode_cell.value, subject
                )
            student_sheet.cell(row=classcode_cell.row, column=col_index).value = content

    student_wb.save(config["file"]["result"])
