# score2grade

`score2grade` is a program what will use the exam scores to predict DSE grades.

## Subjects

### Core Subjects
Core subjects include _Chinese_, _English_, _Mathematics_ and _Liberal Study_. The full mark of these subjects is 300. Make sure the grading ranges are in scale with 300 as full mark.

### Electives
The full mark of electives subjects is 200.

### Combined Science
Combined science (`cscb`) is an elective subjects, which have 2 parts, _Biology_ and _Chemistry_. The exam score of each part have to be grade separately. The final grade of this subject will simply take the rounded average of two parts.

### M2
M2 will be treated as a elective subject, but the full mark of M2 is 100.

## Requirements

Users have to provide THREE (3) xlsx documents for converting the exam scores to DSE grades.

### `grading_ranges.xlsx`

The file should have the format as below:

| grade | chi      | eng      | math     | ls       | bafs     | bio      | ...      |
| ----- | -------- | -------- | -------- | -------- | -------- | -------- | -------- |
| 0     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 1     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 2     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 3     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 4     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 5     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 6     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |
| 7     | min, max | min, max | min, max | min, max | min, max | min, max | min, max |

Each column (beyond the first) corresponds to a the grade ranges for an exam.

Scores have to be rounded first, the grading ranges are not continuous. So both min and max are inclusive.

## `scores.xlsx`

The score file stores the all exam scores of all students. Each sheet only contains score of one subject, the sheet name should be named with subject code. The following table is an example of Subject Chinese in sheet named with `chi`


| classcode | ename | cname | chi |
| --------- | ----- | ----- | --- |
| 6A01      | xx    | xx    | xx  |
| 6A02      | xx    | xx    | xx  |
| 6A03      | xx    | xx    | xx  |
| 6A04      | xx    | xx    | xx  |
| 6A05      | xx    | xx    | xx  |
| 6A06      | xx    | xx    | xx  |
| 6A07      | xx    | xx    | xx  |


# Report

The output will be an xlsx document with the following information:

- classcode
- ename
- cname
- ...grades of the student for subjects
