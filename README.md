
# score2grade

`score2grade` is a program what will use the exam scores to predict DSE grades.

# Requirements

Users have to provide THREE (3) xlsx documents for converting the exam scores to DSE grades.

## Combined subjects

Some subjects consists of separated papers / parts, for example Combined Science (`cscb`) have 2 parts, _Biology_ and _Chemistry_; English have several papers.
The papers / parts have separated assessments, so the exam score of each should be graded separately.
Furthermore, the subject grade may be a weighted average the constituent papers / parts.
We will use _exam code_ as the ID for each exam score, hence a exam code can correspond to a paper or a part for a subject.

If a subject does not have scores in `scores.xlsx`, it must be a combined subjects.
Note the combined subject's score may be precomputed and the program will not treat it as a combined subject (the following section is not applicable)

### Constituents and weighting

The constituents of a combined subject will be:

1. defined in `weighting` field in `config.yaml`
2. otherwise, looked up in `scores.xlsx` with subject code as prefix (e.g. `cscb.*`, `eng.*`)
   the weighting will be simple average

### Computing grade of combined subject

If a computed subject do not have entries in `grading_ranges.xlsx`, the grades of its constituents will be used with weighting to get the subject grade.
If the grading ranges are available, the individual scores of the constituents will be weighted and graded according to the subject's grading ranges.

## `grading_ranges.xlsx`

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
For grade 0 - 6, min is inclusive and max is exclusive.
For grade 7, both min and max are inclusive.

## `scores.xlsx`

The score file stores the exam scores of all students with the format below.

| regno     | chi | eng | math | ls  | subject1 | subject2 | subject3 |
| --------- | --- | --- | ---- | --- | -------- | -------- | -------- |
| lp0000001 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000001 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000002 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000003 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000004 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000005 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000006 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |
| lp0000007 | xx  | xx  | xx   | xx  | xx       | xx       | xx       |

## students.xlsx

students.xlsx stored the student information, formatted as below. This is used to render the final report where the column subjects is the list of subjects students have.

| regno     | classcode | classno | ename | cname |
| --------- | --------- | ------- | ----- | ----- |
| lp0000001 | xx        | xx      | xx    | xx    |
| lp0000001 | xx        | xx      | xx    | xx    |
| lp0000002 | xx        | xx      | xx    | xx    |
| lp0000003 | xx        | xx      | xx    | xx    |
| lp0000004 | xx        | xx      | xx    | xx    |
| lp0000005 | xx        | xx      | xx    | xx    |
| lp0000006 | xx        | xx      | xx    | xx    |
| lp0000007 | xx        | xx      | xx    | xx    |

# Report

The output will be an xlsx document with the following information:

- regno
- classcode
- classno
- ename
- cname
- ...grades of the student for subjects specified in yaml
