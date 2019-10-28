# Student assessment tool


## Usage

To use the evaluator, simply import the `Evaluator` class from the `evaluation_rubric` module (located in the `evaluation_rubric.py` file). Then, create an `Evaluator` instance by feeding it the three mandatory arguments: 1) a file path to an Excel sheet containing the evaluation rubric(s), 2) the course code (a string), and 3) the semester (a string).

For example, to generate reports based on the `FAG123_H2019_vurdering.xslx` file, do the following.

```python
evaluator = Evaluator("FAG123_H2019_vurdering.xslx", "FAG123", "H2019")

# Generate reports for activity 3
evaluator.generate_reports(3)
```

## Rubric format

- Evaluation rubrics for separate activities is organised in a separate spreadsheet. The rubrics corresponding to a specific activity are identified based on the spreadsheet names. By default, spreadsheets should be named `Aktivitet N`, where `N` is the activity number. The spreadsheet naming can be changed by providing the optional `str_activity` argument when constructing the  `Evaluator` instance. For example, if your spreadsheets are numbered in English as `Activity 1`, `Activity 2`, etc, then do `evaluator = Evaluator(file, course, semester, str_activity = 'Activity'`.
- Spreadsheets should be numbered sequentially using integers. `Activity 1`, `Activity 2` will work, `Activity 1A` and `Activity 1B` will not.