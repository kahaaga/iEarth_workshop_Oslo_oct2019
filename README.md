# Student assessment tool

## Usage

To use the evaluator, simply import the `Evaluator` class from the `evaluation_rubric` module (located in the `evaluation_rubric.py` file). Then, create an `Evaluator` instance by feeding it the three mandatory arguments: 1) a file path to an Excel sheet containing the evaluation rubric(s), 2) the course code (a string), and 3) the semester (a string).

For example, to generate reports based on the `FAG123_H2019_vurdering.xslx` file, do the following.

```python
evaluator = Evaluator("FAG123_H2019_vurdering.xslx", "FAG123", "H2019")

# Generate reports for activity 3
evaluator.generate_reports(3)
```

This will generate a folder in the working directory for the course `FAG123`, with a subfolder for the semester `H2019` and subfolder for activity 3, which then contains a subfolder for each student that has been assessed.

## Rubric format

- Evaluation rubrics for separate activities is organised in separate spreadsheets. The rubrics corresponding to a specific activity are identified based on the spreadsheet names. By default, spreadsheets should be named `Aktivitet N`, where `N` is the activity number. The spreadsheet naming can be changed by providing the optional `str_activity` argument when constructing the  `Evaluator` instance. For example, if your spreadsheets are numbered in English as `Activity 1`, `Activity 2`, etc, then do `evaluator = Evaluator(file, course, semester, str_activity = 'Activity'`.
- Spreadsheets should be numbered sequentially using integers. `Activity 1`, `Activity 2` will work, `Activity 1A` and `Activity 1B` will not.
- The `Kategori`-, `Kriterietype`- og `Vurderingskriterium`-columns has to be standard text (no special characters like `_`).
- In the`standardkommentar`-columns, values are intepreted as markdown. In the feedback to the students, you can thus write `kode = sin(x)`, and use **bold** or *italic* fonts, etc.

## Assessment design

Assessment design should:

"EMPOWER"
    - Engage students actively in identifying or formulating criteria
    - Facilitate opportunities for self-assessment and reflection
    - Deliver feedback that helps students self-correct
    - Provide opportunities for feedback dialogue (peer and tutor-student)
    - Encourage positive motivational beliefs and self-esteem
    - Provide opportunities to apply what is learned in new tasks
    - Yield information that teachers can use to help shape teaching
"ENGAGE"
    - Capture sufficient study time and effort in and out of class
    - Distribute students’ effort evenly across topics and weeks.
    - Engage students in deep not just shallow learning activity
    - Communicates clear and high expectations to students.
Adapted from Nicol and Macfarlane-Dick (2006) and Gibbs and Simpson (2004)

- [Holder et al. 2017, Student Learning of Complex Earth Systems: A Model to Guide Development of Student Expertise in Problem-Solving](https://www.tandfonline.com/doi/full/10.5408/17-261.1)
- [Hesse et al. 2014, A Framework for Teachable Collaborative Problem Solving Skills](https://link.springer.com/chapter/10.1007/978-94-017-9395-7_2)
