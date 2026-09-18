# Kahoot! Import File Auditor

`audit-kahoot-xlsx` is a Codex skill for reviewing existing Kahoot! spreadsheet import files (`.xlsx`). It checks each question for curriculum alignment, competency-oriented assessment features, answer correctness, misconception-based distractors, age-appropriate wording, and Kahoot! import compatibility.

## What it reviews

- Taiwan 108 Curriculum alignment and learning-stage fit
- Competency-oriented assessment and whether a context is necessary for solving the question
- Correct answer uniqueness, multi-select completeness, scientific accuracy, units, and conditions
- Distractor plausibility and the misconceptions or reasoning errors they may represent
- Wording, reading load, ambiguity, missing figures or experimental conditions, and expected response time
- Current Kahoot! spreadsheet structure, answer indices, time values, and conservative length checks

The skill separates confirmed evidence, reasoned inference, pending verification, and insufficient data. Without student response data it reports expected distractor quality only; it does not claim measured difficulty or discrimination.

## Inputs and outputs

The skill accepts one `.xlsx`, multiple workbooks, a folder, or a ZIP containing workbooks. It preserves source file, worksheet, Excel row, and question identifiers.

The default output is a Chinese-language review summary plus an independent Excel review report. The original workbook is preserved. When a corrected import workbook is requested, the revised file is saved separately and rechecked against the current official Kahoot! template.

The bundled read-only extractor can be run directly:

```text
python scripts/inspect_kahoot.py questions.xlsx --output extracted.json
```

Providing a verified platform profile enables mechanical checks for the profile's limits. Omitting it leaves platform limits unverified rather than guessing them:

```text
python scripts/inspect_kahoot.py questions.xlsx --rules verified-rules.json --output checked.json
```

The extractor does not replace content review, curriculum research, or an actual Kahoot! upload test. It never modifies the source workbook.

## Use in Codex

Invoke the skill explicitly with:

```text
$audit-kahoot-xlsx
```

Example:

> Use `$audit-kahoot-xlsx` to review this Kahoot! workbook for Grade 6 elementary science and produce a question-by-question Excel report.

## Version

Current version: **1.0.0**

Repository: [prayer168/kahoot-import-file-auditor](https://github.com/prayer168/kahoot-import-file-auditor)
