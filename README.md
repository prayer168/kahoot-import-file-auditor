# Kahoot! Import File Auditor

> English version. The Traditional Chinese version is included below.

`audit-kahoot-xlsx` is a Codex skill for reviewing existing Kahoot! spreadsheet import files (`.xlsx`). It checks each question for curriculum alignment, competency-oriented assessment features, answer correctness, misconception-based distractors, age-appropriate wording, and Kahoot! import compatibility.

## What it reviews

- Taiwan 108 Curriculum alignment and learning-stage fit
- Competency-oriented assessment and whether a context is necessary for solving the question
- Correct answer uniqueness, multi-select completeness, scientific accuracy, units, and conditions
- Distractor plausibility and the misconceptions or reasoning errors they may represent
- Wording, reading load, ambiguity, missing figures or experimental conditions, and expected response time
- Current Kahoot! spreadsheet structure, answer indices, time values, and conservative length checks
- A/B/C/待確認 classifications by question and review dimension
- Correction logs, corrected import workbooks, and post-correction re-audits when requested

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

Current version: **2.0.0**

Repository: [prayer168/kahoot-import-file-auditor](https://github.com/prayer168/kahoot-import-file-auditor)

---

# Kahoot! 匯入檔案審查師（繁體中文）

> English version is above. 以下為繁體中文說明。

`audit-kahoot-xlsx` 是一個 Codex 技能，用來查核既有的 Kahoot! Excel 匯入檔案（`.xlsx`）。它會逐題檢查 108 課綱與單元對應、素養導向特徵、答案正確性、迷思概念與誘答性、年段用詞，以及 Kahoot! 匯入格式。

## 查核內容

- 臺灣 108 課綱對應與年段適切性
- 素養導向評量，以及題目是否需要情境才能作答
- 正確答案唯一性、多選答案完整性、科學正確性、單位與作答條件
- 錯誤選項的合理性，以及可能反映的迷思概念或推理錯誤
- 題目用詞、閱讀負荷、歧義、缺少圖表或實驗條件，以及預期作答時間
- Kahoot! Excel 結構、答案編號、時間值與保守字數限制
- 依題號與查核向度列出 A／B／C／待確認分級
- 使用者要求時輸出訂正對照、修正版匯入檔與修正後複審結果

技能會區分已確認證據、合理推論、待查證項目與資料不足。若沒有學生作答資料，只會報告預期的誘答品質，不會宣稱已測得難度或鑑別度。

## 輸入與輸出

技能接受單一 `.xlsx`、多個活頁簿、資料夾或包含活頁簿的 ZIP，並保留來源檔案、工作表、Excel 列號與題目識別資訊。

預設輸出為繁體中文查核摘要與獨立 Excel 審查報告，原始檔案不會被修改。若要求產生修正版匯入檔，會另存新檔，並依目前查證過的 Kahoot! 範本重新檢查。

內建的唯讀擷取器可直接執行：

```text
python scripts/inspect_kahoot.py questions.xlsx --output extracted.json
```

若提供已查證的平台規則檔，可執行字數、時間與結構檢查：

```text
python scripts/inspect_kahoot.py questions.xlsx --rules verified-rules.json --output checked.json
```

擷取器不能取代內容審查、課綱研究或實際 Kahoot! 上傳測試，也不會修改來源活頁簿。

## 在 Codex 中使用

明確指定技能：

```text
$audit-kahoot-xlsx
```

使用範例：

> 請使用 `$audit-kahoot-xlsx` 查核這份六年級自然科 Kahoot! Excel 匯入檔，並產生逐題 Excel 審查報告。

## 版本

目前版本：**2.0.0**

Repository：[prayer168/kahoot-import-file-auditor](https://github.com/prayer168/kahoot-import-file-auditor)
