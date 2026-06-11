# TMIC Import Format Reference

This reference is based on `问卷格式.xlsx`. Use the actual workbook in `assets/` when available. The upload-only import workbook must contain only respondent-facing questionnaire content.

## Basic Pattern

TMIC upload format is a simple sheet where respondent-facing content appears directly in cells.

For simple question types:

```text
[单选题]请问您的性别是
男

女
```

Recommended clean version:

```text
[单选题]请问您的性别是
男
女
```

Use one blank row between questions for readability.

## Supported Question Headings Seen In The Template

- `[单行文本题]`
- `[单选题]`
- `[多选题]`
- `[排序题]`
- `[单选矩阵题]`
- `[多选矩阵题]`
- `[多行文本题]`
- `[评分题]`
- `[评分矩阵题]`
- `[地址题]`
- `[附件题]`

## Matrix Pattern

For matrix questions, the first row is the question title. The next row contains column labels from column B onward, while the following rows contain row labels in column A.

Example pattern:

```text
[评分矩阵题]请从多个角度给这款新产品打分
        颜色    气味    质感
非常满意
比较满意
一般满意
不满意
很不满意
```

## What Not To Put In The Upload Sheet

Do not include:

- `分页`
- `分页：第2页`
- `第2页`
- `第3页`
- Internal page names.
- Operator instructions.
- Logic notes that are not respondent-facing.
- Source-only labels such as `必答题`.
- Internal template IDs, source row labels, debug notes, or source-only question numbers.

These may be imported as question text or options.

Put page ranges only in the complete workbook's separate `分页指导` sheet, for example `第一页：Q1-Q11`. Never put page labels in the import sheet or upload-only workbook.

## Complete Workbook Review Sheet

The complete workbook's `完整TMIC问卷导入表格` sheet may include review columns:

```text
题号    问卷导入内容          逻辑
Q1      [单选题]请问您的性别
        男                    终止
```

Column A shows final question numbers on question rows, such as `Q1` and `Q24`; option rows leave it blank. Column B remains the questionnaire/import content. Column C is only for human review, uses gray font, and must not be copied into the upload-only workbook. When Column C references another question, use final question numbers such as `Q24.题目原文`, not internal template IDs.

When processing a user-modified complete questionnaire, rebuild the upload-only workbook from Column B only:

- Keep question rows and option rows exactly as edited by the user.
- Recalculate final question numbers only in the complete workbook, not in the upload-only workbook.
- Do not include `题号`, `逻辑`, `逻辑校验结果`, `分页指导`, or `变更识别`.
- If a question or option was deleted from the modified workbook, it must also disappear from the upload-only workbook.
- If a new question or option was added, include it as-is in the upload-only workbook and flag uncertain logic needs in the complete workbook.

## Logic Handling

The upload-only sheet is for question structure. Put these in separate output sheets or the complete workbook review column:

- Termination.
- Display logic.
- Required status.
- Option randomization.
- Mutual exclusion.
- Conditional follow-up.
- Image attachment.
- Page creation and page movement.

Required companion sheets are:

- `逻辑校验结果`
- `整卷答题逻辑`
- `整卷甄别逻辑`
- `分页指导`

Use final question order numbers in the two logic sheets. Do not expose internal template IDs.
