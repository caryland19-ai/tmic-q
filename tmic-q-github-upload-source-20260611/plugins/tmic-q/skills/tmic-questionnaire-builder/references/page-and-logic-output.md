# Page And Logic Output Reference

Use this reference for final workbook structure, logic validation, answer logic, screener logic, and page guidance.

## Required Workbook Sheets

Every final `.xlsx` must include five independent parts:

1. `完整TMIC问卷导入表格`
2. `逻辑校验结果`
3. `整卷答题逻辑`
4. `整卷甄别逻辑`
5. `分页指导`

Optional sheets are allowed, such as:

- `信息点确认`
- `变更识别`

Do not mix page guidance into the import sheet. The complete workbook review sheet may include `题号` and `逻辑` columns, but upload-only workbooks must remain single-column.

## Modified Questionnaire Workflow

When a user uploads a modified complete questionnaire workbook after receiving the original output:

1. Treat the modified workbook as the final questionnaire source.
2. Read `完整TMIC问卷导入表格`; use `问卷导入内容` as respondent-facing content and `逻辑` as the editable logic source.
3. Recalculate final question numbers from the modified question order.
4. Update recognizable old question references in logic text to the new final question numbers.
5. Remove deleted questions and deleted options from both the complete workbook and upload-only workbook.
6. Do not invent logic for new or rewritten questions/options. Flag uncertain logic needs in `逻辑校验结果`.
7. Produce a refreshed complete workbook and a refreshed upload-only workbook.

If the previous complete workbook is also provided, include a `变更识别` sheet for added, deleted, and moved questions. If it is not provided, state in `逻辑校验结果` that validation is based only on the modified workbook.

## Complete Workbook Import Sheet

The complete workbook's `完整TMIC问卷导入表格` sheet contains three columns:

```text
题号    问卷导入内容    逻辑
```

Column A contains final question numbers on question rows, such as `Q1`, `Q2`, `Q24`. Option rows leave Column A blank.

Column B contains:

- Question rows such as `[单选题]题目文本`, `[多选题]题目文本`, `[多行文本题]题目文本`.
- Option rows under choice questions.
- Blank rows only when they are safe for the import format.

Column C contains human-readable logic markers copied or converted from the template:

- Question-level logic such as option randomization and display conditions.
- Option-level logic such as `终止`, `互斥`, and project-specific screeners.
- Generated follow-up display logic, using final question numbers.

Format Column C in gray font.

When referencing a question, use final question numbers such as `Q24.题目原文`. Do not use internal template IDs.

The upload-only workbook must use only the `问卷导入内容` column values and must not include this logic column.

The upload-only workbook must not include `题号` or `逻辑` columns.

Questionnaire content must not contain:

- Page markers such as `分页`, `第2页`, `第3页`.
- Operator instructions.
- Internal template IDs, source labels, row numbers, or debug text.
- Logic notes that are not respondent-facing.

## Page Guidance Sheet

The complete workbook must include a separate `分页指导` sheet. Keep it concise:

```text
分页指导
第一页：Q1-Q11
第二页：Q12-Q14
第三页：Q15-Q16
```

Use final question order numbers with the `Q` prefix. Do not use internal template IDs. Do not place these page markers in `完整TMIC问卷导入表格` or the upload-only workbook.

## Logic Validation

Run validation before producing answer and screener logic.

Validation must check:

- Display rules point to existing final questions and options.
- Termination rules point to existing final questions and options.
- Display targets come after the control question.
- The same condition does not both display a target and terminate the survey unless explicitly confirmed.
- Not-credible follow-up count equals confirmed information-point count.
- Every not-credible follow-up has one corresponding display rule.
- No internal template ID appears in public sheets.
- Deleted or modified questions/options are not referenced by stale logic.
- New or modified questions that need logic are flagged when logic cannot be inferred.
- Remaining template placeholders in question text or options, such as `XX`, `XXX`, `品牌1`, `品牌2`, `品牌3`, and `……`.
- Logic notes that require manual confirmation, such as target brand definitions, project-specific screeners, income threshold screeners, option randomization, and mutual exclusion settings.

Use three columns:

```text
等级    检查项    意见
```

Color levels:

- `重大问题`: red font. Use for broken generation, missing source/option/target, count mismatch, stale references, and other blocking logic errors.
- `一般问题`: orange font. Use for template placeholders, ambiguous conditions, project-specific screeners, option randomization, mutual exclusion, and any item needing manual setup or confirmation.
- `无问题`: black font. Use for checks that pass.

If the template says a rule should terminate or screen respondents but does not name the exact option or condition, do not output a fake rule such as `对应选项或条件`; mark it as `一般问题` in `逻辑校验结果`.

## Answer Logic Format

Each display rule is one sentence:

```text
如果「题目序号.题目原文」选择「选项或条件」则显示「目标题目序号.目标题目原文」，否则不显示。
```

Not-credible follow-up display rules use:

```text
如果「题目序号.控制题原文」选择「XXX」则显示「题目序号.请问您不相信“XXX”词句的原因是？」，否则不显示。
```

Use final output question order. Do not use internal template IDs.

## Screener Logic Format

Each termination rule is one sentence:

```text
如果「题目序号.题目原文」选择「选项或条件」则问卷结束，否则继续。
```

Rules that terminate, screen out, or define unsuitable respondents belong here, not in answer logic.

Special rule:

```text
您家里有多少位小孩？（包括已出生和怀孕中的孩子数量）
```

must not appear as answer/display logic. If it has exclusion or termination behavior, put it in screener logic and mark any corrected ownership in validation.

## Missing Logic

If no logic can be read for a category, write:

```text
未从表格中读取到对应逻辑。
```

Do not invent logic. Missing target criteria should be flagged for manual confirmation.
