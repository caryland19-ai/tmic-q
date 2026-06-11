# Concept Test Template Reference

Use this reference when reading `TMIC模版.xlsx` and converting a product concept test into TMIC import content. Internal IDs may be used for parsing only; never expose them to the user or respondent.

## Workbook Sources

Expected assets:

- `assets/TMIC模版.xlsx`: source questionnaire template.
- `assets/问卷格式.xlsx`: TMIC import layout reference.

If either file is missing, ask the user to provide it or report the missing asset as a blocking issue.

## Source Columns

The source concept-test template usually uses:

- Column A: section label or required marker.
- Column B: internal locator, page text, or question identifier.
- Column C: respondent-facing question text or option text.
- Column D: question type, option code, or scale value.
- Column E: logic note.

Question rows are rows where the locator/type columns indicate a real question. Rows below them with option text are options until the next question or display row. Rows with instruction text but no question type are display/instruction items and should be preserved in order when the template requires them.

## Type Mapping

- `单选` -> `[单选题]`
- `多选` -> `[多选题]`
- `填空` -> `[多行文本题]` for reasons or long answers; `[单行文本题]` only for short single-value answers.
- `编码（便于做均值分析）` -> `[单选题]`; preserve scale order and codes in the logic/check sheet if useful for analysis.

## Concept Information Points

Use confirmed flat information points, not title/description pairs.

Correct option shape:

```text
信息点原文
```

Incorrect option shape:

```text
标题
说明
```

Do not merge adjacent title and description into one option. If a card contains four information points, the target questions receive four information-point options.

## Concept Replacement Targets

After the user confirms information points, replace only the placeholder concept options in the three target concept-selection questions:

- Demand-met description selection.
- Hard-to-understand description selection.
- Unbelievable / not-credible description selection.

Use every confirmed information point as an independent option in the same visual order. Then keep or append these fixed mutually exclusive options:

- Demand-met question: `以上描述都不满足我的需求`
- Hard-to-understand question: `没有不太理解的词句`
- Not-credible question: `没有不太可信的描述`

## Not-credible Follow-up Questions

For the not-credible concept-selection question, create one multi-line text follow-up for each confirmed information point.

Question text:

```text
请问您不相信“XXX”词句的原因是？
```

Rules:

- Replace `XXX` with the exact confirmed information point.
- The number of follow-up questions must equal the number of confirmed information points.
- Each follow-up question appears immediately after the not-credible control question and before subsequent template questions.
- Each follow-up displays only when the corresponding information point is selected.
- Do not create a follow-up for `没有不太可信的描述`.
- If TMIC import cannot preserve display logic, still generate the follow-up questions and put exact one-to-one display rules in `整卷答题逻辑`.

## Logic Handling

Put logic in separate output sheets, not in the import sheet.

Required logic categories:

- Logic validation result.
- Full answer/display logic.
- Full screener/termination logic.

Common logic mapping:

- `选项随机`: option randomization note.
- `互斥`: mutually exclusive option note.
- `终止`: screener/termination logic.
- Notes like `针对...提问`: answer/display logic.
- Notes like selected-option follow-up: generate one follow-up display rule per information point.

Do not invent target criteria for age, income, brand, family stage, or purchase behavior. If target criteria are missing, mark them as needing manual confirmation.

## Public Output Constraint

Final import sheets, user messages, validation reports, file names, examples, and generated follow-up texts must not expose internal template IDs or source row labels. Use final respondent-facing question text and final question sequence numbers only where required by logic output.
