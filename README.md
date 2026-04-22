# test_codex

A tiny Python CLI project for combining right-click Excel notes from cells.

## What It Does

The CLI reads attached Excel note/comment objects from specific cells in one `.xlsx` workbook, combines the non-empty note text, and writes the result into a target sheet and cell in an existing output workbook.

## Setup

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install the project with test dependencies:

```bash
python -m pip install -e ".[test]"
```

## Run The CLI

For one transfer:

```bash
test-codex source.xlsx output.xlsx --cells A1 B1 C1 --target-sheet Summary --target-cell D1
```

Choose a source worksheet:

```bash
test-codex source.xlsx output.xlsx --source-sheet Sheet1 --cells A1 B1 C1 --target-sheet Summary --target-cell D1
```

Use a custom separator:

```bash
test-codex source.xlsx output.xlsx --cells A1 B1 C1 --target-sheet Summary --target-cell D1 --separator " | "
```

Read visible cell values instead of right-click Excel notes:

```bash
test-codex source.xlsx output.xlsx --cells A1 B1 C1 --target-sheet Summary --target-cell D1 --cell-values
```

For many transfers, create a CSV mapping file:

```csv
source_sheet,source_cells,target_sheet,target_cell
전체 점수,"C9 C11",Sheet2,B2
그룹C,"C4 C9 C11",Summary,D36
그룹D,"C4 C9 C11",Summary,D37
```

Then run one command:

```bash
test-codex source.xlsx output.xlsx --mapping mappings.csv
```

Send a prompt range to ChatGPT and write the response to a cell:

```bash
export OPENAI_API_KEY="enter_key"
test-codex assessment_tool.xlsx test.xlsx --prompt-sheet Prompt --prompt-range A1:D8 --response-sheet Results --response-cell B2
```

In ChatGPT mode, the first workbook is where the prompt is read from, and the second workbook is where the response is written.

Use a different model or add extra instructions:

```bash
test-codex assessment_tool.xlsx test.xlsx --prompt-sheet Prompt --prompt-range A1:D8 --response-sheet Results --response-cell B2 --model gpt-4.1-mini --instructions "Answer in Korean and keep it concise."
```

You can also run the CLI without using the installed script:

```bash
python -m test_codex.cli source.xlsx output.xlsx --cells A1 B1 C1 --target-sheet Summary --target-cell D1
```

## Run Tests

```bash
PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 pytest
```

That environment variable keeps pytest focused on this project's tests and avoids auto-loading unrelated plugins from global tools such as ROS.

# Actual Run: move comments for team 1
test-codex 7~8주차_지능1_A_B그룹_평가.xlsx assessment_tool.xlsx --mapping t1_mappings.csv --prompt-jobs prompt_jobs.csv

# Change mapping to t2_mappings.csv and t3_mappings.csv for team 2 and team 3.
test-codex 7~8주차_지능1_A_B그룹_평가.xlsx assessment_tool.xlsx --mapping t2_mappings.csv --prompt-jobs prompt_jobs.csv
test-codex 7~8주차_지능1_A_B그룹_평가.xlsx assessment_tool.xlsx --mapping t3_mappings.csv --prompt-jobs prompt_jobs.csv