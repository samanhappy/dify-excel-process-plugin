# AGENTS.md

## Project overview

This repository is a Dify plugin for extracting text and images from Excel workbooks.

Key implementation files:
- `tools/excel_extractor.py` — main extraction logic for `.xlsx` and `.xls`
- `tests/test_excel_extractor.py` — regression tests for Excel text rendering behavior
- `requirements.txt` — Python dependencies

## Environment and dependency installation

Use the project virtual environment at `.venv`.

### Verified install flow

The repository dependencies were successfully installed with:

- `PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 ./.venv/bin/python -m pip install -r requirements.txt`

### Why this matters

`dify_plugin` pulls in `tiktoken`, which builds through PyO3. With Python `3.14`, installation can fail because the bundled PyO3 version only officially supports up to Python `3.13` unless forward compatibility is enabled.

If dependency installation fails around `tiktoken`/`PyO3`, retry with:

- `PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1`

### Known caveat

A plain `uv sync` may fail in this environment for the same dependency-chain reason. If that happens, prefer installing via the project interpreter in `.venv` using `pip` as shown above.

## Test commands

### Verified regression test command

Run the current tests with:

- `./.venv/bin/python -m pytest`

For a focused file run:

- `./.venv/bin/python -m pytest tests/test_excel_extractor.py`

This command was verified to pass after the blank-cell fix.

## Debugging and regression guidance

### Blank-cell column-shift bug

Issue context already investigated in this repository:
- Empty cells in a row were previously filtered out during text extraction.
- That caused later columns to shift left in the rendered output.

The fix is implemented in `tools/excel_extractor.py` by routing both `.xlsx` and `.xls` row rendering through `_render_row_text()`.

Expected behavior:
- preserve leading and internal blank cells when they are between populated cells
- trim only trailing empty cells
- skip rows that are completely empty

Example:
- input row: `A`, empty, `C`
- expected output: `A |  | C`

## Test-writing notes for this repository

### Current test style

The repository currently uses `pytest`.

### When `dify_plugin` gets in the way

For focused logic tests, it is acceptable to stub `dify_plugin` imports inside the test file instead of requiring a full plugin runtime boot.

The current regression test demonstrates this in `tests/test_excel_extractor.py`.

Prefer:
- module-level `test_...` functions
- `pytest` fixtures for lightweight setup
- plain `assert` statements

### Minimal-object construction pattern

For narrow unit tests of `ExcelExtractorTool`, this pattern is acceptable:

- create the tool with `object.__new__(ExcelExtractorTool)`
- call internal pure-ish helper methods directly when runtime/session wiring is not needed

This keeps regression tests small and avoids unnecessary plugin bootstrapping.

## Type-checking and editor-error guidance

This codebase may hit editor/type-checker complaints even when runtime behavior is correct.

### Patterns that were verified to satisfy static analysis

#### 1. Module stubs in tests

When creating module stubs with `types.ModuleType`, prefer:
- `setattr(module, "Name", value)`

instead of direct dynamic attribute assignment like:
- `module.Name = value`

This avoids Pylance complaints about unknown module attributes.

#### 2. `Workbook.active`

`openpyxl.Workbook.active` may be treated as nullable by static analysis.

When a worksheet is definitely expected, use an explicit cast:
- cast to `Worksheet`

#### 3. Third-party type definitions may be stricter than runtime

`ToolParameter` accepted runtime arguments that static typing did not recognize, including `file_accepts`.

For this codebase, a practical workaround is:
- `tool_parameter_cls = cast(Any, ToolParameter)`
- then instantiate through `tool_parameter_cls(...)`

Use this only when runtime behavior is verified and the library type stubs are incomplete.

#### 4. Prefer broader sequence types for row data

`xlrd` row values can come through as a general sequence rather than a concrete `list`.

For helper methods like `_render_row_text()`, prefer:
- `Sequence[Any]`

instead of restricting to only `list[Any] | tuple[Any, ...]`.

## Safe change guidelines

When modifying extraction behavior:
- keep `.xlsx` and `.xls` rendering behavior aligned
- add or update a regression test first when fixing row-formatting bugs
- validate with the repository test command after changes
- avoid changing image extraction paths unless the issue specifically involves embedded media

## Files to inspect first for related issues

For text extraction bugs:
- `tools/excel_extractor.py`
- `tests/test_excel_extractor.py`

For plugin/runtime entrypoints:
- `main.py`
- `provider/excel_process.py`
