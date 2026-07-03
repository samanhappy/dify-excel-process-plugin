# Excel Process Plugin for Dify

A lightweight Excel-focused plugin for Dify that extracts every piece of
textual cell content together with any embedded images inside `.xlsx`
workbooks.

## Overview

Excel Process bundles tools purpose-built for reading spreadsheets. The current
tool transparently handles modern `.xlsx` and legacy `.xls` files, iterates
over all sheets and rows, produces a human-readable text representation, and
surfaces embedded media (`xl/media` assets in `.xlsx` files or OLE streams in
`.xls`) so downstream workflows can reason over both data and imagery.

## Features

- Extract and format text from every worksheet in an Excel file (.xlsx/.xls)
- Export every embedded image as a separate binary response
- Preserve sheet names and row ordering for clarity

## Available Tools

### Excel Extractor

Reads an Excel workbook and streams the parsed content.

- **Excel Content**: `.xlsx` or `.xls` file whose text and images should be extracted

## Author

Created by [samanhappy](https://github.com/samanhappy)

## Repository

https://github.com/samanhappy/dify-excel-process-plugin

## Publishing

This repository includes a GitHub Actions workflow for publishing Marketplace
updates through `langgenius/dify-plugins`.

Before the first automated release:

- Fork `langgenius/dify-plugins` to `samanhappy/dify-plugins`.
- Add a repository secret named `PLUGIN_ACTION` with permission to push to the
  fork and create pull requests against `langgenius/dify-plugins`.
- Confirm `manifest.yaml` has the intended `author`, `name`, and `version`.

To publish a new version, update `manifest.yaml`, push to `main`, and wait for
the `Publish Dify Plugin` workflow. The workflow runs tests, packages
`excel_process-<version>.difypkg`, pushes it to
`samanhappy/dify-plugins`, and opens the Marketplace pull request.
