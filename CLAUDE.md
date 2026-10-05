# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Biopharma financial models (e.g. rNPV, DCF, revenue forecasts, comparables). The repo is newly scaffolded: there is no build system, toolchain, or test suite yet. Update this file once one is chosen.

- `models/`: model files (Excel, Python, etc.)
- `data/`: input data and assumptions
- `docs/`: notes, write-ups, and outputs

## Working with the Excel models

- Python is not installed on this machine. Workbooks are created and edited through Excel COM automation from PowerShell (`New-Object -ComObject Excel.Application`).
- The user often has workbooks open in Excel. If a lock file `models/~$<name>.xlsx` exists, COM opens the file read-only and `Save()` silently does nothing. Check for the lock file (or `$wb.ReadOnly`) before editing, and ask the user to close the file.
- Price history comes from Nasdaq's JSON API, which the nasdaq.com charts use: `https://api.nasdaq.com/api/quote/<TICKER>/historical?assetclass=stocks&fromdate=YYYY-MM-DD&todate=YYYY-MM-DD&limit=9999`. It requires a browser `User-Agent` header. Rows come back newest-first. Older prices may come back adjusted, with 4 decimal places.
- `models/NVS.xlsx`, `Stock Price` tab: the data table is in `A4:F`, sorted oldest-first. The workbook names `PxDates`/`PxClose` grow to fit rows appended to the table. The chart controls (inputs `I4:I5`, calculations `I8:I16`, result sentence `H18`) drive the names `ChartDates`/`ChartClose`, which feed the chart.

## Version control workflow

The user's main goal for this repo is that every version is saved and easy to revert. So:

- After each meaningful change, commit with a descriptive message and push to `origin main` (https://github.com/HerbieRun/biopharma-financial-models).
- Use a feature branch for larger experiments, then merge back into `main`.
- Undo with `git revert <sha>`, or restore one file with `git checkout <sha> -- <file>`. Never rewrite history or force-push `main`.
- `.gitattributes` marks Office files and PDFs as binary, so git can't diff them. Commit messages for these files should state what changed (assumptions, tabs, outputs).

## Public repository

The GitHub repo is **public**. Before committing, check that files don't contain confidential deal data, non-public company information, or credentials, and ask the user if unsure.
