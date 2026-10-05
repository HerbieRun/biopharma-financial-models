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
- Yahoo's chart API (`https://query1.finance.yahoo.com/v8/finance/chart/<SYMBOL>?period1=<unix>&period2=<unix>&interval=1d`, browser `User-Agent`) is used for ROG (RHHBY) and BAYNr (BAYN.DE, EUR). Reuters pages return 401 to scripts. Yahoo may include a partial bar for a session that is still trading, so drop it.
- Each `models/<TICKER>.xlsx` (NVS, AZN, GSK, ROG, SNY, BAYNr, MRK, PFE, LLY, BMY, JnJ) has an identical `Stock Price` tab. The data table is in `A4:F`, sorted oldest-first, and cell `A2` holds the source and caveats. The workbook names `PxDates`/`PxClose` grow to fit rows appended to the table. The chart controls (inputs `I4:I5`, calculations `I8:I16`, result sentence `H18`) drive the names `ChartDates`/`ChartClose`, which feed the chart.
- `NVS.xlsx` also has a "Peer Comps" chart (below the price chart) and a table (`U46:W59`), both driven by the same `I4:I5`. Their data comes from the `Peer Data` tab:
  - Peer closes are **values copied** from the other workbooks and aligned to NVS trading days (the last close on or before each date). Refresh them whenever a peer's data changes.
  - Columns `N:X` hold the % change for each ticker. The names `PeerPct_<TICKER>` feed the chart lines.
  - `Peer Data!Z:AK` is a helper that sorts the end values and spaces out the end-of-line labels. The chart's label series read their positions from there.
- COM gotcha: `Range(...).Value2 = <number>` sometimes throws "Unable to cast ... to String" inside a long-lived PowerShell session. Run Excel scripts with `powershell -NoProfile -File <script>`, or set the value through `.Formula`. Also, PowerShell variable names are case-insensitive (`$PD` and `$pd` are the same variable).

## Version control workflow

The user's main goal for this repo is that every version is saved and easy to revert. So:

- After each meaningful change, commit with a descriptive message and push to `origin main` (https://github.com/HerbieRun/biopharma-financial-models).
- Use a feature branch for larger experiments, then merge back into `main`.
- Undo with `git revert <sha>`, or restore one file with `git checkout <sha> -- <file>`. Never rewrite history or force-push `main`.
- `.gitattributes` marks Office files and PDFs as binary, so git can't diff them. Commit messages for these files should state what changed (assumptions, tabs, outputs).

## Public repository

The GitHub repo is **public**. Before committing, check that files don't contain confidential deal data, non-public company information, or credentials, and ask the user if unsure.
