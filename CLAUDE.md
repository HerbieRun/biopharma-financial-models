# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Biopharma financial models (e.g. rNPV, DCF, revenue forecasts, comparables). The repo is newly scaffolded: there is no build system, toolchain, or test suite yet. Update this file once one is chosen.

- `models/`: model files (Excel, Python, etc.)
- `data/`: input data and assumptions
- `docs/`: notes, write-ups, and outputs

## Version control workflow

The user's main goal for this repo is that every version is saved and easy to revert. So:

- After each meaningful change, commit with a descriptive message and push to `origin main` (https://github.com/HerbieRun/biopharma-financial-models).
- Use a feature branch for larger experiments, then merge back into `main`.
- Undo with `git revert <sha>`, or restore one file with `git checkout <sha> -- <file>`. Never rewrite history or force-push `main`.
- `.gitattributes` marks Office files and PDFs as binary, so git can't diff them. Commit messages for these files should state what changed (assumptions, tabs, outputs).

## Public repository

The GitHub repo is **public**. Before committing, check that files don't contain confidential deal data, non-public company information, or credentials, and ask the user if unsure.
