# Biopharma Financial Models

Financial models for biopharma companies and assets (e.g. rNPV, DCF, revenue forecasts, comparables).

## Project layout

| Folder    | Contents                                   |
|-----------|--------------------------------------------|
| `models/` | Financial model files (Excel, Python, etc.) |
| `data/`   | Input data and assumptions                 |
| `docs/`   | Notes, write-ups, and outputs              |

> This repository is **public**: don't commit confidential data, credentials, or non-public information.

## Saving and reverting with git

```bash
git status                       # see what changed
git add .                        # stage all changes
git commit -m "Describe change"  # save a version locally
git push                         # back it up to GitHub

git log --oneline                # list saved versions
git revert <commit>              # undo a commit (keeps history)
git checkout <commit> -- <file>  # restore one file from an older version
```
