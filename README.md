# klexer-office-skills

Command-line tools for Word, PDF, Excel and PowerPoint files, built to be used by AI agents and people.

| Skill | What it does | Tool |
|---|---|---|
| `docx/` | Read, create and change Word files; tracked changes, comments, fill, replace, to-PDF | `scripts/docx_tool.py` |
| `pdf/` | Read, check, fill, merge, split, rotate, watermark and create PDF files | `scripts/pdf_tool.py` |
| `xlsx/` | Read and write workbooks, recalculate formulas, check values | `scripts/xlsx_tool.py` |
| `pptx/` | Create, read, edit and check PowerPoint decks | `scripts/pptx_tool.py` |

Each skill folder has:
- `SKILL.md`: plain instructions for an agent (commands, inputs, outputs, exit codes).
- `README.md`: maintainer notes, libraries used, decisions, and what is not done.
- `scripts/`: the tool and its requirements.
- `tests/`: automated tests.
- `benchmark/`: accuracy tasks with expected results, a runner, and the latest results.

Exit codes: 0 success; 1 a user-fixable error (bad path, locked file, unknown command); 2 refused or not supported.

Requirements: Python 3.11 or newer. Run `pip install -r scripts/requirements.txt` in each skill folder. Microsoft Word, Excel or PowerPoint are optional; when present, some checks use them.

Licence: MIT (see `LICENSE`).
