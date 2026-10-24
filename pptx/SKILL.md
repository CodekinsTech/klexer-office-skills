# PPTX skill

A command-line tool for PowerPoint `.pptx` decks. One program, `scripts/pptx_tool.py`, with these commands:
create, read, replace, add, delete, reorder, duplicate, split, replace-image, chart-data, check-text, validate.

Run it from any folder with Python 3.11 and the packages in `requirements.txt`:

    python "<path to>/pptx/scripts/pptx_tool.py" COMMAND [options]

Paths can be absolute or relative to the current folder. Paths with spaces and non-English text work.

## Rules that apply to every command

- Slides are numbered from 1 in deck order.
- A slide list is written like `1-3,5,8-`. A trailing dash means "to the last slide".
- Read commands (`read`, `check-text`, `validate`) take `--json`. The JSON is one object on stdout. Without `--json` you get plain text.
- `--out FILE` on a read command writes the JSON result to a new file. It refuses an existing file unless `--force` is given.
- Errors and failed checks go to stderr. The tool never asks for input.
- A missing required option is exit 1, and the message names the option.
- Commands that change a deck write a new file by default: `<name>.<command>.pptx` next to the input (for example `deck.replace.pptx`). Use `--out FILE` to choose the file. `--force` replaces an existing output. Only `--in-place` changes the input file itself.
- Every output is saved to a temporary file, re-opened and checked, and only then moved into place. If a check fails, nothing is written and the exit code is 1.
- Legacy `.ppt`, macro-enabled `.pptm` and `.potm`, and password-protected decks are refused with exit 2. They are never unlocked or stripped.
- A request for a feature the tool does not provide is refused with exit 2 and prints `Not supported: <feature>. <alternative>.` Refused names: `render`, `export` and `pdf` (slides to PDF or images), `animate`, `transition`, `theme` and `master`. Nothing is written.
- An unknown command name or a usage typo (a wrong or missing option) exits 1.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Done (and verified). For check commands: everything found. |
| 1 | Problem with input, options or the result: file missing, bad slide list, verification failed, check found missing text or a failed check; also an unknown command name or a usage typo. |
| 2 | The file, content or feature is not supported: legacy, encrypted, macro-enabled, a non-category chart, content that duplicate cannot copy, or a refused feature such as `render` (prints `Not supported: ...`). |

## Commands

### create
Make a deck from a plain text outline.

- Input: `OUTLINE.txt` (UTF-8). Options: `--out FILE.pptx` (required), `--template FILE.pptx`, `--force`, `--json`.
- Outline format:
  - `# Title` starts a slide with the "Title and Content" layout.
  - `- text` is a bullet at level 1. Two spaces then `- text` is level 2.
  - Any other non-empty line after a title is a level 1 paragraph.
  - `Notes: text` adds speaker notes. Several `Notes:` lines are joined with new lines.
  - Blank lines are ignored. Text before the first `#` line is an error.
- With `--template`, existing slides in the template are removed and its layouts are used. If "Title and Content" is not in the template, the first layout with a title and a body is used, and the result says so.
- Output: `{"output", "slides", "layout", "layout_note", "verified"}`.
- Exit: 0 created and verified; 1 outline missing, empty, not UTF-8, text before the first title, or the template has no usable layout; 2 template is legacy or encrypted.

### read
Print the text of every slide.

- Input: `DECK.pptx`. Options: `--slides LIST` (default all), `--json`, `--out FILE`, `--force`.
- JSON shape:
  `{"slide_count": N, "slides": [{"number", "layout", "title", "texts", "tables", "charts", "notes"}]}`
  - `texts`: text of every shape that holds text, groups expanded, in stacking order (the title is included).
  - `tables`: a list of tables, each a list of rows of cell text.
  - `charts`: each chart has `type`, `categories`, and `series` (each with `name` and `values`; a blank point is `null`).
  - `notes`: speaker notes text, or an empty string. Reading never creates a notes page.
- Exit: 0 read; 1 file missing, corrupt, or bad slide list; 2 legacy, encrypted, or macro-enabled.

Example:

    python pptx_tool.py read deck.pptx --slides 1-2 --json

### replace
Replace text across slides, tables and speaker notes.

- Input: `DECK.pptx`. Options: either `--find TEXT --replace TEXT`, or `--pairs FILE.json` (an object that maps find text to replacement text). `--include-notes` / `--no-include-notes` (default on). `--out`, `--in-place`, `--force`, `--json`.
- Matching is case-sensitive and works inside one paragraph. A match may span several runs of text.
- Searched: text of all shapes (groups expanded), table cells, and speaker notes when notes are included.
- Not searched: chart text, SmartArt, alt text, text in masters and layouts. The result lists these in `not_searched`.
- Output: `{"output", "replacements", "per_slide", "notes_replacements", "not_searched", "verification", "verified"}`. `per_slide` maps slide numbers (as text) to hits on that slide, including that slide's notes hits. Only slides with hits are listed. `replacements` counts everything, including notes.
- Zero hits is exit 0, and the copy is still written.
- Exit: 0 done and verified; 1 options missing, find text empty, or bad JSON; 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py replace deck.pptx --find ACME --replace Globex --json

### add
Add one slide.

- Input: `DECK.pptx`. Options: `--at N` (insert before slide N; default at the end), `--layout NAME` (default "Title and Content"), `--title TEXT`, `--body LINE` (repeatable; `- ` is level 1, two spaces plus `- ` is level 2, other text is level 1), `--notes TEXT`, `--out`, `--in-place`, `--force`, `--json`.
- Empty placeholders the slide does not use are removed.
- Output: `{"output", "slide_count", "position", "verified"}`.
- Exit: 0 added and verified; 1 layout not found, position out of range (1 to count+1), layout has no matching placeholder, or output exists; 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py add deck.pptx --at 2 --title "Next steps" --body "- Hire" --body "  - two engineers"

### delete
Delete slides.

- Input: `DECK.pptx`. Options: `--slides LIST` (required), `--out`, `--in-place`, `--force`, `--json`.
- The deck must keep at least one slide.
- Output: `{"output", "deleted", "slide_count", "verified"}`.
- Exit: 0 deleted and verified; 1 slide list invalid or it would delete every slide; 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py delete deck.pptx --slides 2,4-5

### reorder
Change slide order.

- Input: `DECK.pptx`. Options: `--order LIST` (every slide number exactly once, for example `3,1,2`), `--out`, `--in-place`, `--force`, `--json`.
- Output: `{"output", "order", "verified"}`. Slide content is not changed.
- Exit: 0 reordered and verified; 1 not a complete order, or bad list; 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py reorder deck.pptx --order 3,1,2

### duplicate
Copy one slide to the position right after it.

- Input: `DECK.pptx`. Options: `--slide N` (required), `--out`, `--in-place`, `--force`, `--json`.
- Copies the shapes, pictures (the same image is used), background and notes text. Animations and transitions are not copied.
- Refused with exit 2 if the slide holds a chart, an OLE object, a video or audio clip, or a SmartArt diagram. The message names which one.
- Output: `{"output", "source_slide", "new_slide", "verified"}`.
- Exit: 0 duplicated and verified; 1 slide number out of range or output exists; 2 unsupported content, legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py duplicate deck.pptx --slide 1 --out deck_copy.pptx

### split
Split a deck into several decks.

- Input: `DECK.pptx`. Options: exactly one of `--every N` or `--ranges "1-3,4-6"`, `--out-dir DIR` (default `<name>.split` next to the input; must not exist or must be empty), `--json`.
- Parts are named `<name>.part01.pptx`, `<name>.part02.pptx` and so on. Each part is checked before any file is moved into the folder.
- Ranges must cover every slide exactly once.
- Output: `{"out_dir", "parts": [{"file", "first_slide", "last_slide", "slides"}], "verified"}`.
- Exit: 0 split and verified; 1 range invalid, slides uncovered, or folder not empty; 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py split deck.pptx --every 5 --out-dir parts

### replace-image
Swap one picture on a slide for a new image.

- Input: `DECK.pptx`. Options: `--slide N` (required), `--shape NAME` or `--shape-index K` (K counts pictures on that slide from 1, in stacking order). If the slide has only one picture, neither option is needed. `--image FILE` (required; PNG, JPEG, GIF, BMP or TIFF), `--out`, `--in-place`, `--force`, `--json`.
- The new picture takes the old picture's position, size and stacking place, and keeps its name. The new image is stretched to the old picture's box.
- Cropping on the old picture is not carried over. The result says so.
- Output: `{"output", "slide", "shape", "cropped_dropped", "verified"}`.
- Exit: 0 replaced and verified; 1 slide or picture not found, not a picture, image missing or unsupported; 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py replace-image deck.pptx --slide 3 --image logo.png

### chart-data
Replace the data of a category chart (bar, column, line or pie). The chart type and series formatting are kept.

- Input: `DECK.pptx`. Options: `--slide N` (required), `--shape NAME` (needed when the slide has more than one chart), `--data FILE.json`, `--out`, `--in-place`, `--force`, `--json`.
- Data file:

      {"categories": ["Q1", "Q2"], "series": [{"name": "Sales", "values": [10, 12.5]}]}

  Every series needs one value per category. `null` is a blank point. Category names are stored as text.
- Output: `{"output", "slide", "series", "categories", "verified"}`.
- Exit: 0 updated and verified; 1 slide or chart not found, data file invalid, value count mismatch, or a value that is not a number; 2 chart is not a bar, column, line or pie chart, or legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py chart-data deck.pptx --slide 2 --data new_data.json

### check-text
Check that phrases appear somewhere in a deck.

- Input: `DECK.pptx`. Options: `--expect FILE.txt` (required; one phrase per line, blank lines ignored), `--include-notes` / `--no-include-notes` (default on), `--out FILE`, `--force`, `--json`.
- Searches slide text, table cells, chart categories and series names, and speaker notes. Runs of spaces, tabs and line breaks count as one space. Matching is case-sensitive.
- Output: `{"found": [{"phrase", "slide"}], "missing": [phrase, ...]}`. `slide` is the first slide where the phrase appears.
- Exit: 0 all phrases found; 1 one or more missing (the list is printed to stderr, and the JSON still lists every phrase); 2 legacy, encrypted or macro-enabled.

Example:

    python pptx_tool.py check-text deck.pptx --expect must_have.txt --json

### validate
Check that the deck's package is well formed.

- Input: `DECK.pptx`. Options: `--out FILE`, `--force`, `--json`.
- Checks, in order: readable zip; content types and presentation part present; every XML and relationship part parses; every internal relationship target exists; python-pptx opens the deck; slide ids are unique and at least 256; every slide has a layout relationship; each notes page points back to its slide; each picture's image relationship resolves. A check that cannot run because an earlier one failed is listed as `skipped`.
- Output: `{"valid": true|false, "checks": [{"name", "status": "pass"|"fail"|"skipped", "detail"}]}`.
- Exit: 0 valid; 1 any check failed or file missing; 2 legacy, encrypted or macro-enabled (not checked).

Example:

    python pptx_tool.py validate deck.pptx --json

## Common messages

- `Input file not found: <path>`
- `File is not a readable .pptx (corrupt or truncated): <path>`
- `Output already exists: <path>. Choose another --out or pass --force.`
- `Refusing to overwrite the input. Use --in-place to change it.`
- `Slide <N> does not exist. The deck has <count> slides.`
- `Slide list is invalid: <list>.`
- `The file is open in another program. Close it and try again.` (exit 1; the file is locked, usually by PowerPoint)
- `Internal error (<exception type>). No output was kept.` (exit 1)
