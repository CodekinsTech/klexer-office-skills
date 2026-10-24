# PPTX skill: developer notes

Built from `specs/pptx.md` (v2 draft). Commands and options are described for agents in `SKILL.md`.

## Layout

- `scripts/pptx_tool.py` - the one command-line entry point (all commands).
- `requirements.txt` - runtime and test packages.
- `tests/test_pptx_tool.py` - pytest suite. One test per acceptance test in spec section 3, plus tests for the open items in spec section 7, the general rules, and the refused features.
- `tests/test_powerpoint_com.py` - opens each written deck in Microsoft PowerPoint through COM (WithWindow false) and checks it opens without a repair prompt and shows the slide count and text the tool reports. Skips when Windows, PowerPoint or pywin32 is missing.
- `SKILL.md` - instructions for AI agents.

## Libraries and licences

Checked from the installed packages in the test virtual environment (Python 3.11.9, Windows 11).

| Package | Version tested | Licence | Used for |
|---|---|---|---|
| python-pptx | 1.0.2 | MIT | all deck reading and writing |
| lxml | 6.1.3 | BSD-3-Clause | XML parsing in `validate`, and python-pptx's element tree |
| Pillow | 13.0.0 | MIT-CMU (PIL licence) | image format detection (python-pptx depends on it, so it is always installed) |
| XlsxWriter | 3.3.0 | BSD-2-Clause | chart workbooks written by python-pptx (its dependency) |
| pytest | 9.1.1 | MIT | tests only |

Note on spec section 7: the spec says replace-image needs no Pillow. That is not true with python-pptx 1.0.2, which requires Pillow. The tool uses python-pptx's own image handling; Pillow is installed as a dependency.

## Running the tests

From this folder, with a venv that has the packages in `requirements.txt`:

    python -m pytest tests -q

Last run: 55 passed, 0 failed, 0 skipped. The 15 PowerPoint tests take most of the run time (PowerPoint start-up and one 90 second control). Tool tests build decks with python-pptx inside each test, in pytest's temporary folders. Without pywin32 or PowerPoint, the PowerPoint tests skip and the rest still run.

PowerPoint tests: each PowerPoint call runs in a child process with a 90 second timeout. A repair prompt blocks PowerPoint's COM call, so a timeout fails the test. On timeout only the POWERPNT.EXE instances started during that test are stopped. A control test checks that a deck with a dangling picture relationship is flagged. Some decks in these tests are authored by PowerPoint itself (sections, notes, a table, a picture).

## Decisions where the spec was open

1. Default output name is `<stem>.<command>.pptx` next to the input. The default split folder is `<stem>.split`.
2. `--in-place` exists (the spec's error table names it). Combining it with `--out` is an error (exit 1).
3. `--no-include-notes` is the off switch for notes (the spec's `--include-notes` default on needs a negative form).
4. Outline: indentation of two or more spaces before `- ` is level 2. Any other non-empty line after a title is a level 1 paragraph. Empty bullets are skipped. A UTF-8 byte-order mark is accepted.
5. `create --template` removes the template's existing slides before building.
6. `read` includes the title in `texts`. Empty text shapes are skipped. Groups are expanded for text, tables and charts. `--out` always writes JSON.
7. `replace`: `per_slide` includes notes hits of that slide. `replacements` = slide hits + notes hits, so the spec's example gives 4. `per_slide` lists only slides with hits. When several pairs are given (`--pairs`), they run in the JSON file's order, and a later find can match an earlier replacement.
8. `replace` verification: with one pair whose replacement does not contain the find text, the tool checks that the find text is gone and that the replacement count rose by the number of hits. When the replacement contains the find text, only the find-text remainder is checked, and the result says so. With several pairs, the file is compared with the in-memory result.
9. `not_searched` always lists chart text, SmartArt, alt text and masters/layouts, plus speaker notes when notes are off.
10. `add`: unused empty placeholders are removed. If the layout has no body (for `--body`) or no title (for `--title`) placeholder, the tool exits 1 with a message.
11. `delete` also removes the deleted slides from section lists and custom shows. After deleting, slide part names are renumbered, because python-pptx would otherwise pick a name that is already in use when a slide is added later. Slide-jump hyperlinks on remaining slides that point at a deleted slide are not cleaned up.
12. `duplicate` copies the shape XML, background and notes text, and re-links pictures and hyperlinks through the relationship of the same target. Animations, transitions and comments are not copied.
13. `replace-image`: the new picture is stretched to the old picture's box, so aspect ratio is not kept. The old picture's name is kept. With no `--shape` or `--shape-index`, a slide with exactly one picture uses it. Only top-level pictures count (not pictures inside groups). The old image relationship is removed only when no other element on the slide uses it.
14. `chart-data`: editable charts are bar, column, line and pie only (python-pptx's BarPlot, LinePlot, PiePlot). Area, doughnut, radar, XY and bubble charts are refused with exit 2. Categories are stored as text. The test confirms a series' fill colour survives `replace_data`.
15. `check-text` joins a slide's text with spaces before searching, so a phrase can match across two shapes on the same slide. Each slide's notes count as that slide's text.
16. `validate`: a check that cannot run after an earlier failure is reported as `skipped`, and it does not count as a failure.
17. Legacy `.ppt` and encrypted decks are detected by file signature (an OLE compound file). An encrypted deck is one with an `EncryptedPackage` stream. A macro-enabled deck is one with `vbaProject.bin` or a macro-enabled content type. The file extension alone also triggers the `.ppt`/`.pptm`/`.potm` refusal.
18. Messages not in the spec's error table (all exit 1 unless noted):
    - `Output must end in .pptx: <path>`
    - `Output folder does not exist: <folder>`
    - `Use --out or --in-place, not both.`
    - `Position must be between 1 and <n>. The deck has <count> slides.`
    - `Layout "<name>" has no body placeholder for --body lines.` (and the title equivalent)
    - `Give --shape NAME or --shape-index K. Slide <N> has <count> pictures.`
    - `Give --find and --replace, or --pairs FILE.json.`
    - `Give exactly one of --every or --ranges.`
    - `--every must be a whole number of 1 or more.`
    - `Output path is not a folder: <path>`
    - `The deck has no slides to split.`
    - `Ranges must cover each slide once. Slide <N> is covered more than once.`
    - `Outline is not UTF-8 text: <path>`
    - `The template has no usable layout. A layout needs a title and a body placeholder.`
    - `Could not read <file>: expected ...` for JSON of the wrong shape (the file-missing case uses the spec's `Input file not found` message, which is also used for JSON and expect files).
    - `Could not read <file>: <reason>` for a file that is not valid JSON or UTF-8 is from the spec.
    - An argparse error (missing or unknown option) is exit 1, not argparse's default 2.
19. Locked files: any `PermissionError` (on read, on write or on the final move) gives the spec's locked-file message, exit 1.
20. Set `PPTX_TOOL_DEBUG=1` to print the traceback for an internal error.
21. Refused features (decision): a request for a feature the tool does not provide exits 2 with `Not supported: <feature>. <alternative>.` and writes nothing. The names are in `UNSUPPORTED_COMMANDS` in `pptx_tool.py`: `render`, `export` and `pdf` (slides to PDF or images), `animate`, `transition`, `theme` and `master`. Those names are hidden from `--help`. An unknown command name or a usage typo is an argparse error and exits 1.

## Not done

- No render or PDF output. `render`, `export` and `pdf` are refused with exit 2 and the `Not supported` message. No renderer is detected.
- PowerPoint checks cover the decks the tool writes in the test runs, not every deck a user might bring. A repair that PowerPoint does without a prompt would not be seen. None were seen in the runs.
- Tested on Windows 11 only. macOS and Linux paths (the LibreOffice lookup) are not implemented, because no command needs LibreOffice.
- Not implemented: animations and transitions, comments, slide master or theme editing, editing text inside charts or SmartArt, alt text in `replace`.
- `validate` does not check every content type entry or every relationship type name. It checks the items listed in the spec.
- `split` never copies slide content (it deletes slides from copies of the deck), so the duplicate refusals do not apply to it.
- Deleting a slide does not clean up hyperlinks from other slides that jump to it.
