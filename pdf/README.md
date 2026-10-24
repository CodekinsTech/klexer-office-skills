# PDF skill (v2): build notes

Agent-facing instructions are in `SKILL.md`. This file is for maintainers.

## Layout

- `SKILL.md`: plain-language instructions for an AI agent (commands, inputs, outputs, exit codes, examples).
- `scripts\pdf_tool.py`: the single command-line entry point. Twelve subcommands, as the spec lists them.
- `scripts\requirements.txt`: pinned versions, with licences noted.
- `tests\test_pdf_tool.py`: one pytest test per acceptance test in spec section 3 (15 tests), plus 8 tests for `--in-place` and the locked-file message (23 in all). Every input PDF is built inside the test with reportlab or pypdf.

## Libraries and licences

Versions are the ones installed in the build venv (Python 3.11.9, Windows 11).

| Package | Version | Licence | Used for |
|---|---|---|---|
| pypdf | 6.21.2 | BSD-3-Clause | page structure, forms, merge, split, rotate, reorder, watermark (stamp), check, check-text, verification |
| pdfplumber | 0.11.10 | MIT | tables |
| pdfminer.six | 20260107 | MIT (checked in its installed METADATA) | pulled in by pdfplumber |
| reportlab | 5.1.0 | BSD (ReportLab licence, checked in its METADATA) | create, and the text watermark |
| cryptography | 50.0.2 | Apache-2.0 OR BSD-3-Clause | AES decryption for pypdf (crypto provider) |
| pypdfium2 | 5.14.0 | Apache-2.0 / BSD-3-Clause | installed in the venv as the brief asked; not imported by the tool |
| Pillow | 13.0.0 | MIT-CMU | installed; pulled in by pdfplumber; not imported by the tool |
| lxml | 6.1.3 | BSD-3-Clause | installed as the brief asked; not used by the tool |
| pytest | 9.1.1 | MIT | tests only |

Build venv: `<your-venv-folder>\venv-v2-pdf`

## How to run the tests

From the `pdf` folder:

    "<your-venv-folder>\venv-v2-pdf\Scripts\python.exe" -m pytest tests -q

Each test runs `scripts\pdf_tool.py` as a subprocess, so the tests exercise the command line exactly as an agent would. Tests write only to pytest's temporary folders.

## Decisions where the spec was open or unclear

1. **reportlab is used** for create and the text watermark. The brief approved it (option a in spec section 7).
2. **cryptography is added** as pypdf's AES crypto provider. Without it, AES-encrypted files cannot be opened. It is in `requirements.txt`, not in the brief's list.
3. **Conflicting "no text layer" messages.** Spec section 2.1 gives `no text layer; scanned PDF not supported`. The section 4 error table gives `No text layer found. Scanned PDFs need OCR, which this skill does not do.` The tool uses the error-table message everywhere.
4. **Missing stamp file.** The section 4 row "Stamp file missing or no option given" gives the message `Give exactly one of --stamp or --text.`, so a missing stamp file gets that message, as the table says.
5. **Command inputs.** `create` takes the text file as a positional argument (spec section 3 names the input file but gives no option for it). The file's name is not part of the options list.
6. **Required options.** `--out` is required on every single-output write command, plus `--out-dir`, `--values`, `--angle`, `--order` and `--expect`. A missing option then gets the argparse message that names it, exit 1, as section 0 requires. On fields-fill, rotate, reorder and watermark, `--out` and `--in-place` form a required mutually exclusive group, so argparse names both when neither or both are given (decision 29).
7. **`--out` and `--force` on all read commands.** Section 0 makes `--out` a common rule, so it applies to fields-list, check and check-text as well as text and tables.
8. **Read command result shapes.** `fields-list` returns `{"fields": [...]}`. `tables` uses a 1-based `index` per page. `text` adds a `no_text_pages` list. `check` and `check-text` use the shapes in the spec.
9. **Default password name.** With no `--password-env`, the password message names `PDF_PASSWORD`.
10. **Empty user password.** If no password is supplied, the tool tries an empty password. A PDF that only has an owner password (restrictions, no user password) therefore opens with no password. Only a PDF whose user password is set returns exit 2. Revisit this if you want stricter behaviour.
11. **Unencrypted output from an encrypted input.** When the caller supplies the password, outputs are written unencrypted. The caller asked for the operation, so this does not count as decrypting without a request. Outputs are never re-encrypted.
12. **`--in-place` is implemented** on fields-fill, rotate, reorder and watermark (decisions 29 to 31). With `--out` set to an input path, the section 4 refusal still applies.
13. **Form field types.** A text field is `text`. A checkbox is `checkbox`. A combo box is `dropdown`. A list box and a push button are `other`, and fill refuses them. A radio group is `radio`. Checkbox `options` is `[]`. Radio `options` are the widget state names without the slash. Dropdown `options` are the export values from `/Opt`.
14. **Field-fill values.** A checkbox takes JSON true or false, and the on-state name is read from the widget. A text field takes a string. A value of the wrong type gets `Value for <field> must be text.` or `must be one of: true, false.`. These messages are not in the table, so they are new.
15. **Form field page.** `page` is the first page with a widget for the field. A field with widgets on several pages reports the first.
16. **Split.** Parts are named `<stem>.partNN.pdf` and are all checked before any is moved into place. If publishing fails part way, the parts already moved are removed. New messages: `Give exactly one of --every or --ranges.`, `--every must be a whole number of 1 or more.`, and `Ranges must cover each page once. Page <N> is covered more than once.` The output folder is created if it does not exist. A path that exists but is a file gets the "not empty" message.
17. **Reorder verification.** Text is compared after whitespace is collapsed, only when both the input page and the output page have text, as section 2.8 says.
18. **Watermark verification.** The check passes when every word of the original page text still appears in order. Watermark words may sit between them, since text extraction can interleave them.
19. **Text watermark.** Helvetica, grey, 45 degrees, centred on the page, sized to fit at most 72 pt. Characters must be in WinAnsi (cp1252). Control characters are refused. Pages with `/Rotate` or a non-zero mediabox origin are drawn in the unrotated frame.
20. **Create.** A UTF-8 byte-order mark is accepted. CRLF and CR line endings become LF. A tab is written as four spaces. A form feed starts a new page, but a form feed on an untouched page, or at the end of the text, does not create an empty page. An empty input gives one blank page. The output must end in `.pdf`. Margin and font size must be numbers, and the margin must leave room on the page. New messages: `Output must end in .pdf: <path>`, `Could not read <file>: not UTF-8 text.`, `Font size must be greater than 0.`, `Margin must be 0 or more.`, `Margin is too large for the page.`.
21. **Create verification.** Whitespace is removed from both the output text and each input line before the in-order search. This is looser than section 2.10's "whitespace-normalised" wording, because it also ignores spaces inside a line. Wrapping can split a long word across lines, so keeping spaces would give false failures.
22. **Missing reportlab.** Create and the text watermark return exit 2 with `The text-drawing engine (reportlab) is not installed. Run: pip install -r requirements.txt`. This message is new.
23. **Page lists** are returned sorted with duplicates removed.
24. **Read command with `--out`.** Nothing is printed to stdout. The file gets the same text or JSON, and a re-read confirms it.
25. **Check.** The file is opened in strict mode, so a broken cross-reference table or trailer is an error. Warnings that pypdf logs during the check are added to `warnings`.
26. **Reading.** Input files are read fully into memory. This avoids Windows keeping a file locked after reading it.
27. **Unexpected errors** print `Internal error (<ExceptionType>). No output was kept.` and exit 1. No traceback is printed.
28. **No-command call** prints usage and `A COMMAND is required...`, exit 1.
29. **`--in-place` scope.** The spec names `--in-place` in section 5 and the error table but does not list the commands. It is offered on the four commands that change one input: fields-fill, rotate, reorder, watermark. It is not offered on split, merge or create, as section 5 says. The result is written to a temp file in the input's folder, checked, then moved onto the input with `os.replace`. If that move fails, the input is unchanged and the temp file is removed.
30. **`--in-place` and encryption.** Refused for a password-protected input (`--in-place is refused for a password-protected PDF. Write to a new file with --out.`, exit 1). Writing back an unencrypted copy would silently remove the protection. The spec does not cover this case, so this is a choice; revisit it if you want in-place allowed on encrypted files.
31. **No backup for `--in-place`.** The spec does not ask for one, so none is kept. The result is checked before it replaces the input, so a failed check leaves the original as it was. A `.bak` copy would be a spec change.
32. **Locked input.** A PermissionError while reading the input now gives the locked-file message (exit 1). Before, the read sat inside the parse-error handler, so a locked file was reported as corrupt. A locked target on replace gives the same message, as the section 4 row "Target file locked" says.

## Locked-file and viewer checks (Windows 11, Python 3.11.9)

- **Locked input, read.** `test_locked_input_reports_open_in_another_program`: the test holds the input with an exclusive `msvcrt.locking` byte-range lock. `rotate --in-place` and `text` both exit 1 with exactly `The file is open in another program. Close it and try again.`. The input bytes are unchanged and no temp file is left.
- **Locked input, replace.** `test_open_handle_blocks_in_place_replace_with_locked_message`: the test holds the input open with a plain handle (no delete sharing). Reading works, the replace fails with WinError 5, and `rotate --in-place` exits 1 with the same message. The input is unchanged and no temp file is left.
- **Acrobat.** Not installed (no Adobe program folder, no uninstall entry, no `AcroExch.AVDoc` COM class), so no COM check was run.
- **Edge PDF viewer.** Edge is installed, but its viewer is a GUI that cannot be driven here without opening a tab in the user's running browser, so it was not opened. As a proxy, all 13 outputs (every write command, plus the four `--in-place` results) were opened with pypdf in strict mode, pdfplumber and pdfium. pdfium is the PDF engine Chromium uses, and Edge is Chromium-based. Page counts matched in all 13 and page 1 rendered in all 13. This is not a visual check in Edge.

## What is not done

- Optional tools (Poppler, Tesseract): the spec never uses them, since OCR is out of scope, so nothing is detected.
- macOS and Linux are not tested. The code has no Windows-only calls, but only Windows was run.
- Text watermark does not handle rotated pages (decision 19).
- Bookmarks and links after `reorder` were not checked.
- Only Helvetica is used. Thai, Arabic and other scripts are refused by design.
- AES-256 encryption is what the tests use. RC4 and AES-128 files were not run through the tool.
- `pypdfium2`, `Pillow` and `lxml` are installed but not used by the tool.
