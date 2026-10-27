# Changelog

SHEETDIFF follows [semantic versioning](https://semver.org): `MAJOR.MINOR.PATCH`.
The version lives in one place, `sheetdiff/__init__.py`. The build script, the window title,
the reports and `--version` all read it from there. Each release gets a git tag `vX.Y.Z` and a
section in this file.

## 0.4.1 (2026-10-27)

* Fixed: the exe (and the window started from Python) crashed when files were dragged into the
  window. The window procedure called Tk while Tk was handling the message; it now only stores
  the dropped paths and the window picks them up with its own timer.
* The comparison thread no longer calls Tk directly either; it hands the result over through a
  queue.
* New test `tests/gui_drop.py`: drops files through the Windows message queue, the same way
  Explorer does, in a separate process.

## 0.4.0 (2026-10-27)

First public version.

* English user interface and reports next to Czech (`--lang en`, language switch in the window,
  remembered in `%APPDATA%\SHEETDIFF\settings.json`).
* Drag and drop: files can be dropped into the window; two files dropped on the exe are compared
  right away. Two files are ordered by save date, the older one becomes A.
* Single-file Windows build with Nuitka (`build_exe.bat`), one exe for the window and the command
  line (`--cli`).
* Better pairing when rows are inserted next to edited rows: rows in such a block are matched by
  content similarity instead of by position.
* Moved shapes are reported with their address in the old file.
* Invented example workbooks (`examples/`), screenshots and tests that generate their own data.

## 0.3.0 (2026-10-27)

* Conditional formatting (classic and x14 rules), data validation and hyperlinks are compared.
  Their ranges and formulas follow inserted rows and renamed sheets.
* VBA macros are reported as changed, added or removed (content hash).
* Insertion and deletion markers in the word diff: the old text shows where the new version
  inserted text, the new text shows where text was deleted.

## 0.2.0 (2026-10-27)

* Cell formatting is compared by appearance (number format, font, fill, border, alignment,
  protection) and grouped into ranges. Merged cells, row heights, column widths, hidden and
  grouped rows and columns.
* Inserted and deleted rows and columns are detected by content. Formula references, named
  ranges and chart series of the old file are shifted the way Excel shifts them.
* Renamed sheets are detected, references to them are followed.
* Comments (classic and threaded).
* Optional Excel report.

## 0.1.0 (2026-10-27)

* Package, structure, cell values and formulas (shared formulas included), calculated values,
  text boxes and shapes, pictures and charts.
* HTML report with a verdict, overview by sheet, filters, search and a word diff.
* Command line with exit codes, a small window started from `SHEETDIFF.bat`.
