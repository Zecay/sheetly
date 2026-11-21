# Changelog

All notable changes to Sheetly are listed here, newest first.
Sheetly is in early **alpha**: expect bugs, and please report them in the GitHub issues.

## 0.1.0-alpha (first public release)

First public version.

- Virtualized grid of 1,048,576 rows x 16,384 columns, selection, keyboard, editing, resizing, freeze panes.
- Formula engine with dependency graph, about 130 functions, autocomplete with argument hints, coloured references, named ranges.
- Undo/redo, clipboard (also tab-separated text), fill handle, insert/delete, find/replace, sort and filter.
- Formatting, Format Cells, format painter, conditional formatting, cell styles.
- Multiple sheets, `.xlsx` / `.csv` / `.sheetly` files, autosave recovery, recent files.
- Zoom, context menus, tooltips, shortcuts, print to PDF, dark mode.
- Hardening: limits on nested formulas, huge selections and hostile files, atomic saves, safer recovery.
- Borders menu with icons; "Merge cells" option in Format Cells > Alignment; borders on merged cells draw correctly.
- Windows installer about 1.2 MB, portable exe about 3.4 MB.

Known gaps: no charts; Windows only (Linux, macOS and phones planned); the UI has only been tested by hand on a few files.

