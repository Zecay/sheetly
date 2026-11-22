# Changelog

All notable changes to Sheetly are listed here, newest first.
Sheetly is in early **alpha**: expect bugs, and please report them in the GitHub issues.

## 0.2.0-alpha

Added:

- Appearance panel (File > Options > Appearance, or View > Appearance) that changes the real app live.
- 28 pre-built themes, one-click themes from a single color, "Surprise me", and your own saved themes.
- Full color control over the top bar, ribbon, sheet, headers, gridlines, selection, menus and text, with gradients.
- Wallpapers (gradient or your own picture), 12 textures, and adjustable see-through sheet and toolbars.
- 10 animated backgrounds including Black Hole, Aurora and Retro Sun, with speed, amount and performance settings. They pause when Sheetly is minimized or not focused.
- Interface font, text size, rounded corners, and optional menu and dialog animations.

Fixed:

- The first cell of a multi-cell selection is now highlighted with the rest, with a thin outline marking the active cell.
- The sheet no longer flashes black when it is resized, such as when opening the Appearance panel.
- Borders on merged cells no longer disappear: fills are painted first, so a neighbouring cell or merge cannot erase them. Borders added in dark mode are now stored as black and shown light on the dark background, so they stay visible on coloured fills and in other apps.
- Saved `.xlsx` files avoid more causes of Excel's repair prompt: invalid or too-long sheet names, overlapping merges, invalid names, unsupported error codes, bad colour scales, and values over Excel's row height, column width and cell text limits.
- Opening `.xlsx` files: 1904-date workbooks, ISO date cells, sheet-scoped names and default column widths are now read correctly.
- Hyperlinks, comments and data validation from opened `.xlsx` files are now kept when saving. Sheetly does not show or edit them yet.

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

