# Sheetly

A small, fast, offline spreadsheet for Windows. It looks and feels like a modern desktop spreadsheet:
ribbon tabs, a formula bar with `fx`, name box, a virtualized grid of 1,048,576 rows × 16,384 columns (A to XFD),
sheet tabs, a status bar with Average / Count / Sum, zoom, and light and dark mode.

> **Early alpha (0.2.0-alpha).** Expect bugs. Please open an issue if you find one. See [CHANGELOG.md](CHANGELOG.md).

Sheetly is an independent project and is not affiliated with Microsoft. It never connects to the internet and collects no data.
Released under the [MIT license](LICENSE).

## Install (friends: read this)

1. Get `Sheetly_0.2.0_x64-setup.exe` (the installer) from whoever sent it to you. It is about 1.2 MB.
2. Double-click it. It installs for your user only, so no administrator password is needed.
3. **Windows SmartScreen will warn you.** Sheetly is not code-signed (a signing certificate costs money), so Windows
   shows *"Windows protected your PC"*. This is expected for any unsigned app.
   - Click **More info**
   - Click **Run anyway**
4. Start Sheetly from the Start menu.

Requirements: Windows 10 or 11 with the Microsoft Edge WebView2 runtime. Windows 11 and up-to-date Windows 10 already have it.
If the installer says it is missing, install the "Evergreen" WebView2 runtime from Microsoft once.

If you received the bare `sheetly.exe` instead, just double-click it (same SmartScreen steps). It needs no installer.

## What you can do

- **Grid**: smooth scrolling, click / drag / Shift / Ctrl selection, arrow keys, Ctrl+arrows, Home, Ctrl+Home/End,
  PageUp/PageDown, F2 to edit, type to replace, Esc to cancel, Delete to clear, resizable rows and columns
  (double-click the border to autofit), freeze panes.
- **Formulas**: `+ - * / ^ & = <> < > <= >= %`, `$A$1` references, ranges, `Sheet2!A1`, named ranges, the errors
  `#DIV/0! #REF! #NAME? #VALUE! #N/A #NUM! #CIRC!`, and about 130 functions (SUM, AVERAGE, COUNTIF(S), SUMIF(S),
  IF, IFS, IFERROR, VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH, TEXT, TEXTJOIN, DATE, TODAY, RAND and many more).
  Typing `=` shows an autocomplete list with argument hints, and references are colour-coded while you edit.
  Only the cells affected by a change are recalculated.
- **Editing**: undo/redo (200 steps), copy / cut / paste (also to and from other programs as tab-separated text),
  fill handle (copies, adjusts references, continues 1,2,3 / Mon,Tue / Jan,Feb / Item 1, Item 2),
  insert and delete rows and columns with formulas updated, Find & Replace, sort, AutoFilter.
- **Formatting**: font, size, bold / italic / underline / strikethrough, colours, borders, alignment, wrap, merge,
  number formats (General, Number, Currency, Accounting, Percent, Date, Time, Fraction, Scientific, Text, Custom),
  the Format Cells dialog (Ctrl+1), Format Painter, conditional formatting (highlight rules, top/bottom, data bars,
  colour scales) and cell styles.
- **Files**: open and save `.xlsx` and `.csv` (and Sheetly's own `.sheetly`), multiple sheets (add, rename, delete,
  reorder, colour), recent files, autosave recovery every minute, print or save as PDF.
- **Polish**: zoom 10–400 % (also Ctrl + mouse wheel), right-click menus, tooltips on ribbon buttons, Excel-style
  keyboard shortcuts, dark mode (View tab).

Charts are not included yet.

## Build it yourself

Needs Node.js 18+, Rust (stable), and the Visual Studio "Desktop development with C++" workload (MSVC + Windows SDK).

```powershell
npm install
npm test                # formula engine, file formats and editing tests
npm run tauri build     # produces the installer and the exe
```

Output:

- `src-tauri/target/release/sheetly.exe` (portable, 3.2 MB)
- `src-tauri/target/release/bundle/nsis/Sheetly_0.2.0_x64-setup.exe` (installer, 1.2 MB)

The release profile uses `opt-level = "s"`, `lto = true`, `codegen-units = 1`, `strip = true`, `panic = "abort"`.

## Project layout

- `src/core/` – the spreadsheet engine with no DOM dependency: parser, evaluator, functions, dependency graph and
  recalculation, number formats, editing operations, `.xlsx` / `.csv` / native readers and writers, conditional formats.
- `src/ui/` – canvas grid, cell editor, ribbon, dialogs, file menu, printing.
- `src-tauri/` – the Rust shell: a window plus a few commands for reading and writing files.
- `tests/` – unit tests (`npm test`).
