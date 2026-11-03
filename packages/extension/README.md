# Office Workbench

View Word, Excel, PowerPoint and PDF files inside VS Code, and edit Word, Excel, PowerPoint and CSV files without leaving the editor. No Microsoft Office installation is needed.

> **Preview.** Office Workbench is in early development. Saving only rewrites the parts of a file you changed, but keep a copy (or use Git) for files that matter.

## Features

### Viewers

| Format | Extensions | What you get |
|---|---|---|
| Word | `.docx` `.docm` | Page view, zoom, fit to width, page numbers |
| Excel | `.xlsx` `.xlsm` | Sheet tabs, number and date formats, fonts, fills, borders, merged cells, column widths and row heights, formula bar, keyboard navigation, virtual scrolling for large sheets |
| PowerPoint | `.pptx` `.pptm` | Slide list, layout and master inheritance, theme colors, shapes, gradients, pictures, table styles, speaker notes, hidden slides |
| PDF | `.pdf` | Page navigation, zoom, text search and selection, outline (bookmarks), Japanese and other CJK character maps |

Charts, SmartArt and conditional formatting are not rendered yet; they are shown as placeholders.

### Excel and CSV editing

- Type into cells (IME supported), edit with F2 or a double-click, or use the formula bar. Numbers, `1,234`, `12.5%`, dates, times, `TRUE`/`FALSE` and `=formulas` are recognized; start with `'` for literal text.
- Select ranges with the mouse or Shift+arrows, clear with Delete, copy and paste with Ctrl+C / Ctrl+X / Ctrl+V (tab-separated text, compatible with Excel).
- Insert and delete rows and columns, add, rename and delete sheets. Formulas, merged cells, chart ranges and frozen panes follow the change.
- Undo, redo and save with the usual VS Code shortcuts. Unsaved changes survive a restart (Hot Exit).
- Formulas are not calculated by the editor; Excel recalculates them when the file is next opened.

`.csv`, `.tsv` and `.tab` files open as text by default. Use **Open as Table** in the editor title bar or the Explorer context menu to edit them in the grid. UTF-8 (with or without BOM), Shift_JIS and UTF-16 are detected and kept on save, along with the delimiter and line endings; only edited rows are rewritten. To always open CSV files as a table, add `"*.csv": "officeWorkbench.csv"` to `workbench.editorAssociations`.

### Word editing

Click **Edit** in the toolbar to open an edit panel next to the page view. Edit paragraph text, press Enter to split a paragraph, Backspace on an empty paragraph to delete it. Paragraph styles, lists and alignment are kept. Paragraphs that contain fields, hyperlinks, pictures, tracked changes, comments or content controls are locked so they cannot be broken.

### PowerPoint editing

Double-click a shape, text box or table cell on a slide to edit its text in place. Right-click a slide thumbnail to duplicate, delete or move it, or drag thumbnails to reorder.

### Text view and Git diff

- **Open as Markdown** shows a Word document, workbook or presentation as read-only Markdown that updates when the file changes.
- **Compare with HEAD as Text** diffs the working copy against the Git `HEAD` version as Markdown, so you can review changes to Office files line by line. Available from the editor title bar, the Explorer and the Source Control view.

### Language model tools

Office Workbench adds tools that chat agents in VS Code (for example GitHub Copilot in agent mode) can use. Reference them in chat as `#readOffice`, `#editSpreadsheet`, `#editDocument` and `#editPresentation`.

| Tool | What it does |
|---|---|
| Read Office file | Reads Word, Excel, PowerPoint and CSV files as Markdown, including unsaved changes in open editors. Excel reads can be limited to a sheet or range. |
| Edit spreadsheet | Sets and clears cells, inserts and deletes rows and columns, adds, renames and deletes sheets. |
| Edit Word document | Rewrites, inserts and deletes paragraphs, and replaces text. |
| Edit PowerPoint presentation | Sets text in shapes and table cells, replaces text, duplicates, deletes and moves slides. |

You are asked to confirm every edit. Edits are applied in the editor as undoable steps and are left unsaved for you to review unless the tool is told to save.

## Requirements

VS Code 1.95 or later. Works on Windows, macOS and Linux; Microsoft Office is not required.

## Opening files with another editor

If another extension also handles `.pdf` (or other) files, VS Code asks which editor to use. Switch at any time with **Reopen Editor With...** from the editor title context menu.

## License

[MIT](LICENSE). Third-party components and their licenses are listed in [ThirdPartyNotices.txt](ThirdPartyNotices.txt).
