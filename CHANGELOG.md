# Changelog

## [1.0.0] — 2026-10-07

First public release.

- **PDF → Markdown** via PyMuPDF4LLM: headings, bold/italic, tables, lists, pictures saved as PNG.
  - Repeated page headers/footers are detected (same text on at least half of the pages) and cut off.
  - Table-of-contents lines and "Page N of M" lines are removed.
  - Optional page-break markers, picture quality 100–300 dpi.
  - Warning for scanned PDFs without a text layer; password-protected PDFs are reported.
- **Word .docx → Markdown** via mammoth + markdownify.
  - Headers/footers and review comments are skipped; table of contents removed.
  - Letters typed in the Symbol font (µ, Ω, α …) become real Unicode characters.
  - Caption numbers (`Figure 3`, `Table 1`) are kept; auto-numbered headings get their numbers back (1, 1.1, 1.1.1).
  - Tables with an empty header row get their first row as the header.
  - EMF/WMF pictures are converted to PNG on Windows.
- **Window** (`pdf2md_gui.py`): file list, folders, options, progress, remembers settings.
- **Command line** (`pdf2md.py`): files and folders, `--out`, `--no-images`, `--keep-headers`, `--keep-toc`,
  `--no-numbering`, `--page-sep`, `--dpi`, `--version`.
- `Manual.pdf` and `Manual.docx` converted together no longer overwrite each other (`Manual_pdf.md`, `Manual_docx.md`).
- PyMuPDF4LLM pinned to 0.0.27 … 0.3.x (classic engine, no 200 MB layout/AI dependencies); a 1.x version, if
  installed, is switched to the classic engine automatically.
