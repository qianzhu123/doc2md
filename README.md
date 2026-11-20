# doc2md

Convert common document formats (docx, pdf, pptx) into agent-friendly Markdown,
with images extracted to disk and a metadata sidecar for quick triage.


Built for AI agent workflows: the output is a predictable product directory
(`article.md` + `images/` + `meta.json`) so an agent can read the converted
markdown instead of guessing at binary formats.

## Supported formats

| Format | Status | Engine |
|---|---|---|
| `.docx` | v1 (current) | mammoth (semantic) + python-docx (tables) + raw zip (media) |
| `.pdf` | v1.1 (current) | PyMuPDF: text layer + font-size heading detection + image block rendering; scanned-PDF detection with OCR guidance |
| `.pptx` | v1.2 (current) | python-pptx: per-slide dump with title headings, bullet levels, tables, pictures, speaker notes |

## Install

```bash
pip install python-docx mammoth pymupdf python-pptx
```

Uses only libraries already present in most Anaconda distributions.
No Word / Office installation required (no COM automation).

## Usage

```bash
python doc2md.py <input.docx> [-o <output-dir>]
```

- `-o` defaults to `./output` relative to the input file.
- Exit code `0` on success; non-zero with a message on failure.

### Product layout

```
output/<name>/
  article.md      # body: heading levels, paragraphs, tables, lists, image refs
  images/         # all embedded images, stable filenames
  meta.json       # counts (paragraphs/tables/images), style list, degradation notes
```

`meta.json` flags anything that could not be represented in markdown
(e.g. equations, text boxes) as a `degraded` entry, so the agent knows to
inspect the raw XML for those parts instead of silently losing content.

## Design notes

- **No Word COM.** docx is a zip of XML; converting via COM automation is
  fragile (modal dialogs hang headless automation). This tool reads the
  package directly.
- **Tables are first-class.** Cell text is read cell-by-cell via python-docx
  and emitted as GitHub-flavored pipe tables, which survive better than
  HTML-to-markdown conversion of table markup.
- **Images stay on disk.** docx images are extracted from `word/media/`;
  pdf image blocks are rendered as page-position PNGs (vertically-adjacent
  decorative fragments merged). Nothing is inlined as base64.
- **Scanned PDFs are detected, not silently mangled.** If a pdf yields
  near-zero text, `meta.json` flags it as `degraded` with a pointer to OCR
  (tesseract) instead of producing an empty markdown.

## License

MIT
