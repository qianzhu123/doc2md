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
| `.pdf` | planned | text layer extraction; OCR fallback for scans (tesseract) |
| `.pptx` | planned | per-slide text dump + media extraction |

## Install

```bash
pip install python-docx mammoth
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
- **Images stay on disk.** They are extracted from `word/media/` and
  referenced with relative paths; nothing is inlined as base64.

## License

MIT
