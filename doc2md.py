#!/usr/bin/env python3
"""doc2md - Convert documents (docx first, pdf/pptx planned) to agent-friendly Markdown.

Product layout per document:
    output/<name>/
        article.md   - body markdown (headings, paragraphs, tables, image refs)
        images/      - extracted embedded images
        meta.json    - counts, style list, degradation notes

Usage:
    python doc2md.py <input.docx> [-o <output-dir>]
"""

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

try:
    import mammoth
    import docx
except ImportError as e:
    print(f"missing dependency: {e.name} (pip install python-docx mammoth)", file=sys.stderr)
    sys.exit(2)


# ---------------------------------------------------------------------------
# image extraction
# ---------------------------------------------------------------------------

def extract_images(docx_path: Path, images_dir: Path) -> dict:
    """Extract word/media/* into images_dir. Returns {media_filename: disk_path}."""
    images_dir.mkdir(parents=True, exist_ok=True)
    mapping = {}
    with zipfile.ZipFile(docx_path) as zf:
        for info in zf.infolist():
            name = info.filename
            if not name.startswith("word/media/"):
                continue
            base = Path(name).name
            if not base:
                continue
            target = images_dir / base
            if not target.exists():
                target.write_bytes(zf.read(name))
            mapping[base] = str(target)
    return mapping


# ---------------------------------------------------------------------------
# tables via python-docx (precise cell-by-cell -> pipe tables)
# ---------------------------------------------------------------------------

def cell_text(cell) -> str:
    parts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    text = " ".join(parts)
    # pipe chars break GFM tables; escape them
    return text.replace("|", "\\|").replace("\n", " ")


def table_to_markdown(tbl) -> str:
    rows = []
    for row in tbl.rows:
        rows.append([cell_text(c) for c in row.cells])
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([""] * (width - len(r)))
    lines = ["| " + " | ".join(rows[0]) + " |",
             "| " + " | ".join(["---"] * width) + " |"]
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# mammoth HTML -> markdown shaping
# ---------------------------------------------------------------------------

def shape_mammoth_html(html: str) -> str:
    """Light shaping of mammoth's HTML output into markdown.

    mammoth already emits h1-h6/p/ul/ol/strong/em/a/img; we normalize
    whitespace and drop empty paragraphs. Table markup is removed here because
    tables are re-emitted separately from python-docx.
    """
    # non-greedy per-table match; placeholder must NOT look like an HTML tag,
    # otherwise the generic tag stripper later would eat it
    html = re.sub(r"<table>.*?</table>", "\x00TABLE_SLOT\x00", html, flags=re.S)
    # headings
    for level in range(1, 7):
        html = re.sub(rf"<h{level}>(.*?)</h{level}>", rf"{'#'*level} \1\n", html, flags=re.S)
    # images -> relative refs (mammoth emits <img src="data:...">; replace whole tag)
    html = re.sub(r"<img\b[^>]*>", "<!--IMAGE-->", html)
    html = re.sub(r"</?p>", "\n", html)
    html = re.sub(r"<strong>(.*?)</strong>", r"**\1**", html, flags=re.S)
    html = re.sub(r"<em>(.*?)</em>", r"*\1*", html, flags=re.S)
    html = re.sub(r"<li>", "- ", html)
    html = re.sub(r"</?[uo]l>", "\n", html)
    html = re.sub(r"<a [^>]*href=\"([^\"]*)\"[^>]*>(.*?)</a>", r"[\2](\1)", html, flags=re.S)
    # strip any remaining tags
    html = re.sub(r"<[^>]+>", "", html)
    # unescape a handful of common entities
    for ent, ch in (("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                    ("&quot;", '"'), ("&#39;", "'"), ("&nbsp;", " ")):
        html = html.replace(ent, ch)
    # collapse whitespace
    html = re.sub(r"[ \t]+\n", "\n", html)
    html = re.sub(r"\n{3,}", "\n\n", html)
    return html.strip()


# ---------------------------------------------------------------------------
# main conversion
# ---------------------------------------------------------------------------

def convert_docx(docx_path: Path, out_root: Path) -> dict:
    name = docx_path.stem
    out_dir = out_root / name
    images_dir = out_dir / "images"
    out_dir.mkdir(parents=True, exist_ok=True)

    image_map = extract_images(docx_path, images_dir)

    doc = docx.Document(str(docx_path))
    n_tables = len(doc.tables)

    table_md = [table_to_markdown(t) for t in doc.tables]

    result = mammoth.convert_to_html(str(docx_path))
    html = result.value
    mammoth_messages = [m.message for m in result.messages]

    body = shape_mammoth_html(html)

    # re-insert tables at the first slots in order of appearance
    for t in table_md:
        body = body.replace("\x00TABLE_SLOT\x00", t, 1)
    leftover_slots = body.count("\x00TABLE_SLOT\x00")
    body = body.replace("\x00TABLE_SLOT\x00", "")  # extra slots (nested/unsupported)
    body = re.sub(r"\n{3,}", "\n\n", body)

    md_path = out_dir / "article.md"
    md_path.write_text(body + "\n", encoding="utf-8")

    meta = {
        "source": str(docx_path),
        "format": "docx",
        "paragraphs": len(doc.paragraphs),
        "tables": n_tables,
        "tables_emitted": n_tables - leftover_slots,
        "images": len(image_map),
        "styles": sorted({p.style.name for p in doc.paragraphs if p.style is not None}),
        "mammoth_messages": mammoth_messages[:20],
        "degraded": [],
    }
    if leftover_slots:
        meta["degraded"].append(
            f"{leftover_slots} table slot(s) not re-emitted (nested tables or "
            "tables in unusual positions) - inspect word/document.xml directly")
    (out_dir / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return meta


def main() -> int:
    ap = argparse.ArgumentParser(description="Convert docx to agent-friendly markdown")
    ap.add_argument("input", help="path to .docx file")
    ap.add_argument("-o", "--output", help="output root dir (default: ./output next to input)")
    args = ap.parse_args()

    src = Path(args.input).expanduser().resolve()
    if not src.is_file():
        print(f"input not found: {src}", file=sys.stderr)
        return 1
    if src.suffix.lower() != ".docx":
        print(f"unsupported format for now (v1 is docx-only): {src.suffix}", file=sys.stderr)
        return 1

    out_root = Path(args.output).expanduser().resolve() if args.output else src.parent / "output"
    meta = convert_docx(src, out_root)
    out_dir = out_root / src.stem
    print(f"converted: {out_dir / 'article.md'}")
    print(f"meta: paragraphs={meta['paragraphs']} tables={meta['tables_emitted']}/{meta['tables']} "
          f"images={meta['images']} degraded={len(meta['degraded'])}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
