#!/usr/bin/env python3
"""doc2md - Convert documents (docx, pdf, pptx) to agent-friendly Markdown.

Product layout per document:
    output/<name>/
        article.md   - body markdown (headings, paragraphs, tables, image refs)
        images/      - extracted embedded images
        meta.json    - counts, style list, degradation notes

Usage:
    python doc2md.py <input.docx> [-o <output-dir>]
    python doc2md.py <input.pdf>  [-o <output-dir>]
    python doc2md.py <input.pptx> [-o <output-dir>]
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


# ---------------------------------------------------------------------------
# pdf conversion via PyMuPDF (fitz)
# ---------------------------------------------------------------------------

def pdf_escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def pdf_detect_heading_sizes(page_count: int, sizes: dict) -> dict:
    """Map body font size -> heading level. Returns {size: level} for the
    largest distinct sizes (level 1 for biggest, max 3 levels)."""
    if not sizes:
        return {}
    ranked = sorted(sizes, key=lambda s: -s)
    # skip sizes that look like body text duplicates (within 0.5pt)
    distinct = []
    for s in ranked:
        if not distinct or (distinct[-1] - s) > 0.5:
            distinct.append(s)
    levels = {}
    for i, s in enumerate(distinct[:3]):
        levels[s] = i + 1
    return levels


def get_fitz():
    """Lazy import of pymupdf. Returns the module or None with an error printed."""
    try:
        import fitz  # pymupdf
        return fitz
    except ImportError as e:
        print(f"missing dependency: {e.name} (pip install pymupdf)", file=sys.stderr)
        return None


def pdf_extract_images(fitz_doc, images_dir: Path) -> int:
    """Extract embedded raster images. Returns count. Skips tiny icons (<10KB)."""
    fitz = get_fitz()
    if fitz is None:
        return 0
    images_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    seen_xrefs = set()
    for pno in range(len(fitz_doc)):
        for info in fitz_doc[pno].get_images(full=True):
            xref = info[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            try:
                pix = fitz.Pixmap(fitz_doc, xref)
                if pix.width < 50 or pix.height < 50 or len(pix.samples) < 10_000:
                    continue
                if pix.colorspace and pix.colorspace.n > 3:  # CMYK etc -> RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                ext = "png" if pix.alpha or (pix.colorspace and pix.colorspace.n != 3) else "jpeg"
                target = images_dir / f"p{pno + 1}_img{count + 1}.{ext}"
                target.write_bytes(pix.tobytes(ext if ext == "png" else "jpg"))
                count += 1
            except Exception as e:  # unreadable image: note and move on
                print(f"warn: image xref {xref} on page {pno + 1}: {e}", file=sys.stderr)
    return count


def convert_pdf(pdf_path: Path, out_root: Path) -> dict:
    name = pdf_path.stem
    out_dir = out_root / name
    images_dir = out_dir / "images"
    out_dir.mkdir(parents=True, exist_ok=True)

    fitz = get_fitz()
    if fitz is None:
        raise SystemExit(2)
    doc = fitz.open(str(pdf_path))
    n_pages = len(doc)

    # first pass: font size histogram for heading detection
    size_hist = {}
    for page in doc:
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    if span["text"].strip():
                        size = round(span["size"], 1)
                        size_hist[size] = size_hist.get(size, 0) + 1
    # body size = most frequent; headings = sizes clearly above body
    body_size = max(size_hist, key=lambda s: size_hist[s]) if size_hist else 0
    heading_sizes = {s: lvl for s, lvl in pdf_detect_heading_sizes(n_pages, size_hist).items()
                     if s > body_size + 1.0}

    n_images = pdf_extract_images(doc, images_dir)

    # second pass: emit markdown per page
    parts = []
    total_chars = 0
    img_seq = 0
    for pno, page in enumerate(doc, 1):
        if pno > 1:
            parts.append("\n---\n")  # page separator
        page_md = []
        last_img_rect = None  # merge vertically-adjacent image blocks (decorative splits)
        for block in page.get_text("dict")["blocks"]:
            if block["type"] == 1:  # image block: save as positional PNG
                rect = fitz.Rect(block["bbox"])
                if rect.width < 20 or rect.height < 20:
                    continue  # tiny decoration, skip entirely
                if (last_img_rect is not None and rect.y0 - last_img_rect.y1 < 5
                        and abs(rect.x0 - last_img_rect.x0) < 5):
                    # continuation of previous image strip: don't re-reference
                    last_img_rect = rect
                    continue
                last_img_rect = rect
                img_seq += 1
                target = images_dir / f"p{pno}_img{img_seq}.png"
                target.parent.mkdir(parents=True, exist_ok=True)
                pix = page.get_pixmap(clip=rect, dpi=150)
                target.write_bytes(pix.tobytes("png"))
                page_md.append(f"![](images/{target.name})")
                continue
            text = ""
            sizes_seen = []
            for line in block["lines"]:
                line_text = "".join(span["text"] for span in line["spans"]).strip()
                if line_text:
                    text += line_text + " "
                    sizes_seen.extend(round(sp["size"], 1) for sp in line["spans"]
                                      if sp["text"].strip())
            text = text.strip()
            if not text:
                continue
            total_chars += len(text)
            top = max(sizes_seen) if sizes_seen else body_size
            if top in heading_sizes:
                page_md.append("#" * heading_sizes[top] + " " + text)
            else:
                page_md.append(text)
        if page_md:
            parts.append("\n\n".join(page_md))
        # vector/table regions can't be split reliably; PyMuPDF text blocks
        # already flatten table rows into text lines, which is acceptable

    doc.close()

    body = "\n\n".join(parts)
    body = re.sub(r"\n{3,}", "\n\n", body)
    (out_dir / "article.md").write_text(body + "\n", encoding="utf-8")

    meta = {
        "source": str(pdf_path),
        "format": "pdf",
        "pages": n_pages,
        "images": n_images,
        "image_blocks_rendered": img_seq,
        "body_font_size": body_size,
        "heading_sizes": heading_sizes,
        "text_chars": total_chars,
        "degraded": [],
    }
    if total_chars < 100 and n_pages > 0:
        meta["degraded"].append(
            "near-zero text extracted - likely a scanned PDF; run OCR "
            "(tesseract -l chi_sim) on page images instead")
    (out_dir / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return meta


# ---------------------------------------------------------------------------
# pptx conversion via python-pptx
# ---------------------------------------------------------------------------

def pptx_shape_text(shape) -> str:
    """Extract paragraph text from a shape, preserving list levels as bullets."""
    if not shape.has_text_frame:
        return ""
    lines = []
    for para in shape.text_frame.paragraphs:
        text = "".join(run.text for run in para.runs).strip()
        if not text:
            continue
        indent = "  " * min(para.level, 3) if para.level else ""
        lines.append(f"{indent}- {text}")
    return "\n".join(lines)


def convert_pptx(pptx_path: Path, out_root: Path) -> dict:
    name = pptx_path.stem
    out_dir = out_root / name
    images_dir = out_dir / "images"
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        from pptx import Presentation
    except ImportError as e:
        print(f"missing dependency: {e.name} (pip install python-pptx)", file=sys.stderr)
        raise SystemExit(2)

    prs = Presentation(str(pptx_path))
    parts = []
    n_tables = 0
    n_images = 0
    n_notes = 0
    degraded = []

    for idx, slide in enumerate(prs.slides, 1):
        slide_md = [f"## Slide {idx}"]
        # title placeholder first, as the slide's real heading
        title = slide.shapes.title
        if title is not None and title.text.strip():
            slide_md.append(f"# {title.text.strip()}")
        for shape in slide.shapes:
            if shape == title:
                continue
            # pictures
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                try:
                    image = shape.image
                    n_images += 1
                    target = images_dir / f"slide{idx}_img{n_images}.{image.ext}"
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(image.blob)
                    slide_md.append(f"![](images/{target.name})")
                except Exception as e:  # grouped/linked shapes without blob
                    degraded.append(f"slide {idx}: unrecoverable picture ({e})")
                continue
            # tables
            if shape.has_table:
                n_tables += 1
                tbl = shape.table
                rows = []
                for row in tbl.rows:
                    rows.append([pdf_escape(c.text.strip()) for c in row.cells])
                if rows:
                    width = max(len(r) for r in rows)
                    for r in rows:
                        r.extend([""] * (width - len(r)))
                    md = ["| " + " | ".join(rows[0]) + " |",
                          "| " + " | ".join(["---"] * width) + " |"]
                    for r in rows[1:]:
                        md.append("| " + " | ".join(r) + " |")
                    slide_md.append("\n".join(md))
                continue
            # grouped shapes: flatten one level
            if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
                for sub in shape.shapes:
                    t = pptx_shape_text(sub)
                    if t:
                        slide_md.append(t)
                continue
            text = pptx_shape_text(shape)
            if text:
                slide_md.append(text)
        # speaker notes
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                n_notes += 1
                slide_md.append(f"> notes: {notes}")
        parts.append("\n\n".join(slide_md))

    body = "\n\n".join(parts)
    body = re.sub(r"\n{3,}", "\n\n", body)
    (out_dir / "article.md").write_text(body + "\n", encoding="utf-8")

    meta = {
        "source": str(pptx_path),
        "format": "pptx",
        "slides": len(prs.slides),
        "tables": n_tables,
        "images": n_images,
        "slides_with_notes": n_notes,
        "degraded": degraded[:20],
    }
    (out_dir / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return meta


def main() -> int:
    ap = argparse.ArgumentParser(description="Convert documents (docx/pdf/pptx) to agent-friendly markdown")
    ap.add_argument("input", help="path to .docx, .pdf or .pptx file")
    ap.add_argument("-o", "--output", help="output root dir (default: ./output next to input)")
    args = ap.parse_args()

    src = Path(args.input).expanduser().resolve()
    if not src.is_file():
        print(f"input not found: {src}", file=sys.stderr)
        return 1

    fmt = src.suffix.lower()
    if fmt not in (".docx", ".pdf", ".pptx"):
        print(f"unsupported format (docx, pdf, pptx supported): {fmt}", file=sys.stderr)
        return 1

    out_root = Path(args.output).expanduser().resolve() if args.output else src.parent / "output"
    if fmt == ".pdf":
        meta = convert_pdf(src, out_root)
        print(f"meta: pages={meta['pages']} images={meta['images']} "
              f"chars={meta['text_chars']} degraded={len(meta['degraded'])}")
    elif fmt == ".pptx":
        meta = convert_pptx(src, out_root)
        print(f"meta: slides={meta['slides']} tables={meta['tables']} "
              f"images={meta['images']} degraded={len(meta['degraded'])}")
    else:
        meta = convert_docx(src, out_root)
        print(f"meta: paragraphs={meta['paragraphs']} tables={meta['tables_emitted']}/{meta['tables']} "
              f"images={meta['images']} degraded={len(meta['degraded'])}")
    out_dir = out_root / src.stem
    print(f"converted: {out_dir / 'article.md'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
