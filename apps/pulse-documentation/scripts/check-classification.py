#!/usr/bin/env python3
"""Verify every controlled PDF carries the classification marking on every page.

The externally authored documents (01-08) arrived already marked. The generated
documents (09-12) carry the marking natively from build-source-document.mjs. This
script is the check that neither situation regresses — a generated document built
before the marking was added, or a replacement external PDF that lacks it, both
show up here.

Pass --stamp to overlay the marking on any unmarked file. That is a repair path,
not part of the normal build: a document that needs stamping should ideally be
reissued with the marking built in.

Usage:  python3 scripts/check-classification.py [--stamp]
Exit 1 if any page is unmarked.
"""
import io
import os
import re
import sys
import tempfile
from pathlib import Path

import fitz  # PyMuPDF, for reading text per page
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas

CLASSIFICATION = "UNCLASSIFIED"
INK = Color(7 / 255, 7 / 255, 8 / 255)
GENERATED = re.compile(r"^(09|10|11|12)_")
SOURCE = Path(__file__).resolve().parent.parent / "public" / "source-pdfs"


def unmarked_pages(path):
    """Page numbers (1-based) with no classification marking."""
    document = fitz.open(str(path))
    missing = []
    for index, page in enumerate(document, start=1):
        if CLASSIFICATION not in page.get_text().replace(" ", "").upper():
            missing.append(index)
    return missing, document.page_count


def overlay(width, height):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    c.setFont("Helvetica-Bold", 8.5)
    c.setFillColor(INK)
    spaced = " ".join(CLASSIFICATION)
    for y in (height - 22, 14):
        c.drawCentredString(width / 2, y, spaced)
    c.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]


def stamp(path):
    reader = PdfReader(str(path))
    writer = PdfWriter()
    for page in reader.pages:
        box = page.mediabox
        page.merge_page(overlay(float(box.width), float(box.height)))
        writer.add_page(page)
    handle, tmp = tempfile.mkstemp(suffix=".pdf")
    try:
        with os.fdopen(handle, "wb") as stream:
            writer.write(stream)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise


def main():
    do_stamp = "--stamp" in sys.argv
    failures = 0
    for path in sorted(SOURCE.glob("*.pdf")):
        missing, total = unmarked_pages(path)
        origin = "generated" if GENERATED.match(path.name) else "external"
        if not missing:
            print(f"  {path.name[:52]:54} {total:2}/{total:2}  {origin}")
            continue
        if do_stamp:
            stamp(path)
            missing, total = unmarked_pages(path)
            state = "stamped" if not missing else f"STILL UNMARKED {missing}"
            print(f"  {path.name[:52]:54} {total - len(missing):2}/{total:2}  {origin}  {state}")
            failures += 1 if missing else 0
        else:
            print(f"  {path.name[:52]:54} {total - len(missing):2}/{total:2}  {origin}  UNMARKED pages {missing}")
            failures += 1
    if failures:
        print(f"\n{failures} document(s) missing the {CLASSIFICATION} marking."
              f"{'' if do_stamp else '  Re-run with --stamp to overlay it.'}")
        return 1
    print(f"\nAll documents carry {CLASSIFICATION} on every page.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
