"""Read structural cues from a Word document without changing the source."""
import re
from docx import Document

PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_. -]+)\s*\}\}")


def _walk_tables(tables):
    for table in tables:
        yield table
        for row in table.rows:
            for cell in row.cells:
                yield from _walk_tables(cell.tables)


def inspect_docx(file_obj):
    doc = Document(file_obj)
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    headings = [p.text.strip() for p in paragraphs if p.style and p.style.name.lower().startswith("heading")]
    tables = []
    for idx, table in enumerate(_walk_tables(doc.tables), start=1):
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        tables.append({"index": idx, "rows": rows, "columns": rows[0] if rows else []})

    all_text = [p.text for p in paragraphs]
    for table in _walk_tables(doc.tables):
        all_text.extend(cell.text for row in table.rows for cell in row.cells)
    placeholders = sorted({m.group(1).strip() for text in all_text for m in PLACEHOLDER_RE.finditer(text)})
    return {"headings": headings, "paragraphs": [p.text.strip() for p in paragraphs], "tables": tables,
            "placeholders": placeholders}

