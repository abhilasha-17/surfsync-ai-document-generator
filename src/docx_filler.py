"""Fill DOCX placeholders and omit values that were not supplied."""

import re
from copy import deepcopy
from io import BytesIO

from docx import Document
from docx.table import _Row
from docx.text.paragraph import Paragraph


PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_. -]+)\s*\}\}")
LABEL_ONLY_RE = re.compile(r"^\s*[\w .()/&-]+\s*:\s*$")


def _normalize_key(key: str) -> str:
    """Normalize equivalent names such as 'Client Name' and 'client_name'."""

    return re.sub(r"[^a-z0-9]+", "_", str(key).strip().lower()).strip("_")


def _clean_value(value) -> str:
    """Convert a value to text and remove any placeholder tokens it contains."""

    if value is None:
        return ""

    text = str(value).strip()

    # If the model echoed a marker such as {{client_name}}, treat it as missing.
    text = PLACEHOLDER_RE.sub("", text).strip()

    return text


def _make_lookup(values: dict) -> dict[str, str]:
    """Build exact and normalized lookup keys for field values."""

    lookup = {}

    for key, value in (values or {}).items():
        cleaned_key = str(key).strip()
        cleaned_value = _clean_value(value)

        lookup[cleaned_key] = cleaned_value
        lookup[_normalize_key(cleaned_key)] = cleaned_value

    return lookup


def _get_value(lookup: dict[str, str], key: str) -> str:
    """Look up a placeholder exactly, then try its normalized spelling."""

    if key in lookup:
        return lookup[key]

    return lookup.get(_normalize_key(key), "")


def _iter_paragraphs(container):
    """Walk paragraphs in a document, section, cell, or nested table."""

    for paragraph in getattr(container, "paragraphs", []):
        yield paragraph

    for table in getattr(container, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_paragraphs(cell)


def _remove_paragraph(paragraph, remove_preceding_heading=False):
    """Remove a paragraph and optionally its immediately preceding heading."""

    element = paragraph._p
    parent = element.getparent()

    if remove_preceding_heading:
        previous = element.getprevious()

        if previous is not None and previous.tag.endswith("}p"):
            previous_paragraph = Paragraph(previous, paragraph._parent)
            style = getattr(previous_paragraph, "style", None)
            style_name = getattr(style, "name", "") if style else ""

            if style_name.lower().startswith("heading") and parent is not None:
                parent.remove(previous)

    if parent is not None:
        parent.remove(element)


def _replace_in_paragraph(paragraph, lookup: dict[str, str]):
    """Replace all markers in a paragraph, using empty text for missing values."""

    original = paragraph.text

    if not PLACEHOLDER_RE.search(original):
        return

    replaced = PLACEHOLDER_RE.sub(
        lambda match: _get_value(lookup, match.group(1).strip()),
        original,
    )

    was_only_placeholders = PLACEHOLDER_RE.sub("", original).strip() == ""

    # Remove empty optional paragraphs and their directly preceding headings.
    if not replaced.strip() and was_only_placeholders:
        _remove_paragraph(
            paragraph,
            remove_preceding_heading=True,
        )
        return

    # Remove lines such as "Version: " when the value was omitted.
    if LABEL_ONLY_RE.fullmatch(replaced):
        _remove_paragraph(paragraph)
        return

    # Keep the first run's styling and handle markers split over multiple runs.
    if paragraph.runs:
        paragraph.runs[0].text = replaced

        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(replaced)


def _remove_table(table):
    """Remove a table from its document or containing cell."""

    element = table._tbl
    parent = element.getparent()

    if parent is not None:
        parent.remove(element)


def _expand_items_table(table, items):
    """Expand a prototype row marked with {{items}}."""

    marker_row = next(
        (
            row
            for row in table.rows
            if any("{{items}}" in cell.text for cell in row.cells)
        ),
        None,
    )

    if marker_row is None:
        return False

    # Remove the whole dynamic table when there are no generated rows.
    if not items:
        _remove_table(table)
        return True

    row_element = marker_row._tr
    parent = row_element.getparent()
    insert_at = parent.index(row_element)
    template_xml = deepcopy(row_element)
    parent.remove(row_element)

    inserted_rows = 0

    for item in items:
        item_lookup = _make_lookup(item)
        item_lookup["items"] = ""
        item_lookup[_normalize_key("items")] = ""

        if not any(
            value.strip()
            for key, value in item_lookup.items()
            if key != "items"
        ):
            continue

        new_xml = deepcopy(template_xml)
        parent.insert(insert_at + inserted_rows, new_xml)
        new_row = _Row(new_xml, table)
        inserted_rows += 1

        for cell in new_row.cells:
            for paragraph in cell.paragraphs:
                _replace_in_paragraph(paragraph, item_lookup)

    if inserted_rows == 0:
        _remove_table(table)

    return True


def _expand_nested_tables(tables, items):
    """Expand dynamic tables recursively, including tables inside cells."""

    for table in list(tables):
        _expand_items_table(table, items)

        # A dynamic table with no items may have been removed.
        if table._tbl.getparent() is None:
            continue

        for row in list(table.rows):
            for cell in row.cells:
                _expand_nested_tables(cell.tables, items)


def _remove_rows_with_missing_values(tables, lookup):
    """Remove metadata rows when all their placeholder values are missing."""

    for table in list(tables):
        for row in list(table.rows):
            row_markers = []

            for cell in row.cells:
                row_markers.extend(
                    match.group(1).strip()
                    for match in PLACEHOLDER_RE.finditer(cell.text)
                    if match.group(1).strip().lower() != "items"
                )

            if row_markers and all(
                not _get_value(lookup, marker)
                for marker in row_markers
            ):
                row_element = row._tr
                parent = row_element.getparent()

                if parent is not None:
                    parent.remove(row_element)

        for row in list(table.rows):
            for cell in row.cells:
                _remove_rows_with_missing_values(cell.tables, lookup)


def _all_document_paragraphs(doc):
    """Yield body, header, footer, and nested-table paragraphs."""

    yield from _iter_paragraphs(doc)

    for section in doc.sections:
        yield from _iter_paragraphs(section.header)
        yield from _iter_paragraphs(section.footer)
        yield from _iter_paragraphs(section.first_page_header)
        yield from _iter_paragraphs(section.first_page_footer)
        yield from _iter_paragraphs(section.even_page_header)
        yield from _iter_paragraphs(section.even_page_footer)


def fill_docx(
    template_bytes: bytes,
    fields: dict[str, str],
    items: list[dict[str, str]] | None = None,
) -> bytes:
    """Fill a copy of a DOCX template and remove unresolved marker text."""

    doc = Document(BytesIO(template_bytes))
    item_rows = items or []
    field_lookup = _make_lookup(fields or {})
    field_lookup["items"] = ""
    field_lookup[_normalize_key("items")] = ""

    # Expand dynamic prototype rows before filling scalar placeholders.
    _expand_nested_tables(doc.tables, item_rows)

    # Remove empty metadata rows such as Client or Version when unspecified.
    _remove_rows_with_missing_values(doc.tables, field_lookup)

    # Fill body content and page furniture.
    seen_paragraphs = set()

    for paragraph in list(_all_document_paragraphs(doc)):
        paragraph_id = id(paragraph._p)

        if paragraph_id in seen_paragraphs:
            continue

        seen_paragraphs.add(paragraph_id)
        _replace_in_paragraph(paragraph, field_lookup)

    # Final cleanup pass: no raw {{...}} markers should reach the download.
    cleanup_lookup = {"": ""}

    for paragraph in list(_all_document_paragraphs(doc)):
        if PLACEHOLDER_RE.search(paragraph.text):
            _replace_in_paragraph(paragraph, cleanup_lookup)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()