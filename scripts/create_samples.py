"""Create SurfSync letterhead templates and filled example outputs."""

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.docx_filler import fill_docx  # noqa: E402


SAMPLES = ROOT / "samples"
LOGO_PATH = ROOT / "assets" / "surfsync_logo.png"
SAMPLES.mkdir(exist_ok=True)


def base_doc(title: str) -> Document:
    """Create a branded letterhead document with a title and footer."""

    doc = Document()
    section = doc.sections[0]

    section.top_margin = Inches(1.3)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    section.header_distance = Inches(0.12)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT

    if LOGO_PATH.is_file():
        header.add_run().add_picture(str(LOGO_PATH), width=Inches(2.0))

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.text = "SurfSync Infotech  |  Confidential client document"

    for run in footer.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(20, 69, 112)

    title_paragraph = doc.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_paragraph.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(18)
    title_run.font.color.rgb = RGBColor(20, 69, 112)

    return doc


def add_heading(doc: Document, text: str):
    """Add a styled section heading."""

    paragraph = doc.add_paragraph(style="Heading 2")
    paragraph.add_run(text)
    return paragraph


def style_table(table):
    """Apply a consistent style to a table and its header row."""

    table.style = "Light Shading Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for cell in table.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(20, 69, 112)


def add_dynamic_table(doc: Document, headers: list[str], row_placeholders: list[str]):
    """Add a header row and one prototype row for dynamic expansion."""

    table = doc.add_table(rows=2, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = header

    for cell, placeholder in zip(table.rows[1].cells, row_placeholders):
        cell.text = placeholder

    style_table(table)
    return table


def make_quotation() -> Document:
    doc = base_doc("COMMERCIAL QUOTATION")

    doc.add_paragraph(
        "Quotation No: SSI-{{quotation_number}}     Date: {{date}}"
    )
    doc.add_paragraph("Prepared for: {{client_name}}")

    add_heading(doc, "1. Project Overview")
    doc.add_paragraph("{{project_overview}}")

    add_heading(doc, "2. Scope and Pricing")
    add_dynamic_table(
        doc,
        ["Description", "Quantity", "Unit Price", "Total"],
        [
            "{{items}}{{description}}",
            "{{quantity}}",
            "{{unit_price}}",
            "{{total}}",
        ],
    )

    doc.add_paragraph("Grand total: {{grand_total}}")

    add_heading(doc, "3. Assumptions and Exclusions")
    doc.add_paragraph("{{assumptions}}")

    add_heading(doc, "4. Terms and Conditions")
    doc.add_paragraph(
        "This quotation is valid for 30 days. Work begins after written approval. "
        "Payment milestones will be agreed before project commencement."
    )

    return doc


def make_brd() -> Document:
    doc = base_doc("BUSINESS REQUIREMENTS DOCUMENT")

    doc.add_paragraph("Client: {{client_name}}     Version: {{version}}")

    add_heading(doc, "1. Executive Summary")
    doc.add_paragraph("{{executive_summary}}")

    add_heading(doc, "2. Business Objectives")
    doc.add_paragraph("{{business_objectives}}")

    add_heading(doc, "3. Scope of Work")
    doc.add_paragraph("{{scope_of_work}}")

    add_heading(doc, "4. Business Requirements")
    add_dynamic_table(
        doc,
        ["ID", "Requirement", "Priority"],
        ["{{items}}{{id}}", "{{requirement}}", "{{priority}}"],
    )

    add_heading(doc, "5. Assumptions and Dependencies")
    doc.add_paragraph("{{assumptions}}")

    add_heading(doc, "6. Approval")
    doc.add_paragraph(
        "Client representative: ____________________    Date: __________"
    )

    return doc


def make_frd() -> Document:
    doc = base_doc("FUNCTIONAL REQUIREMENTS DOCUMENT")

    doc.add_paragraph("Client: {{client_name}}     Version: {{version}}")

    add_heading(doc, "1. Purpose")
    doc.add_paragraph("{{purpose}}")

    add_heading(doc, "2. System Overview")
    doc.add_paragraph("{{system_overview}}")

    add_heading(doc, "3. User Roles")
    doc.add_paragraph("{{user_roles}}")

    add_heading(doc, "4. Functional Requirements")
    add_dynamic_table(
        doc,
        ["ID", "Feature", "Functional Description", "Acceptance Criteria", "Priority"],
        [
            "{{items}}{{id}}",
            "{{feature}}",
            "{{description}}",
            "{{acceptance_criteria}}",
            "{{priority}}",
        ],
    )

    add_heading(doc, "5. Assumptions and Dependencies")
    doc.add_paragraph("{{assumptions}}")

    add_heading(doc, "6. Approval")
    doc.add_paragraph(
        "Client representative: ____________________    Date: __________"
    )

    return doc


def make_srs() -> Document:
    doc = base_doc("SOFTWARE REQUIREMENTS SPECIFICATION")

    doc.add_paragraph("Client: {{client_name}}     Version: {{version}}")

    add_heading(doc, "1. Purpose and Scope")
    doc.add_paragraph("{{purpose_and_scope}}")

    add_heading(doc, "2. Product Overview")
    doc.add_paragraph("{{product_overview}}")

    add_heading(doc, "3. User Classes")
    doc.add_paragraph("{{user_classes}}")

    add_heading(doc, "4. Requirements Catalogue")
    add_dynamic_table(
        doc,
        ["Type", "ID", "Requirement", "Priority", "Acceptance Criteria / Metric"],
        [
            "{{items}}{{type}}",
            "{{id}}",
            "{{requirement}}",
            "{{priority}}",
            "{{acceptance_criteria}}",
        ],
    )

    add_heading(doc, "5. Interfaces and Constraints")
    doc.add_paragraph("{{interfaces_and_constraints}}")

    add_heading(doc, "6. Assumptions")
    doc.add_paragraph("{{assumptions}}")

    add_heading(doc, "7. Approval")
    doc.add_paragraph(
        "Client representative: ____________________    Date: __________"
    )

    return doc


def write_example(
    template_doc: Document,
    name: str,
    fields: dict[str, str],
    items: list[dict[str, str]],
):
    """Write a template and its filled example output."""

    template_path = SAMPLES / f"{name}_template.docx"
    output_path = SAMPLES / f"{name}_example_output.docx"

    template_doc.save(template_path)
    output_path.write_bytes(
        fill_docx(template_path.read_bytes(), fields, items)
    )


def build_quotation():
    write_example(
        make_quotation(),
        "quotation",
        {
            "quotation_number": "DEMO-001",
            "date": "23 November 2026",
            "client_name": "Example Retail Co.",
            "project_overview": (
                "A responsive product catalogue and enquiry portal "
                "for the client's retail team."
            ),
            "grand_total": "INR 1,50,000 (illustrative)",
            "assumptions": (
                "Prices are sample values for demonstration only. Confirm "
                "scope, taxes, and rates before use."
            ),
        },
        [
            {
                "description": "Discovery and requirements",
                "quantity": "1",
                "unit_price": "INR 25,000",
                "total": "INR 25,000",
            },
            {
                "description": "Portal design and development",
                "quantity": "1",
                "unit_price": "INR 95,000",
                "total": "INR 95,000",
            },
            {
                "description": "Testing and handover",
                "quantity": "1",
                "unit_price": "INR 30,000",
                "total": "INR 30,000",
            },
        ],
    )


def build_brd():
    write_example(
        make_brd(),
        "brd",
        {
            "client_name": "Example Retail Co.",
            "version": "1.0",
            "executive_summary": (
                "The project will provide a central online catalogue and "
                "enquiry workflow for retail customers."
            ),
            "business_objectives": (
                "Improve product discovery, reduce manual enquiry handling, "
                "and give staff a consistent view of incoming requests."
            ),
            "scope_of_work": (
                "Build a responsive catalogue, product detail pages, an "
                "enquiry form, and a staff view for submitted enquiries."
            ),
            "assumptions": (
                "The client supplies product descriptions and images. Payment "
                "processing and inventory integrations are outside this sample."
            ),
        },
        [
            {
                "id": "BR-01",
                "requirement": "Visitors can browse and search published products.",
                "priority": "High",
            },
            {
                "id": "BR-02",
                "requirement": "Visitors can submit enquiries about products.",
                "priority": "High",
            },
            {
                "id": "BR-03",
                "requirement": "Staff can review and update enquiry status.",
                "priority": "Medium",
            },
        ],
    )


def build_frd():
    write_example(
        make_frd(),
        "frd",
        {
            "client_name": "Example Retail Co.",
            "version": "1.0",
            "purpose": (
                "This document describes the expected user-facing and staff "
                "functions for the product catalogue and enquiry portal."
            ),
            "system_overview": (
                "The system provides a public product catalogue, product "
                "search, enquiry submission, and a staff enquiry view."
            ),
            "user_roles": (
                "Visitor: browses products and submits enquiries. "
                "Staff member: reviews enquiries and updates their status."
            ),
            "assumptions": (
                "The client supplies product information and images. "
                "Inventory and payment integrations are excluded from this sample."
            ),
        },
        [
            {
                "id": "FR-01",
                "feature": "Product browsing",
                "description": (
                    "The system displays published products with their "
                    "name, description, and image."
                ),
                "acceptance_criteria": (
                    "A visitor can open the catalogue and view each published "
                    "product's available details."
                ),
                "priority": "High",
            },
            {
                "id": "FR-02",
                "feature": "Product search",
                "description": (
                    "The system lets visitors search products by name."
                ),
                "acceptance_criteria": (
                    "Submitting a search shows matching products or a clear "
                    "no-results message."
                ),
                "priority": "High",
            },
            {
                "id": "FR-03",
                "feature": "Enquiry submission",
                "description": (
                    "The system accepts a visitor's contact details and "
                    "selected product."
                ),
                "acceptance_criteria": (
                    "The system validates required fields and displays a "
                    "submission confirmation."
                ),
                "priority": "High",
            },
        ],
    )


def build_srs():
    write_example(
        make_srs(),
        "srs",
        {
            "client_name": "Example Retail Co.",
            "version": "1.0",
            "purpose_and_scope": (
                "This SRS describes a responsive product catalogue and "
                "enquiry portal. It covers product discovery, enquiry "
                "submission, and staff review."
            ),
            "product_overview": (
                "The product is a web application for retail customers and "
                "staff. It includes a public catalogue and a staff-facing "
                "enquiry view."
            ),
            "user_classes": (
                "Visitor: browses products and submits enquiries. "
                "Staff member: reviews and updates enquiries."
            ),
            "interfaces_and_constraints": (
                "The application is accessed through a web browser. Product "
                "content and images are supplied by the client. No external "
                "inventory or payment integration is assumed in this sample."
            ),
            "assumptions": (
                "Hosting, browser support targets, security controls, and "
                "availability commitments must be confirmed with the client."
            ),
        },
        [
            {
                "type": "Functional",
                "id": "SRS-F-01",
                "requirement": (
                    "The system shall display published products with their "
                    "available names, descriptions, and images."
                ),
                "priority": "High",
                "acceptance_criteria": (
                    "A visitor can open a product and see the published "
                    "information provided by the client."
                ),
            },
            {
                "type": "Functional",
                "id": "SRS-F-02",
                "requirement": (
                    "The system shall allow a visitor to submit an enquiry "
                    "associated with a selected product."
                ),
                "priority": "High",
                "acceptance_criteria": (
                    "Required fields are validated and a successful submission "
                    "produces a confirmation."
                ),
            },
            {
                "type": "Non-functional",
                "id": "SRS-NF-01",
                "requirement": (
                    "The interface shall adapt to desktop and mobile screen sizes."
                ),
                "priority": "Medium",
                "acceptance_criteria": (
                    "Core catalogue and enquiry actions remain usable at the "
                    "screen sizes agreed during acceptance."
                ),
            },
        ],
    )


if __name__ == "__main__":
    build_quotation()
    build_brd()
    build_frd()
    build_srs()

    print("Created four branded templates and their example outputs in samples/.")