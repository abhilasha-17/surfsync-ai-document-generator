# SurfSync Infotech AI Document Generator

A Streamlit application that uses an OpenAI-compatible language model to draft client documents from a project brief and a branded Word template. It supports BRD, FRD, SRS, and quotation templates, and provides generated files as Word documents and, when LibreOffice is available, PDFs.

The app keeps the uploaded template as the basis for the output. It replaces marked placeholders and can expand a table row marked with `{{items}}`.

## Features

- Upload a `.docx` template.
- Select BRD, FRD, SRS, or Quotation as the document type.
- Describe the client, goals, scope, constraints, timeline, deliverables, and any confirmed pricing.
- Inspect detected headings, tables, and body placeholders before generating.
- Generate structured content through an OpenAI-compatible chat completions API.
- Validate the response shape with Pydantic.
- Replace placeholders in document paragraphs, tables, headers, and footers.
- Expand a formatted table prototype row marked with `{{items}}`.
- Download the generated Word document.
- Download a PDF when LibreOffice is installed on the machine running the app.
- Keep provider settings and the API key out of the app interface.

## Project structure

```text
.
├── app.py                         Streamlit interface and generation workflow
├── src/
│   ├── __init__.py
│   ├── template_reader.py         Reads headings, tables, and body placeholders
│   ├── generator.py               Calls the model API and validates JSON with Pydantic
│   ├── docx_filler.py              Replaces placeholders and expands marked table rows
│   └── pdf_export.py               Converts a generated DOCX to PDF using LibreOffice
├── assets/
│   └── surfsync_logo.png           SurfSync Infotech logo
├── samples/
│   ├── brd_template.docx
│   ├── brd_example_output.docx
│   ├── frd_template.docx
│   ├── frd_example_output.docx
│   ├── srs_template.docx
│   ├── srs_example_output.docx
│   ├── quotation_template.docx
│   └── quotation_example_output.docx
├── scripts/
│   └── create_samples.py           Rebuilds the sample templates and output documents
├── .streamlit/
│   └── config.toml                 Streamlit theme and upload-size configuration
├── .env.example                    Environment variable names; contains no API key
├── .gitignore
├── requirements.txt
└── README.md


## Links

- **Live application:** https://surfsync-document-generator.streamlit.app/
- **Demo video:** [Add a public or view-accessible video link here]
- **Source code:** https://github.com/abhilasha-17/surfsync-ai-document-generator
