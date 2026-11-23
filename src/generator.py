"""Generate and strictly validate content for DOCX templates."""

import json
import re
from typing import Any

import requests
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    StrictStr,
    ValidationError,
    create_model,
)


PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_. -]+)\s*\}\}")


class GeneratedContent(BaseModel):
    """Validated values ready for the DOCX filler."""

    fields: dict[str, str] = Field(default_factory=dict)
    items: list[dict[str, str]] = Field(default_factory=list)


DOCUMENT_GUIDANCE = {
    "brd": """
Create a Business Requirements Document (BRD). Focus on business context,
objectives, scope, stakeholders, business requirements, priorities, assumptions,
dependencies, and approvals when those sections exist in the template.
""",
    "frd": """
Create a Functional Requirements Document (FRD). Describe system behavior in
clear, testable terms. Include user roles, features, inputs, processing, outputs,
validation, error handling, acceptance criteria, and priorities where the template
asks for them.
""",
    "srs": """
Create a Software Requirements Specification (SRS). Cover purpose, scope,
product context, user classes, functional and non-functional requirements,
interfaces, constraints, assumptions, and measurable acceptance criteria where
the template asks for them.
""",
    "quotation": """
Create a commercial quotation. Describe scope, deliverables, quantities, rates,
totals, validity, assumptions, and exclusions where the template asks for them.
Never invent a price, tax, date, payment term, or legal commitment.
""",
}


def _get_dynamic_row_fields(tables: list[dict[str, Any]]) -> list[dict[str, str]]:
    """Return exact placeholder keys from prototype rows marked with {{items}}."""

    result = []
    seen = set()

    for table in tables:
        rows = table.get("rows", [])
        columns = table.get("columns", [])

        for row in rows:
            if not any("{{items}}" in cell for cell in row):
                continue

            for cell_index, cell_text in enumerate(row):
                header = (
                    columns[cell_index]
                    if cell_index < len(columns)
                    else f"Column {cell_index + 1}"
                )

                for match in PLACEHOLDER_RE.finditer(cell_text):
                    key = match.group(1).strip()

                    if key.lower() == "items" or key in seen:
                        continue

                    seen.add(key)
                    result.append({
                        "placeholder": key,
                        "column": header,
                    })

    return result


def _make_object_model(
    model_name: str,
    keys: list[str],
) -> type[BaseModel]:
    """Build a strict Pydantic object with aliases matching template placeholders."""

    model_fields = {
        f"value_{index}": (
            StrictStr,
            Field(..., alias=placeholder),
        )
        for index, placeholder in enumerate(keys)
    }

    return create_model(
        model_name,
        __config__=ConfigDict(extra="forbid"),
        **model_fields,
    )


def _make_response_model(
    scalar_fields: list[str],
    item_fields: list[str],
) -> type[BaseModel]:
    """Build a strict Pydantic model matching one uploaded template."""

    FieldsModel = _make_object_model(
        "TemplateScalarFields",
        scalar_fields,
    )

    ItemModel = _make_object_model(
        "TemplateTableItem",
        item_fields,
    )

    return create_model(
        "TemplateGenerationResponse",
        __config__=ConfigDict(extra="forbid"),
        fields=(FieldsModel, ...),
        items=(list[ItemModel], ...),
    )


def _parse_and_validate(
    raw_content: str,
    response_model: type[BaseModel],
    scalar_fields: list[str],
    item_fields: list[str],
) -> GeneratedContent:
    """Validate model JSON and return ordinary dictionaries for the DOCX filler."""

    parsed = response_model.model_validate_json(raw_content)
    dumped = parsed.model_dump(by_alias=True)

    fields = dumped["fields"]
    items = dumped["items"]

    if not item_fields and items:
        raise ValueError(
            "The template has no dynamic table row, but the model returned items."
        )

    for item_index, item in enumerate(items, start=1):
        missing = [key for key in item_fields if key not in item]
        if missing:
            raise ValueError(
                f"Table item {item_index} is missing fields: {', '.join(missing)}"
            )

    if set(fields) != set(scalar_fields):
        raise ValueError(
            "The generated scalar field names do not match the template placeholders."
        )

    return GeneratedContent(
        fields=fields,
        items=items,
    )


def generate_content(
    *,
    base_url: str,
    model: str,
    api_key: str,
    document_type: str,
    project_prompt: str,
    template_info: dict,
) -> GeneratedContent:
    """Generate content and enforce the uploaded template's exact field structure."""

    if not api_key.strip():
        raise ValueError(
            "The server has no LLM_API_KEY configured. Ask the app administrator."
        )

    placeholders = [
        str(value).strip()
        for value in template_info.get("placeholders", [])
        if str(value).strip()
    ]
    tables = template_info.get("tables", [])
    headings = template_info.get("headings", [])
    paragraphs = template_info.get("paragraphs", [])

    dynamic_fields = _get_dynamic_row_fields(tables)
    dynamic_keys = {item["placeholder"] for item in dynamic_fields}

    scalar_fields = [
        name
        for name in placeholders
        if name.lower() != "items" and name not in dynamic_keys
    ]
    item_fields = [item["placeholder"] for item in dynamic_fields]

    response_model = _make_response_model(
        scalar_fields,
        item_fields,
    )

    schema_hint = {
        "fields": {name: "string" for name in scalar_fields},
        "items": [
            {name: "string" for name in item_fields}
        ] if item_fields else [],
    }

    guidance = DOCUMENT_GUIDANCE.get(
        document_type.strip().lower(),
        "Create professional content matching the uploaded template.",
    )

    prompt = f"""
Prepare a draft {document_type} using only the facts in the project brief.

DOCUMENT-TYPE GUIDANCE:
{guidance}

PROJECT BRIEF:
{project_prompt}

TEMPLATE HEADINGS:
{json.dumps(headings, ensure_ascii=False)}

TEMPLATE PARAGRAPHS:
{json.dumps(paragraphs, ensure_ascii=False)}

TEMPLATE TABLES:
{json.dumps(tables, ensure_ascii=False)}

EXACT SCALAR PLACEHOLDERS:
{json.dumps(scalar_fields, ensure_ascii=False)}

DYNAMIC TABLE PLACEHOLDERS:
{json.dumps(dynamic_fields, ensure_ascii=False)}

RETURN ONLY THIS JSON STRUCTURE:
{json.dumps(schema_hint, ensure_ascii=False)}

STRICT CONTENT RULES:
- Use every exact scalar placeholder name as a key in "fields".
- Do not add keys that are not listed in the template placeholders.
- Use the exact dynamic row placeholder names in every object in "items".
- Do not replace a placeholder key with its table heading. For example, if the
  placeholder is "description" and its column heading is "Functional Description",
  the JSON key must still be "description".
- Every value must be a string.
- For optional scalar information absent from the brief, use an empty string.
- For dynamic table rows, include only requirements or items supported by the brief.
- If a field is unknown, do not invent a value.
- If the template has no dynamic table placeholders, return an empty "items" list.
- Do not rewrite static boilerplate or legal text from the template.
- Do not invent client names, version numbers, prices, dates, integrations,
  performance targets, security certifications, or legal terms.

The application validates your response against this exact structure. Return valid JSON only.
"""

    normalized_type = document_type.strip().lower()
    system_prompt = (
        f"You generate accurate {normalized_type} content. "
        "Return valid JSON only. Never invent facts. "
        "Use the exact template placeholder names."
    )

    validation_feedback = ""

    for attempt in range(2):
        user_prompt = prompt

        if validation_feedback:
            user_prompt += (
                "\n\nYour previous response failed strict validation for this reason:\n"
                f"{validation_feedback}\n"
                "Return a corrected JSON response with the exact required keys and "
                "string values. Do not add explanations."
            )

        try:
            response = requests.post(
                base_url.rstrip("/") + "/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": model,
                    "temperature": 0.2,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {
                            "role": "system",
                            "content": system_prompt,
                        },
                        {
                            "role": "user",
                            "content": user_prompt,
                        },
                    ],
                },
                timeout=90,
            )
        except requests.RequestException as exc:
            raise RuntimeError(
                f"Could not contact the language model service: {exc}"
            ) from exc

        if not response.ok:
            try:
                error = response.json().get("error", {})
                detail = error.get("message") or response.text[:500]
                code = error.get("code") or error.get("type")
                suffix = f" [{code}]" if code else ""
            except (ValueError, AttributeError):
                detail = response.text[:500]
                suffix = ""

            raise RuntimeError(
                f"API error {response.status_code}{suffix}: {detail}"
            )

        try:
            raw_content = response.json()["choices"][0]["message"]["content"]

            if not isinstance(raw_content, str):
                raise ValueError("The model response content was not text.")

            return _parse_and_validate(
                raw_content,
                response_model,
                scalar_fields,
                item_fields,
            )

        except (
            KeyError,
            IndexError,
            TypeError,
            ValueError,
            ValidationError,
        ) as exc:
            validation_feedback = str(exc)

            if attempt == 1:
                raise ValueError(
                    "The model returned content that did not match the uploaded "
                    f"template after one correction attempt: {exc}"
                ) from exc

    raise RuntimeError("Document generation did not complete.")