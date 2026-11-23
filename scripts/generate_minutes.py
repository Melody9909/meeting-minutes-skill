#!/usr/bin/env python3
"""Generate an editable Tutor-format Meeting Minutes DOCX from validated JSON.

This program validates structure and field rules from template-schema.json. It
does not verify whether meeting content is true to a transcript; that check must
already have been completed by the meeting-minutes Skill extraction workflow.
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any

try:
    from docx import Document
    from docx.document import Document as DocumentObject
    from docx.table import Table, _Row
    from docx.text.paragraph import Paragraph
except ImportError as exc:  # pragma: no cover - exercised when dependency is absent
    print(
        "Dependency error: python-docx is required. "
        "Install it with: python -m pip install python-docx",
        file=sys.stderr,
    )
    raise SystemExit(4) from exc


SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_SCHEMA_PATH = SKILL_DIR / "references" / "template-schema.json"
DEFAULT_RULES_PATH = SKILL_DIR / "references" / "meeting-rules.md"
ASSETS_DIR = SKILL_DIR / "assets"


class ValidationError(Exception):
    """Raised when schema, input JSON, or template structure is invalid."""

    def __init__(self, errors: list[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate an editable Meeting Minutes DOCX from structured JSON."
    )
    parser.add_argument("meeting_json", type=Path, help="Structured meeting JSON path")
    parser.add_argument("output_docx", type=Path, help="Output DOCX path")
    parser.add_argument(
        "--schema",
        type=Path,
        default=DEFAULT_SCHEMA_PATH,
        help=f"Schema path (default: {DEFAULT_SCHEMA_PATH})",
    )
    parser.add_argument(
        "--rules",
        type=Path,
        default=DEFAULT_RULES_PATH,
        help=f"Business rules path (default: {DEFAULT_RULES_PATH})",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=None,
        help="Tutor DOCX path; defaults to the single DOCX in the Skill assets folder",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacement of an existing output DOCX",
    )
    return parser.parse_args()


def read_json(path: Path, label: str) -> Any:
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            return json.load(handle)
    except FileNotFoundError as exc:
        raise ValidationError([f"{label}: file not found: {path}"]) from exc
    except json.JSONDecodeError as exc:
        raise ValidationError(
            [f"{label}: invalid JSON at line {exc.lineno}, column {exc.colno}: {exc.msg}"]
        ) from exc
    except OSError as exc:
        raise ValidationError([f"{label}: cannot read {path}: {exc}"]) from exc


def read_rules(path: Path) -> str:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError as exc:
        raise ValidationError([f"meeting rules: file not found: {path}"]) from exc
    except OSError as exc:
        raise ValidationError([f"meeting rules: cannot read {path}: {exc}"]) from exc
    if not text.strip():
        raise ValidationError([f"meeting rules: file is empty: {path}"])
    return text


def schema_sections(schema: dict[str, Any]) -> list[dict[str, Any]]:
    sections = schema.get("sections")
    if not isinstance(sections, list) or not sections:
        raise ValidationError(["schema.sections: expected a non-empty array"])
    return sections


def validate_schema(schema: Any) -> dict[str, Any]:
    errors: list[str] = []
    if not isinstance(schema, dict):
        raise ValidationError(["schema: expected an object"])

    document = schema.get("document")
    if not isinstance(document, dict):
        errors.append("schema.document: expected an object")
    else:
        for key in ("identifier", "title", "format", "editable"):
            if key not in document:
                errors.append(f"schema.document.{key}: missing")
        if document.get("format") != "docx":
            errors.append("schema.document.format: expected 'docx'")
        if document.get("editable") is not True:
            errors.append("schema.document.editable: expected true")

    try:
        sections = schema_sections(schema)
    except ValidationError as exc:
        errors.extend(exc.errors)
        sections = []

    seen_keys: set[str] = set()
    supported_types = {"fields", "table", "list"}
    for index, section in enumerate(sections):
        path = f"schema.sections[{index}]"
        if not isinstance(section, dict):
            errors.append(f"{path}: expected an object")
            continue
        for key in ("heading", "type", "data_key"):
            value = section.get(key)
            if not isinstance(value, str) or not value.strip():
                errors.append(f"{path}.{key}: expected a non-empty string")
        section_type = section.get("type")
        if section_type not in supported_types:
            errors.append(f"{path}.type: unsupported value {section_type!r}")
        data_key = section.get("data_key")
        if isinstance(data_key, str):
            if data_key in seen_keys:
                errors.append(f"{path}.data_key: duplicate value {data_key!r}")
            seen_keys.add(data_key)

        definitions_key = {"fields": "fields", "table": "columns"}.get(section_type)
        if definitions_key:
            definitions = section.get(definitions_key)
            if not isinstance(definitions, list) or not definitions:
                errors.append(f"{path}.{definitions_key}: expected a non-empty array")
                continue
            seen_definition_keys: set[str] = set()
            for def_index, definition in enumerate(definitions):
                def_path = f"{path}.{definitions_key}[{def_index}]"
                if not isinstance(definition, dict):
                    errors.append(f"{def_path}: expected an object")
                    continue
                for key in ("label", "key"):
                    value = definition.get(key)
                    if not isinstance(value, str) or not value.strip():
                        errors.append(f"{def_path}.{key}: expected a non-empty string")
                definition_key = definition.get("key")
                if isinstance(definition_key, str):
                    if definition_key in seen_definition_keys:
                        errors.append(f"{def_path}.key: duplicate value {definition_key!r}")
                    seen_definition_keys.add(definition_key)
                if "generated" not in definition and definition.get("type") != "string":
                    errors.append(f"{def_path}.type: only 'string' is supported")
        elif section_type == "list" and section.get("item_type") != "string":
            errors.append(f"{path}.item_type: only 'string' is supported")

    required_data_keys = {
        "meeting",
        "agenda",
        "open_actions",
        "questions_for_mentor",
        "questions_for_client",
    }
    missing_data_keys = sorted(required_data_keys - seen_keys)
    if missing_data_keys:
        errors.append(
            "schema.sections: missing required data_key values: "
            + ", ".join(missing_data_keys)
        )

    if errors:
        raise ValidationError(errors)
    return schema


def type_matches(value: Any, declared_type: str) -> bool:
    return declared_type == "string" and isinstance(value, str)


def validate_fields_section(
    section: dict[str, Any], value: Any, errors: list[str]
) -> dict[str, str]:
    data_key = section["data_key"]
    if not isinstance(value, dict):
        errors.append(f"{data_key}: expected an object")
        return {}

    definitions = section["fields"]
    allowed_keys = {definition["key"] for definition in definitions}
    actual_keys = set(value)
    for key in sorted(allowed_keys - actual_keys):
        errors.append(f"{data_key}.{key}: missing field")
    for key in sorted(actual_keys - allowed_keys):
        errors.append(f"{data_key}.{key}: unexpected field")

    normalized: dict[str, str] = {}
    for definition in definitions:
        key = definition["key"]
        field_path = f"{data_key}.{key}"
        if key not in value:
            continue
        field_value = value[key]
        if not type_matches(field_value, definition.get("type", "")):
            errors.append(f"{field_path}: expected string")
            continue
        if definition.get("required") and not field_value.strip():
            errors.append(f"{field_path}: required value must not be empty")
        if "fixed_value" in definition and field_value != definition["fixed_value"]:
            errors.append(
                f"{field_path}: expected fixed value {definition['fixed_value']!r}"
            )
        normalized[key] = field_value
    return normalized


def validate_table_section(
    section: dict[str, Any], value: Any, errors: list[str]
) -> list[dict[str, str]]:
    data_key = section["data_key"]
    if not isinstance(value, list):
        errors.append(f"{data_key}: expected an array")
        return []

    definitions = section["columns"]
    input_definitions = [item for item in definitions if "generated" not in item]
    generated_keys = {item["key"] for item in definitions if "generated" in item}
    allowed_keys = {item["key"] for item in input_definitions}
    normalized_rows: list[dict[str, str]] = []

    for row_index, row in enumerate(value):
        row_path = f"{data_key}[{row_index}]"
        if not isinstance(row, dict):
            errors.append(f"{row_path}: expected an object")
            continue
        actual_keys = set(row)
        for key in sorted(actual_keys & generated_keys):
            errors.append(f"{row_path}.{key}: generated fields must not be supplied")
        for key in sorted(actual_keys - allowed_keys - generated_keys):
            errors.append(f"{row_path}.{key}: unexpected field")

        normalized_row: dict[str, str] = {}
        for definition in input_definitions:
            key = definition["key"]
            field_path = f"{row_path}.{key}"
            if key not in row:
                if "default" in definition:
                    normalized_row[key] = definition["default"]
                elif definition.get("required"):
                    errors.append(f"{field_path}: missing field")
                continue
            field_value = row[key]
            if not type_matches(field_value, definition.get("type", "")):
                errors.append(f"{field_path}: expected string")
                continue
            if definition.get("non_empty") and not field_value.strip():
                errors.append(f"{field_path}: value must not be empty")
            if "default" in definition and field_value != definition["default"]:
                errors.append(
                    f"{field_path}: expected {definition['default']!r} "
                    "according to the schema default/generation rule"
                )
            normalized_row[key] = field_value
        normalized_rows.append(normalized_row)
    return normalized_rows


def validate_list_section(
    section: dict[str, Any], value: Any, errors: list[str]
) -> list[str]:
    data_key = section["data_key"]
    if not isinstance(value, list):
        errors.append(f"{data_key}: expected an array")
        return []
    normalized: list[str] = []
    for index, item in enumerate(value):
        item_path = f"{data_key}[{index}]"
        if not isinstance(item, str):
            errors.append(f"{item_path}: expected string")
        elif not item.strip():
            errors.append(f"{item_path}: value must not be empty")
        else:
            normalized.append(item)
    return normalized


def validate_meeting_json(
    payload: Any, schema: dict[str, Any]
) -> dict[str, Any]:
    errors: list[str] = []
    if not isinstance(payload, dict):
        raise ValidationError(["meeting JSON root: expected an object"])

    sections = schema_sections(schema)
    expected_keys = {section["data_key"] for section in sections}
    actual_keys = set(payload)
    for key in sorted(expected_keys - actual_keys):
        errors.append(f"{key}: missing root field")
    for key in sorted(actual_keys - expected_keys):
        errors.append(f"{key}: unexpected root field")

    normalized: dict[str, Any] = {}
    for section in sections:
        data_key = section["data_key"]
        if data_key not in payload:
            continue
        section_type = section["type"]
        if section_type == "fields":
            normalized[data_key] = validate_fields_section(
                section, payload[data_key], errors
            )
        elif section_type == "table":
            normalized[data_key] = validate_table_section(
                section, payload[data_key], errors
            )
        elif section_type == "list":
            normalized[data_key] = validate_list_section(
                section, payload[data_key], errors
            )

    if errors:
        raise ValidationError(errors)
    return normalized


def resolve_template(explicit_path: Path | None) -> Path:
    if explicit_path is not None:
        template = explicit_path
    else:
        candidates = sorted(ASSETS_DIR.glob("*.docx"))
        if len(candidates) != 1:
            raise ValidationError(
                [
                    "template: expected exactly one DOCX in "
                    f"{ASSETS_DIR}, found {len(candidates)}; use --template"
                ]
            )
        template = candidates[0]
    if not template.is_file():
        raise ValidationError([f"template: file not found: {template}"])
    if template.suffix.lower() != ".docx":
        raise ValidationError([f"template: expected a .docx file: {template}"])
    return template.resolve()


def clean_text(value: str) -> str:
    return value.replace("\r", "").replace("\x07", "").strip()


def paragraph_text(paragraph: Paragraph) -> str:
    return clean_text(paragraph.text)


def cell_text(cell: Any) -> str:
    return clean_text(cell.text)


def set_cell_text(cell: Any, text: str) -> None:
    cell.text = text


def find_exact_paragraph(document: DocumentObject, text: str) -> Paragraph:
    matches = [p for p in document.paragraphs if paragraph_text(p) == text]
    if len(matches) != 1:
        raise ValidationError(
            [f"template paragraph {text!r}: expected one match, found {len(matches)}"]
        )
    return matches[0]


def find_fields_table(document: DocumentObject, section: dict[str, Any]) -> Table:
    expected_labels = [field["label"] for field in section["fields"]]
    matches: list[Table] = []
    for table in document.tables:
        labels = [cell_text(row.cells[0]) for row in table.rows if len(row.cells) >= 2]
        if labels == expected_labels:
            matches.append(table)
    if len(matches) != 1:
        raise ValidationError(
            [
                f"template section {section['heading']!r}: expected one fields table "
                f"with labels {expected_labels!r}, found {len(matches)}"
            ]
        )
    return matches[0]


def table_headers(table: Table) -> list[str]:
    if not table.rows:
        return []
    return [cell_text(cell) for cell in table.rows[0].cells]


def find_data_table(document: DocumentObject, section: dict[str, Any]) -> Table:
    expected_headers = [column["label"] for column in section["columns"]]
    matches = [
        table for table in document.tables if table_headers(table) == expected_headers
    ]
    if len(matches) != 1:
        raise ValidationError(
            [
                f"template section {section['heading']!r}: expected one table with "
                f"headers {expected_headers!r}, found {len(matches)}"
            ]
        )
    return matches[0]


def replace_table_body(table: Table, row_count: int) -> list[_Row]:
    if row_count < 1:
        raise ValueError("row_count must be at least 1")
    if len(table.rows) > 1:
        row_template = copy.deepcopy(table.rows[1]._tr)
    else:
        created = table.add_row()
        row_template = copy.deepcopy(created._tr)
    for row in list(table.rows[1:]):
        table._tbl.remove(row._tr)
    for _ in range(row_count):
        table._tbl.append(copy.deepcopy(row_template))
    return list(table.rows[1:])


def populate_fields_table(
    document: DocumentObject, section: dict[str, Any], values: dict[str, str]
) -> None:
    table = find_fields_table(document, section)
    for row, definition in zip(table.rows, section["fields"], strict=True):
        set_cell_text(row.cells[1], values[definition["key"]])


def populate_data_table(
    document: DocumentObject, section: dict[str, Any], rows: list[dict[str, str]]
) -> None:
    table = find_data_table(document, section)
    output_rows = replace_table_body(table, max(1, len(rows)))
    columns = section["columns"]

    if not rows:
        for cell in output_rows[0].cells:
            set_cell_text(cell, "")
        content_column = next(
            index
            for index, column in enumerate(columns)
            if "generated" not in column and column["key"] != "deadline"
        )
        set_cell_text(
            output_rows[0].cells[content_column], section.get("empty_row_value", "None")
        )
        return

    for row_number, (output_row, values) in enumerate(zip(output_rows, rows), start=1):
        for column_index, definition in enumerate(columns):
            if definition.get("generated") == "sequential_number":
                text = str(row_number)
            else:
                text = values.get(definition["key"], definition.get("default", ""))
            set_cell_text(output_row.cells[column_index], text)


def has_numbering(paragraph: Paragraph) -> bool:
    properties = paragraph._p.pPr
    return properties is not None and properties.numPr is not None


def set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    for child in list(paragraph._p):
        if child is not paragraph._p.pPr:
            paragraph._p.remove(child)
    paragraph.add_run(text)


def section_paragraphs(
    document: DocumentObject, heading: str, all_headings: set[str]
) -> tuple[Paragraph, list[Paragraph]]:
    paragraphs = document.paragraphs
    heading_indexes = [
        index for index, paragraph in enumerate(paragraphs) if paragraph_text(paragraph) == heading
    ]
    if len(heading_indexes) != 1:
        raise ValidationError(
            [f"template heading {heading!r}: expected one match, found {len(heading_indexes)}"]
        )
    start = heading_indexes[0]
    end = len(paragraphs)
    for index in range(start + 1, len(paragraphs)):
        if paragraph_text(paragraphs[index]) in all_headings:
            end = index
            break
    return paragraphs[start], paragraphs[start + 1 : end]


def populate_list_section(
    document: DocumentObject,
    section: dict[str, Any],
    values: list[str],
    all_headings: set[str],
) -> None:
    _, paragraphs = section_paragraphs(document, section["heading"], all_headings)
    numbered = [paragraph for paragraph in paragraphs if has_numbering(paragraph)]
    if not numbered:
        raise ValidationError(
            [f"template section {section['heading']!r}: numbered list paragraph not found"]
        )

    base = numbered[0]
    for extra in numbered[1:]:
        extra._p.getparent().remove(extra._p)
    output_values = values or [section.get("empty_value", "None")]
    set_paragraph_text(base, output_values[0])

    previous_xml = base._p
    for value in output_values[1:]:
        cloned_xml = copy.deepcopy(base._p)
        previous_xml.addnext(cloned_xml)
        cloned_paragraph = Paragraph(cloned_xml, base._parent)
        set_paragraph_text(cloned_paragraph, value)
        previous_xml = cloned_xml


def populate_document(
    document: DocumentObject, schema: dict[str, Any], data: dict[str, Any]
) -> None:
    fixed_document = schema["document"]
    find_exact_paragraph(document, fixed_document["identifier"])
    find_exact_paragraph(document, fixed_document["title"])

    sections = schema_sections(schema)
    all_headings = {section["heading"] for section in sections}
    for section in sections:
        find_exact_paragraph(document, section["heading"])
        values = data[section["data_key"]]
        if section["type"] == "fields":
            populate_fields_table(document, section, values)
        elif section["type"] == "table":
            populate_data_table(document, section, values)
        elif section["type"] == "list":
            populate_list_section(document, section, values, all_headings)


def verify_generated_document(
    path: Path, schema: dict[str, Any], data: dict[str, Any]
) -> None:
    try:
        document = Document(path)
    except Exception as exc:
        raise ValidationError([f"generated DOCX cannot be reopened: {exc}"]) from exc

    fixed_document = schema["document"]
    find_exact_paragraph(document, fixed_document["identifier"])
    find_exact_paragraph(document, fixed_document["title"])
    sections = schema_sections(schema)
    all_headings = {section["heading"] for section in sections}

    for section in sections:
        find_exact_paragraph(document, section["heading"])
        values = data[section["data_key"]]
        if section["type"] == "fields":
            table = find_fields_table(document, section)
            for row, definition in zip(table.rows, section["fields"], strict=True):
                expected = values[definition["key"]]
                actual = cell_text(row.cells[1])
                if actual != expected:
                    raise ValidationError(
                        [
                            f"verification {section['data_key']}.{definition['key']}: "
                            f"expected {expected!r}, found {actual!r}"
                        ]
                    )
        elif section["type"] == "table":
            table = find_data_table(document, section)
            expected_rows = max(1, len(values))
            if len(table.rows) - 1 != expected_rows:
                raise ValidationError(
                    [
                        f"verification {section['data_key']}: expected {expected_rows} "
                        f"body rows, found {len(table.rows) - 1}"
                    ]
                )
            if not values:
                expected_empty = section.get("empty_row_value", "None")
                body_text = [cell_text(cell) for cell in table.rows[1].cells]
                if expected_empty not in body_text:
                    raise ValidationError(
                        [
                            f"verification {section['data_key']}: empty row does not "
                            f"contain {expected_empty!r}"
                        ]
                    )
            else:
                for row_number, (row, row_data) in enumerate(
                    zip(table.rows[1:], values), start=1
                ):
                    for column_index, definition in enumerate(section["columns"]):
                        expected = (
                            str(row_number)
                            if definition.get("generated") == "sequential_number"
                            else row_data.get(
                                definition["key"], definition.get("default", "")
                            )
                        )
                        actual = cell_text(row.cells[column_index])
                        if actual != expected:
                            raise ValidationError(
                                [
                                    f"verification {section['data_key']}[{row_number - 1}]."
                                    f"{definition['key']}: expected {expected!r}, "
                                    f"found {actual!r}"
                                ]
                            )
        elif section["type"] == "list":
            _, paragraphs = section_paragraphs(
                document, section["heading"], all_headings
            )
            actual = [paragraph_text(p) for p in paragraphs if has_numbering(p)]
            expected = values or [section.get("empty_value", "None")]
            if actual != expected:
                raise ValidationError(
                    [
                        f"verification {section['data_key']}: expected {expected!r}, "
                        f"found {actual!r}"
                    ]
                )


def generate_docx(
    template: Path,
    output: Path,
    schema: dict[str, Any],
    data: dict[str, Any],
    overwrite: bool,
) -> None:
    output = output.resolve()
    if output.suffix.lower() != ".docx":
        raise ValidationError([f"output: expected a .docx path: {output}"])
    if output == template.resolve():
        raise ValidationError(["output: must not overwrite the original Tutor template"])
    if output.exists() and not overwrite:
        raise ValidationError(
            [f"output: file already exists: {output}; use --overwrite to replace it"]
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        document = Document(template)
        populate_document(document, schema, data)
        with tempfile.NamedTemporaryFile(
            prefix=f".{output.stem}.", suffix=".docx", dir=output.parent, delete=False
        ) as temporary:
            temporary_path = Path(temporary.name)
        document.save(temporary_path)
        verify_generated_document(temporary_path, schema, data)
        os.replace(temporary_path, output)
        temporary_path = None
    except ValidationError:
        raise
    except Exception as exc:
        raise RuntimeError(f"DOCX generation failed: {exc}") from exc
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


def main() -> int:
    args = parse_args()
    try:
        schema = validate_schema(read_json(args.schema, "schema"))
        read_rules(args.rules)
        payload = read_json(args.meeting_json, "meeting JSON")
        data = validate_meeting_json(payload, schema)
        template = resolve_template(args.template)
        generate_docx(template, args.output_docx, schema, data, args.overwrite)
    except ValidationError as exc:
        print("Validation failed; no DOCX was generated:", file=sys.stderr)
        for error in exc.errors:
            print(f"- {error}", file=sys.stderr)
        return 2
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 3

    print(f"Created editable DOCX: {args.output_docx.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
