#!/usr/bin/env python3

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any
from openpyxl import load_workbook


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")



def _load_workbook_with_validation(file_path: Path, result: dict[str, Any]) -> Any | None:
    if not file_path.exists():
        result["errors"].append(f"Input file does not exist: {file_path}")
        return None

    if file_path.suffix.lower() != ".xlsx":
        result["errors"].append(f"Input file must be an .xlsx file: {file_path.name}")
        return None

    try:
        return load_workbook(file_path, read_only=True, data_only=True)
    except Exception as exc:  # pragma: no cover - defensive branch
        result["errors"].append(f"Unable to open workbook: {exc}")
        return None


def validate_batch_template(file_path: Path, file_type: str = "batch") -> dict[str, Any]:
    result: dict[str, Any] = {
        "valid": False,
        "file": str(file_path),
        "file_type": file_type,
        "errors": [],
        "warnings": [],
        "worksheets": [],
    }

    workbook = _load_workbook_with_validation(file_path, result)
    if workbook is None:
        return result

    result["worksheets"] = workbook.sheetnames

    if not workbook.sheetnames:
        result["errors"].append("Workbook contains no worksheets.")
        return result

    interface_reference_matches = [
        sheet_name for sheet_name in workbook.sheetnames
        if "Interface Reference" in sheet_name
    ]

    if not interface_reference_matches:
        result["errors"].append(
            "Batch template validation failed: the workbook must contain a worksheet whose name includes 'Interface Reference'."
        )
        return result

    for sheet_name in interface_reference_matches:
        worksheet = workbook[sheet_name]
        rows = list(worksheet.iter_rows(values_only=True))

        if not rows:
            result["errors"].append(f"Worksheet '{sheet_name}' is empty.")
            continue

        for row_index, row in enumerate(rows[1:], start=2):
            if not row or len(row) < 4:
                continue

            mandatory_value = row[3]
            column_b_value = row[1]

            if str(mandatory_value).strip().lower() == "mandatory" and _is_blank(column_b_value):
                result["errors"].append(
                    f"Worksheet '{sheet_name}' row {row_index} is marked 'Mandatory' in column D but column B is blank."
                )

    if not result["errors"]:
        result["valid"] = True

    return result


def validate_workbook(file_path: Path, file_type: str = "batch") -> dict[str, Any]:
    if file_type.lower() == "batch":
        return validate_batch_template(file_path, file_type=file_type)

    result: dict[str, Any] = {
        "valid": False,
        "file": str(file_path),
        "file_type": file_type,
        "errors": [],
        "warnings": [],
        "worksheets": [],
    }

    workbook = _load_workbook_with_validation(file_path, result)
    if workbook is None:
        return result

    result["worksheets"] = workbook.sheetnames

    if not workbook.sheetnames:
        result["errors"].append("Workbook contains no worksheets.")
        return result

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        rows = list(worksheet.iter_rows(values_only=True))

        if not rows:
            result["warnings"].append(f"Worksheet '{sheet_name}' is empty.")
            continue

        header_row = rows[0]
        if all(_is_blank(value) for value in header_row):
            result["errors"].append(f"Worksheet '{sheet_name}' does not contain headers.")
            continue

        for row_index, row in enumerate(rows[1:], start=2):
            if all(_is_blank(value) for value in row):
                continue

            if len(row) < len(header_row):
                result["warnings"].append(
                    f"Worksheet '{sheet_name}' row {row_index} has fewer values than the header row."
                )

    result["valid"] = not result["errors"]
    return result


def _find_sheet_by_name(workbook: Any, target_name: str) -> str | None:
    normalized_target = " ".join(str(target_name).strip().lower().split())
    for sheet_name in workbook.sheetnames:
        normalized_sheet = " ".join(str(sheet_name).strip().lower().split())
        if normalized_target in normalized_sheet:
            return sheet_name
    return None


def _get_source_system_name(interface_metadata: dict[str, Any]) -> str:
    for key, value in interface_metadata.items():
        normalized_key = " ".join(str(key).strip().lower().split())
        if normalized_key in {"source system", "source_system", "source-system", "system"}:
            text = str(value).strip()
            if not _is_blank(text):
                return text
    return "Unknown Source System"


def build_interface_agreement_standard_json(file_path: Path, output_dir: Path | None = None) -> Path:
    if not file_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {file_path}")

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    interface_reference_sheet = _find_sheet_by_name(workbook, "Interface Reference")

    if interface_reference_sheet is None:
        raise ValueError(f"Workbook does not contain an 'Interface Reference' sheet: {file_path}")

    interface_worksheet = workbook[interface_reference_sheet]
    interface_rows = list(interface_worksheet.iter_rows(values_only=True))
    if not interface_rows:
        raise ValueError(f"Interface Reference sheet is empty: {file_path}")

    interface_metadata: dict[str, Any] = {}
    for row in interface_rows[1:]:
        if not row or len(row) < 2:
            continue

        key = str(row[0]).strip()
        value = row[1]

        if _is_blank(key):
            continue

        if key.lower() in {"key", "field", "column", "attribute"}:
            continue

        interface_metadata[key] = value

    datasets_sheet = _find_sheet_by_name(workbook, "Datasets")
    datasets: list[dict[str, Any]] = []
    if datasets_sheet is not None:
        dataset_worksheet = workbook[datasets_sheet]
        dataset_rows = list(dataset_worksheet.iter_rows(values_only=True))

        if dataset_rows:
            headers = [
                str(value).strip() if not _is_blank(value) else f"column_{idx + 1}"
                for idx, value in enumerate(dataset_rows[2])
            ]

            for row in dataset_rows[4:]:
                if not row or all(_is_blank(value) for value in row):
                    continue

                dataset_record: dict[str, Any] = {}
                for idx, value in enumerate(row):
                    column_name = headers[idx] if idx < len(headers) else f"column_{idx + 1}"
                    dataset_record[column_name] = value
                datasets.append(dataset_record)

    dataset_attributes_sheet = _find_sheet_by_name(workbook, "Dataset Attributes")
    dataset_attributes: list[dict[str, Any]] = []
    if dataset_attributes_sheet is not None:
        attributes_worksheet = workbook[dataset_attributes_sheet]
        attribute_rows = list(attributes_worksheet.iter_rows(values_only=True))

        if attribute_rows:
            header_row_index = 2
            headers = [
                str(value).strip() if not _is_blank(value) else f"column_{idx + 1}"
                for idx, value in enumerate(attribute_rows[header_row_index][1:10])
            ]

            grouped_attribute_rows: dict[str, list[dict[str, Any]]] = {}
            for row in attribute_rows[4:]:
                if not row or all(_is_blank(value) for value in row[1:10]):
                    continue

                dataset_name = str(row[1]).strip() if len(row) > 1 else ""
                if _is_blank(dataset_name):
                    continue

                attribute_record: dict[str, Any] = {}
                for idx, value in enumerate(row[2:10]):
                    column_name = headers[idx + 1] if idx + 1 < len(headers) else f"column_{idx + 2}"
                    attribute_record[column_name] = value

                grouped_attribute_rows.setdefault(dataset_name, []).append(attribute_record)

            dataset_attributes = [
                {"dataset_name": dataset_name, "attributes": attributes}
                for dataset_name, attributes in grouped_attribute_rows.items()
            ]

    payload: dict[str, Any] = {
        "source_file": file_path.name,
        "interface_agreement": interface_metadata,
        "datasets": datasets,
        "dataset_attributes": dataset_attributes,
    }

    source_system_name = _get_source_system_name(interface_metadata)
    safe_source_system = re.sub(r"[^A-Za-z0-9 _\-.]", "", source_system_name).strip()
    safe_source_system = " ".join(safe_source_system.split()) or "Unknown Source System"

    destination_dir = output_dir or Path("build/generated")
    destination_dir.mkdir(parents=True, exist_ok=True)
    output_file = destination_dir / f"Interface Agreement - {safe_source_system}.json"
    output_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_file


def convert_workbook_to_records(file_path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    records: list[dict[str, Any]] = []

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        rows = list(worksheet.iter_rows(values_only=True))

        if not rows:
            continue

        headers = [str(value).strip() if not _is_blank(value) else f"column_{idx + 1}" for idx, value in enumerate(rows[0])]
        for row in rows[1:]:
            if all(_is_blank(value) for value in row):
                continue

            item: dict[str, Any] = {"worksheet": sheet_name}
            for idx, value in enumerate(row):
                header = headers[idx] if idx < len(headers) else f"column_{idx + 1}"
                item[header] = value
            records.append(item)

    return records


def write_report(report_dir: Path, source_file: Path, validation: dict[str, Any], generated_json: Path | None = None) -> Path:
    report_dir.mkdir(parents=True, exist_ok=True)
    report_data: dict[str, Any] = {
        "source_file": str(source_file),
        "status": "valid" if validation["valid"] else "invalid",
        "validation": validation,
        "generated_json": str(generated_json) if generated_json else None,
    }
    report_file = report_dir / f"{source_file.stem}_validation_report.json"
    report_file.write_text(json.dumps(report_data, ensure_ascii=False, indent=2), encoding="utf-8")
    return report_file


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate and convert Interface Agreement Excel files into JSON.")
    parser.add_argument("--input", required=True, help="Path to the source .xlsx file")
    parser.add_argument("--output-dir", default="build/generated", help="Directory for generated JSON output")
    parser.add_argument("--report-dir", default="build/reports", help="Directory for validation reports")
    parser.add_argument("--file-type", default="batch", help="Workbook file type used to select the correct sheet metadata contract")
    parser.add_argument("--validate-only", action="store_true", help="Only validate the workbook and write evidence; do not output JSON")
    return parser.parse_args(argv)


def run_interface_agreement_pipeline(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    input_file = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    report_dir = Path(args.report_dir).expanduser().resolve()

    validation = validate_workbook(input_file, file_type=args.file_type)
    report_path = write_report(report_dir, input_file, validation)

    if not validation["valid"]:
        print("Validation failed.", file=sys.stderr)
        for error in validation["errors"]:
            print(f"- {error}", file=sys.stderr)
        print(f"Validation evidence saved to: {report_path}", file=sys.stderr)
        return 1

    print(f"Validation passed for: {input_file}")
    print(f"Validation evidence saved to: {report_path}")

    interface_reference_output = build_interface_agreement_standard_json(input_file, output_dir)
    print(f"Saved Interface Reference standard JSON: {interface_reference_output}")

    if args.validate_only:
        print("Validation-only mode enabled; no full canonical JSON output generated.")
        return 0

    # Refresh the report with the generated JSON path after conversion.
    write_report(report_dir, input_file, validation, generated_json=interface_reference_output)

    print(f"Converted workbook to canonical JSON: {interface_reference_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run_interface_agreement_pipeline())
