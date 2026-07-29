from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


@dataclass(frozen=True)
class ReferenceRule:
    start: int
    step: int
    total: int


REFERENCE_RULES = {
    "y007": ReferenceRule(1, 1, 1),
    "y015": ReferenceRule(4, 4, 13),
    "y016": ReferenceRule(3, 1, 5),
    "y017": ReferenceRule(2, 1, 5),
    "y018": ReferenceRule(1, 1, 4),
    "y020": ReferenceRule(1, 4, 12),
    "y025": ReferenceRule(1, 1, 1),
    "y029": ReferenceRule(1, 1, 4),
    "y037": ReferenceRule(3, 1, 3),
    "y040": ReferenceRule(2, 4, 10),
    "y046": ReferenceRule(4, 1, 5),
    "y047": ReferenceRule(3, 1, 6),
    "y048": ReferenceRule(1, 4, 8),
    "y049": ReferenceRule(1, 4, 10),
}


def clean_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""

    return str(value).strip()


def normalize_reference(value: Any) -> str:
    reference = clean_value(value).lower()

    return re.sub(
        r"-\d+/\d+$",
        "",
        reference,
    )


def split_labels(value: Any) -> list[str]:
    text = clean_value(value)

    if not text:
        return []

    items = re.split(
        r"\s*[,，]\s*",
        text,
    )

    return [
        item.strip()
        for item in items
        if item.strip()
    ]


def calculate_visual_width(value: Any) -> int:
    text = clean_value(value)
    width = 0

    for character in text:
        character_width = unicodedata.east_asian_width(
            character
        )

        if character_width in {"W", "F"}:
            width += 2
        else:
            width += 1

    return width


def determine_label_size(
    label: str,
    large_width: int,
    medium_width: int,
) -> tuple[str, int]:
    width = calculate_visual_width(label)

    if width <= large_width:
        return "large", width

    if width <= medium_width:
        return "medium", width

    return "small", width


def generate_reference(
    base_reference: str,
    label_index: int,
) -> tuple[str, str]:
    base_reference = normalize_reference(
        base_reference
    )

    if not base_reference:
        return "", "Reference is empty."

    rule = REFERENCE_RULES.get(base_reference)

    if rule is None:
        return (
            base_reference,
            f"No reference rule exists for {base_reference}.",
        )

    position = (
        rule.start
        + label_index * rule.step
    )

    full_reference = (
        f"{base_reference}-"
        f"{position:02d}/"
        f"{rule.total:02d}"
    )

    warning = ""

    if position > rule.total:
        warning = (
            f"Generated reference {full_reference} "
            "exceeds the configured total."
        )

    return full_reference, warning


def validate_columns(
    dataframe: pd.DataFrame,
) -> None:
    required_columns = {
        "sp",
        "ref",
        "label",
    }

    missing_columns = (
        required_columns
        - set(dataframe.columns)
    )

    if missing_columns:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(
                sorted(missing_columns)
            )
        )


def expand_source_data(
    source_df: pd.DataFrame,
    large_width: int,
    medium_width: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    expanded_rows = []
    validation_rows = []

    for index, source_row in source_df.iterrows():
        excel_row = index + 2

        sp = clean_value(
            source_row["sp"]
        )

        base_reference = normalize_reference(
            source_row["ref"]
        )

        labels = split_labels(
            source_row["label"]
        )

        if not sp:
            validation_rows.append(
                {
                    "severity": "Warning",
                    "source_row": excel_row,
                    "sp": sp,
                    "ref": base_reference,
                    "message": "The sp value is empty.",
                }
            )

        if not base_reference:
            validation_rows.append(
                {
                    "severity": "Warning",
                    "source_row": excel_row,
                    "sp": sp,
                    "ref": base_reference,
                    "message": "The ref value is empty.",
                }
            )

        if not labels:
            validation_rows.append(
                {
                    "severity": "Warning",
                    "source_row": excel_row,
                    "sp": sp,
                    "ref": base_reference,
                    "message": "No valid labels were found.",
                }
            )
            continue

        for label_index, label in enumerate(labels):
            full_reference, warning = generate_reference(
                base_reference=base_reference,
                label_index=label_index,
            )

            label_size, visual_width = (
                determine_label_size(
                    label=label,
                    large_width=large_width,
                    medium_width=medium_width,
                )
            )

            expanded_rows.append(
                {
                    "source_row": excel_row,
                    "sp": sp,
                    "base_ref": base_reference,
                    "ref": full_reference,
                    "label": label,
                    "label_large": (
                        label
                        if label_size == "large"
                        else ""
                    ),
                    "label_medium": (
                        label
                        if label_size == "medium"
                        else ""
                    ),
                    "label_small": (
                        label
                        if label_size == "small"
                        else ""
                    ),
                    "label_size": label_size,
                    "label_visual_width": visual_width,
                }
            )

            if warning:
                validation_rows.append(
                    {
                        "severity": "Warning",
                        "source_row": excel_row,
                        "sp": sp,
                        "ref": base_reference,
                        "message": warning,
                    }
                )

    expanded_df = pd.DataFrame(
        expanded_rows,
        columns=[
            "source_row",
            "sp",
            "base_ref",
            "ref",
            "label",
            "label_large",
            "label_medium",
            "label_small",
            "label_size",
            "label_visual_width",
        ],
    )

    validation_df = pd.DataFrame(
        validation_rows,
        columns=[
            "severity",
            "source_row",
            "sp",
            "ref",
            "message",
        ],
    )

    return expanded_df, validation_df


def create_mail_merge_data(
    expanded_df: pd.DataFrame,
    group_size: int,
) -> pd.DataFrame:
    output_rows = []

    for start_index in range(
        0,
        len(expanded_df),
        group_size,
    ):
        group = expanded_df.iloc[
            start_index:start_index + group_size
        ]

        output_row = {}

        for position in range(group_size):
            number = position + 1

            field_names = {
                "sp": f"sp{number}",
                "ref": f"ref{number}",
                "label": f"label{number}",
                "large": f"label{number}_large",
                "medium": f"label{number}_medium",
                "small": f"label{number}_small",
                "size": f"label{number}_size",
            }

            if position < len(group):
                record = group.iloc[position]

                output_row[
                    field_names["sp"]
                ] = record["sp"]

                output_row[
                    field_names["ref"]
                ] = record["ref"]

                output_row[
                    field_names["label"]
                ] = record["label"]

                output_row[
                    field_names["large"]
                ] = record["label_large"]

                output_row[
                    field_names["medium"]
                ] = record["label_medium"]

                output_row[
                    field_names["small"]
                ] = record["label_small"]

                output_row[
                    field_names["size"]
                ] = record["label_size"]

            else:
                for field_name in field_names.values():
                    output_row[field_name] = ""

        output_rows.append(output_row)

    return pd.DataFrame(output_rows)


def format_excel_workbook(
    source_buffer: io.BytesIO,
) -> io.BytesIO:
    source_buffer.seek(0)

    workbook = load_workbook(
        source_buffer
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for row in worksheet.iter_rows(
            min_row=2,
        ):
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

        for column_index in range(
            1,
            worksheet.max_column + 1,
        ):
            column_letter = get_column_letter(
                column_index
            )

            worksheet.column_dimensions[
                column_letter
            ].width = 22

    output_buffer = io.BytesIO()

    workbook.save(
        output_buffer
    )

    output_buffer.seek(0)

    return output_buffer


def convert_excel(
    source_file,
    group_size: int = 4,
    large_width: int = 16,
    medium_width: int = 28,
) -> dict:
    source_file.seek(0)

    source_df = pd.read_excel(
        source_file,
        dtype=str,
    ).fillna("")

    source_df.columns = [
        clean_value(column).lower()
        for column in source_df.columns
    ]

    validate_columns(source_df)

    expanded_df, validation_df = (
        expand_source_data(
            source_df=source_df,
            large_width=large_width,
            medium_width=medium_width,
        )
    )

    if expanded_df.empty:
        raise ValueError(
            "No output records were created."
        )

    mail_merge_df = create_mail_merge_data(
        expanded_df=expanded_df,
        group_size=group_size,
    )

    raw_buffer = io.BytesIO()

    with pd.ExcelWriter(
        raw_buffer,
        engine="openpyxl",
    ) as writer:
        mail_merge_df.to_excel(
            writer,
            sheet_name="MailMerge",
            index=False,
        )

        expanded_df.to_excel(
            writer,
            sheet_name="ExpandedData",
            index=False,
        )

        source_df.to_excel(
            writer,
            sheet_name="OriginalData",
            index=False,
        )

        validation_df.to_excel(
            writer,
            sheet_name="Validation",
            index=False,
        )

    formatted_buffer = format_excel_workbook(
        raw_buffer
    )

    return {
        "output": formatted_buffer,
        "original": source_df,
        "expanded": expanded_df,
        "mail_merge": mail_merge_df,
        "validation": validation_df,
    }
