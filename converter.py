from __future__ import annotations

import io
import re
import unicodedata
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def clean_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""

    return str(value).strip()


def normalize_reference(value: Any) -> str:
    reference = clean_value(value)

    return re.sub(
        r"-\d+/\d+$",
        "",
        reference,
    ).strip()


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
    label_number: int,
    total_labels: int,
) -> str:
    """
    Generate a reference such as:

        y050-01/08
        y050-02/08
        y050-08/08

    The number width automatically expands when there are
    more than 99 labels.
    """

    base_reference = normalize_reference(
        base_reference
    )

    if not base_reference:
        return ""

    if total_labels < 1:
        return base_reference

    number_width = max(
        2,
        len(str(total_labels)),
    )

    return (
        f"{base_reference}-"
        f"{label_number:0{number_width}d}/"
        f"{total_labels:0{number_width}d}"
    )

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

    # -----------------------------------------------------
    # Step 1: Expand every source row into individual labels
    # -----------------------------------------------------
    for source_index, source_row in source_df.iterrows():
        excel_row = source_index + 2

        sp = clean_value(
            source_row.get("sp", "")
        )

        base_reference = normalize_reference(
            source_row.get("ref", "")
        )

        labels = split_labels(
            source_row.get("label", "")
        )

        if not sp:
            validation_rows.append(
                {
                    "severity": "Warning",
                    "source_row": excel_row,
                    "sp": "",
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
                    "ref": "",
                    "message": "The reference is empty.",
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

        for label_order_in_source, label in enumerate(
            labels,
            start=1,
        ):
            label_size, visual_width = determine_label_size(
                label=label,
                large_width=large_width,
                medium_width=medium_width,
            )

            expanded_rows.append(
                {
                    "source_order": source_index,
                    "source_row": excel_row,
                    "label_order_in_source": label_order_in_source,
                    "sp": sp,
                    "base_ref": base_reference,
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

    expanded_df = pd.DataFrame(expanded_rows)

    if expanded_df.empty:
        expanded_df = pd.DataFrame(
            columns=[
                "source_order",
                "source_row",
                "label_order_in_source",
                "sp",
                "base_ref",
                "ref",
                "label_number",
                "total_labels",
                "label",
                "label_large",
                "label_medium",
                "label_small",
                "label_size",
                "label_visual_width",
            ]
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

    # -----------------------------------------------------
    # Step 2: Keep the original workbook order
    # -----------------------------------------------------
    expanded_df = expanded_df.sort_values(
        by=[
            "source_order",
            "label_order_in_source",
        ],
        kind="stable",
    ).reset_index(drop=True)

    # -----------------------------------------------------
    # Step 3: Number all labels belonging to the same ref
    # -----------------------------------------------------
    valid_reference = (
        expanded_df["base_ref"]
        .astype(str)
        .str.strip()
        .ne("")
    )

    expanded_df["label_number"] = ""
    expanded_df["total_labels"] = ""
    expanded_df["ref"] = ""

    valid_df = expanded_df.loc[valid_reference].copy()

    valid_df["label_number"] = (
        valid_df.groupby(
            "base_ref",
            sort=False,
        ).cumcount()
        + 1
    )

    valid_df["total_labels"] = (
        valid_df.groupby(
            "base_ref",
            sort=False,
        )["base_ref"].transform("size")
    )

    expanded_df.loc[
        valid_reference,
        "label_number",
    ] = valid_df["label_number"]

    expanded_df.loc[
        valid_reference,
        "total_labels",
    ] = valid_df["total_labels"]

    # -----------------------------------------------------
    # Step 4: Generate references such as y050-01/08
    # -----------------------------------------------------
    def build_full_reference(row) -> str:
        base_reference = clean_value(
            row["base_ref"]
        )

        if not base_reference:
            return ""

        label_number = int(
            row["label_number"]
        )

        total_labels = int(
            row["total_labels"]
        )

        number_width = max(
            2,
            len(str(total_labels)),
        )

        return (
            f"{base_reference}-"
            f"{label_number:0{number_width}d}/"
            f"{total_labels:0{number_width}d}"
        )

    expanded_df.loc[
        valid_reference,
        "ref",
    ] = expanded_df.loc[
        valid_reference
    ].apply(
        build_full_reference,
        axis=1,
    )

    # -----------------------------------------------------
    # Step 5: Arrange final columns
    # -----------------------------------------------------
    expanded_df = expanded_df[
        [
            "source_row",
            "sp",
            "base_ref",
            "ref",
            "label_number",
            "total_labels",
            "label",
            "label_large",
            "label_medium",
            "label_small",
            "label_size",
            "label_visual_width",
        ]
    ]

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
    """
    Group expanded label records into configurable output rows.

    For group_size=4, the output contains:
        sp1, ref1, label1, ...
        sp2, ref2, label2, ...
        sp3, ref3, label3, ...
        sp4, ref4, label4, ...

    For group_size=6, the output continues through:
        sp6, ref6, label6, ...
    """

    if group_size < 1:
        raise ValueError(
            "Records per row must be at least 1."
        )

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

            if position < len(group):
                record = group.iloc[position]

                output_row[f"sp{number}"] = clean_value(
                    record.get("sp", "")
                )

                output_row[f"ref{number}"] = clean_value(
                    record.get("ref", "")
                )

                output_row[f"label{number}"] = clean_value(
                    record.get("label", "")
                )

                output_row[
                    f"label{number}_large"
                ] = clean_value(
                    record.get("label_large", "")
                )

                output_row[
                    f"label{number}_medium"
                ] = clean_value(
                    record.get("label_medium", "")
                )

                output_row[
                    f"label{number}_small"
                ] = clean_value(
                    record.get("label_small", "")
                )

                output_row[
                    f"label{number}_size"
                ] = clean_value(
                    record.get("label_size", "")
                )

                output_row[
                    f"label{number}_visual_width"
                ] = record.get(
                    "label_visual_width",
                    "",
                )

            else:
                output_row[f"sp{number}"] = ""
                output_row[f"ref{number}"] = ""
                output_row[f"label{number}"] = ""

                output_row[
                    f"label{number}_large"
                ] = ""

                output_row[
                    f"label{number}_medium"
                ] = ""

                output_row[
                    f"label{number}_small"
                ] = ""

                output_row[
                    f"label{number}_size"
                ] = ""

                output_row[
                    f"label{number}_visual_width"
                ] = ""

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
