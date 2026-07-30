import streamlit as st

from converter import convert_excel


st.set_page_config(
    page_title="Excel Mail Merge Converter",
    page_icon="📄",
    layout="wide",
)

st.title("Excel Mail Merge Converter")

st.write(
    "Upload an Excel file containing the columns "
    "`sp`, `ref`, and `label`."
)

# =========================================================
# Step 1: Upload
# =========================================================

st.header("Step 1 — Upload Excel")

uploaded_file = st.file_uploader(
    "Choose the source Excel file",
    type=["xlsx"],
    help="The Excel file must contain sp, ref, and label columns.",
)

with st.expander("View required source format"):
    st.markdown(
        """
| sp | ref | label |
|---|---|---|
| name001 | y007 | label001 |
| name002 | y015 | label002, label003 |
| name003 | y020 | label004, label005, label006 |
"""
    )

# =========================================================
# Step 2: Settings
# =========================================================

st.header("Step 2 — Conversion Settings")

settings_col1, settings_col2, settings_col3 = st.columns(3)

with settings_col1:
    records_per_row = st.number_input(
        "Records on each output row",
        min_value=1,
        max_value=20,
        value=4,
        step=1,
        help=(
            "For example, selecting 4 creates "
            "sp1/ref1/label1 through sp4/ref4/label4."
        ),
    )

with settings_col2:
    large_width = st.number_input(
        "Large-font maximum visual width",
        min_value=1,
        max_value=100,
        value=16,
        step=1,
    )

with settings_col3:
    medium_width = st.number_input(
        "Medium-font maximum visual width",
        min_value=2,
        max_value=150,
        value=28,
        step=1,
    )

if medium_width <= large_width:
    st.warning(
        "Medium-font maximum width must be greater "
        "than large-font maximum width."
    )

st.info(
    f"The output will contain {int(records_per_row)} records "
    "on each Excel row."
)

# Show the dynamic field names.
with st.expander("View generated mail-merge columns"):
    generated_columns = []

    for number in range(1, int(records_per_row) + 1):
        generated_columns.extend(
            [
                f"sp{number}",
                f"ref{number}",
                f"label{number}",
                f"label{number}_large",
                f"label{number}_medium",
                f"label{number}_small",
                f"label{number}_size",
            ]
        )

    st.code("\n".join(generated_columns))

# =========================================================
# Step 3: Convert
# =========================================================

st.header("Step 3 — Convert")

can_convert = (
    uploaded_file is not None
    and medium_width > large_width
)

convert_clicked = st.button(
    "Convert Excel",
    type="primary",
    disabled=not can_convert,
)

if uploaded_file is None:
    st.info("Upload an Excel file before converting.")

if convert_clicked:
    try:
        with st.spinner("Converting the Excel file..."):
            result = convert_excel(
                source_file=uploaded_file,
                group_size=int(records_per_row),
                large_width=int(large_width),
                medium_width=int(medium_width),
            )

        st.session_state["conversion_result"] = {
            "output": result["output"].getvalue(),
            "original": result["original"],
            "expanded": result["expanded"],
            "mail_merge": result["mail_merge"],
            "validation": result["validation"],
            "records_per_row": int(records_per_row),
        }

        st.success("Conversion completed successfully.")

    except Exception as error:
        st.error(f"Conversion failed: {error}")
        st.exception(error)

# =========================================================
# Step 4: Preview
# =========================================================

if "conversion_result" in st.session_state:
    result = st.session_state["conversion_result"]

    st.header("Step 4 — Review Results")

    metric1, metric2, metric3, metric4 = st.columns(4)

    metric1.metric(
        "Original rows",
        len(result["original"]),
    )

    metric2.metric(
        "Expanded labels",
        len(result["expanded"]),
    )

    metric3.metric(
        "Mail merge rows",
        len(result["mail_merge"]),
    )

    metric4.metric(
        "Records per row",
        result["records_per_row"],
    )

    preview_tab, expanded_tab, validation_tab = st.tabs(
        [
            "Mail Merge Preview",
            "Expanded Labels",
            "Validation",
        ]
    )

    with preview_tab:
        preview_columns = []

        for number in range(
            1,
            result["records_per_row"] + 1,
        ):
            preview_columns.extend(
                [
                    f"sp{number}",
                    f"ref{number}",
                    f"label{number}",
                ]
            )

        existing_preview_columns = [
            column
            for column in preview_columns
            if column in result["mail_merge"].columns
        ]

        st.dataframe(
            result["mail_merge"][existing_preview_columns],
            use_container_width=True,
            hide_index=True,
        )

    with expanded_tab:
        expanded_columns = [
            "source_row",
            "sp",
            "base_ref",
            "ref",
            "label_number",
            "total_labels",
            "label",
            "label_size",
            "label_visual_width",
        ]

        existing_expanded_columns = [
            column
            for column in expanded_columns
            if column in result["expanded"].columns
        ]

        st.dataframe(
            result["expanded"][existing_expanded_columns],
            use_container_width=True,
            hide_index=True,
        )

    with validation_tab:
        if result["validation"].empty:
            st.success("No validation warnings.")
        else:
            st.dataframe(
                result["validation"],
                use_container_width=True,
                hide_index=True,
            )

    # =====================================================
    # Step 5: Download
    # =====================================================

    st.header("Step 5 — Download")

    st.download_button(
        label="Download Converted Excel",
        data=result["output"],
        file_name="mail_merge_output.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
