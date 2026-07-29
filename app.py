import streamlit as st

from converter import convert_excel


st.set_page_config(
    page_title="Excel Mail Merge Converter",
    page_icon="📄",
    layout="wide",
)

st.title("Excel Mail Merge Converter")

st.write(
    "Convert the original Excel file into a four-column "
    "Word Mail Merge recipient format."
)

# ---------------------------------------------------------
# Step 1: Upload
# ---------------------------------------------------------

st.header("Step 1 — Upload Excel")

uploaded_file = st.file_uploader(
    "Choose the original Excel file",
    type=["xlsx"],
    help="The Excel file must contain sp, ref, and label columns.",
)

with st.expander("View required Excel format"):
    st.markdown(
        """
        | sp | ref | label |
        |---|---|---|
        | 江素 | y007 | 江氏歷代祖先 |
        | Sabrina Lam 林玉瑩 | y015 | 多生父母師長..., 歷劫怨親債主... |
        """
    )

# ---------------------------------------------------------
# Step 2: Settings
# ---------------------------------------------------------

st.header("Step 2 — Conversion Settings")

column1, column2, column3 = st.columns(3)

with column1:
    records_per_row = st.number_input(
        "Labels per output row",
        min_value=1,
        max_value=10,
        value=4,
        step=1,
    )

with column2:
    large_width = st.number_input(
        "Large-font maximum width",
        min_value=1,
        max_value=100,
        value=16,
        step=1,
    )

with column3:
    medium_width = st.number_input(
        "Medium-font maximum width",
        min_value=1,
        max_value=150,
        value=28,
        step=1,
    )

st.caption(
    "Labels up to the large-width limit use the large Word field. "
    "Labels up to the medium-width limit use the medium Word field. "
    "Longer labels use the small Word field."
)

# ---------------------------------------------------------
# Step 3: Convert
# ---------------------------------------------------------

st.header("Step 3 — Convert")

convert_clicked = st.button(
    "Convert Excel",
    type="primary",
    disabled=uploaded_file is None,
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

        output_buffer = result["output"]
        original_df = result["original"]
        expanded_df = result["expanded"]
        mail_merge_df = result["mail_merge"]
        validation_df = result["validation"]

        st.success("Conversion completed successfully.")

        # Save the result in session state so it remains
        # available after Streamlit reruns the page.
        st.session_state["conversion_result"] = {
            "output": output_buffer.getvalue(),
            "original": original_df,
            "expanded": expanded_df,
            "mail_merge": mail_merge_df,
            "validation": validation_df,
        }

    except Exception as error:
        st.error(f"Conversion failed: {error}")
        st.exception(error)

# ---------------------------------------------------------
# Step 4: Preview and download
# ---------------------------------------------------------

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

    warning_count = len(
        result["validation"][
            result["validation"]["severity"] == "Warning"
        ]
    )

    metric4.metric(
        "Warnings",
        warning_count,
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

        for column in result["mail_merge"].columns:
            if (
                column.startswith("sp")
                or column.startswith("ref")
                or (
                    column.startswith("label")
                    and "_" not in column
                )
            ):
                preview_columns.append(column)

        st.dataframe(
            result["mail_merge"][preview_columns],
            use_container_width=True,
            hide_index=True,
        )

    with expanded_tab:
        st.dataframe(
            result["expanded"][
                [
                    "sp",
                    "base_ref",
                    "ref",
                    "label",
                    "label_size",
                    "label_visual_width",
                ]
            ],
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
