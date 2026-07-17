import re
from datetime import datetime

import streamlit as st

from combiner_engine import (
    InputFile,
    collect_input_files,
    combine_files,
    detect_header_rows,
    preview_rows,
)

st.set_page_config(page_title="File Combiner", page_icon="📎", layout="centered")

st.title("📎 Split File Combiner")
st.caption(
    "Upload split files that share the same header/template — a zip full of "
    "them, or several files directly — and get one combined workbook back."
)

uploads = st.file_uploader(
    "Upload files",
    type=["zip", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
    help="Zips are unpacked automatically. You can also drop individual .xlsx/.xls/.csv files directly.",
)

if uploads:
    upload_tuples = [(uf.name, uf.getvalue()) for uf in uploads]
    try:
        files: list[InputFile] = collect_input_files(upload_tuples)
    except Exception as e:
        st.error(f"Couldn't read the upload: {e}")
        files = []

    if not files:
        st.warning("No supported files (.xlsx, .xls, .csv) found in what you uploaded.")
    else:
        st.write(f"**Found {len(files)} file(s) to combine:**")
        st.write(", ".join(f.name for f in files))

        auto_detected = detect_header_rows(files) if len(files) >= 2 else 1

        col1, col2 = st.columns([1, 2])
        with col1:
            header_rows = st.number_input(
                "Header rows to keep once",
                min_value=1,
                max_value=50,
                value=auto_detected,
                help="Number of leading rows (titles/column headers/notes) that are "
                "identical across files. Auto-detected from your files, but you can override it.",
            )
        with col2:
            if len(files) >= 2:
                st.caption(f"Auto-detected: **{auto_detected}** header row(s) common to all files.")
            else:
                st.caption("Only one file uploaded — set header rows manually if needed.")

        with st.expander("Preview first file's rows (to double check header count)"):
            preview = preview_rows(files[0], n=min(header_rows + 3, 15))
            for i, row in enumerate(preview, start=1):
                marker = "🔷 header" if i <= header_rows else "⬜ data"
                st.text(f"{marker}  row {i}: {row}")

        output_filename = st.text_input("Output filename", value="COMBINED.xlsx")

        if st.button("Combine", type="primary"):
            with st.spinner("Combining..."):
                try:
                    result = combine_files(
                        files,
                        header_rows=int(header_rows),
                        output_filename=output_filename,
                    )
                except Exception as e:
                    st.error(f"❌ {e}")
                    result = None

            if result:
                st.success(f"✅ Combined {len(result.files_processed)} file(s) — {result.total_rows} total data rows.")

                cols = st.columns(2)
                with cols[0]:
                    st.write("**Files combined:**")
                    for name in result.files_processed:
                        st.write(f"- {name}: {result.rows_per_file.get(name, 0)} rows")
                with cols[1]:
                    if result.files_skipped_empty:
                        st.write("**Skipped (no data rows):**")
                        for name in result.files_skipped_empty:
                            st.write(f"- {name}")

                for w in result.warnings:
                    st.warning(w)

                st.download_button(
                    label=f"⬇️ Download {result.output_filename}",
                    data=result.output_bytes,
                    file_name=result.output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{datetime.now().timestamp()}",
                )

st.divider()
with st.expander("ℹ️ How this works"):
    st.markdown(
        """
1. Upload a zip of split files, or select several files directly —
   `.xlsx`, `.xls`, and `.csv` are all supported, and can be mixed.
2. The app compares the leading rows across all your files and
   auto-detects how many rows make up the shared header/template
   (title rows, column headers, notes, etc). You can override this
   if it guesses wrong.
3. It keeps that header once, then appends every data row from every
   file underneath, in filename order (`1`, `2`, `3`, ... sorted
   naturally, not alphabetically).
4. Files that turn out to be empty templates (header only, no data)
   are automatically skipped and reported separately.
5. Known export quirks — like Shopee's broken `activePane` XML
   attribute — are patched automatically and don't require any
   special handling from you.
        """
    )
