"""Streamlit UI for converting scanned bank statements to Excel."""
from __future__ import annotations

import hashlib

import streamlit as st

from statement_converter import (
    StatementError,
    create_workbook,
    ocr_document,
    open_document,
    parse_transactions,
)

st.set_page_config(page_title="Bank Statement to Excel", page_icon="🏦", layout="centered")
st.title("🏦 Scanned Bank Statement → Excel")
st.caption("Runs OCR on a PDF or image, detects transaction rows, and creates an editable .xlsx workbook.")

with st.expander("Privacy and accuracy", expanded=False):
    st.markdown(
        "- The app processes the upload in memory; it does not save the statement or password.\n"
        "- The password is used only to open the PDF. It is never included in Excel.\n"
        "- OCR is not perfect. Always compare dates, amounts, and balances with the original."
    )

uploaded = st.file_uploader(
    "Select a scanned statement",
    type=["pdf", "png", "jpg", "jpeg", "tif", "tiff"],
    help="PDF (including password-protected PDFs), PNG, JPG, or multi-page TIFF",
)

if uploaded:
    file_bytes = uploaded.getvalue()
    password = st.text_input(
        "PDF password (leave blank if not protected)",
        type="password",
        help="The password remains in this browser session and is only passed to the PDF reader.",
    )
    try:
        preview_doc = open_document(file_bytes, uploaded.name, password)
        page_count = preview_doc.page_count
        preview_doc.close()
        st.success(f"File opened successfully — {page_count} page{'s' if page_count != 1 else ''} found.")

        col1, col2 = st.columns(2)
        with col1:
            first_page = st.number_input("First page", 1, page_count, 1)
        with col2:
            last_page = st.number_input("Last page", int(first_page), page_count, page_count)
        quality = st.select_slider(
            "Scan quality",
            options=["Faster", "Balanced", "High accuracy"],
            value="Balanced",
            help="Higher accuracy uses more memory and takes longer.",
        )
        dpi_by_quality = {"Faster": 180, "Balanced": 250, "High accuracy": 300}

        if st.button("Convert to Excel", type="primary", use_container_width=True):
            progress = st.progress(0, text="Starting OCR…")
            try:
                doc = open_document(file_bytes, uploaded.name, password)
                pages = ocr_document(
                    doc,
                    int(first_page),
                    int(last_page),
                    dpi=dpi_by_quality[quality],
                    progress=lambda done, total: progress.progress(
                        done / total, text=f"Reading page {done} of {total}…"
                    ),
                )
                doc.close()
                progress.progress(1.0, text="Building Excel workbook…")
                transactions = parse_transactions(pages)
                excel = create_workbook(transactions, pages, uploaded.name)
                # Key prevents an old result being mistaken for a newly selected file.
                st.session_state["result"] = {
                    "key": hashlib.sha256(file_bytes).hexdigest(),
                    "excel": excel,
                    "count": len(transactions),
                    "pages": len(pages),
                    "name": uploaded.name.rsplit(".", 1)[0] + ".xlsx",
                }
                progress.empty()
            except StatementError as exc:
                progress.empty()
                st.error(str(exc))
            except Exception as exc:
                progress.empty()
                st.error(f"Conversion failed: {exc}")

        result = st.session_state.get("result")
        if result and result["key"] == hashlib.sha256(file_bytes).hexdigest():
            if result["count"]:
                st.success(f"Done — found {result['count']} transaction rows on {result['pages']} processed pages.")
            else:
                st.warning("OCR completed, but no transaction rows were detected. The OCR Text sheet is included for review.")
            st.download_button(
                "⬇️ Download Excel file",
                data=result["excel"],
                file_name=result["name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
    except StatementError as exc:
        st.warning(str(exc))
    except Exception as exc:
        st.error(f"Could not inspect the file: {exc}")
else:
    st.info("Choose a statement above to begin.")

st.divider()
st.caption("Designed for common HDFC statement layouts. The workbook also includes all raw OCR text for checking or manual correction.")
