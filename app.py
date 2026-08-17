"""Streamlit interface for the HDFC statement converter."""

from __future__ import annotations

import hashlib
from pathlib import Path

import streamlit as st

from hdfc_converter import StatementError, convert_statement


st.set_page_config(page_title="HDFC Statement to Excel", page_icon="📊", layout="centered")
st.markdown("""
<style>
.block-container {max-width: 850px; padding-top: 2.5rem;}
[data-testid="stFileUploaderDropzone"] {border: 2px dashed #2c6e91; background: #f6fafc;}
.small-note {color: #58636d; font-size: .9rem;}
</style>
""", unsafe_allow_html=True)

st.title("HDFC Statement → Excel")
st.write("Select an HDFC Bank statement PDF, optionally enter its password, and download a formatted Excel file.")

with st.form("converter"):
    uploaded = st.file_uploader("Select bank statement", type=["pdf"], help="Your file is processed only by this app and is not saved by the converter.")
    password = st.text_input("PDF password (only if protected)", type="password", placeholder="Leave blank for an unprotected PDF")
    submitted = st.form_submit_button("Convert to Excel", type="primary", use_container_width=True)

if submitted:
    if uploaded is None:
        st.warning("Please select a PDF statement first.")
    else:
        file_bytes = uploaded.getvalue()
        fingerprint = hashlib.sha256(file_bytes + password.encode()).hexdigest()
        try:
            with st.spinner("Reading transactions and preparing Excel…"):
                result = convert_statement(file_bytes, password)
            st.session_state["result"] = result
            st.session_state["fingerprint"] = fingerprint
            st.session_state["filename"] = f"{Path(uploaded.name).stem}_transactions.xlsx"
        except StatementError as exc:
            st.session_state.pop("result", None)
            st.error(str(exc))
        except Exception:
            st.session_state.pop("result", None)
            st.error("Conversion failed unexpectedly. Please confirm this is an HDFC statement PDF and try again.")

if "result" in st.session_state:
    result = st.session_state["result"]
    st.success(f"Excel is ready — {len(result.transactions)} transactions extracted.")
    if result.warnings:
        for warning in result.warnings:
            st.warning(warning)
    st.download_button(
        "Download Excel file",
        data=result.workbook,
        file_name=st.session_state["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    with st.expander("Conversion details"):
        cols = st.columns(2)
        cols[0].metric("Transactions", len(result.transactions))
        cols[1].metric("Rows to review", sum(bool(row.warnings) for row in result.transactions))
        if result.metadata:
            st.write("**Statement information found**")
            st.json(result.metadata)
        st.caption("OCR-based PDFs can contain recognition mistakes. Rows needing attention are marked in the Excel “Review Notes” column.")

st.divider()
st.markdown('<p class="small-note">Privacy: conversion happens in memory. Passwords and uploaded PDFs are not written to disk by the application.</p>', unsafe_allow_html=True)
