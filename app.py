import streamlit as st
import zipfile, io, os
import pandas as pd
from pypdf import PdfReader

from agents.expression_agent import *
from agents.context_agent import detect_context
from agents.page_agent import calculate_pages
from agents.aggregate_agent import aggregate

st.title("📦 인쇄 자동 정산 Agent Team")

uploaded = st.file_uploader("ZIP 업로드", type="zip")

if uploaded:
    results = []

    with zipfile.ZipFile(uploaded) as z:
        for p in z.namelist():
            if not p.lower().endswith(".pdf"): continue

            text = p.lower()
            ctx = detect_context(text)
            if ctx.get("ignore"): continue

            with z.open(p) as f:
                raw_pages = len(PdfReader(io.BytesIO(f.read())).pages)

            pps = extract_pages_per_sheet(text)
            copies = extract_copies(text)
            mats = extract_materials(text)

            final_pages, detail = calculate_pages(raw_pages, pps, copies)

            results.append({
                "folder": p.split("/")[0],
                "file": os.path.basename(p),
                "print_type": ctx.get("print_type"),
                "pages": final_pages,
                "detail": detail,
                "materials": mats
            })

    summary = aggregate(results)

    st.subheader("📊 요약")
    st.dataframe(pd.DataFrame(summary).T)

    st.subheader("🔍 상세")
    st.dataframe(pd.DataFrame(results))
