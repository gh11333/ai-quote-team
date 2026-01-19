import streamlit as st
import zipfile, io, os
import pandas as pd
from pypdf import PdfReader

from agents.expression_agent import *
from agents.context_agent import detect_context
from agents.page_agent import calculate_pages
from agents.material_agent import extract_folder_materials
from agents.aggregate_agent import aggregate

st.set_page_config(layout="wide")
st.title("📦 인쇄 자동 정산 Agent Team (안정판)")

uploaded = st.file_uploader("ZIP 업로드", type="zip")

if uploaded:
    results = []
    folder_files, folder_txts = {}, {}

    with zipfile.ZipFile(uploaded) as z:
        for p in z.namelist():
            if p.startswith("__MACOSX") or p.endswith("/"):
                continue

            folder = p.split("/")[0]
            folder_files.setdefault(folder, [])
            folder_txts.setdefault(folder, [])

            if p.lower().endswith(".txt"):
                folder_txts[folder].append(p)
            else:
                folder_files[folder].append(p)

        # 자재 먼저 계산
        folder_materials = {}
        for folder in folder_files:
            folder_materials[folder] = extract_folder_materials(
                folder,
                folder_files[folder],
                folder_txts.get(folder, [])
            )

        # 인쇄 계산
        for p in z.namelist():
            if not p.lower().endswith(".pdf"):
                continue

            folder = p.split("/")[0]
            filename = os.path.basename(p)
            text = p.lower()

            ctx = detect_context(text)
            if ctx.get("ignore"):
                continue

            with z.open(p) as f:
                raw_pages = len(PdfReader(io.BytesIO(f.read())).pages)

            pps = extract_pages_per_sheet(text)
            copies = extract_copies(text)
            final_pages, detail = calculate_pages(raw_pages, pps, copies)

            results.append({
                "folder": folder,
                "file": filename,
                "print_type": ctx["print_type"],
                "pages": final_pages,
                "detail": detail
            })

    summary = aggregate(results, folder_materials)

    st.subheader("📊 요약")
    st.dataframe(pd.DataFrame(summary).T)

    st.subheader("🔍 상세")
    st.dataframe(pd.DataFrame(results))
