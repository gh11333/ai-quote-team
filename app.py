import streamlit as st
import zipfile, os, io, re, math
from pypdf import PdfReader
import pandas as pd

# ===============================
# 유틸
# ===============================

def extract_pages_per_sheet(text):
    patterns = [
        r'(\d+)\s*up',
        r'한면\s*(\d+)\s*페이지',
        r'한면(\d+)페이지',
        r'(\d+)\s*페이지\s*출력',
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return int(m.group(1))
    return None

def extract_copies(text):
    m = re.search(r'(\d+)\s*(부|장)', text)
    return int(m.group(1)) if m else None

def is_color(text):
    return any(k in text for k in ['컬러', '칼라', 'color'])

def folder_path_text(folder):
    parts = []
    while folder:
        parts.append(os.path.basename(folder))
        folder = os.path.dirname(folder)
    return " ".join(parts)

def top_level_folder(path):
    return path.split('/')[0] if '/' in path else 'ROOT'

# ===============================
# Streamlit
# ===============================

st.set_page_config(layout="wide")
st.title("ZIP 인쇄 자동 정산기 (정확도 최우선 v1.1)")

uploaded_zip = st.file_uploader("ZIP 업로드", type="zip")

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as z:
        files = [f for f in z.namelist() if not f.endswith('/')]

        # txt 수집
        folder_txt = {}
        for f in files:
            if f.lower().endswith('.txt'):
                folder = os.path.dirname(f)
                with z.open(f) as tf:
                    folder_txt.setdefault(folder, "")
                    folder_txt[folder] += " " + tf.read().decode('utf-8', errors='ignore').lower()

        def collect_txt(folder):
            texts = []
            while True:
                if folder in folder_txt:
                    texts.append(folder_txt[folder])
                if not folder:
                    break
                folder = os.path.dirname(folder)
            return " ".join(texts)

        results = []

        for f in files:
            if not f.lower().endswith('.pdf'):
                continue

            folder = os.path.dirname(f)
            filename = os.path.basename(f)
            top_folder = top_level_folder(f)

            context = " ".join([
                filename.lower(),
                folder_path_text(folder).lower(),
                collect_txt(folder)
            ])

            pps = extract_pages_per_sheet(context) or 1
            copies = extract_copies(context) or 1
            color = "컬러" if is_color(context) else "흑백"

            with z.open(f) as pf:
                reader = PdfReader(io.BytesIO(pf.read()))
                raw = len(reader.pages)

            final_pages = math.ceil(raw / pps) * copies

            results.append({
                "폴더": top_folder,
                "파일명": filename,
                "구분": color,
                "원본페이지": raw,
                "한면": pps,
                "부수": copies,
                "최종페이지": final_pages
            })

    df = pd.DataFrame(results)

    summary = (
        df.groupby(["폴더", "구분"])["최종페이지"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    st.subheader("📊 폴더별 요약")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📄 상세")
    st.dataframe(df, use_container_width=True)
