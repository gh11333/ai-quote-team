import streamlit as st
import zipfile, os, io, re, math
import pandas as pd
from pypdf import PdfReader

# =====================
# 유틸 함수
# =====================

def top_level_folder(path):
    return path.split('/')[0] if '/' in path else 'ROOT'

def folder_path_text(folder):
    parts = []
    while folder:
        parts.append(os.path.basename(folder))
        folder = os.path.dirname(folder)
    return " ".join(parts)

def extract_first_pages_per_sheet(text):
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
    return int(m.group(1)) if m else 1

def is_color(context):
    return any(k in context for k in ['컬러', '칼라', 'color'])

def is_page_excluded(context):
    exclude_keywords = [
        'usb', 'cd', '제작',
        'binder', 'face', 'spine',
        'toc', '목차'
    ]
    return any(k in context for k in exclude_keywords)

def has_vinyl_pdf(filename):
    return '비닐내지' in filename

# =====================
# Streamlit UI
# =====================

st.set_page_config(layout="wide")
st.title("ZIP 인쇄 페이지 정산기 (1단계 안정판)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    results = []

    with zipfile.ZipFile(uploaded_zip) as z:
        files = [f for f in z.namelist() if not f.endswith('/')]

        # TXT 수집
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

        for f in files:
            if not f.lower().endswith('.pdf'):
                continue

            folder = os.path.dirname(f)
            filename = os.path.basename(f)
            top_folder = top_level_folder(f)

            # context 구성
            context = " ".join([
                filename.lower(),
                folder_path_text(folder).lower(),
                collect_txt(folder)
            ])

            # 페이지 제외 대상
            if is_page_excluded(context):
                continue

            # PDF 페이지 수
            with z.open(f) as pf:
                reader = PdfReader(io.BytesIO(pf.read()))
                raw_pages = len(reader.pages)

            # 한면 n페이지 (가장 먼저 발견된 것 1개)
            pps = (
                extract_first_pages_per_sheet(filename.lower())
                or extract_first_pages_per_sheet(folder_path_text(folder).lower())
                or extract_first_pages_per_sheet(collect_txt(folder))
                or 1
            )

            copies = extract_copies(context)
            final_pages = math.ceil(raw_pages / pps) * copies

            results.append({
                "폴더": top_folder,
                "파일명": filename,
                "구분": "컬러" if is_color(context) else "흑백",
                "원본페이지": raw_pages,
                "한면": pps,
                "부수": copies,
                "최종페이지": final_pages,
                "비닐": 1 if has_vinyl_pdf(filename) else 0
            })

    df = pd.DataFrame(results)

    summary = (
        df.groupby("폴더")[["최종페이지", "비닐"]]
        .sum()
        .reset_index()
    )

    st.subheader("📊 폴더별 요약")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📄 상세 내역")
    st.dataframe(df, use_container_width=True)
