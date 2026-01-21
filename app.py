import streamlit as st
import zipfile
import os
import io
import re
import pandas as pd
from collections import defaultdict
from pypdf import PdfReader

st.set_page_config(page_title="STEP1 흑백·비닐 안정화", layout="wide")
st.title("📦 STEP 1 · 흑백 / 비닐 계산 엔진")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

# ---------- 유틸 함수 ----------

def is_txt(name):
    return name.lower().endswith(".txt")

def is_pdf(name):
    return name.lower().endswith(".pdf")

def has_vinyl_keyword(text):
    return "비닐" in text

def has_usb(text):
    t = text.lower()
    return ("usb" in t) or ("cd" in t)

def extract_vinyl_from_txt(text):
    """
    TXT에서만 숫자 허용
    '비닐내지 5장' → 5
    숫자 없고 비닐만 있으면 → 1
    """
    m = re.search(r'비닐.*?(\d+)\s*장', text)
    if m:
        return int(m.group(1))
    if "비닐" in text:
        return 1
    return 0

def count_pdf_pages(file_bytes):
    reader = PdfReader(io.BytesIO(file_bytes))
    return len(reader.pages)

# ---------- 메인 로직 ----------

if uploaded_zip:
    result = defaultdict(lambda: {"흑백": 0, "비닐": 0, "파일수": 0})

    with zipfile.ZipFile(uploaded_zip, "r") as z:
        for raw_path in z.namelist():
            if raw_path.endswith("/") or "__MACOSX" in raw_path:
                continue

            # 경로 정규화 (윈도우/맥 호환)
            path = raw_path.replace("\\", "/")
            filename = os.path.basename(path)
            folder = path.split("/")[0]
            lower_name = filename.lower()

            result[folder]["파일수"] += 1

            # -------------------------
            # 1️⃣ TXT 처리 (페이지 X)
            # -------------------------
            if is_txt(filename):
                try:
                    text = z.read(raw_path).decode("utf-8", errors="ignore")
                except:
                    text = ""

                vinyl = extract_vinyl_from_txt(text)
                result[folder]["비닐"] += vinyl
                continue

            # -------------------------
            # 2️⃣ USB 관련 PDF
            # 페이지 X / 비닐 +1
            # -------------------------
            if is_pdf(filename) and has_usb(lower_name):
                result[folder]["비닐"] += 1
                continue

            # -------------------------
            # 3️⃣ PDF + 비닐 키워드
            # 페이지 O / 비닐 +1
            # -------------------------
            if is_pdf(filename) and has_vinyl_keyword(lower_name):
                try:
                    pages = count_pdf_pages(z.read(raw_path))
                except:
                    pages = 0

                result[folder]["흑백"] += pages
                result[folder]["비닐"] += 1
                continue

            # -------------------------
            # 4️⃣ 일반 PDF
            # 페이지 O
            # -------------------------
            if is_pdf(filename):
                try:
                    pages = count_pdf_pages(z.read(raw_path))
                except:
                    pages = 0

                result[folder]["흑백"] += pages
                continue

    # ---------- 결과 출력 ----------
    st.subheader("📊 STEP 1 결과")

    df = (
        pd.DataFrame.from_dict(result, orient="index")
        .reset_index()
        .rename(columns={"index": "상위폴더"})
    )

    st.dataframe(df, use_container_width=True)
