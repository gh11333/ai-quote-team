import streamlit as st
import zipfile, os, io, re, math
import pandas as pd
from pypdf import PdfReader

st.set_page_config(layout="wide")
st.title("📂 견적 자동화 - 안정화 STEP 1 (폴더 규칙 고정)")

uploaded_zip = st.file_uploader("ZIP 업로드", type="zip")

def extract_number(text, keyword):
    m = re.search(rf'{keyword}.*?(\d+)', text)
    return int(m.group(1)) if m else None

def extract_up(text):
    m = re.search(r'(\d+)\s*(?:up|페이지|쪽)', text)
    return int(m.group(1)) if m else None

if uploaded_zip:
    result = {}
    folder_rules = {}

    with zipfile.ZipFile(uploaded_zip) as z:
        files = [f for f in z.namelist() if not f.endswith("/") and "__MACOSX" not in f]

        # 1️⃣ 상위폴더 목록
        top_folders = sorted(set(f.split("/")[0] for f in files))

        # 2️⃣ 폴더별 규칙 1회만 추출
        for folder in top_folders:
            rule_text = ""
            for f in files:
                if f.startswith(folder) and f.lower().endswith(".txt"):
                    rule_text += " " + f.lower()

            up = extract_up(rule_text) or 1
            vinyl = extract_number(rule_text, "비닐내지") or 0

            folder_rules[folder] = {
                "up": up,
                "vinyl": vinyl
            }

            result[folder] = {
                "흑백": 0,
                "컬러": 0,
                "비닐": vinyl,   # ✅ 딱 1번만
                "USB": 0,
                "바인더": 1
            }

        # 3️⃣ 파일 처리
        for f in files:
            folder = f.split("/")[0]
            fname = os.path.basename(f).lower()

            # USB
            if "usb" in fname:
                result[folder]["USB"] += 1
                continue

            # PDF만 페이지 계산
            if not f.lower().endswith(".pdf"):
                continue

            up = folder_rules[folder]["up"]

            try:
                with z.open(f) as fp:
                    raw = len(PdfReader(io.BytesIO(fp.read())).pages)
                    pages = math.ceil(raw / up)
            except:
                continue

            if "컬러" in fname or "color" in fname:
                result[folder]["컬러"] += pages
            else:
                result[folder]["흑백"] += pages

            # 📌 파일명에 비닐내지 있으면 추가 1
            if "비닐내지" in fname:
                result[folder]["비닐"] += 1

    st.subheader("📊 STEP 1 결과 (폴더 규칙 1회 적용)")
    df = pd.DataFrame.from_dict(result, orient="index")
    st.dataframe(df, use_container_width=True)
