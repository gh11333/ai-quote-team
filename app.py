import streamlit as st
import zipfile, os, io, re, math
import pandas as pd
from pypdf import PdfReader

st.set_page_config(layout="wide")
st.title("📂 견적 자동화 – 단일 안정 엔진")

uploaded_zip = st.file_uploader("ZIP 업로드", type="zip")

def extract_up(text):
    text = text.replace(" ", "").lower()
    m = re.search(r'(\d+)(?:up|페이지|쪽)', text)
    return int(m.group(1)) if m else None

if uploaded_zip:
    result = {}

    with zipfile.ZipFile(uploaded_zip) as z:
        files = [f for f in z.namelist() if not f.endswith("/") and "__MACOSX" not in f]

        for f in files:
            parts = f.split("/")
            top = parts[0]

            if top not in result:
                result[top] = {
                    "흑백":0,"컬러":0,"비닐":0,"USB":0,"바인더":1
                }

        for f in files:
            parts = f.split("/")
            top = parts[0]
            fname = os.path.basename(f).lower()

            # txt 수집
            rules = []
            for p in files:
                if p.startswith(top) and p.lower().endswith(".txt"):
                    rules.append(p.lower())

            rule_text = " ".join(rules) + " " + fname

            # USB
            if "usb" in rule_text:
                result[top]["USB"] += 1
                continue

            # 비닐내지
            if "비닐내지" in rule_text:
                result[top]["비닐"] += 1
                continue

            # PDF만 페이지 계산
            if not f.lower().endswith(".pdf"):
                continue

            up = extract_up(rule_text) or 1

            try:
                with z.open(f) as fp:
                    pages = len(PdfReader(io.BytesIO(fp.read())).pages)
                    pages = math.ceil(pages / up)
            except:
                continue

            if "컬러" in rule_text or "color" in rule_text:
                result[top]["컬러"] += pages
            else:
                result[top]["흑백"] += pages

    df = pd.DataFrame.from_dict(result, orient="index")
    st.dataframe(df, use_container_width=True)
