import streamlit as st
import zipfile, os, io, re, math
import pandas as pd
from pypdf import PdfReader

st.set_page_config(layout="wide")
st.title("📂 견적 자동화 - 안정판")

uploaded_zip = st.file_uploader("ZIP 업로드", type="zip")

def extract_up(text):
    text = text.replace(" ", "").lower()
    m = re.search(r'(\d+)(?:up|페이지|쪽)', text)
    return int(m.group(1)) if m else 1

def extract_vinyl_count(text):
    m = re.search(r'비닐내지.*?(\d+)', text)
    return int(m.group(1)) if m else 1

if uploaded_zip:
    result = {}

    with zipfile.ZipFile(uploaded_zip) as z:
        files = [f for f in z.namelist() if not f.endswith("/") and "__MACOSX" not in f]

        # 폴더 초기화
        for f in files:
            top = f.split("/")[0]
            if top not in result:
                result[top] = {
                    "흑백":0,"컬러":0,"비닐":0,"USB":0,"바인더":1
                }

        for f in files:
            fname = os.path.basename(f).lower()
            top = f.split("/")[0]

            # 규칙 텍스트 수집
            rule_text = fname
            for p in files:
                if p.startswith(top) and p.lower().endswith(".txt"):
                    rule_text += " " + p.lower()

            # USB
            if "usb" in rule_text:
                result[top]["USB"] += 1
                continue

            is_pdf = f.lower().endswith(".pdf")

            # ▶ PDF면 무조건 페이지 계산
            if is_pdf:
                up = extract_up(rule_text)

                try:
                    with z.open(f) as fp:
                        raw = len(PdfReader(io.BytesIO(fp.read())).pages)
                        pages = math.ceil(raw / up)
                except:
                    continue

                if "컬러" in rule_text or "color" in rule_text:
                    result[top]["컬러"] += pages
                else:
                    result[top]["흑백"] += pages

                # PDF + 비닐내지 → 비닐 추가
                if "비닐내지" in rule_text:
                    result[top]["비닐"] += extract_vinyl_count(rule_text)

            # ▶ TXT는 페이지 계산 ❌, 자재만
            else:
                if "비닐내지" in rule_text:
                    result[top]["비닐"] += extract_vinyl_count(rule_text)

    st.subheader("📊 최종 집계")
    df = pd.DataFrame.from_dict(result, orient="index")
    st.dataframe(df, use_container_width=True)
