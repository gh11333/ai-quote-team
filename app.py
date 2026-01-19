import streamlit as st
import zipfile, os, io, re, math
import pandas as pd
from pypdf import PdfReader

st.set_page_config(layout="wide")
st.title("📂 견적 자동화 - 디버그 확인용")

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
            top = f.split("/")[0]
            if top not in result:
                result[top] = {
                    "흑백":0,"컬러":0,"비닐":0,"USB":0,"바인더":1
                }

        st.write("===== 📁 파일 목록 =====")
        for f in files:
            st.write(f)

        for f in files:
            fname = os.path.basename(f).lower()
            top = f.split("/")[0]

            rules = []
            for p in files:
                if p.startswith(top) and p.lower().endswith(".txt"):
                    rules.append(p.lower())

            rule_text = " ".join(rules) + " " + fname

            st.write("──────────────")
            st.write("📁 상위폴더:", top)
            st.write("📄 파일명:", fname)
            st.write("📝 규칙텍스트:", rule_text)

            if "usb" in rule_text:
                st.write("👉 USB로 판단 → 페이지 계산 안 함")
                result[top]["USB"] += 1
                continue

            if "비닐내지" in rule_text:
                st.write("👉 비닐내지 → 페이지 계산 안 함 / 비닐 +1")
                result[top]["비닐"] += 1
                continue

            if not f.lower().endswith(".pdf"):
                st.write("👉 PDF 아님 → 무시")
                continue

            up = extract_up(rule_text) or 1
            st.write("➗ 한면 n페이지:", up)

            try:
                with z.open(f) as fp:
                    pages_raw = len(PdfReader(io.BytesIO(fp.read())).pages)
                    pages = math.ceil(pages_raw / up)
            except:
                st.write("❌ PDF 읽기 실패")
                continue

            st.write("📄 원본 페이지:", pages_raw)
            st.write("📄 계산 후 페이지:", pages)

            if "컬러" in rule_text or "color" in rule_text:
                result[top]["컬러"] += pages
                st.write("🎨 컬러로 합산")
            else:
                result[top]["흑백"] += pages
                st.write("🖤 흑백으로 합산")

    st.write("===== 📊 최종 집계 =====")
    df = pd.DataFrame.from_dict(result, orient="index")
    st.dataframe(df, use_container_width=True)
