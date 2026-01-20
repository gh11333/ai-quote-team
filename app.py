import streamlit as st
import zipfile
import tempfile
import os
import re
from pypdf import PdfReader

st.set_page_config(page_title="문서 수량 자동 계산", layout="wide")

# -----------------------------
# 유틸
# -----------------------------
def read_pdf_pages(path):
    try:
        return len(PdfReader(path).pages)
    except:
        return 0

def read_txt(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except:
        return ""

def extract_n_up(text):
    patterns = [
        r"한면\s*(\d+)\s*페이지",
        r"(\d+)\s*up",
        r"한면\s*(\d+)",
    ]
    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return int(m.group(1))
    return 1

def extract_vinyl(text):
    m = re.search(r"비닐내지.*?(\d+)", text)
    if m:
        return int(m.group(1))
    if "비닐내지" in text:
        return 1
    return 0

def is_usb(text):
    return any(k in text.lower() for k in ["usb", "전자파일"])

def is_page_excluded(text, pages):
    # 6페이지 이하 + TOC/표지 계열이면 제외
    if pages <= 6:
        if any(k in text.lower() for k in ["toc", "table of contents", "표지", "index"]):
            return True
    return False

# -----------------------------
# 메인
# -----------------------------
st.title("📦 문서 출력 수량 자동 계산 (최종 안정화)")

uploaded = st.file_uploader("ZIP 업로드", type=["zip"])

if uploaded:
    tmpdir = tempfile.mkdtemp()

    with zipfile.ZipFile(uploaded, "r") as z:
        z.extractall(tmpdir)

    result = {}

    for root, dirs, files in os.walk(tmpdir):
        for file in files:
            path = os.path.join(root, file)
            upper = os.path.relpath(root, tmpdir).split(os.sep)[0]

            if upper not in result:
                result[upper] = {
                    "흑백": 0,
                    "컬러": 0,
                    "비닐": 0,
                    "USB": 0
                }

            name = file.lower()

            # TXT 먼저
            if file.lower().endswith(".txt"):
                txt = read_txt(path)

                if is_usb(txt):
                    result[upper]["USB"] += 1
                    continue

                vinyl = extract_vinyl(txt)
                result[upper]["비닐"] += vinyl
                continue

            # PDF
            if file.lower().endswith(".pdf"):
                pages = read_pdf_pages(path)
                text = file.lower()

                # USB 제작이면 페이지 제외
                if is_usb(text):
                    result[upper]["USB"] += 1
                    continue

                # 비닐내지만 있는 파일
                if "비닐내지" in text and pages <= 1:
                    result[upper]["비닐"] += 1
                    continue

                # 제외 판단
                if is_page_excluded(text, pages):
                    continue

                n_up = extract_n_up(text)
                sheets = (pages + n_up - 1) // n_up

                # 컬러/흑백
                if "컬러" in text:
                    result[upper]["컬러"] += sheets
                else:
                    result[upper]["흑백"] += sheets

    st.subheader("📊 결과")

    rows = []
    for k, v in result.items():
        rows.append({
            "폴더": k,
            "흑백": v["흑백"],
            "컬러": v["컬러"],
            "비닐": v["비닐"],
            "USB": v["USB"]
        })

    st.dataframe(rows, use_container_width=True)
