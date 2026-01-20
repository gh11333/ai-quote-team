import streamlit as st
import os
import zipfile
import tempfile
import re

st.set_page_config(page_title="출력물 계산기", layout="wide")
st.title("📦 ZIP 업로드 출력물 계산기")

# -----------------------------
# 유틸 함수
# -----------------------------

def safe_int(x, default=0):
    try:
        return int(x)
    except:
        return default


def extract_vinyl_count(text):
    if not text:
        return 0

    # "비닐내지 5장"
    m = re.search(r"비닐내지[^0-9]*(\d+)\s*장", text)
    if m:
        return safe_int(m.group(1))

    # 숫자 없는 비닐내지 → 1
    if "비닐내지" in text:
        return 1

    return 0


def extract_print_mode(text):
    if "양면" in text:
        return "double"
    if "단면" in text:
        return "single"
    return "single"


def extract_nup(text):
    m = re.search(r"1면에\s*(\d+)\s*페이지", text)
    if m:
        return safe_int(m.group(1), 1)
    return 1


def read_pdf_pages_safe(path):
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(path)
        return len(reader.pages)
    except:
        return 0


# -----------------------------
# 파일 처리
# -----------------------------

def process_file(path):
    filename = os.path.basename(path)

    result = {"bw": 0, "vinyl": 0}

    # TXT
    if filename.lower().endswith(".txt"):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
        except:
            content = ""

        result["vinyl"] += extract_vinyl_count(content)
        return result

    # PDF
    if filename.lower().endswith(".pdf"):
        pages = read_pdf_pages_safe(path)

        mode = extract_print_mode(filename)
        nup = extract_nup(filename)

        if mode == "double":
            pages = (pages + 1) // 2

        pages = (pages + nup - 1) // nup
        result["bw"] += pages

        # PDF + 비닐내지
        if "비닐내지" in filename:
            result["vinyl"] += 1

    return result


def process_folder(folder_path):
    bw = 0
    vinyl = 0

    for root, _, files in os.walk(folder_path):
        for f in files:
            r = process_file(os.path.join(root, f))
            bw += r["bw"]
            vinyl += r["vinyl"]

    return bw, vinyl


# -----------------------------
# UI (ZIP 업로드)
# -----------------------------

uploaded_zip = st.file_uploader("📦 ZIP 파일 업로드", type=["zip"])

if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, uploaded_zip.name)

        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.getbuffer())

        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(tmpdir)

        rows = []

        for name in sorted(os.listdir(tmpdir)):
            folder_path = os.path.join(tmpdir, name)
            if not os.path.isdir(folder_path):
                continue

            bw, vinyl = process_folder(folder_path)

            rows.append({
                "상위폴더": name,
                "흑백 페이지": bw,
                "비닐내지": vinyl
            })

        st.success("✅ 계산 완료")
        st.table(rows)
