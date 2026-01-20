import streamlit as st
import os
import re

st.set_page_config(page_title="출력물 계산기", layout="wide")
st.write("✅ 앱 정상 실행됨")

# -----------------------------
# 유틸 함수
# -----------------------------

def safe_int(x, default=0):
    try:
        return int(x)
    except:
        return default


def extract_vinyl_count(text):
    """
    비닐내지 숫자 추출
    - '비닐내지(3공) 5장' → 5
    - '비닐내지 10장' → 10
    - '비닐내지 안에 넣어주세요' → 1
    - 연도 숫자(2024 등) 절대 제외
    """
    if not text:
        return 0

    m = re.search(r"비닐내지[^0-9]*(\d+)\s*장", text)
    if m:
        return safe_int(m.group(1))

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
    except Exception:
        st.warning(f"PDF 읽기 실패: {os.path.basename(path)}")
        return 0


# -----------------------------
# 파일 처리
# -----------------------------

def process_file(path):
    filename = os.path.basename(path)

    result = {
        "bw": 0,
        "vinyl": 0
    }

    # TXT 처리
    if filename.lower().endswith(".txt"):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
        except:
            content = ""

        result["vinyl"] += extract_vinyl_count(content)
        return result

    # PDF 처리
    if filename.lower().endswith(".pdf"):
        pages = read_pdf_pages_safe(path)

        mode = extract_print_mode(filename)
        nup = extract_nup(filename)

        if mode == "double":
            pages = (pages + 1) // 2

        pages = (pages + nup - 1) // nup
        result["bw"] += pages

        # 비닐내지 PDF 규칙
        if "비닐내지" in filename:
            result["vinyl"] += 1

        return result

    return result


def process_folder(folder_path):
    total_bw = 0
    total_vinyl = 0

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            path = os.path.join(root, file)
            r = process_file(path)
            total_bw += r["bw"]
            total_vinyl += r["vinyl"]

    return total_bw, total_vinyl


# -----------------------------
# UI
# -----------------------------

st.title("📄 출력물 페이지 / 비닐내지 계산기")

base_folder = st.text_input(
    "📁 최상위 폴더 경로 입력",
    placeholder="예: /mount/src/data"
)

if base_folder and os.path.isdir(base_folder):
    rows = []

    for name in sorted(os.listdir(base_folder)):
        folder_path = os.path.join(base_folder, name)
        if not os.path.isdir(folder_path):
            continue

        bw, vinyl = process_folder(folder_path)

        rows.append({
            "폴더명": name,
            "흑백 페이지": bw,
            "비닐내지": vinyl
        })

    st.table(rows)
else:
    st.info("폴더 경로를 입력하세요.")
