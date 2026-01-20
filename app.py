import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

st.set_page_config(page_title="출력물 자동 정산기", layout="wide")
st.title("📦 ZIP 출력물 자동 정산 (페이지 + 비닐 통합 판단)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type=["zip"])

# ---------------------------
# 유틸 함수
# ---------------------------

def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.lower())

def extract_vinyl_qty(text: str) -> int:
    """
    비닐/비닐내지 수량 추출
    규칙:
    - 비닐 키워드 없으면 0
    - 숫자 있으면 그 숫자
    - 숫자 없으면 1
    - (3공)의 3은 무시
    """
    if "비닐" not in text:
        return 0

    # 장 / 개 와 붙은 숫자 우선
    nums = re.findall(r"(\d+)\s*(?:장|개)", text)
    if nums:
        return sum(int(n) for n in nums)

    # 그 외 숫자 (단, 3공 제외)
    nums = re.findall(r"\d+", text)
    filtered = [int(n) for n in nums if int(n) != 3]
    if filtered:
        return max(filtered)

    return 1

def extract_up_divisor(text: str) -> int:
    """
    한면 n페이지 / nup / n페이지씩 → n
    없으면 1
    """
    patterns = [
        r"(\d+)\s*up",
        r"한면\s*(\d+)\s*페이지",
        r"1면\s*(\d+)\s*페이지",
        r"(\d+)\s*페이지씩"
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return int(m.group(1))
    return 1

def is_page_excluded(text: str) -> bool:
    """
    페이지 계산 제외 조건
    """
    exclude_keywords = [
        "비닐만",
        "비닐내지만",
        "출력없음",
        "페이지 계산 안함"
    ]
    return any(k in text for k in exclude_keywords)

# ---------------------------
# 메인 처리
# ---------------------------

if uploaded_zip:
    summary = {}
    details = []

    with zipfile.ZipFile(uploaded_zip, "r") as z:
        all_files = [f for f in z.namelist() if not f.endswith("/")]

        # TXT 내용 미리 읽기
        txt_contents = {}
        for f in all_files:
            if f.lower().endswith(".txt"):
                with z.open(f) as tf:
                    try:
                        txt_contents[os.path.dirname(f)] = normalize(
                            tf.read().decode("utf-8", errors="ignore")
                        )
                    except:
                        txt_contents[os.path.dirname(f)] = ""

        for f in all_files:
            if not f.lower().endswith(".pdf"):
                continue

            top_folder = f.split("/")[0]
            folder = os.path.dirname(f)
            filename = os.path.basename(f)

            if top_folder not in summary:
                summary[top_folder] = {
                    "흑백페이지": 0,
                    "비닐": 0
                }

            # ---------------------------
            # 1️⃣ 텍스트 수집
            # ---------------------------
            texts = [
                normalize(filename),
                normalize(folder),
                txt_contents.get(folder, "")
            ]
            full_text = " ".join(texts)

            # ---------------------------
            # 2️⃣ 비닐 판단
            # ---------------------------
            vinyl_qty = extract_vinyl_qty(full_text)
            summary[top_folder]["비닐"] += vinyl_qty

            # ---------------------------
            # 3️⃣ 페이지 계산 여부
            # ---------------------------
            if is_page_excluded(full_text):
                page_count = 0
            else:
                with z.open(f) as pdf_file:
                    reader = PdfReader(io.BytesIO(pdf_file.read()))
                    raw_pages = len(reader.pages)

                up = extract_up_divisor(full_text)
                page_count = math.ceil(raw_pages / up)

            summary[top_folder]["흑백페이지"] += page_count

            details.append({
                "상위폴더": top_folder,
                "파일명": filename,
                "원본페이지": raw_pages if page_count else 0,
                "UP": up if page_count else "-",
                "최종페이지": page_count,
                "비닐": vinyl_qty
            })

    df_summary = pd.DataFrame(summary).T.reset_index().rename(columns={"index": "폴더"})
    df_detail = pd.DataFrame(details)

    st.subheader("📊 폴더별 요약")
    st.dataframe(df_summary, use_container_width=True)

    st.subheader("📄 상세 내역")
    st.dataframe(df_detail, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="요약", index=False)
        df_detail.to_excel(writer, sheet_name="상세", index=False)

    st.download_button(
        "📥 엑셀 다운로드",
        data=output.getvalue(),
        file_name="정산결과.xlsx"
    )
