import streamlit as st
import zipfile
import os
import io
import re
import math
from pypdf import PdfReader
import pandas as pd

# ===============================
# [1] 정규식 유틸
# ===============================

def extract_pages_per_sheet(text: str) -> int | None:
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

def extract_copies(text: str) -> int | None:
    m = re.search(r'(\d+)\s*(부|장)', text)
    if m:
        return int(m.group(1))
    return None

def is_color(text: str) -> bool:
    return any(k in text for k in ['컬러', '칼라', 'color'])

# ===============================
# [2] Streamlit 기본
# ===============================

st.set_page_config(page_title="AI 견적 엔진 v1", layout="wide")
st.title("📦 ZIP 인쇄 자동 정산기 (정확도 우선 v1)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

# ===============================
# [3] ZIP 처리
# ===============================

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as z:
        all_files = [f for f in z.namelist() if not f.endswith('/')]

        # --------------------------------
        # 3-1. 폴더별 txt 내용 수집
        # --------------------------------
        folder_txt = {}

        for f in all_files:
            if f.lower().endswith('.txt'):
                folder = os.path.dirname(f)
                with z.open(f) as tf:
                    content = tf.read().decode('utf-8', errors='ignore').lower()
                    folder_txt.setdefault(folder, "")
                    folder_txt[folder] += " " + content

        # 상위 폴더 상속용 함수
        def collect_txt_context(folder):
            texts = []
            while True:
                if folder in folder_txt:
                    texts.append(folder_txt[folder])
                if not folder or folder == ".":
                    break
                folder = os.path.dirname(folder)
            return " ".join(texts)

        # --------------------------------
        # 3-2. PDF 계산
        # --------------------------------
        results = []

        for f in all_files:
            if not f.lower().endswith('.pdf'):
                continue

            folder = os.path.dirname(f)
            filename = os.path.basename(f)

            # 컨텍스트 합치기 (🔥 핵심)
            context = (
                filename.lower()
                + " "
                + collect_txt_context(folder)
            )

            # 인쇄 조건 추출
            pps = extract_pages_per_sheet(context) or 1
            copies = extract_copies(context) or 1
            color = "컬러" if is_color(context) else "흑백"

            # 페이지 수
            with z.open(f) as pdf_file:
                reader = PdfReader(io.BytesIO(pdf_file.read()))
                raw_pages = len(reader.pages)

            final_pages = math.ceil(raw_pages / pps) * copies

            results.append({
                "폴더": folder if folder else "ROOT",
                "파일명": filename,
                "구분": color,
                "원본페이지": raw_pages,
                "한면": pps,
                "부수": copies,
                "최종페이지": final_pages
            })

    # ===============================
    # [4] 결과 출력
    # ===============================

    df = pd.DataFrame(results)

    summary = (
        df.groupby(["폴더", "구분"])["최종페이지"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    st.subheader("📊 폴더별 요약")
    st.dataframe(summary, use_container_width=True)

    st.subheader("📄 상세 내역")
    st.dataframe(df, use_container_width=True)

    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="요약", index=False)
        df.to_excel(writer, sheet_name="상세", index=False)

    st.download_button(
        "📥 엑셀 다운로드",
        data=output.getvalue(),
        file_name="정산결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
