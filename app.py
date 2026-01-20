import streamlit as st
import zipfile
import io
import os
import re
import math
import random
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation

st.set_page_config(page_title="인쇄 계산 검증기", layout="wide")
st.title("📊 인쇄 페이지 계산 검증 (원본 vs 계산)")

uploaded = st.file_uploader("ZIP 파일 업로드", type="zip")

def extract_nup(text):
    text = text.lower().replace(" ", "")
    m = re.search(r'(\d+)(?:up|페이지|면)', text)
    return int(m.group(1)) if m else 1

def extract_copies(text):
    text = text.lower().replace(" ", "")
    m = re.search(r'(\d+)(?:부|장)', text)
    return int(m.group(1)) if m else 1

if uploaded:
    folder_stats = {}
    sample_rows = []

    with zipfile.ZipFile(uploaded) as z:
        files = [f for f in z.namelist() if not f.endswith('/')]

        # txt 규칙 수집 (상위폴더 기준)
        folder_rules = {}
        for f in files:
            if f.lower().endswith('.txt'):
                top = f.split('/')[0]
                with z.open(f) as t:
                    content = t.read().decode(errors="ignore")
                folder_rules.setdefault(top, []).append(f + " " + content)

        for f in files:
            if not f.lower().endswith(('.pdf', '.pptx')):
                continue

            top = f.split('/')[0]
            name = os.path.basename(f)

            folder_stats.setdefault(top, {
                "원본페이지": 0,
                "계산페이지": 0,
                "파일수": 0
            })

            context = name
            for rule in folder_rules.get(top, []):
                context += " " + rule
            context = context.lower()

            # USB / CD → 페이지 제외
            if any(k in context for k in ["usb", "cd제작", "cd 제작"]):
                continue

            # 비닐내지 → 페이지 제외
            if "비닐내지" in context or "비닐 내지" in context:
                continue

            # 원본 페이지
            raw_pages = 0
            with z.open(f) as fs:
                data = io.BytesIO(fs.read())
                try:
                    if f.lower().endswith('.pdf'):
                        raw_pages = len(PdfReader(data).pages)
                    else:
                        raw_pages = len(Presentation(data).slides)
                except:
                    continue

            nup = extract_nup(context)
            copies = extract_copies(context)
            calc_pages = math.ceil(raw_pages / nup) * copies

            folder_stats[top]["원본페이지"] += raw_pages
            folder_stats[top]["계산페이지"] += calc_pages
            folder_stats[top]["파일수"] += 1

            # 샘플 5개만 저장
            if len(sample_rows) < 5 and random.random() < 0.2:
                sample_rows.append({
                    "폴더": top,
                    "파일명": name,
                    "원본": raw_pages,
                    "n-up": nup,
                    "부수": copies,
                    "계산결과": calc_pages
                })

    # 결과 테이블
    df = pd.DataFrame.from_dict(folder_stats, orient="index")
    df["차이율(%)"] = ((df["계산페이지"] - df["원본페이지"]) / df["원본페이지"] * 100).round(1)

    st.subheader("📁 상위폴더별 요약 (이것만 보면 됨)")
    st.dataframe(df, use_container_width=True)

    st.subheader("🔍 랜덤 샘플 (검증용, 최대 5개)")
    if sample_rows:
        st.dataframe(pd.DataFrame(sample_rows), use_container_width=True)
    else:
        st.write("샘플 없음")

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="요약")
        if sample_rows:
            pd.DataFrame(sample_rows).to_excel(writer, sheet_name="샘플", index=False)

    st.download_button(
        "📥 엑셀 다운로드",
        data=out.getvalue(),
        file_name="검증_리포트.xlsx"
    )
