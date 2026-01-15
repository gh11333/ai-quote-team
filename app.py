import streamlit as st
import zipfile
import os
import io
import re
from pypdf import PdfReader
import pandas as pd

# --- [에이전트 팀의 규칙 엔진] ---
def calculate_multiplier(text):
    # 1. "1면 X페이지" 규칙 (나누기)
    div_match = re.search(r'1면\s*(\d+)페이지', text)
    if div_match:
        return 1 / int(div_match.group(1)), f"1면 {div_match.group(1)}페이지(÷)"

    # 2. "X장" 규칙 (곱하기)
    mul_match = re.search(r'(\d+)장', text)
    if mul_match:
        return float(mul_match.group(1)), f"{mul_match.group(1)}장(×)"
    
    return 1.0, "기본(1:1)"

# --- [웹 화면 설계] ---
st.set_page_config(page_title="AI 견적 팀", layout="wide")
st.title("📂 무결점 AI 견적 에이전트 팀 (V2.1)")
st.write("규칙: '1면 2페이지'는 0.5배, '3장'은 3배로 자동 계산하며 '비닐' 파일을 별도 체크합니다.")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 올려주세요", type="zip")

if uploaded_zip:
    results = []
    vinyl_count = 0
    
    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        for f in z.namelist():
            if f.startswith('__MACOSX') or not f.lower().endswith('.pdf'): continue
            
            filename = os.path.basename(f)
            foldername = os.path.dirname(f)
            
            # 에이전트 1: 폴더명과 파일명 우선 검토
            # 폴더명에서 먼저 규칙을 찾고, 파일명에 규칙이 있으면 파일명 규칙을 우선합니다.
            multiplier, rule_name = calculate_multiplier(foldername)
            file_multiplier, file_rule_name = calculate_multiplier(filename)
            
            if file_multiplier != 1.0: # 파일명에 규칙이 있으면 덮어쓰기
                multiplier = file_multiplier
                rule_name = file_rule_name

            # 에이전트 2: 비닐 단어 체크
            is_vinyl = "비닐" in filename
            if is_vinyl: vinyl_count += 1
            
            # 에이전트 3: PDF 페이지 추출 및 계산
            try:
                with z.open(f) as pdf_file:
                    reader = PdfReader(io.BytesIO(pdf_file.read()))
                    raw_pages = len(reader.pages)
                    final_pages = raw_pages * multiplier
            except:
                raw_pages, final_pages = 0, 0

            results.append({
                "폴더명": foldername,
                "파일명": filename,
                "물리 페이지": raw_pages,
                "적용 규칙": rule_name,
                "최종 계산": final_pages,
                "비닐 여부": "O" if is_vinyl else "X"
            })

    # 결과 요약
    df = pd.DataFrame(results)
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("총 파일", f"{len(df)}개")
    c2.metric("비닐 포함", f"{vinyl_count}개")
    c3.metric("최종 페이지 합계", f"{df['최종 계산'].sum()}p")

    st.table(df) # 상세 내역 출력

    # 엑셀 다운로드
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    st.download_button("📊 엑셀 견적서 다운로드", data=output.getvalue(), file_name="견적결과.xlsx")
