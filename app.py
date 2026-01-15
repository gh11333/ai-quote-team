import streamlit as st
import zipfile
import os
import io
import re
from pypdf import PdfReader
import pandas as pd

# --- [에이전트 규칙 엔진: 더 정교한 패턴 매칭] ---
def get_multiplier(text):
    text = text.lower().replace(" ", "") # 공백 제거 및 소문자화로 오차 감소
    
    # 1. 나누기 규칙 (1면 2페이지, 한면에4페이지, 4up 등)
    div_patterns = [r'(\d+)up', r'1면(\d+)페이지', r'한면에(\d+)페이지']
    for p in div_patterns:
        match = re.search(p, text)
        if match:
            val = int(match.group(1))
            return 1 / val, f"{val}분할(÷{val})"

    # 2. 곱하기 규칙 (X장, X회)
    mul_match = re.search(r'(\d+)장', text)
    if mul_match:
        val = int(mul_match.group(1))
        return float(val), f"{val}장(×{val})"
    
    return 1.0, "기본"

# --- [웹 화면 설계] ---
st.set_page_config(page_title="AI 견적 마스터 팀", layout="wide")
st.title("📂 사내 업무 자동화: 견적 에이전트 팀 V3.0")
st.markdown("#### 1. 출력X 제외 | 2. 페이지 분할/배수 적용 | 3. 폴더별 상세 분류 (컬러/색지/비닐)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 선택하세요", type="zip")

if uploaded_zip:
    # 폴더별로 결과를 담을 사전 (Dictionary)
    folder_data = {}

    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        for f in z.namelist():
            # PDF만 처리하며, 맥용 시스템 파일이나 디렉토리 자체는 제외
            if f.startswith('__MACOSX') or not f.lower().endswith('.pdf'): continue
            
            filename = os.path.basename(f)
            foldername = os.path.dirname(f) if os.path.dirname(f) else "루트폴더"
            
            # [규칙 1] 출력X 항목은 계산에서 완전 제외
            if "출력x" in filename.lower(): continue
            
            # [규칙 2] 페이지 배수 계산 (파일명 우선, 없으면 폴더명)
            multiplier, rule_name = get_multiplier(filename)
            if multiplier == 1.0:
                multiplier, rule_name = get_multiplier(foldername)

            # [규칙 3] 분류 에이전트 (카테고리 결정)
            category = "일반(흑백)"
            fn_low = filename.lower()
            if any(k in fn_low for k in ["칼라", "컬러", "color"]):
                category = "컬러"
            elif any(k in fn_low for k in ["색지", "색간지"]):
                category = "색지/간지"
            elif "비닐" in fn_low:
                category = "비닐내지"

            # [규칙 4] 페이지 추출
            try:
                with z.open(f) as pdf_file:
                    reader = PdfReader(io.BytesIO(pdf_file.read()))
                    raw_pages = len(reader.pages)
                    final_pages = raw_pages * multiplier
            except:
                raw_pages, final_pages = 0, 0

            # 폴더별 데이터 합산
            if foldername not in folder_data:
                folder_data[foldername] = {"일반(흑백)": 0, "컬러": 0, "색지/간지": 0, "비닐내지": 0, "파일수": 0}
            
            folder_data[foldername][category] += final_pages
            folder_data[foldername]["파일수"] += 1

    # 데이터프레임 변환 및 출력
    if folder_data:
        df = pd.DataFrame.from_dict(folder_data, orient='index').reset_index()
        df.columns = ["폴더명", "일반(흑백)", "컬러", "색지/간지", "비닐내지", "파일수"]
        
        st.divider()
        st.subheader("📊 폴더별 상세 견적 리포트")
        st.dataframe(df, use_container_width=True)

        # 전체 합계 계산
        total_sum = df.sum(numeric_only=True)
        st.info(f"✅ **전체 합계** | 흑백: {total_sum['일반(흑백)']}p, 컬러: {total_sum['컬러']}p, 색지: {total_sum['색지/간지']}p, 비닐: {total_sum['비닐내지']}p")

        # 엑셀 다운로드
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        st.download_button("📂 폴더별 견적 엑셀 받기", data=output.getvalue(), file_name="최종_폴더별_견적.xlsx")
    else:
        st.warning("분석할 수 있는 PDF 파일이 없습니다. (출력X 제외됨)")
