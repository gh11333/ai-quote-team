import streamlit as st
import zipfile
import os
import io
import re
from pypdf import PdfReader
from pptx import Presentation
import pandas as pd

# --- [에이전트 규칙 엔진: 고도화] ---
def parse_instructions(text):
    text = text.lower().replace(" ", "")
    # 1. 나누기 규칙 (1면 X페이지)
    multiplier = 1.0
    div_match = re.search(r'1면(\d+)페이지', text)
    if div_match:
        multiplier = 1 / int(div_match.group(1))
    elif "양면" in text or "2up" in text:
        multiplier = 0.5
    elif "4up" in text:
        multiplier = 0.25

    # 2. '장' 수량 추출 (자재용)
    count_match = re.search(r'(\d+)장', text)
    count = int(count_match.group(1)) if count_match else 1
    return multiplier, count

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 견적 에이전트 V4.0", layout="wide")
st.title("📂 사내 견적 에이전트 팀 (최상위 폴더별 합산)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 업로드하세요", type="zip")

if uploaded_zip:
    summary_data = {} # {최상위폴더: {흑백: 0, 컬러: 0, 비닐: 0, 색지: 0}}

    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        all_files = [f for f in z.namelist() if not f.startswith('__MACOSX') and not f.endswith('/')]
        
        for f in all_files:
            path_parts = f.split('/')
            if len(path_parts) < 1: continue
            
            # 최상위 폴더명 추출 (예: 12 Site Training)
            top_folder = path_parts[0]
            if top_folder not in summary_data:
                summary_data[top_folder] = {"흑백(기본)": 0, "컬러": 0, "비닐": 0, "색지": 0}
            
            filename = os.path.basename(f)
            fn_low = filename.lower()
            if "출력x" in fn_low: continue # 출력X 제외

            # 규칙 및 수량 파악
            multiplier, count = parse_instructions(fn_low)

            # [자재 체크] 비닐이나 색지는 페이지가 아니라 '개수'로만 더함
            is_material = False
            if "비닐" in fn_low:
                summary_data[top_folder]["비닐"] += count
                is_material = True
            if any(k in fn_low for k in ["색지", "색간지", "간지"]):
                summary_data[top_folder]["색지"] += count
                is_material = True

            # [페이지 체크] PDF, PPTX 파일인 경우에만 흑백/컬러 페이지 합산
            ext = os.path.splitext(f)[1].lower()
            if ext in ['.pdf', '.pptx'] and not is_material:
                try:
                    with z.open(f) as file_data:
                        f_stream = io.BytesIO(file_data.read())
                        pages = 0
                        if ext == '.pdf':
                            pages = len(PdfReader(f_stream).pages)
                        elif ext == '.pptx':
                            pages = len(Presentation(f_stream).slides)
                        
                        final_calc = (pages * multiplier) * count
                        
                        if any(k in fn_low for k in ["컬러", "칼라", "color"]):
                            summary_data[top_folder]["컬러"] += final_calc
                        else:
                            summary_data[top_folder]["흑백(기본)"] += final_calc
                except:
                    pass
            
            # [메모장 체크] .txt 파일에 적힌 비닐/색지 수량 합산
            if ext == '.txt':
                try:
                    with z.open(f) as txt_f:
                        content = txt_f.read().decode('utf-8', errors='ignore')
                        _, txt_count = parse_instructions(content)
                        if "비닐" in content: summary_data[top_folder]["비닐"] += txt_count
                        if "색지" in content or "색간지" in content: summary_data[top_folder]["색지"] += txt_count
                except:
                    pass

    # 결과 테이블 출력
    if summary_data:
        df = pd.DataFrame.from_dict(summary_data, orient='index').reset_index()
        df.columns = ["최상위 카테고리", "흑백(기본)", "컬러", "비닐", "색지"]
        
        st.divider()
        st.subheader("📋 최상위 폴더별 견적 요약 결과")
        st.table(df) # 사용자가 요청한 깔끔한 요약 표

        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        st.download_button("📊 엑셀 견적서 다운로드", data=output.getvalue(), file_name="최종_견적서.xlsx")
