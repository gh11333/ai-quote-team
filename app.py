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
    
    # 1. 나누기 규칙 (분할/양면)
    multiplier = 1.0
    if "양면" in text or "2up" in text or "1면2페이지" in text:
        multiplier = 0.5
    elif "4up" in text or "1면4페이지" in text:
        multiplier = 0.25
        
    # 2. 장수(복사본) 또는 자재 수량 추출
    count_match = re.search(r'(\d+)장', text)
    count = int(count_match.group(1)) if count_match else 1
    
    return multiplier, count

# --- [메인 애플리케이션] ---
st.set_page_config(page_title="AI 견적 팀 V4.0", layout="wide")
st.title("🚀 무결점 사내 견적 에이전트 팀 (최상위 폴더별 합산)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 업로드하세요", type="zip")

if uploaded_zip:
    # 데이터 저장 구조: {최상위폴더: {흑백: 0, 컬러: 0, 비닐: 0, 색지: 0}}
    summary_data = {}

    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        all_files = [f for f in z.namelist() if not f.startswith('__MACOSX') and not f.endswith('/')]
        
        for f in all_files:
            path_parts = f.split('/')
            if len(path_parts) < 2: continue # 최상위 폴더가 없는 경우 제외
            
            top_folder = path_parts[0] # 예: "1. ISF Regulatory Binder"
            if top_folder not in summary_data:
                summary_data[top_folder] = {"흑백": 0, "컬러": 0, "비닐": 0, "색지": 0}
            
            filename = os.path.basename(f)
            fn_low = filename.lower()
            
            # [규칙 1] 출력X 제외
            if "출력x" in fn_low: continue

            # [규칙 2] 자재(비닐/색지) 수량 파악
            # 메모장(.txt)이나 파일명에서 '비닐/색지'와 함께 적힌 '장' 수 추출
            multiplier, count = parse_instructions(fn_low)
            
            is_material = False
            if "비닐" in fn_low:
                summary_data[top_folder]["비닐"] += count
                is_material = True
            if any(k in fn_low for k in ["색지", "색간지", "간지"]):
                summary_data[top_folder]["색지"] += count
                is_material = True

            # [규칙 3] 문서 페이지 계산 (PDF, PPTX)
            # 텍스트 파일은 자재 수량만 체크하고 페이지 계산은 건너뜀
            ext = os.path.splitext(f)[1].lower()
            if ext in ['.pdf', '.pptx']:
                try:
                    with z.open(f) as file_data:
                        f_stream = io.BytesIO(file_data.read())
                        pages = 0
                        if ext == '.pdf':
                            pages = len(PdfReader(f_stream).pages)
                        elif ext == '.pptx':
                            pages = len(Presentation(f_stream).slides)
                        
                        # 실제 출력 페이지 = (물리 페이지 * 분할배수) * 출력장수
                        final_pages = (pages * multiplier) * count
                        
                        # 컬러/흑백 분류 (파일명에 컬러/칼라가 없으면 흑백)
                        if any(k in fn_low for k in ["컬러", "칼라", "color"]):
                            summary_data[top_folder]["컬러"] += final_pages
                        else:
                            summary_data[top_folder]["흑백"] += final_pages
                except:
                    pass

    # 결과 출력
    if summary_data:
        df = pd.DataFrame.from_dict(summary_data, orient='index').reset_index()
        df.columns = ["최상위 카테고리", "흑백(기본)", "컬러", "비닐(속지)", "색지(간지)"]
        
        st.divider()
        st.subheader("📋 최상위 폴더별 견적 요약")
        st.table(df) # 사용자가 원하는 형태의 깔끔한 표

        # 다운로드 버튼
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        st.download_button("📊 최종 견적서 엑셀 다운로드", data=output.getvalue(), file_name="최종_업무_견적서.xlsx")
