import streamlit as st
import zipfile
import os
import io
import re
from pypdf import PdfReader
from pptx import Presentation
import pandas as pd

# --- [에이전트 규칙 엔진: 지능형 해석] ---
def get_multiplier(text):
    text = text.lower().replace(" ", "")
    
    # 1. 분할 인쇄 (나누기)
    div_val = 1.0
    if any(k in text for k in ['4up', '1면4페이지', '4쪽모아']): div_val = 0.25
    elif any(k in text for k in ['2up', '1면2페이지', '2쪽모아', '양면인쇄']): div_val = 0.5 # 양면인쇄가 나누기 의미일 때 대비
    
    # 2. 부수/장수 (곱하기)
    mul_val = 1
    mul_match = re.search(r'(\d+)(부|장)', text)
    if mul_match: mul_val = int(mul_match.group(1))
    
    return div_val, mul_val

def get_category(filename):
    fn = filename.lower()
    if any(k in fn for k in ['cover', 'spine', 'face', '표지']): return "바인더세트"
    if any(k in fn for k in ['toc', 'tableofcontents', '목차']): return "TOC"
    if any(k in fn for k in ['명함', '라벨']): return "특수출력"
    if any(k in fn for k in ['컬러', '칼라', 'color']): return "컬러"
    return "흑백"

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V6.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 시스템 (에이전트 팀 V6.0)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 올려주세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {} # {폴더: {흑백:0, 컬러:0, 비닐:0, 클립:0, TOC:0, 바인더:0, 특수:0}}

    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        # 1. 파일 스캔 및 워드 제거
        all_files = [f for f in z.namelist() if not f.startswith('__MACOSX') and not f.endswith('/')]
        valid_files = [f for f in all_files if not f.lower().endswith(('.doc', '.docx'))]
        
        for f in valid_files:
            path_parts = f.split('/')
            top_folder = path_parts[0] if path_parts else "Root"
            if top_folder not in summary:
                summary[top_folder] = {"흑백":0, "컬러":0, "비닐":0, "클립":0, "USB":0, "CD":0, "TOC":0, "바인더":0, "특수":0}
            
            filename = os.path.basename(f)
            foldername = os.path.dirname(f)
            fn_low = filename.lower()
            if "출력x" in fn_low: continue

            # 배수 및 규칙 추출
            f_div, f_mul = get_multiplier(filename)
            fold_div, fold_mul = get_multiplier(foldername)
            
            final_mul = f_mul if f_mul > 1 else fold_mul
            final_div = f_div if f_div < 1.0 else fold_div
            
            cat = get_category(filename)
            ext = os.path.splitext(f)[1].lower()
            
            p_bw, p_color, m_vinyl, m_clip, m_toc, m_binder, m_special = 0, 0, 0, 0, 0, 0, 0

            # [자재 및 특수 카테고리 정산]
            if cat == "바인더세트": m_binder = final_mul
            elif cat == "TOC": m_toc = final_mul
            elif cat == "특수출력": m_special = final_mul
            
            # [부자재 체크]
            if "비닐" in fn_low:
                # '각' 또는 '각각'이 있으면 부수만큼 비닐도 추가
                m_vinyl = final_mul if any(k in fn_low for k in ['각', '각각', '하나씩']) else 1
            if "클립" in fn_low: m_clip = 1
            
            # [페이지 계산] - 일반 문서인 경우에만
            if ext in ['.pdf', '.pptx'] and cat in ["흑백", "컬러"]:
                try:
                    with z.open(f) as fd:
                        stream = io.BytesIO(fd.read())
                        raw_p = len(PdfReader(stream).pages) if ext=='.pdf' else len(Presentation(stream).slides)
                        calc_p = (raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = calc_p
                        else: p_bw = calc_p
                except: raw_p = 0
            else: raw_p = 0

            # 요약 데이터 합산
            summary[top_folder]["흑백"] += p_bw
            summary[top_folder]["컬러"] += p_color
            summary[top_folder]["비닐"] += m_vinyl
            summary[top_folder]["클립"] += m_clip
            summary[top_folder]["TOC"] += m_toc
            summary[top_folder]["바인더"] += m_binder
            summary[top_folder]["특수"] += m_special

            detailed_log.append({
                "폴더": top_folder, "파일명": filename, "카테고리": cat, "원본P": raw_p,
                "배수": f"{final_div}x{final_mul}", "결과P": p_bw + p_color, "비닐": m_vinyl, "TOC": m_toc
            })

    # 화면 표시
    st.subheader("📊 1. 최상위 폴더별 최종 견적 요약")
    st.dataframe(pd.DataFrame.from_dict(summary, orient='index'))
    
    st.subheader("🔍 2. 상세 계산 근거 (검증용)")
    st.dataframe(pd.DataFrame(detailed_log))

    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
        pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
    st.download_button("📂 검증 완료 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_리포트.xlsx")
