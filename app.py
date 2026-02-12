import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import openpyxl

# --- [설정 및 상수] ---
VERSION = "V40.0-PRO"
SUPPORTED_EXTS = ('.pdf', '.pptx', '.xlsx', '.xls')
CATEGORY_KEYWORDS = {
    "바인더": ['face', 'spine', 'cover', '표지', 'binder', '세로형'],
    "TOC": ['toc', '목차'],
}

# --- [핵심 로직 함수] ---

def get_number_from_text(text, patterns):
    """다양한 패턴에서 숫자를 추출하는 유틸리티"""
    text = text.lower().replace(" ", "")
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return int(match.group(1))
    return None

def analyze_file_context(filename, folder_instrs):
    """
    파일명과 폴더 지시사항을 분석하여 인쇄 옵션 결정
    우선순위: 파일명 > 현재 폴더 > 상위 폴더
    """
    # 1. n-up (페이지 축약) 추출
    up_patterns = [r'(\d+)up', r'(\d+)쪽모아', r'(\d+)분할', r'(\d+)페이지(?:당|씩)']
    n_up = get_number_from_text(filename, up_patterns)
    
    # 2. 부수(Copies) 추출
    copy_patterns = [r'(\d+)부', r'(\d+)세트', r'(\d+)장씩']
    copies = get_number_from_text(filename, copy_patterns)

    # 3. 폴더 지시사항에서 상속 (파일 이름에 없을 경우)
    for instr in reversed(folder_instrs):
        if n_up is None: n_up = get_number_from_text(instr, up_patterns)
        if copies is None: copies = get_number_from_text(instr, copy_patterns)

    return (n_up or 1), (copies or 1)

def get_page_count(file_content, ext):
    """파일 타입별 실제 페이지/슬라이드 수 계산"""
    try:
        f_stream = io.BytesIO(file_content)
        if ext == '.pdf':
            return len(PdfReader(f_stream).pages)
        elif ext == '.pptx':
            return len(Presentation(f_stream).slides)
        elif ext in ['.xlsx', '.xls']:
            wb = openpyxl.load_stream(f_stream) if ext == '.xlsx' else None
            return len(wb.sheetnames) if wb else 1
    except Exception:
        return 0
    return 0

# --- [메인 서비스 클래스] ---

class QuotationEngine:
    def __init__(self):
        self.summary = {}
        self.detailed_logs = []
        self.processed_fixed = set()

    def process_zip(self, uploaded_file):
        with zipfile.ZipFile(uploaded_file, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX') and not p.endswith('/')]
            
            # 폴더별 지시서(txt) 및 폴더명 미리 로드
            db = {}
            for p in z.namelist():
                dir_name = os.path.dirname(p)
                if dir_name not in db: db[dir_name] = [os.path.basename(dir_name)]
                if p.lower().endswith('.txt'):
                    with z.open(p) as f:
                        db[dir_name].append(f.read().decode('utf-8', errors='ignore'))

            for path in all_paths:
                filename = os.path.basename(path)
                ext = os.path.splitext(filename)[1].lower()
                folder_path = os.path.dirname(path)
                top_folder = path.split('/')[0] if '/' in path else "Root"
                
                if top_folder not in self.summary:
                    self.summary[top_folder] = {"흑백": 0, "컬러": 0, "색간지": 0, "비닐": 0, "USB": 0, "TOC": 0, "바인더": 0, "파일수": 0}

                # 1. 지시사항 상속 (상위 폴더 트리 탐색)
                folder_nodes = []
                curr = folder_path
                while True:
                    folder_nodes.append(db.get(curr, []))
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)
                
                flat_instrs = [item for sublist in folder_nodes for item in sublist]
                n_up, copies = analyze_file_context(filename, flat_instrs)

                # 2. 카테고리 분류
                cat = "인쇄"
                if any(k in filename.lower() for k in CATEGORY_KEYWORDS["바인더"]): cat = "바인더"
                elif any(k in filename.lower() for k in CATEGORY_KEYWORDS["TOC"]): cat = "TOC"
                
                # 컬러 여부 판단 (Context 기반)
                context_str = (filename + " ".join(flat_instrs)).lower()
                is_color = any(k in context_str for k in ['컬러', '칼라', 'color'])
                if cat == "인쇄": cat = "컬러" if is_color else "흑백"

                # 3. 페이지 계산
                final_p = 0
                if ext in SUPPORTED_EXTS and cat in ["흑백", "컬러"]:
                    raw_p = get_page_count(z.read(path), ext)
                    # 계산 공식: ceil(원본 / N-up) * 부수
                    final_p = math.ceil(raw_p / n_up) * copies
                    self.summary[top_folder][cat] += final_p
                    self.summary[top_folder]["파일수"] += 1

                # 4. 자재 정산 (비닐/간지)
                m_vinyl, m_divider = 0, 0
                for item, key in {"비닐": "비닐", "색간지": "간지"}.items():
                    if any(k in context_str for k in [f'{key}각', f'{key}각각']):
                        val = copies
                        if item == "비닐": m_vinyl = val
                        else: m_divider = val
                
                self.summary[top_folder]["비닐"] += m_vinyl
                self.summary[top_folder]["색간지"] += m_divider

                # 로그 기록
                self.detailed_logs.append({
                    "상위폴더": top_folder,
                    "파일명": filename,
                    "분류": cat,
                    "설정": f"{n_up}UP / {copies}부",
                    "최종P": final_p,
                    "비닐": m_vinyl,
                    "간지": m_divider
                })

# --- [Streamlit UI] ---
st.set_page_config(page_title=f"무결점 엔진 {VERSION}", layout="wide")
st.title(f"🚀 견적 자동화 시스템 {VERSION}")
st.markdown("---")

uploaded_file = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_file:
    engine = QuotationEngine()
    with st.spinner("파일 분석 중..."):
        engine.process_zip(uploaded_file)
    
    st.subheader("📊 폴더별 정산 요약")
    df_summary = pd.DataFrame.from_dict(engine.summary, orient='index')
    st.dataframe(df_summary, use_container_width=True)

    st.subheader("📑 상세 내역 로그")
    df_details = pd.DataFrame(engine.detailed_logs)
    st.dataframe(df_details, use_container_width=True)

    # 엑셀 다운로드
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='요약')
        df_details.to_excel(writer, sheet_name='상세내역')
    
    st.download_button(
        label="📂 엑셀 정산서 다운로드",
        data=output.getvalue(),
        file_name=f"견적정산_{VERSION}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
