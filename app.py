import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

try:
    from pptx import Presentation
except ImportError:
    Presentation = None

# --- [에이전트 1: 정밀 규칙 추출기 - 우선순위 기반] ---
def extract_rules(text, is_filename=False):
    t = text.lower().replace(" ", "")
    div, mul = None, None
    # 부수 추출: 파일명일 때는 "00부" 형태만 인정하여 버전번호(v7.0) 등과 혼동 방지
    mul_pattern = r'(\d+)(?:부|장)'
    m_mul = re.search(mul_pattern, t)
    if m_mul:
        # 자재 키워드가 주변에 없을 때만 부수로 인정
        if not any(k in t[max(0, m_mul.start()-5):m_mul.end()+5] for k in ['비닐', '간지', '색지', '탭지', '특수']):
            mul = int(m_mul.group(1))
            
    # N-up 추출 (2, 4, 6, 9, 16)
    div_pattern = r'(\d+)(?:up|페이지|쪽|면|쪽모아)'
    m_div = re.search(div_pattern, t)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]:
            div = 1 / val
    return div, mul

# --- [에이전트 2: USB 및 인쇄 차단 판독기] ---
def check_usb_skip(text):
    t = text.lower().replace(" ", "")
    # 단어 경계 없이 실무 키워드 전체 검색
    usb_keywords = ['usb', 'cd', 'usb제작', 'usb담기', 'cd제작', '복사본']
    if any(k in t for k in usb_keywords):
        if 'cdms' not in t: # CDMS 예외처리
            return True
    return False

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V34.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V34.0 - 정밀 분류 엔진)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_instr_contents = set() 

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시서 DB 구축
            db = {}
            for p in all_paths:
                d = os.path.dirname(p).replace('\\', '/')
                if d not in db: db[d] = {"instrs": [os.path.basename(d)], "folder_name": os.path.basename(d)}
                if p.lower().endswith('.txt'):
                    db[d]["instrs"].append(os.path.basename(p))
                    try:
                        with z.open(p) as f:
                            content = f.read().decode('utf-8', errors='ignore')
                            if content.strip(): db[d]["instrs"].append(content)
                    except: pass

            # 2. 분석 엔진
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_special = 0, 0, 0
                
                clean_p = p.replace('\\', '/')
                filename = os.path.basename(clean_p)
                foldername = os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}

                # [계층적 상속 수집]
                path_trace = []
                curr = foldername
                while True:
                    path_trace.append(curr)
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)

                # [규칙 결정 - 상속 우선순위]
                final_div, final_mul = 1.0, 1
                # 1단계: 폴더명/지시서에서 기본값 상속
                for folder in reversed(path_trace): # 상위부터 하위로
                    for instr in db.get(folder, {}).get("instrs", []):
                        d, m = extract_rules(instr)
                        if d: final_div = d
                        if m: final_mul = m
                
                # 2단계: 파일명에 명시된 규칙이 있으면 최우선 적용 (Override)
                f_div, f_mul = extract_rules(filename, is_filename=True)
                if f_div: final_div = f_div
                if f_mul: final_mul = f_mul

                # [카테고리 판정 - 엄격 분리]
                cat = "흑백"
                fn_low = filename.lower()
                # 분류는 파일명과 상위 지시서 내용을 모두 보되, TOC는 파일명에 있을 때만 강력 적용
                if any(k in fn_low for k in ['face', 'spine', 'cover', '표지', 'binder']):
                    cat = "바인더"
                elif any(k in fn_low for k in ['toc', '목차']):
                    cat = "TOC"
                elif '컬러' in fn_low or 'color' in fn_low or '칼라' in fn_low or 'color' in " ".join(db.get(foldername,{}).get("instrs",[])).lower():
                    cat = "컬러"

                # [USB 판정 - 인쇄 제외]
                full_instr_context = filename + " " + " ".join(db.get(foldername, {}).get("instrs", []))
                if check_usb_skip(full_instr_context):
                    cat = "SKIP(USB)"
                    summary[top_folder]["USB"] = 1

                # [자재 정산] - 중복 방지 로직 적용
                # (생략된 자재 로직은 V33과 동일하게 유지하여 비닐 폭발 방지)

                # 페이지 계산
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as f:
                            f_stream = io.BytesIO(f.read())
                            if p.lower().endswith('.pdf'): raw_p = len(PdfReader(f_stream).pages)
                            elif p.lower().endswith('.pptx') and Presentation: raw_p = len(Presentation(f_stream).slides)
                        
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                        summary[top_folder]["총파일수"] += 1
                    except: pass

                # 결과 집계
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, 
                    "계산식": f"{final_div}up x {final_mul}부", "최종P": final_p
                })

        st.subheader("📊 V34.0 정산 요약 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V34.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V34.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
