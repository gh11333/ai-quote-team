import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [에이전트 A: 고유 지시 식별기] ---
def extract_material_data(text, keyword):
    t = text.lower().replace(" ", "")
    if keyword not in t: return None, 0
    
    is_each = any(x in t for x in ['각', '각각', '하나씩'])
    m = re.search(rf'{keyword}.*?(\d+)|(\d+).*?{keyword}', t)
    val = int(m.group(1) or m.group(2)) if m else (1 if is_each or keyword in t else 0)
    
    return "EACH" if is_each else "FIXED", val

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 엔진 V37.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V37.0 - 중복 원천 차단)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    
    # [핵심] 중복 계산 방지용 영수증 (Registry)
    processed_fixed_instrs = set() # (지시내용)
    folder_processed_fixed = set() # (폴더경로, 자재명)

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            db = {}
            for p in all_paths:
                d = os.path.dirname(p).replace('\\', '/')
                if d not in db: db[d] = {"instrs": [os.path.basename(d)]}
                if p.lower().endswith('.txt'):
                    db[d]["instrs"].append(os.path.basename(p))
                    try:
                        with z.open(p) as f:
                            content = f.read().decode('utf-8', errors='ignore')
                            if content.strip(): db[d]["instrs"].append(content)
                    except: pass

            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider = 0, 0
                
                filename = os.path.basename(p)
                foldername = os.path.dirname(p).replace('\\', '/')
                top_folder = p.split('/')[0] if '/' in p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0}

                # [1. 규칙 추출: 상속 체계]
                path_nodes = []
                curr = foldername
                while True:
                    path_nodes.append(curr)
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)
                
                final_div, final_mul = 1.0, 1
                # (생략: 기존 규칙 추출 로직 유지)

                # [2. 자재 정산: 영수증 기반 중복 제거]
                for item, key in {"비닐": "비닐", "색간지": "간지"}.items():
                    # A. 폴더/지시서에 적힌 고정 수량 (FIXED)
                    # 지시 내용 자체가 이미 처리되었다면 패스
                    local_instrs = db.get(foldername, {}).get("instrs", [])
                    for instr in local_instrs:
                        mode, val = extract_material_data(instr, key)
                        if mode == "FIXED" and val > 0:
                            instr_fingerprint = f"{foldername}_{instr}_{val}"
                            if instr_fingerprint not in processed_fixed_instrs:
                                if item == "비닐": m_vinyl += val
                                else: m_divider += val
                                processed_fixed_instrs.add(instr_fingerprint)

                    # B. 개별 수량 (EACH) - 파일명에 직접 있거나 상위 지시에 '각'이 있을 때
                    all_parent_texts = " ".join([txt for node in path_nodes for txt in db.get(node,{}).get("instrs",[])])
                    if any(x in (all_parent_texts + filename).lower() for x in ['각', '각각', '하나씩']):
                        if key in (all_parent_texts + filename).lower():
                            if item == "비닐": m_vinyl += (1 * final_mul)
                            else: m_divider += (1 * final_mul)

                # [3. 카테고리 및 인쇄 정산]
                # (생략: 기존 바인더/TOC/인쇄 로직 유지)

                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                # (나머지 집계...)

        st.subheader("📊 V37.0 최종 요약 (오차 제로 도전)")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'))
        
    except Exception as e:
        st.error(f"오류: {e}")
