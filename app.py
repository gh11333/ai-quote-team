import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [1단계: 사전 정의 함수들] ---

def get_clean_num(text, pattern):
    """지정된 패턴에서 숫자만 안전하게 추출"""
    m = re.search(pattern, text.lower().replace(" ", ""))
    if m:
        for g in m.groups():
            if g is not None: return int(g)
    return None

def get_material_info(text_list, keyword):
    """지시 리스트에서 고정수량(Fixed)과 개별수량(Each)을 분리 추출"""
    fixed_val = 0
    is_each = False
    for t in text_list:
        sl = t.lower().replace(" ", "")
        if keyword not in sl: continue
        if any(x in sl for x in ['각', '각각', '하나씩']): is_each = True
        num = get_clean_num(sl, rf'{keyword}.*?(\d+)|(\d+).*?{keyword}')
        if num: fixed_val += num
    return is_each, fixed_val

# --- [메인 화면 구성] ---
st.set_page_config(page_title="무결점 엔진 V38.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V38.0 - 완전 재설계)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_registry = set() # 중복 정산 방지

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시사항 전수 DB 구축
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

            # 2. 분석 및 계산
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                # 변수 초기화
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider = 0, 0
                
                clean_p = p.replace('\\', '/')
                filename = os.path.basename(clean_p)
                foldername = os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "총파일수":0}

                # [계층 분석]
                path_nodes = []
                curr = foldername
                while True:
                    path_nodes.append(curr)
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)

                # [규칙 확정 - 상속 스택]
                final_div, final_mul = 1.0, 1
                div_f, mul_f = False, False
                # 파일명 규칙 우선 적용
                d_val = get_clean_num(filename, r'(\d+)(?:up|페이지|쪽|면|쪽모아)')
                m_val = get_clean_num(filename, r'(\d+)(?:부|장)')
                if d_val: final_div, div_f = 1/d_val, True
                if m_val: final_mul, mul_f = m_val, True
                
                # 상위 상속
                for node in path_nodes:
                    node_instrs = db.get(node, {}).get("instrs", [])
                    for instr in node_instrs:
                        if not div_f:
                            d = get_clean_num(instr, r'(\d+)(?:up|페이지|쪽|면|쪽모아)')
                            if d: final_div, div_f = 1/d, True
                        if not mul_f:
                            m = get_clean_num(instr, r'(\d+)(?:부|장)')
                            if m: final_mul, mul_f = m, True

                # [카테고리 분류 - 파일명 독립 단어 기준]
                fn_lower = " " + filename.lower().replace("_", " ").replace("-", " ") + " "
                cat = "인쇄"
                if any(re.search(rf'\b{k}\b', fn_lower) for k in ['face', 'spine', 'cover', '표지', 'binder']):
                    cat = "바인더"
                elif any(re.search(rf'\b{k}\b', fn_lower) for k in ['toc', '목차']):
                    cat = "TOC"
                
                # 컬러 판단 (파일명 + 현재 폴더 지시서)
                context = (filename + " " + " ".join(db.get(foldername,{}).get("instrs",[]))).lower()
                if cat == "인쇄":
                    cat = "컬러" if any(k in context for k in ['컬러', '칼라', 'color']) else "흑백"

                # [자재 정산 - 영수증 로직]
                for item, key in {"비닐": "비닐", "색간지": "간지"}.items():
                    # A. 고정 수량 (지시가 있는 폴더에서 딱 한 번만 합산)
                    local_is_each, local_fixed = get_material_info(db.get(foldername,{}).get("instrs",[]), key)
                    if local_fixed > 0:
                        reg_id = f"{foldername}_{item}_{local_fixed}"
                        if reg_id not in processed_fixed_registry:
                            if item == "비닐": m_vinyl += local_fixed
                            else: m_divider += local_fixed
                            processed_fixed_registry.add(reg_id)
                    
                    # B. 개별 수량 (상위 경로 어디든 '각'이 있으면 파일당 합산)
                    all_path_instrs = []
                    for node in path_nodes: all_path_instrs.extend(db.get(node,{}).get("instrs",[]))
                    global_is_each, _ = get_material_info(all_path_instrs + [filename], key)
                    if global_is_each:
                        if item == "비닐": m_vinyl += (1 * final_mul)
                        else: m_divider += (1 * final_mul)

                # [USB 차단]
                if any(k in context for k in ['usb', 'cd제작']) and 'cdms' not in filename.lower():
                    cat = "SKIP"
                    summary[top_folder]["USB"] = 1

                # [페이지 계산]
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as f_in:
                            f_stream = io.BytesIO(f_in.read())
                            if p.lower().endswith('.pdf'):
                                raw_p = len(PdfReader(f_stream).pages)
                            elif p.lower().endswith('.pptx') and Presentation:
                                raw_p = len(Presentation(f_stream).slides)
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                        summary[top_folder]["총파일수"] += 1
                    except: pass

                # 결과 집합
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({"폴더": top_folder, "파일명": filename, "분류": cat, "계산": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl})

        st.subheader("📊 V38.0 최종 정산 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세')
        st.download_button("📂 정산서 다운로드", data=output.getvalue(), file_name="최종_정산_V38.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
