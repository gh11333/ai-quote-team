import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [에이전트 A: 엄격한 규칙 추출기] ---
def extract_print_rule(text):
    t = text.lower().replace(" ", "")
    div, mul = None, None
    
    # 1. 분할 인쇄(Up) 추출
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', t)
    if m_div and int(m_div.group(1)) in [2, 4, 6, 9, 16]:
        div = 1 / int(m_div.group(1))
        
    # 2. 인쇄 부수(Mul) 추출 - 자재 관련 단어가 없을 때만!
    if not any(k in t for k in ['비닐', '간지', '색지', '탭지', '특수', '라벨', '스티커']):
        m_mul = re.search(r'(\d+)(?:부|장)', t)
        if m_mul: mul = int(m_mul.group(1))
            
    return div, mul

# --- [에이전트 B: 분류 및 자재 판독기] ---
def get_file_category(filename):
    """분류는 오직 파일 이름으로만 결정 (폴더명 배제)"""
    fn = filename.lower()
    if any(k in fn for k in ['face', 'spine', 'cover', '표지', 'binder']): return "바인더"
    if any(k in fn for k in ['toc', '목차']): return "TOC"
    return "인쇄"

def get_material_count(segments, keyword):
    """자재 수량은 모든 지시 세그먼트에서 전수 조사"""
    total = 0
    is_each = False
    for s in segments:
        sl = s.lower().replace(" ", "")
        if keyword not in sl: continue
        if any(x in sl for x in ['각', '각각', '하나씩']): is_each = True
        # 숫자 추출
        m = re.search(rf'{keyword}.*?(\d+)|(\d+).*?{keyword}', sl)
        if m: total += int(m.group(1) or m.group(2))
    return is_each, total

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 엔진 V36.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V36.0 - 독립형 엔진)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_items = set() # 중복 합산 방지

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 폴더별 지시서 DB 구축
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

            # 2. 메인 정산
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_special = 0, 0, 0
                
                filename = os.path.basename(p)
                foldername = os.path.dirname(p).replace('\\', '/')
                top_folder = p.split('/')[0] if '/' in p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0}

                # [계층 구조 및 상속]
                path_segments = [filename]
                curr = foldername
                while True:
                    if curr in db: path_segments.extend(db[curr]["instrs"])
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)

                # [규칙 결정: 파일명 우선 -> 상위 상속]
                final_div, final_mul = 1.0, 1
                found_div, found_mul = False, False
                for s in [filename] + db.get(foldername,{}).get("instrs", []) + path_segments:
                    d, m = extract_print_rule(s)
                    if not found_div and d: final_div, found_div = d, True
                    if not found_mul and m: final_mul, found_mul = m, True

                # [카테고리 결정: 오직 파일명으로만!]
                cat = get_file_category(filename)
                if cat == "인쇄":
                    if any(k in filename.lower() or k in " ".join(db.get(foldername,{}).get("instrs",[])).lower() for k in ['컬러', '칼라', 'color']):
                        cat = "컬러"
                    else: cat = "흑백"

                # [자재 정산: FIXED/EACH 분리]
                for item, key in {"비닐": "비닐", "색간지": "간지", "특수": "특수"}.items():
                    is_each, fixed_val = get_material_count(path_segments, key)
                    if is_each: # 개별 모드
                        count = 1 * final_mul
                        if item == "비닐": m_vinyl = count
                        elif item == "색간지": m_divider = count
                    elif fixed_val > 0: # 고정 수량 모드
                        key_id = f"{foldername}_{item}_{fixed_val}"
                        if key_id not in processed_fixed_items:
                            if item == "비닐": m_vinyl = fixed_val
                            elif item == "색간지": m_divider = fixed_val
                            processed_fixed_items.add(key_id)

                # [USB 차단]
                if any(k in (filename + " " + foldername).lower() for k in ['usb', 'cd제작']) and 'cdms' not in filename.lower():
                    cat = "SKIP"
                    summary[top_folder]["USB"] = 1

                # [연산]
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as f:
                            f_stream = io.BytesIO(f.read())
                            if p.lower().endswith('.pdf'): raw_p = len(PdfReader(f_stream).pages)
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                    except: pass

                # 집계
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({"폴더": top_folder, "파일명": filename, "분류": cat, "계산": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl})

        st.subheader("📊 V36.0 최종 요약")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'))
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세')
        st.download_button("📂 정산서 다운로드", data=output.getvalue(), file_name="최종_정산_V36.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
