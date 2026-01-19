import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# PPTX 라이브러리 체크
try:
    from pptx import Presentation
except ImportError:
    Presentation = None

# --- [에이전트 1: 정밀 규칙 추출기] ---
def extract_print_rule(text):
    """배수(부) 및 분할(up) 규칙 추출 (독립 단어 판정)"""
    t = " " + text.lower().replace(" ", " ") + " "
    div, mul = None, None
    
    # 1. 분할 인쇄(Up) 추출
    m_div = re.search(r'(\d+)\s*(?:up|페이지|쪽|면|쪽모아)', t)
    if m_div and int(m_div.group(1)) in [2, 4, 6, 9, 16]:
        div = 1 / int(m_div.group(1))
        
    # 2. 인쇄 부수(Mul) 추출 - 자재 관련 단어가 주변에 없을 때만
    if not any(k in t for k in ['비닐', '간지', '색지', '탭지', '특수', '라벨', '스티커', '카드', '클립']):
        m_mul = re.search(r'(\d+)\s*(?:부|장)', t)
        if m_mul: mul = int(m_mul.group(1))
            
    return div, mul

# --- [에이전트 2: 카테고리 및 자재 판독기] ---
def get_file_category(filename):
    """분류는 오직 파일 이름의 '독립 단어'로만 결정"""
    fn = " " + filename.lower().replace("_", " ").replace("-", " ") + " "
    # 바인더 부속 (Face, Spine, Cover)
    if any(re.search(rf'\b{k}\b', fn) for k in ['face', 'spine', 'cover', '표지', 'binder']):
        return "바인더"
    # TOC (목차) - Protocol 내의 toc 방지 위해 단어 경계(\b) 필수
    if any(re.search(rf'\b{k}\b', fn) for k in ['toc', '목차']):
        return "TOC"
    return "인쇄"

def get_material_data(text_list, keyword):
    """자재 수량 판별 (EACH vs FIXED)"""
    is_each = False
    fixed_val = 0
    keyword_found = False
    
    for text in text_list:
        t = text.lower().replace(" ", "")
        if keyword not in t: continue
        keyword_found = True
        
        # '각' 모드 판별
        if any(x in t for x in ['각', '각각', '하나씩']): is_each = True
        
        # 숫자 추출
        m = re.search(rf'{keyword}.*?(\d+)|(\d+).*?{keyword}', t)
        if m: fixed_val += int(m.group(1) or m.group(2))
            
    return is_each, fixed_val, keyword_found

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 엔진 V37.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V37.0 - 중복 차단 및 정밀 분류)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_items = set() # (폴더명, 자재명, 값) 기준 중복 방지

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

            # 2. 메인 정산 엔진
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                # 변수 초기화
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_special = 0, 0, 0
                
                clean_p = p.replace('\\', '/')
                filename = os.path.basename(clean_p)
                foldername = os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}

                # [계층 구조 상속 수집]
                path_segments = []
                curr = foldername
                while True:
                    if curr in db: path_segments.extend(db[curr]["instrs"])
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)

                # [규칙 결정: 상속 스택]
                final_div, final_mul = 1.0, 1
                div_found, mul_found = False, False
                # 파일명 우선 -> 하위 폴더 -> 상위 폴더 순
                for s in [filename] + db.get(foldername,{}).get("instrs", []) + path_segments:
                    d, m = extract_print_rule(s)
                    if not div_found and d: final_div, div_found = d, True
                    if not mul_found and m: final_mul, mul_found = m, True

                # [카테고리 결정: 파일명 독립 단어 기준]
                cat = get_file_category(filename)
                
                # 인쇄물인 경우 컬러 여부 판단 (지시서 포함)
                if cat == "인쇄":
                    context_all = (filename + " " + " ".join(db.get(foldername,{}).get("instrs",[]))).lower()
                    if any(k in context_all for k in ['컬러', '칼라', 'color']): cat = "컬러"
                    else: cat = "흑백"

                # [자재 정산: EACH vs FIXED 격리]
                for item, keys in {"비닐": ["비닐"], "색간지": ["간지", "색지", "탭지"], "특수": ["라벨", "스티커", "카드", "클립"]}.items():
                    # A. FIXED (폴더당 한 번만 합산)
                    local_instrs = db.get(foldername, {}).get("instrs", [])
                    is_each, fixed_val, found = analyze_accessories(local_instrs, keys[0])
                    if fixed_val > 0:
                        key_id = f"{foldername}_{item}_{fixed_val}"
                        if key_id not in processed_fixed_items:
                            if item == "비닐": m_vinyl += fixed_val
                            elif item == "색간지": m_divider += fixed_val
                            else: m_special += fixed_val
                            processed_fixed_items.add(key_id)
                    
                    # B. EACH (상위 어디든 '각'이 있으면 파일당 합산)
                    global_is_each, _, _ = analyze_accessories(path_segments + [filename], keys[0])
                    if global_is_each:
                        val = 1 * final_mul
                        if item == "비닐": m_vinyl += val
                        elif item == "색간지": m_divider += val
                        else: m_special += val

                # [USB/CD 인쇄 제외]
                if re.search(r'\b(usb|cd)\b', (filename + " " + foldername).lower()) and 'cdms' not in filename.lower():
                    cat = "SKIP"
                    summary[top_folder]["USB"] = 1

                # [페이지 연산]
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

                # [집계 반영]
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, "원본P": raw_p, 
                    "계산식": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl
                })

        st.subheader("📊 V37.0 정밀 정산 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세')
        st.download_button("📂 정산서 다운로드", data=output.getvalue(), file_name="최종_정산_V37.xlsx")

    except Exception as e:
        st.error(f"오류 발생: {e}")

# Helper function
def analyze_accessories(text_list, keyword):
    is_each = False
    fixed_val = 0
    found = False
    for txt in text_list:
        t = txt.lower().replace(" ", "")
        if keyword not in t: continue
        found = True
        if any(x in t for x in ['각', '각각', '하나씩']): is_each = True
        m = re.search(rf'{keyword}.*?(\d+)|(\d+).*?{keyword}', t)
        if m: fixed_val += int(m.group(1) or m.group(2))
    return is_each, fixed_val, found
