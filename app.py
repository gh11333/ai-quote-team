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

# --- [에이전트 1: 정밀 규칙 추출기 - 독립 실행] ---
def extract_print_rule(text):
    text = text.lower().replace(" ", "")
    div, mul = None, None
    # 부수 추출 (숫자+부/장)
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지', '특수']):
        m_mul = re.search(r'(\d+)(?:부|장)', text)
        if m_mul: mul = int(m_mul.group(1))
    # N-up 추출
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]: div = 1 / val
    return div, mul

# --- [에이전트 2: 자재 산출기 - 중복 제거] ---
def get_accessory_logic(segments, keyword):
    """지시 세그먼트들을 분석하여 FIXED(수량) 또는 EACH(개별) 판별"""
    is_each = False
    fixed_sum = 0
    seen_contents = set() # 중복 지시서 내용 방지
    
    for s in segments:
        s_low = s.lower().replace(" ", "")
        if keyword not in s_low or s_low in seen_contents: continue
        seen_contents.add(s_low)
        
        if any(x in s_low for x in ['각', '각각', '하나씩']):
            is_each = True
        
        m = re.search(rf'{keyword}.*?(\d+)(?:장|개|매)|(\d+)(?:장|개|매).*?{keyword}', s_low)
        if m: fixed_sum += int(m.group(1) or m.group(2))
    
    if is_each: return "EACH", 1
    if fixed_sum > 0: return "FIXED", fixed_sum
    if seen_contents: return "FIXED", 1 # 키워드는 있는데 숫자가 없으면 기본 1개
    return None, 0

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V32.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V32.0 - 엔진형 구조)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_accessories = set() # (폴더, 자재명) 기준 중복 방지

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시서 DB 구축 (파일 내용/이름 수집)
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

            # 2. 메인 분석 엔진
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

                # [계층적 세그먼트 수집]
                path_segments = [filename]
                curr = foldername
                while True:
                    if curr in db: path_segments.extend(db[curr]["instrs"])
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)

                # [규칙 결정 - 격리 분석]
                final_div, final_mul = 1.0, 1
                for s in path_segments:
                    s_div, s_mul = extract_print_rule(s)
                    if final_div == 1.0 and s_div: final_div = s_div
                    if final_mul == 1 and s_mul: final_mul = s_mul

                # [자재 정산 - EACH/FIXED 엔진]
                for item, keys in {"비닐": "비닐", "색간지": "간지", "특수": "특수"}.items():
                    mode, val = get_accessory_logic(path_segments, keys)
                    if mode == "EACH":
                        # EACH는 파일마다 합산 (파일당 1번만)
                        count = val * final_mul
                        if item == "비닐": m_vinyl = count
                        elif item == "색간지": m_divider = count
                        else: m_special = count
                    elif mode == "FIXED" and f"{foldername}_{item}" not in processed_fixed_accessories:
                        # FIXED는 폴더당 한 번만 합산
                        if item == "비닐": m_vinyl = val
                        elif item == "색간지": m_divider = val
                        else: m_special = val
                        processed_fixed_accessories.add(f"{foldername}_{item}")

                # [카테고리 분류]
                cat = "흑백"
                if any(k in filename.lower() for k in ['face', 'spine', 'cover', '표지', 'binder']): cat = "바인더"
                elif any(k in filename.lower() for k in ['toc', '목차']): cat = "TOC"
                elif any(k in " ".join(path_segments).lower() for k in ['컬러', 'color', '칼라']): cat = "컬러"

                # USB 예외
                if re.search(r'\b(usb|cd)\b', " ".join(path_segments).lower()) and 'cdms' not in filename.lower():
                    cat = "SKIP"
                    summary[top_folder]["USB"] = 1

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

                # 집계
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, 
                    "계산식": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl
                })

        # --- [에이전트 D: 최종 무결성 검사] ---
        for folder, data in summary.items():
            if data["비닐"] > data["총파일수"] * 20 and data["비닐"] > 100:
                st.warning(f"⚠️ [{folder}] 폴더의 비닐 수량({data['비닐']}개)이 파일 수에 비해 과도하게 많습니다. 로직 확인이 필요할 수 있습니다.")

        st.subheader("📊 V32.0 정밀 정산 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V32.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V32.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
