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

# --- [에이전트 1: 정밀 규칙 추출기] ---
def get_rules(text):
    """단일 텍스트 세그먼트에서 규칙 추출"""
    text = text.lower().replace(" ", "")
    div, mul = None, None
    
    # N-up 추출 (2, 4, 6, 9, 16 대응)
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]: div = 1 / val
        
    # 배수(부수) 추출: 해당 세그먼트에 자재 키워드가 없을 때만 인정
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        m_mul = re.search(r'(\d+)(?:부|장)', text)
        if m_mul: mul = int(m_mul.group(1))
        
    return div, mul

# --- [에이전트 2: 지능형 자재 정산기] ---
def get_accessory_info(text_list, keyword):
    """지시 리스트에서 EACH(각)와 FIXED(합산) 판별 및 중복 제거"""
    mode = "FIXED"
    found_values = set() # 동일 폴더 내 중복 숫자(예: 10, 10) 제거용
    has_keyword = False
    
    for text in text_list:
        t = text.lower().replace(" ", "")
        if keyword not in t: continue
        has_keyword = True
        
        # '각' 모드 판별
        if any(x in t for x in ['각', '각각', '하나씩']):
            mode = "EACH"
            
        # 숫자 추출 (3장 + 2장 = 5장 / 10장 + 10장 = 10장 대응)
        matches = re.findall(rf'{keyword}.*?(\d+)(?:장|개|매)|(\d+)(?:장|개|매).*?{keyword}', t)
        for g1, g2 in matches:
            found_values.add(int(g1 or g2))
            
    if not has_keyword: return None, 0
    if mode == "EACH": return "EACH", 1
    return "FIXED", sum(found_values) if found_values else 1

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V31.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V31.0 - 로직 독립화)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_instr = set() # (폴더경로, 자재명, 수량) 기준 중복 방지

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시서 DB 구축
            db = {}
            for p in all_paths:
                d = os.path.dirname(p).replace('\\', '/')
                if d not in db: db[d] = {"txts": [], "folder_name": os.path.basename(d)}
                if p.lower().endswith('.txt'):
                    db[d]["txts"].append(os.path.basename(p))
                    try:
                        with z.open(p) as f:
                            content = f.read().decode('utf-8', errors='ignore')
                            if content.strip(): db[d]["txts"].append(content)
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

                # [계층적 상속 세그먼트 수집]
                segments = [filename]
                curr = foldername
                while True:
                    info = db.get(curr, {})
                    if info:
                        segments.extend(info["txts"])
                        segments.append(info["folder_name"])
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)
                
                # [규칙 확정: 파일명 -> 현재폴더 -> 상위폴더 순으로 첫 발견값 채택]
                final_div, final_mul = 1.0, 1
                for seg in segments:
                    s_div, s_mul = get_rules(seg)
                    if final_div == 1.0 and s_div is not None: final_div = s_div
                    if final_mul == 1 and s_mul is not None: final_mul = s_mul

                # [자재 정산 - 독립 로직]
                for item_name, keys in {"비닐": ["비닐"], "색간지": ["간지", "색지", "탭지"], "특수": ["클립", "카드", "라벨", "스티커"]}.items():
                    mode, val = get_accessory_info(segments, keys[0])
                    if mode == "EACH":
                        # 각비닐 등 개별 모드: 모든 파일에 부수만큼 합산
                        count = val * final_mul
                        if item_name == "비닐": m_vinyl = count
                        elif item_name == "색간지": m_divider = count
                        else: m_special = count
                    elif mode == "FIXED":
                        # 절대 수량 모드: 해당 폴더에서 이 수량이 정산된 적 없으면 합산
                        instr_id = f"{foldername}_{item_name}_{val}"
                        if instr_id not in processed_fixed_instr:
                            if item_name == "비닐": m_vinyl = val
                            elif item_name == "색간지": m_divider = val
                            else: m_special = val
                            processed_fixed_instr.add(instr_id)

                # [카테고리 판정 - 파일명 기반]
                cat = "흑백"
                full_instr_lower = " ".join(segments).lower()
                if any(k in filename.lower() for k in ['face', 'spine', 'cover', '표지', 'binder']): cat = "바인더"
                elif any(k in filename.lower() for k in ['toc', '목차']): cat = "TOC"
                elif any(k in full_instr_lower for k in ['컬러', 'color', '칼라']): cat = "컬러"

                # USB/CD 처리
                if re.search(r'\b(usb|cd)\b', full_instr_lower) and 'cdms' not in full_instr_lower:
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

                # 요약 합산
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

        st.subheader("📊 V31.0 최종 정산 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V31.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V31.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
