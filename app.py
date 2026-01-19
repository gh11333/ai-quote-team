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

# --- [에이전트 1: 시맨틱 추출 엔진] ---
def get_rules(text):
    text = text.lower().replace(" ", "")
    div, mul = 1.0, None
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        m_mul = re.search(r'(\d+)(?:부|장)', text)
        if m_mul: mul = int(m_mul.group(1))
    # 2, 4, 6, 9, 16up 정밀 대응
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]: div = 1 / val
    return div, mul

def get_accessory_info(text_list, keyword):
    """지시 뭉치에서 절대수량(FIXED)과 개별수량(EACH)을 판별"""
    mode = "FIXED"
    fixed_val = 0
    has_keyword = False
    
    for text in text_list:
        t = text.lower().replace(" ", "")
        if keyword not in t: continue
        has_keyword = True
        
        # '각/각각' 키워드 발견 시 모드 전환
        if any(x in t for x in ['각', '각각', '하나씩']):
            mode = "EACH"
        
        # 숫자 추출 (예: 비닐 10장)
        m = re.search(rf'{keyword}.*?(\d+)(?:장|개|매)|(\d+)(?:장|개|매).*?{keyword}', t)
        if m:
            val = int(m.group(1) or m.group(2))
            fixed_val += val
            
    if not has_keyword: return None, 0
    if mode == "EACH": return "EACH", 1
    return "FIXED", max(fixed_val, 1)

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V30.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V30.0 - 상속 및 개별 정산 완결)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    
    # 중복 방지: {top_folder: {지시내용_세트}}
    processed_fixed_instr = {} 

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 전역 지시서 DB 구축 (파일명 + 내용 통합)
            db = {}
            for p in all_paths:
                d = os.path.dirname(p).replace('\\', '/')
                if d not in db: db[d] = {"txts": [], "folder_name": os.path.basename(d)}
                
                if p.lower().endswith('.txt'):
                    # Rule: .txt 파일의 '이름' 자체가 지시인 경우 대응
                    db[d]["txts"].append(os.path.basename(p))
                    try:
                        with z.open(p) as f:
                            content = f.read().decode('utf-8', errors='ignore')
                            if content.strip(): db[d]["txts"].append(content)
                    except: pass

            # 2. 메인 정산 루프
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                # 변수 리셋
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_special = 0, 0, 0
                
                clean_p = p.replace('\\', '/')
                filename = os.path.basename(clean_p)
                foldername = os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}
                    processed_fixed_instr[top_folder] = set()

                # [계층적 상속 구현] 폴더 트리 끝까지 올라가며 지시 수집
                path_instrs = []
                curr = foldername
                while True:
                    info = db.get(curr, {})
                    if info:
                        path_instrs.extend(info.get("txts", []))
                        path_instrs.append(info.get("folder_name", ""))
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)
                
                full_instr_text = " ".join(path_instrs) + " " + filename
                
                # 규칙 확정
                f_div, f_mul = get_rules(filename)
                p_div, p_mul = get_rules(full_instr_text)
                final_div = f_div if f_div < 1.0 else p_div
                final_mul = f_mul if f_mul is not None else (p_mul if p_mul is not None else 1)

                # [자재 정산 - EACH와 FIXED의 엄격한 분리]
                for item_name, keywords in {"비닐": ["비닐"], "색간지": ["간지", "색지", "탭지"], "특수": ["클립", "카드", "라벨", "스티커"]}.items():
                    mode, val = get_accessory_info(path_instrs + [filename], keywords[0])
                    
                    if mode == "EACH":
                        # '각' 모드: 파일당 [1 * 부수] 만큼 무조건 합산
                        if item_name == "비닐": m_vinyl = 1 * final_mul
                        elif item_name == "색간지": m_divider = 1 * final_mul
                        else: m_special = 1 * final_mul
                    elif mode == "FIXED":
                        # 절대 수량 모드: 지시 내용이 중복되지 않을 때만 합산
                        instr_key = f"{item_name}_{val}"
                        if instr_key not in processed_fixed_instr[top_folder]:
                            if item_name == "비닐": m_vinyl = val
                            elif item_name == "색간지": m_divider = val
                            else: m_special = val
                            processed_fixed_instr[top_folder].add(instr_key)

                # 카테고리 판정 (V26 설계도 준수: Binder 우선)
                cat = "흑백"
                if any(k in filename.lower() for k in ['face', 'spine', 'cover', '표지', 'binder']): cat = "바인더"
                elif any(k in filename.lower() for k in ['toc', '목차']): cat = "TOC"
                elif any(k in full_instr_text.lower() for k in ['컬러', 'color', '칼라']): cat = "컬러"

                # 인쇄 차단
                if re.search(r'\b(usb|cd)\b', full_instr_text.lower()) and 'cdms' not in full_instr_text.lower():
                    cat = "SKIP(USB)"
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

                # 요약 결과 반영
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, "계산": f"{final_div}up x {final_mul}부", 
                    "최종P": final_p, "비닐": m_vinyl, "특수": m_special
                })

        st.subheader("📊 V30.0 정밀 요약 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V30.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V30.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
