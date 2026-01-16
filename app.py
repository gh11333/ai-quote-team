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

# --- [정밀 시맨틱 유틸리티] ---
def get_num_with_unit(text, keywords, unit_pattern=r'(\d+)(?:장|개|매)'):
    """지정된 키워드 주변의 숫자를 추출 (절대 수량 합산용)"""
    text = text.lower().replace(" ", "")
    total = 0
    for k in keywords:
        if k in text:
            # 키워드 뒤 숫자 또는 앞 숫자 추출
            matches = re.findall(rf'{k}.*?{unit_pattern}', text)
            matches += re.findall(rf'{unit_pattern}.*?{k}', text)
            total += sum(int(m) for m in matches)
    return total

def get_print_rules(text):
    """배수(부) 및 분할(up) 규칙 추출 (V26.0 엄격 판정)"""
    text = text.lower().replace(" ", "")
    mul = None
    # Rule: 숫자 뒤에 반드시 '부' 또는 '장' (자재 키워드 제외 시)
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        m = re.search(r'(\d+)(?:부|장)', text)
        if m: mul = int(m.group(1))
    
    div = 1.0
    # 29p 2-up -> 0.5배 처리
    m_div = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽|면)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 8]: div = 1 / val
    return div, mul

def is_hard_skip(text):
    """Rule: USB, CD 단어 경계 처리 (CDMS 제외)"""
    t = " " + text.lower() + " "
    if any(k in t for k in [' usb ', ' cd ']):
        if 'cdms' not in t: return True
    if any(k in t for k in ['usb제작', 'cd제작', 'usb에', 'cd에', 'usb담기']): return True
    return False

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V26.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V26.0 - 마스터 설계도 준수)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    
    # 중복 합산 방지 장치
    processed_folders_materials = set()
    usb_counted_top = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 구조 스캔 및 지시서 DB 구축
            folder_db = {} 
            sibling_rules = {} # 형제 상속용
            
            for p in all_paths:
                clean_p = p.replace('\\', '/').rstrip('/')
                d, b = os.path.dirname(clean_p), os.path.basename(clean_p)
                if d not in folder_db: folder_db[d] = {"texts": [os.path.basename(d)], "raw_names": []}
                
                folder_db[d]["raw_names"].append(b)
                if b.lower().endswith('.txt'):
                    try:
                        with z.open(p) as tf:
                            folder_db[d]["texts"].append(tf.read().decode('utf-8', errors='ignore'))
                    except: pass
                
                # 형제 상속: 폴더명 자체가 지시인 경우 (1면 4페이지 폴더 등)
                if p.endswith('/') or '.' not in b:
                    parent = os.path.dirname(d)
                    sibling_rules[d] = b 

            # 2. 메인 분석 엔진
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                # [초기화] 이전 파일의 데이터 잔존 차단
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_usb = 0, 0, 0
                
                clean_f = p.replace('\\', '/').rstrip('/')
                filename = os.path.basename(clean_f)
                foldername = os.path.dirname(clean_f)
                top_folder = clean_f.split('/')[0] if '/' in clean_f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "총파일수":0}

                # 지시 수집 (상위 + 지시서 + 형제)
                inheritance_trace = []
                curr = foldername
                combined_instr = ""
                while True:
                    local_text = " ".join(folder_db.get(curr, {}).get("texts", []))
                    # 형제 상속 추가
                    siblings = " ".join([sibling_rules.get(k, "") for k in sibling_rules if os.path.dirname(k) == os.path.dirname(curr)])
                    local_info = local_text + " " + siblings
                    combined_instr += " " + local_info
                    inheritance_trace.append(curr)
                    if curr == os.path.dirname(curr) or not curr: break
                    curr = os.path.dirname(curr)

                # 규칙 확정 (파일명 우선)
                f_div, f_mul = get_print_rules(filename)
                p_div, p_mul = get_print_rules(combined_instr)
                
                final_div = f_div if f_div < 1.0 else p_div
                final_mul = f_mul if f_mul is not None else (p_mul if p_mul is not None else 1)

                # 자재 정산 (폴더당 1회 합산)
                if foldername not in processed_folders_materials:
                    all_local_txt = " ".join(folder_db.get(foldername, {}).get("texts", []))
                    m_vinyl = get_num_with_unit(all_local_txt, ["비닐"])
                    m_divider = get_num_with_unit(all_local_txt, ["간지", "색지", "탭지"])
                    # '각' 키워드 시 부수 곱하기 (Rule 4)
                    if any(k in all_local_txt.lower() for k in ['각', '각각', '하나씩']):
                        if m_vinyl == 0 and '비닐' in all_local_txt.lower(): m_vinyl = final_mul
                    processed_folders_materials.add(foldername)

                # 카테고리 판정 및 인쇄 차단
                full_scope = (filename + " " + combined_instr).lower()
                cat = "흑백"
                if is_hard_skip(full_scope):
                    cat = "SKIP(USB)"
                    if top_folder not in usb_counted_top:
                        summary[top_folder]["USB"] += 1
                        usb_counted_top.add(top_folder)
                elif any(k in full_scope for k in ['binder', 'face', 'spine', 'cover', '표지']): cat = "바인더"
                elif any(k in full_scope for k in ['목차', 'toc']) and 'protocol' not in full_scope: cat = "TOC"
                elif any(k in full_scope for k in ['컬러', '칼라', 'color']): cat = "컬러"

                # 페이지 계산
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if p.lower().endswith('.pdf'):
                                raw_p = len(PdfReader(f_stream).pages)
                            elif p.lower().endswith('.pptx') and Presentation:
                                raw_p = len(Presentation(f_stream).slides)
                        
                        # Rule: math.ceil(원본*분할)*부수
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                        summary[top_folder]["총파일수"] += 1
                    except: pass

                # 요약 반영
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, "원본P": raw_p, 
                    "상속지시": combined_instr[:100] + "...", "계산식": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl
                })

        st.subheader("📊 V26.0 요약 리포트 (설계도 준수)")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V26.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V26.xlsx")

    except Exception as e:
        st.error(f"시스템 오류: {e}")
