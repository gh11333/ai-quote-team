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

# --- [에이전트 A: 정밀 시맨틱 추출기] ---
def parse_value(text, keywords, unit_pattern):
    """텍스트에서 키워드와 결합된 숫자를 추출"""
    text = text.lower().replace(" ", "")
    results = set()
    for k in keywords:
        if k in text:
            # 패턴: 키워드+숫자+단위 또는 숫자+단위+키워드
            m1 = re.findall(rf'{k}.*?(\d+){unit_pattern}', text)
            m2 = re.findall(rf'(\d+){unit_pattern}.*?{k}', text)
            for val in (m1 + m2): results.add(int(val))
    return results

def extract_printing_rules(text):
    """배수(부) 및 분할(up) 규칙 추출"""
    text = text.lower().replace(" ", "")
    mul = None
    # '부' 또는 '장'이 붙은 숫자 추출 (비닐 등 자재 키워드 제외 시)
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        m = re.search(r'(\d+)(?:부|장)', text)
        if m: mul = int(m.group(1))
    
    div = 1.0
    # 1면 4페이지, 4up, 4쪽모아 등 대응
    m_div = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽|면)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 8]: div = 1 / val
    return div, mul

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V25.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V25.0 - 시맨틱 정밀 감사)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    
    # 중복 방지 캐시
    processed_materials = {} # {folder_path: {item_name: set(values)}}
    usb_counted = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 전수 스캔 및 지시서 데이터베이스화
            folder_db = {} 
            for p in all_paths:
                clean_p = p.replace('\\', '/').rstrip('/')
                d, b = os.path.dirname(clean_p), os.path.basename(clean_p)
                if d not in folder_db: folder_db[d] = {"texts": [os.path.basename(d)], "files": []}
                if b.lower().endswith('.txt'):
                    try:
                        with z.open(p) as tf:
                            folder_db[d]["texts"].append(tf.read().decode('utf-8', errors='ignore'))
                    except: pass
                else:
                    folder_db[d]["files"].append(p)

            # 2. 분석 엔진 가동
            for folder_path, data in folder_db.items():
                top_folder = folder_path.split('/')[0] if '/' in folder_path else "Root"
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0}

                # 해당 폴더의 자재(비닐/간지) 계산 (중복 제거 합산)
                local_texts = " ".join(data["texts"])
                for item, keys in {"비닐": ["비닐"], "색간지": ["간지", "색지", "탭지"]}.items():
                    found_values = parse_value(local_texts, keys, r'(?:장|개|매)')
                    summary[top_folder][item] += sum(found_values)

                # 파일별 상세 분석
                for f_path in data["files"]:
                    if f_path.endswith('/'): continue
                    
                    # 변수 초기화
                    filename = os.path.basename(f_path)
                    raw_p, p_bw, p_color = 0, 0, 0
                    
                    # 상속 규칙 결정
                    # 부모/상위 지시 수집
                    inherited_text = ""
                    curr = os.path.dirname(f_path)
                    while curr:
                        inherited_text += " " + " ".join(folder_db.get(curr, {}).get("texts", []))
                        if curr == os.path.dirname(curr) or not curr: break
                        curr = os.path.dirname(curr)
                    
                    f_div, f_mul = extract_printing_rules(filename)
                    p_div, p_mul = extract_printing_rules(inherited_text)
                    
                    final_div = f_div if f_div < 1.0 else p_div
                    final_mul = f_mul if f_mul is not None else (p_mul if p_mul is not None else 1)

                    # 카테고리 판정 (강력한 우선순위)
                    combined_scope = (f_path + " " + inherited_text).lower()
                    cat = "흑백"
                    if any(k in combined_scope for k in ['binder', 'face', 'spine', 'cover', '표지']): 
                        cat = "바인더"
                    elif any(k in combined_scope for k in ['tableofcontents', '목차', 'toc']) and 'protocol' not in combined_scope: 
                        cat = "TOC"
                    elif any(k in combined_scope for k in ['컬러', '칼라', 'color']): 
                        cat = "컬러"

                    # USB 판정 (최상위 폴더당 1회)
                    if any(k in combined_scope for k in ['usb', 'cd']) and 'cdms' not in combined_scope:
                        if top_folder not in usb_counted:
                            summary[top_folder]["USB"] += 1
                            usb_counted.add(top_folder)
                        cat = "SKIP" # 인쇄 제외

                    # 인쇄 계산
                    if cat in ["흑백", "컬러"]:
                        try:
                            with z.open(f_path) as fd:
                                f_stream = io.BytesIO(fd.read())
                                if f_path.lower().endswith('.pdf'):
                                    raw_p = len(PdfReader(f_stream).pages)
                                elif f_path.lower().endswith('.pptx') and Presentation:
                                    raw_p = len(Presentation(f_stream).slides)
                            
                            p_val = math.ceil(raw_p * final_div) * final_mul
                            if cat == "컬러": p_color = p_val
                            else: p_bw = p_val
                        except: pass
                    
                    # 요약 합산
                    summary[top_folder]["흑백"] += p_bw
                    summary[top_folder]["컬러"] += p_color
                    if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                    if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                    detailed_log.append({
                        "폴더": top_folder, "파일명": filename, "분류": cat, "원본P": raw_p, "규칙": f"{final_div}up x {final_mul}부", "최종P": p_bw + p_color
                    })

        st.subheader("📊 V25.0 정밀 감사 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V25.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V25.xlsx")

    except Exception as e:
        st.error(f"시스템 오류: {e}")
