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

# --- [정밀 추출 엔진] ---
def get_rules(text):
    text = text.lower().replace(" ", "")
    div, mul = 1.0, None
    # 배수 추출 
    m_mul = re.search(r'(\d+)(?:부|장)', text)
    if m_mul: mul = int(m_mul.group(1))
    # N-up 추출 (2, 4, 6, 9, 16 대응)
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]: div = 1 / val
    return div, mul

def get_special_count(text, keywords, default_mul=1):
    text = text.lower().replace(" ", "")
    found_item = ""
    count = 0
    for k in keywords:
        if k in text:
            found_item = k
            # 숫자 추출 시도 
            m = re.search(rf'{k}.*?(\d+)(?:장|개|매)|(\d+)(?:장|개|매).*?{k}', text)
            if m:
                count = int(m.group(1) or m.group(2))
            else:
                count = default_mul if any(x in text for x in ['각', '각각', '하나씩']) else 1
            break
    return count, found_item

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V28.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V28.0 - 특수 자재 정밀 정산)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip") [cite: 5]

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_folders = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')] [cite: 5]
            
            # 1. 지시서 및 구조 DB 구축 [cite: 6, 7]
            db = {}
            for p in all_paths:
                d = os.path.dirname(p)
                if d not in db: db[d] = {"txt": "", "name": os.path.basename(d)}
                if p.lower().endswith('.txt'):
                    try:
                        with z.open(p) as f: db[d]["txt"] += f.read().decode('utf-8', errors='ignore')
                    except: pass

            # 2. 정산 엔진 가동
            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                # [변수 초기화 - 에이전트 5] [cite: 21]
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_special, special_note = 0, ""
                m_vinyl, m_divider = 0, 0
                
                filename = os.path.basename(p)
                foldername = os.path.dirname(p)
                top_folder = p.split('/')[0] if '/' in p else "Root" [cite: 9]
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}

                # 규칙 상속 [cite: 10, 11]
                all_instr = filename + " " + db.get(foldername, {}).get("name", "") + " " + db.get(foldername, {}).get("txt", "")
                f_div, f_mul = get_rules(filename)
                p_div, p_mul = get_rules(db.get(foldername, {}).get("name", "") + " " + db.get(foldername, {}).get("txt", ""))
                
                final_div = f_div if f_div < 1.0 else p_div
                final_mul = f_mul if f_mul is not None else (p_mul if p_mul is not None else 1)

                # [카테고리 판정 - 에이전트 3 (Binder 우선)] 
                cat = "흑백"
                if any(k in filename.lower() for k in ['face', 'spine', 'cover', '표지', 'binder']): cat = "바인더"
                elif any(k in filename.lower() for k in ['toc', '목차']): cat = "TOC"
                elif any(k in all_instr.lower() for k in ['컬러', 'color', '칼라']): cat = "컬러"

                # [인쇄 차단 - USB/CD] [cite: 4, 15]
                if re.search(r'\b(usb|cd)\b', all_instr.lower()) and 'cdms' not in all_instr.lower():
                    cat = "SKIP(USB)"
                    summary[top_folder]["USB"] = 1

                # [자재 및 특수 단어 정산 - 에이전트 4] 
                if foldername not in processed_folders:
                    m_vinyl, _ = get_special_count(all_instr, ["비닐"], final_mul)
                    m_divider, _ = get_special_count(all_instr, ["간지", "색지", "탭지"], final_mul)
                    # 클립, 카드, 라벨, 스티커 정산
                    m_special, special_note = get_special_count(all_instr, ["클립", "카드", "라벨", "스티커"], final_mul)
                    processed_folders.add(foldername)

                # 페이지 계산 [cite: 18, 19]
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as f:
                            f_stream = io.BytesIO(f.read())
                            if p.lower().endswith('.pdf'): raw_p = len(PdfReader(f_stream).pages)
                            elif p.lower().endswith('.pptx') and Presentation: raw_p = len(Presentation(f_stream).slides)
                        
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                        summary[top_folder]["총파일수"] += 1 [cite: 21]
                    except: pass

                # 결과 집계 [cite: 20, 21]
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, 
                    "계산식": f"{final_div}up x {final_mul}부", "최종P": final_p, 
                    "비닐": m_vinyl, "특수항목": special_note, "특수수량": m_special
                })

        st.subheader("📊 V28.0 최종 요약 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True) [cite: 22]
        
        output = io.BytesIO() [cite: 23]
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V28.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V28.xlsx") [cite: 23]

    except Exception as e:
        st.error(f"시스템 오류: {e}")
