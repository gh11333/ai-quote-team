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
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        m_mul = re.search(r'(\d+)(?:부|장)', text)
        if m_mul: mul = int(m_mul.group(1))
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]: div = 1 / val
    return div, mul

def get_accessory_logic(text, keyword):
    """자재 지시 성격 판별 (절대수량 vs 개별수량)"""
    text = text.lower().replace(" ", "")
    if keyword not in text: return None, 0
    
    # 숫자 패턴 검색
    m = re.search(rf'{keyword}.*?(\d+)(?:장|개|매)|(\d+)(?:장|개|매).*?{keyword}', text)
    if m:
        return "FIXED", int(m.group(1) or m.group(2))
    elif any(x in text for x in ['각', '각각', '하나씩']):
        return "EACH", 1
    return "FIXED", 1

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V29.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V29.0 - 개별 자재 정밀 정산)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_materials = set() # 절대수량 중복 방지용
    usb_counted_top = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            db = {}
            for p in all_paths:
                d = os.path.dirname(p)
                if d not in db: db[d] = {"txt": "", "name": os.path.basename(d)}
                if p.lower().endswith('.txt'):
                    try:
                        with z.open(p) as f: db[d]["txt"] += f.read().decode('utf-8', errors='ignore')
                    except: pass

            for p in all_paths:
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_special = 0, 0, 0
                
                clean_p = p.replace('\\', '/')
                filename, foldername = os.path.basename(clean_p), os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}

                all_instr = filename + " " + db.get(foldername, {}).get("name", "") + " " + db.get(foldername, {}).get("txt", "")
                f_div, f_mul = get_rules(filename)
                p_div, p_mul = get_rules(db.get(foldername, {}).get("name", "") + " " + db.get(foldername, {}).get("txt", ""))
                final_div = f_div if f_div < 1.0 else p_div
                final_mul = f_mul if f_mul is not None else (p_mul if p_mul is not None else 1)

                # --- [자재 정산 로직 개선] ---
                for item_name, keys in {"비닐": ["비닐"], "색간지": ["간지", "색지", "탭지"], "특수": ["클립", "카드", "라벨", "스티커"]}.items():
                    mode, val = get_accessory_logic(all_instr, keys[0] if item_name != "색간지" else "간지")
                    
                    if mode == "EACH":
                        # '각' 모드: 모든 파일에 대해 부수만큼 합산
                        count = val * final_mul
                        if item_name == "비닐": m_vinyl = count
                        elif item_name == "색간지": m_divider = count
                        else: m_special = count
                    elif mode == "FIXED" and f"{foldername}_{item_name}" not in processed_fixed_materials:
                        # 절대수량 모드: 폴더당 1회만 합산
                        if item_name == "비닐": m_vinyl = val
                        elif item_name == "색간지": m_divider = val
                        else: m_special = val
                        processed_fixed_materials.add(f"{foldername}_{item_name}")

                # 카테고리 및 인쇄 계산 (기존 유지)
                cat = "흑백"
                if any(k in filename.lower() for k in ['face', 'spine', 'cover', '표지', 'binder']): cat = "바인더"
                elif any(k in filename.lower() for k in ['toc', '목차']): cat = "TOC"
                elif any(k in all_instr.lower() for k in ['컬러', 'color', '칼라']): cat = "컬러"

                if re.search(r'\b(usb|cd)\b', all_instr.lower()) and 'cdms' not in all_instr.lower():
                    cat = "SKIP(USB)"
                    if top_folder not in usb_counted_top:
                        summary[top_folder]["USB"] = 1
                        usb_counted_top.add(top_folder)

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

                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({"폴더": top_folder, "파일명": filename, "분류": cat, "원본P": raw_p, "계산": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl})

        st.subheader("📊 V29.0 요약 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V29.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V29.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
