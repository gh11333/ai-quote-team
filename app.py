import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# PPT 지원 부품
try:
    from pptx import Presentation
    HAS_PPTX = True
except:
    HAS_PPTX = False

# --- [에이전트 지능: 계층적 지휘 엔진 V12.0] ---
def get_multiplier(text):
    if not text: return 1.0, 1
    text = text.lower().replace(" ", "")
    div_val = 1.0
    # '1면4페이지', '4up', '4쪽모아' 등에서 숫자 추출
    div_match = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽)', text)
    if div_match:
        val = int(div_match.group(1))
        if val in [2, 4, 6, 8, 16]: div_val = 1 / val
    
    mul_val = 1
    mul_match = re.search(r'(\d+)(?:부|장)', text)
    if mul_match: mul_val = int(mul_match.group(1))
    return div_val, mul_val

def get_inherited_instructions(foldername, folder_notes):
    """상위 폴더로 거슬러 올라가며 지시사항을 합산함"""
    combined_instr = ""
    current = foldername
    while current:
        if current in folder_notes:
            combined_instr += " " + folder_notes[current]
        parent = os.path.dirname(current)
        if parent == current: break
        current = parent
    return combined_instr

def analyze_file(filename, foldername, inherited_instr):
    fn, fd, fi = filename.lower(), foldername.lower(), inherited_instr.lower()
    combined = f"{fn} {fd} {fi}"
    
    res = {
        "is_usb": any(k in combined for k in ['usb', 'cd']),
        "is_vinyl": "비닐" in combined,
        "is_divider": any(k in combined for k in ['색지', '색간지', '간지', '탭지']),
        "is_divider_each": any(k in combined for k in ['파일마다', '사이사이', '문서마다', '각각', '파일사이']),
        "is_binder": any(k in combined for k in ['cover', 'spine', 'face', '표지']),
        "is_toc": (any(k in fn for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b|_toc|toc_', fn) and 'protocol' not in fn)),
        "is_color": any(k in combined for k in ['컬러', '칼라', 'color'])
    }
    return res

# --- [메인 시스템] ---
st.set_page_config(page_title="사내 견적 에이전트 V12.0", layout="wide")
st.title("🚀 무결점 사내 견적 에이전트 팀 (V12.0 - 계층적 상속 완성)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {} 
    usb_done_folders = set() 

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 전역 지시서 수집 (모든 폴더의 .txt 파일 정보 수집)
            folder_notes = {}
            for p in all_paths:
                if p.lower().endswith('.txt'):
                    d = os.path.dirname(p)
                    folder_notes[d] = folder_notes.get(d, "") + " " + os.path.basename(p)
                    try:
                        with z.open(p) as tf:
                            folder_notes[d] += " " + tf.read().decode('utf-8', errors='ignore')
                    except: pass

            # 2. 파일 분석 및 정산
            valid_files = [f for f in all_paths if not f.endswith('/') and not f.lower().endswith(('.doc', '.docx', '.txt', '.msg'))]
            
            for f in valid_files:
                filename = os.path.basename(f)
                foldername = os.path.dirname(f)
                top_folder = f.split('/')[0] if '/' in f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "특수":0, "총파일수":0}

                if "출력x" in filename.lower(): continue

                # [핵심] 상위 폴더로부터 지시사항 상속
                inherited_instr = get_inherited_instructions(foldername, folder_notes)
                info = analyze_file(filename, foldername, inherited_instr)
                
                # 배수 결정 (파일명 > 상속된 지시사항 > 폴더명)
                f_div, f_mul = get_multiplier(filename)
                txt_div, txt_mul = get_multiplier(inherited_instr)
                fold_div, fold_mul = get_multiplier(foldername)
                
                final_mul = f_mul if f_mul > 1 else (txt_mul if txt_mul > 1 else fold_mul)
                final_div = f_div if f_div < 1.0 else (txt_div if txt_div < 1.0 else fold_div)
                
                ext = os.path.splitext(f)[1].lower()
                p_bw, p_color, m_divider, m_vinyl, m_usb = 0, 0, 0, 0, 0

                # [출력물 판별]
                is_printed = (ext in ['.pdf', '.pptx'] and not info["is_binder"] and not info["is_toc"] and not info["is_divider"] and not info["is_usb"])

                # [USB 및 부자재 정산]
                if info["is_usb"] and foldername not in usb_done_folders:
                    m_usb = 1
                    usb_done_folders.add(foldername)
                
                if info["is_divider"]:
                    m_divider = 1 if info["is_divider_each"] else final_mul
                
                if info["is_vinyl"]:
                    m_vinyl = final_mul if any(k in filename for k in ['각', '각각']) else f_mul

                # [페이지 계산]
                raw_p = 0
                if is_printed:
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                            elif ext == '.pptx' and HAS_PPTX: raw_p = len(Presentation(f_stream).slides)
                            
                            p_val = math.ceil(raw_p * final_div) * final_mul
                            if info["is_color"]: p_color = p_val
                            else: p_bw = p_val
                    except: pass

                # 합산
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["USB or CD"] += m_usb
                if is_printed and (p_bw > 0 or p_color > 0): summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "원본P": raw_p, "최종배수": f"{final_div}x{final_mul}",
                    "결과P": p_bw + p_color, "색간지": m_divider, "비닐": m_vinyl, "인쇄여부": is_printed
                })

        st.subheader("📊 1. 최상위 폴더별 견적 요약 리포트 (V12.0)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        cols = ["흑백", "컬러", "색간지", "비닐", "USB or CD", "특수", "총파일수"]
        st.dataframe(sum_df[cols], use_container_width=True)
        
        st.subheader("🔍 2. 상세 계산 근거")
        st.dataframe(pd.DataFrame(detailed_log), use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df[cols].to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V12.0 최종 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_리포트_V12.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
