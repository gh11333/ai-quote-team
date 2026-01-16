import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [에이전트 지능: 고도화된 전략 엔진 V15.0] ---
def get_multiplier(text):
    if not text: return 1.0, 1
    text = text.lower().replace(" ", "")
    div_val = 1.0
    div_match = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽)', text)
    if div_match:
        val = int(div_match.group(1))
        if val in [2, 4, 6, 8, 16]: div_val = 1 / val
    
    mul_val = 1
    mul_match = re.search(r'(\d+)(?:부|장)', text)
    if mul_match: mul_val = int(mul_match.group(1))
    return div_val, mul_val

def get_category(filename):
    fn = filename.lower()
    if any(k in fn for k in ['cover', 'spine', 'face', '표지']): return "바인더세트"
    if any(k in fn for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b|_toc|toc_', fn) and 'protocol' not in fn):
        return "TOC"
    if any(k in fn for k in ['명함', '라벨']): return "특수출력"
    if any(k in fn for k in ['컬러', '칼라', 'color']): return "컬러"
    return "흑백"

# --- [메인 시스템] ---
st.set_page_config(page_title="사내 견적 에이전트 V15.0", layout="wide")
st.title("🚀 무결점 사내 견적 에이전트 팀 (V15.0 - USB 및 페이지 계산 완결판)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {} 
    # [Q1 해결 핵심] USB 지시가 시작된 근원지를 추적하여 중복 카운트 방지
    usb_sources_counted = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 전역 지시서 정보 수집
            folder_notes = {}
            for p in all_paths:
                if p.lower().endswith('.txt'):
                    d = os.path.dirname(p)
                    folder_notes[d] = folder_notes.get(d, "") + " " + os.path.basename(p)
                    try:
                        with z.open(p) as tf:
                            folder_notes[d] += " " + tf.read().decode('utf-8', errors='ignore')
                    except: pass

            # 2. 파일 분석 시작
            valid_files = [f for f in all_paths if not f.endswith('/') and not f.lower().endswith(('.doc', '.docx', '.txt', '.msg'))]
            
            for f in valid_files:
                filename = os.path.basename(f)
                foldername = os.path.dirname(f)
                top_folder = f.split('/')[0] if '/' in f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "특수":0, "TOC":0, "바인더":0, "총파일수":0}

                # 계층적 지시 상속 및 USB 지시 근원지 찾기
                inherited_instr = ""
                usb_source_path = ""
                curr = foldername
                while True:
                    curr_instr = folder_notes.get(curr, "")
                    if any(k in (curr + curr_instr).lower() for k in ['usb', 'cd']):
                        usb_source_path = curr # USB 지시가 처음 발견된 가장 상위 폴더 기록
                    
                    if curr in folder_notes: inherited_instr += " " + folder_notes[curr]
                    parent = os.path.dirname(curr)
                    if parent == curr or not curr: break
                    curr = parent
                
                combined_low = (filename + " " + foldername + " " + inherited_instr).lower()
                
                # 배수 및 규칙 결정
                f_div, f_mul = get_multiplier(filename)
                txt_div, txt_mul = get_multiplier(inherited_instr)
                fold_div, fold_mul = get_multiplier(foldername)
                
                final_mul = f_mul if f_mul > 1 else (txt_mul if txt_mul > 1 else fold_mul)
                final_div = f_div if f_div < 1.0 else (txt_div if txt_div < 1.0 else fold_div)
                
                cat = get_category(filename)
                ext = os.path.splitext(f)[1].lower()
                p_bw, p_color, m_divider, m_vinyl, m_usb = 0, 0, 0, 0, 0

                # [Q1 해결] USB 정산: 지시 근원지당 딱 1번만 카운트
                if usb_source_path and usb_source_path not in usb_sources_counted:
                    m_usb = 1
                    usb_sources_counted.add(usb_source_path)

                # [부자재 정산]
                is_divider_file = any(k in filename.lower() for k in ['색지', '색간지', '간지', '탭지'])
                has_divider_instr = any(k in (foldername + inherited_instr).lower() for k in ['색지', '색간지', '간지', '탭지', '파일사이'])
                if is_divider_file: m_divider = final_mul
                elif has_divider_instr: m_divider = 1
                if "비닐" in combined_low:
                    m_vinyl = final_mul if any(k in filename.lower() for k in ['각', '각각']) else f_mul

                # [Q2 해결] 페이지 계산 로직
                raw_p = 0
                is_instruction_pdf = any(k in filename for k in ["제작방식", "지시서"])
                is_printed = (ext in ['.pdf', '.pptx'] and cat in ["흑백", "컬러"] and not is_divider_file and not usb_source_path and not is_instruction_pdf)
                
                if is_printed:
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                            # 29p * 0.5 = 14.5 -> ceil(14.5) = 15장 확정
                            p_val = math.ceil(raw_p * final_div) * final_mul
                            if cat == "컬러": p_color = p_val
                            else: p_bw = p_val
                    except: pass

                # 요약 합산
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["USB or CD"] += m_usb
                summary[top_folder]["TOC"] += (1 if cat == "TOC" else 0)
                summary[top_folder]["바인더"] += (1 if cat == "바인더세트" else 0)
                if is_printed and (p_bw > 0 or p_color > 0): summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "카테고리": cat, "원본P": raw_p,
                    "배수": f"{final_div}x{final_mul}", "최종P": p_bw + p_color, "USB": m_usb, "근원지": usb_source_path
                })

        st.subheader("📊 1. 최상위 폴더별 견적 요약 리포트 (V15.0)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        cols = ["흑백", "컬러", "색간지", "비닐", "USB or CD", "특수", "TOC", "바인더", "총파일수"]
        st.dataframe(sum_df[cols], use_container_width=True)
        
        st.subheader("🔍 2. 상세 계산 근거 (검증용)")
        st.dataframe(pd.DataFrame(detailed_log), use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df[cols].to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V15.0 최종 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_리포트_V15.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
