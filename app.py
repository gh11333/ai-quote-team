import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [에이전트 1: 배수 및 분할 판독관] ---
def get_strict_multiplier(text):
    if not text: return 1.0, 1
    t = text.lower().replace(" ", "")
    
    # 1. 분할 인쇄 (나누기)
    div_val = 1.0
    div_match = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽)', t)
    if div_match:
        val = int(div_match.group(1))
        if val in [2, 4, 6, 8]: div_val = 1 / val
        
    # 2. 부수 (곱하기) - 매우 엄격하게 매칭 (예: '20부', '5장')
    mul_val = 1
    mul_match = re.search(r'(\d+)(?:부|장)(?![\d\w])', t) # 숫자+부/장 뒤에 글자가 더 없어야 함
    if mul_match:
        mul_candidate = int(mul_match.group(1))
        # 상식 밖의 배수(예: 100배 이상)는 오인식으로 간주하여 차단
        if mul_candidate < 100: mul_val = mul_candidate
        
    return div_val, mul_val

# --- [에이전트 2: 카테고리 판별관] ---
def get_strict_category(filename):
    fn = filename.lower()
    if any(k in fn for k in ['cover', 'spine', 'face', '표지']): return "바인더세트"
    if any(k in fn for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b|_toc|toc_', fn) and 'protocol' not in fn):
        return "TOC"
    if any(k in fn for k in ['명함', '라벨']): return "특수출력"
    if any(k in fn for k in ['컬러', '칼라', 'color']): return "컬러"
    return "흑백"

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V22.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 시스템 (V22.0 - 무결점 감사 버전)")

uploaded_zip = st.file_uploader("작업 ZIP 파일을 올려주세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    usb_counted_paths = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            raw_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시서 사전 스캔
            folder_notes = {}
            for p in raw_paths:
                clean_p = p.replace('\\', '/').rstrip('/')
                d, b = os.path.dirname(clean_p), os.path.basename(clean_p)
                if b.lower().endswith('.txt'):
                    try:
                        with z.open(p) as tf:
                            folder_notes[d] = folder_notes.get(d, "") + " " + tf.read().decode('utf-8', errors='ignore')
                    except: pass

            # 2. 파일 전수 조사
            valid_files = [p for p in raw_paths if not p.endswith('/') and not p.lower().endswith(('.doc', '.docx', '.txt', '.msg'))]
            
            for f in valid_files:
                clean_f = f.replace('\\', '/').rstrip('/')
                filename, foldername = os.path.basename(clean_f), os.path.dirname(clean_f)
                top_folder = clean_f.split('/')[0] if '/' in clean_f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "TOC":0, "바인더":0, "총파일수":0}

                # 지시 상속 (가장 가까운 폴더 지시 우선)
                inherited_instr = ""
                usb_source = ""
                curr = foldername
                while True:
                    local_info = folder_notes.get(curr, "") + " " + os.path.basename(curr)
                    inherited_instr += " " + local_info
                    if re.search(r'\busb\b|\bcd\b|usb제작|cd제작', local_info.lower()) and not usb_source:
                        usb_source = curr
                    parent = os.path.dirname(curr)
                    if parent == curr or not curr: break
                    curr = parent

                # 배수 결정 (파일 지시가 상위 폴더 지시보다 우선함)
                f_div, f_mul = get_strict_multiplier(filename)
                p_div, p_mul = get_strict_multiplier(inherited_instr)
                
                final_mul = f_mul if f_mul > 1 else p_mul
                final_div = f_div if f_div < 1.0 else p_div
                
                cat = get_strict_category(filename)
                ext = os.path.splitext(clean_f)[1].lower()
                p_bw, p_color, m_divider, m_vinyl, m_usb = 0, 0, 0, 0, 0

                # [정산 로직]
                if usb_source and usb_source not in usb_counted_paths:
                    m_usb = 1
                    usb_counted_paths.add(usb_source)
                
                combined_low = (filename + " " + inherited_instr).lower()
                if "비닐" in combined_low:
                    # '각'이 없으면 묶음 포장(1개), 있으면 개별 포장(final_mul)
                    m_vinyl = final_mul if any(k in combined_low for k in ['각', '각각', '하나씩']) else 1
                
                if any(k in combined_low for k in ['색지', '색간지', '간지', '탭지']):
                    m_divider = final_mul if any(k in filename.lower() for k in ['색지', '간지']) else 1

                # [페이지 계산 - 엄격 분리]
                raw_p = 0
                is_printed = (ext in ['.pdf', '.pptx'] and cat in ["흑백", "컬러"] and not usb_source and "제작방식" not in filename)
                
                if is_printed:
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                            # 계산식: 올림(원본P * 분할) * 부수
                            p_val = math.ceil(raw_p * final_div) * final_mul
                            if cat == "컬러": p_color = p_val
                            else: p_bw = p_val
                    except: pass

                # 결과 합산 (감사 에이전트: TOC/바인더는 일반 페이지에서 완전 제외)
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["USB or CD"] += m_usb
                summary[top_folder]["TOC"] += (final_mul if cat == "TOC" else 0)
                summary[top_folder]["바인더"] += (final_mul if cat == "바인더세트" else 0)
                if is_printed and (p_bw + p_color > 0): summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "원본P": raw_p, "배수": f"{final_div}x{final_mul}", "결과P": p_bw + p_color, "비닐": m_vinyl, "색간지": m_divider
                })

        st.subheader("📊 1. 최종 검증 완료 요약 (V22.0)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        cols = ["흑백", "컬러", "색간지", "비닐", "USB or CD", "TOC", "바인더", "총파일수"]
        st.dataframe(sum_df[cols], use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df[cols].to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V22.0 최종 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V22.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
