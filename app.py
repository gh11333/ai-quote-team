import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [에이전트 지능: 고도화된 전략 엔진 V21.0] ---
def get_multiplier(text):
    if not text: return 1.0, 1
    text = text.lower().replace(" ", "")
    div_val = 1.0
    # 분할 인쇄 패턴 인식 (2up, 4페이지 등)
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

def check_usb_instr(text):
    text_low = text.lower()
    if re.search(r'\busb\b|\bcd\b', text_low): return True
    if any(k in text_low for k in ['usb제작', 'cd제작', 'usb에', 'cd에']): return True
    return False

# --- [메인 시스템] ---
st.set_page_config(page_title="사내 견적 에이전트 V21.0", layout="wide")
st.title("📂 무결점 사내 견적 에이전트 팀 (V21.0 - 비닐 묶음/개별 정밀 정산)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {} 
    usb_counted_paths = set()
    print_keywords = ['up', '페이지', '장', '부', '쪽', '색지', '비닐', '간지', '클립', '양면', '3공']

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            raw_paths = z.namelist()
            
            # 1. 지휘 체계: 폴더별 지시사항 수집
            folder_txt_notes = {}
            folder_sibling_notes = {}
            for p in raw_paths:
                clean_p = p.replace('\\', '/').rstrip('/')
                d = os.path.dirname(clean_p)
                b = os.path.basename(clean_p)
                if not b: continue
                if b.lower().endswith('.txt'):
                    try:
                        with z.open(p) as tf:
                            folder_txt_notes[d] = folder_txt_notes.get(d, "") + " " + tf.read().decode('utf-8', errors='ignore')
                    except: pass
                if any(k in b.lower() for k in print_keywords):
                    folder_sibling_notes[d] = folder_sibling_notes.get(d, "") + " " + b

            # 2. 파일 분석 및 정산
            valid_files = [p for p in raw_paths if not p.endswith('/') and not p.lower().endswith(('.doc', '.docx', '.txt', '.msg'))]
            
            for f in valid_files:
                clean_f = f.replace('\\', '/').rstrip('/')
                filename = os.path.basename(clean_f)
                foldername = os.path.dirname(clean_f)
                top_folder = clean_f.split('/')[0] if '/' in clean_f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "TOC":0, "바인더":0, "총파일수":0}

                # 계층적 상속
                parent_instr = ""
                usb_source = ""
                curr = foldername
                while True:
                    local_instr = folder_txt_notes.get(curr, "") + " " + folder_sibling_notes.get(curr, "")
                    parent_instr += " " + local_instr + " " + os.path.basename(curr)
                    if check_usb_instr(local_instr + " " + os.path.basename(curr)) and not usb_source:
                        usb_source = curr
                    parent = os.path.dirname(curr)
                    if parent == curr or not curr: break
                    curr = parent
                
                combined_low = (filename + " " + parent_instr).lower()
                f_div, f_mul = get_multiplier(filename)
                txt_div, txt_mul = get_multiplier(parent_instr)
                
                final_mul = f_mul if f_mul > 1 else txt_mul
                final_div = f_div if f_div < 1.0 else txt_div
                
                cat = get_category(filename)
                ext = os.path.splitext(clean_f)[1].lower()
                p_bw, p_color, m_divider, m_vinyl, m_usb = 0, 0, 0, 0, 0

                # [USB 정산]
                if usb_source and usb_source not in usb_counted_paths:
                    m_usb = 1
                    usb_counted_paths.add(usb_source)
                
                # [부자재 정산 - 비닐 로직 핵심 수정]
                if "비닐" in combined_low:
                    # '각', '각각', '하나씩' 키워드가 있으면 부수만큼 곱함
                    if any(k in combined_low for k in ['각', '각각', '하나씩']):
                        m_vinyl = final_mul
                    else:
                        # 파일명에 '비닐 3장'처럼 직접 숫자가 붙어있는지 체크
                        vinyl_num_match = re.search(r'비닐(?:내지)?\s*(\d+)장', combined_low)
                        if vinyl_num_match:
                            m_vinyl = int(vinyl_num_match.group(1))
                        else:
                            m_vinyl = 1 # 별도 지시 없으면 묶음당 1개 (사용자 요청 반영)

                if any(k in combined_low for k in ['색지', '색간지', '간지', '탭지']):
                    m_divider = final_mul if any(k in filename.lower() for k in ['색지', '간지']) else 1

                # [페이지 계산]
                raw_p = 0
                is_printed = (ext in ['.pdf', '.pptx'] and cat in ["흑백", "컬러"] and not usb_source and "제작방식" not in filename)
                if cat in ["바인더세트", "TOC"]: is_printed = False

                if is_printed:
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                            p_val = math.ceil(raw_p * final_div) * final_mul
                            if cat == "컬러": p_color = p_val
                            else: p_bw = p_val
                    except: pass

                # 결과 합산
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["USB or CD"] += m_usb
                summary[top_folder]["TOC"] += (final_mul if cat == "TOC" else 0)
                summary[top_folder]["바인더"] += (final_mul if cat == "바인더세트" else 0)
                if is_printed and (p_bw + p_color > 0): summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "최종P": p_bw + p_color, "비닐": m_vinyl, "배수": f"{final_div}x{final_mul}"
                })

        st.subheader("📊 1. 최종 검증 요약 리포트 (V21.0)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        cols = ["흑백", "컬러", "색간지", "비닐", "USB or CD", "TOC", "바인더", "총파일수"]
        st.dataframe(sum_df[cols], use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df[cols].to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V21.0 최종 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_V21.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
