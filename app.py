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

# --- [지능형 에이전트 1: 수량 및 자재 추출기] ---
def extract_quantities(text):
    text = text.lower().replace(" ", "")
    mul_val = 1
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        mul_match = re.search(r'(\d+)(?:부|장)', text)
        if mul_match: mul_val = int(mul_match.group(1))

    div_val = 1.0
    div_match = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽)', text)
    if div_match:
        val = int(div_match.group(1))
        if val in [2, 4, 6, 8]: div_val = 1 / val
        
    return div_val, mul_val

def get_accessory_list(text, item_name):
    """텍스트에서 특정 자재의 숫자 지시를 모두 찾아 리스트로 반환 (중복 검토용)"""
    text = text.lower().replace(" ", "")
    if item_name not in text: return []
    
    # '비닐10장', '10장비닐' 패턴 모두 추출
    matches = re.findall(rf'{item_name}.*?(\d+)(?:장|개|매)', text)
    matches += re.findall(rf'(\d+)(?:장|개|매).*?{item_name}', text)
    
    return [int(m) for m in matches]

def is_skip_printing(text):
    t = text.lower()
    if any(k in t for k in ['usb제작', 'cd제작', 'usb에', 'cd에', 'usb담기']): return True
    if re.search(r'[^a-z]usb[^a-z]|[^a-z]cd[^a-z]', " " + t + " "):
        if 'cdms' not in t: return True
    return False

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V24.1", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V24.1 - 자재 중복 제거 버전)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    
    # 중복 계산 방지용 셋
    processed_folder_accessories = set() 
    usb_counted_top_folders = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            raw_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 사전 스캔: 폴더별 지시서 및 이름 수집
            folder_info = {} # {폴더경로: [지시문구들]}
            sibling_names = {}

            for p in raw_paths:
                clean_p = p.replace('\\', '/').rstrip('/')
                d, b = os.path.dirname(clean_p), os.path.basename(clean_p)
                
                if d not in folder_info: folder_info[d] = [os.path.basename(d)]
                
                if b.lower().endswith('.txt'):
                    try:
                        with z.open(p) as tf:
                            content = tf.read().decode('utf-8', errors='ignore')
                            folder_info[d].append(content)
                    except: pass
                
                if not ('.' in b): # 폴더인 경우 형제 정보 수집
                    parent_dir = os.path.dirname(d)
                    sibling_names[parent_dir] = sibling_names.get(parent_dir, "") + " " + b

            # 2. 메인 분석 엔진
            valid_files = [p for p in raw_paths if not p.endswith('/') and not p.lower().endswith(('.doc', '.docx', '.txt', '.msg'))]
            
            for f in valid_files:
                clean_f = f.replace('\\', '/').rstrip('/')
                filename, foldername = os.path.basename(clean_f), os.path.dirname(clean_f)
                top_folder = clean_f.split('/')[0] if '/' in clean_f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "TOC":0, "바인더":0, "총파일수":0}

                # 계층적 지시 수집
                inherited_instr = ""
                skip_reason = ""
                curr = foldername
                path_trace = []
                while True:
                    path_trace.append(curr)
                    local_info = " ".join(folder_info.get(curr, [])) + " " + sibling_names.get(os.path.dirname(curr), "")
                    inherited_instr += " " + local_info
                    if is_skip_printing(local_info) and not skip_reason: skip_reason = curr
                    parent = os.path.dirname(curr)
                    if parent == curr or not curr: break
                    curr = parent

                # 배수 산출
                combined_low = (filename + " " + inherited_instr).lower()
                f_div, f_mul = extract_quantities(filename)
                p_div, p_mul = extract_quantities(inherited_instr)
                final_mul = f_mul if f_mul > 1 else p_mul
                final_div = f_div if f_div < 1.0 else p_div

                # --- [자재 정산: 폴더당 1회만 합산] ---
                m_vinyl, m_divider = 0, 0
                for path in path_trace:
                    access_key = f"{path}_{item_name if 'item_name' in locals() else 'access'}"
                    if access_key not in processed_folder_accessories:
                        # 해당 폴더의 모든 지시문구에서 숫자 추출
                        raw_instrs = folder_info.get(path, [])
                        unique_counts = set()
                        for instr in raw_instrs:
                            counts = get_accessory_list(instr, '비닐')
                            for c in counts: unique_counts.add(c) # 중복 숫자(예: 10, 10)는 하나로 취급
                        
                        m_vinyl += sum(unique_counts)
                        
                        # 간지/색지도 동일 로직
                        div_counts = set()
                        for instr in raw_instrs:
                            for k in ['색지', '간지']:
                                for c in get_accessory_list(instr, k): div_counts.add(c)
                        m_divider += sum(div_counts)
                        
                        processed_folder_accessories.add(access_key)

                # USB 정산: 최상위 폴더당 1개로 제한
                m_usb = 0
                if skip_reason and top_folder not in usb_counted_top_folders:
                    m_usb = 1
                    usb_counted_top_folders.add(top_folder)

                # 카테고리 분류 및 인쇄 계산
                fn_low = filename.lower()
                cat = "흑백"
                if any(k in fn_low for k in ['cover', 'spine', 'face', '표지']): cat = "바인더"
                elif any(k in fn_low for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b', fn_low) and 'protocol' not in fn_low): cat = "TOC"
                elif any(k in fn_low for k in ['컬러', '칼라', 'color']): cat = "컬러"

                raw_p, p_bw, p_color = 0, 0, 0
                is_printed = (cat in ["흑백", "컬러"] and not skip_reason and "제작방식" not in filename)
                
                if is_printed:
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if clean_f.lower().endswith('.pdf'):
                                raw_p = len(PdfReader(f_stream).pages)
                            elif clean_f.lower().endswith('.pptx') and Presentation:
                                raw_p = len(Presentation(f_stream).slides)
                        p_val = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = p_val
                        else: p_bw = p_val
                    except: pass

                # 결과 집계
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["USB or CD"] += m_usb
                summary[top_folder]["TOC"] += (final_mul if cat == "TOC" else 0)
                summary[top_folder]["바인더"] += (final_mul if cat == "바인더" else 0)
                if is_printed and (p_bw + p_color > 0): summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "원본P": raw_p, "배수": f"{final_div}x{final_mul}", "최종P": p_bw + p_color, "비닐": m_vinyl, "간지": m_divider
                })

        st.subheader("📊 V24.1 정밀 감사 리포트 (자재 중복 제거 적용)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        st.dataframe(sum_df[["흑백", "컬러", "색간지", "비닐", "USB or CD", "TOC", "바인더", "총파일수"]], use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df.to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V24.1 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V24_1.xlsx")

    except Exception as e:
        st.error(f"시스템 오류: {e}")
