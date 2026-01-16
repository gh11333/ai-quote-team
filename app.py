import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# --- [지능형 에이전트 1: 의미론적 수량 추출기] ---
def extract_quantities(text):
    text = text.lower().replace(" ", "")
    # 1. 문서 배수 (x부, x장) - '비닐' 등과 붙어있지 않을 때만 배수로 인정
    mul_val = 1
    # '비닐'이나 '간지'가 없는 상태에서 '부' 또는 '장'이 오면 문서 배수
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        mul_match = re.search(r'(\d+)(?:부|장)', text)
        if mul_match: mul_val = int(mul_match.group(1))

    # 2. 분할 인쇄 (up/페이지)
    div_val = 1.0
    div_match = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽)', text)
    if div_match:
        val = int(div_match.group(1))
        if val in [2, 4, 6, 8]: div_val = 1 / val
        
    return div_val, mul_val

# --- [지능형 에이전트 2: 자재 수량 정밀 정산기] ---
def get_accessory_count(text, item_name, default_mul=1):
    text = text.lower().replace(" ", "")
    if item_name not in text: return 0
    # '아이템명 + 숫자 + 장/개' 패턴 검색 (예: 비닐10장)
    num_match = re.search(rf'{item_name}.*?(\d+)(?:장|개|매)', text)
    if num_match: return int(num_match.group(1))
    # '숫자 + 장/개 + 아이템명' 패턴 검색 (예: 10장비닐)
    num_match_rev = re.search(rf'(\d+)(?:장|개|매).*?{item_name}', text)
    if num_match_rev: return int(num_match_rev.group(1))
    # 키워드만 있으면: '각각'일 땐 배수만큼, 아니면 1개
    if any(k in text for k in ['각', '각각', '하나씩']): return default_mul
    return 1

# --- [지능형 에이전트 3: 인쇄 차단 판독기] ---
def is_skip_printing(text):
    t = text.lower()
    # CDMS, DOCX 등 일반 단어 속의 알파벳은 제외하고 '제작' 의미가 강할 때 skip
    if any(k in t for k in ['usb제작', 'cd제작', 'usb에', 'cd에', 'usb담기']): return True
    # 독립된 단어로서의 usb, cd
    if re.search(r'[^a-z]usb[^a-z]|[^a-z]cd[^a-z]', " " + t + " "):
        if 'cdms' not in t: return True
    return False

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V23.0", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V23.0 - 세만틱 정밀 감사 버전)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    usb_counted_paths = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            raw_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 사전 지휘소: 지시서 전수 스캔
            folder_notes = {}
            for p in raw_paths:
                clean_p = p.replace('\\', '/').rstrip('/')
                d, b = os.path.dirname(clean_p), os.path.basename(clean_p)
                if b.lower().endswith('.txt'):
                    try:
                        with z.open(p) as tf:
                            folder_notes[d] = folder_notes.get(d, "") + " " + tf.read().decode('utf-8', errors='ignore')
                    except: pass

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
                while True:
                    local_info = folder_notes.get(curr, "") + " " + os.path.basename(curr)
                    inherited_instr += " " + local_info
                    if is_skip_printing(local_info) and not skip_reason: skip_reason = curr
                    parent = os.path.dirname(curr)
                    if parent == curr or not curr: break
                    curr = parent

                # 배수 및 자재 산출
                combined_low = (filename + " " + inherited_instr).lower()
                f_div, f_mul = extract_quantities(filename)
                p_div, p_mul = extract_quantities(inherited_instr)
                
                final_mul = f_mul if f_mul > 1 else p_mul
                final_div = f_div if f_div < 1.0 else p_div
                
                # 자재 수량 (숫자 지시 우선)
                m_vinyl = get_accessory_count(combined_low, '비닐', final_mul)
                m_divider = get_accessory_count(combined_low, '색지', final_mul) or get_accessory_count(combined_low, '간지', final_mul)
                
                m_usb = 0
                if skip_reason and skip_reason not in usb_counted_paths:
                    m_usb = 1
                    usb_counted_paths.add(skip_reason)

                # 카테고리 및 인쇄 계산
                fn_low = filename.lower()
                cat = "흑백"
                if any(k in fn_low for k in ['cover', 'spine', 'face', '표지']): cat = "바인더"
                elif any(k in fn_low for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b', fn_low) and 'protocol' not in fn_low): cat = "TOC"
                elif any(k in fn_low for k in ['컬러', '칼라', 'color']): cat = "컬러"

                raw_p, p_bw, p_color = 0, 0, 0
                is_printed = (cat in ["흑백", "컬러"] and not skip_reason and "제작방식" not in filename)
                
                if is_printed and clean_f.lower().endswith(('.pdf', '.pptx')):
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            raw_p = len(PdfReader(f_stream).pages) if clean_f.lower().endswith('.pdf') else 0 # PPT는 HAS_PPTX 생략시 0
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

        st.subheader("📊 1. V23.0 정밀 감사 요약 리포트")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        st.dataframe(sum_df[["흑백", "컬러", "색간지", "비닐", "USB or CD", "TOC", "바인더", "총파일수"]], use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df.to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V23.0 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V23.xlsx")

    except Exception as e:
        st.error(f"시스템 오류: {e}")
