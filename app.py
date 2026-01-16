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

# --- [에이전트 지능: 고도화된 규칙 엔진 V9.0] ---
def get_multiplier(text):
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

def get_prefix(filename):
    # 파일명 앞의 숫자 그룹 추출 (예: 01.01-1 -> 01.01)
    match = re.match(r'^([\d\.]+)', filename)
    if match:
        prefix = match.group(1).rstrip('.')
        # 하이픈이 있다면 하이픈 앞까지만 (예: 1-1 -> 1)
        if '-' in prefix: prefix = prefix.split('-')[0]
        return prefix
    return None

def analyze_file(filename, foldername, folder_instructions):
    fn = filename.lower()
    fd = foldername.lower()
    combined = fn + " " + fd + " " + folder_instructions.lower()
    
    res = {
        "is_usb": any(k in combined for k in ['usb', 'cd']),
        "is_vinyl": "비닐" in combined,
        "is_group_vinyl": any(k in combined for k in ['앞숫자', '앞번호', '앞뒤로', '같은문서']),
        "is_divider": any(k in combined for k in ['색지', '색간지', '간지', '탭지']),
        "is_special": any(k in combined for k in ['클립', '스테플러', '집게', '핀']),
        "is_binder": any(k in combined for k in ['cover', 'spine', 'face', '표지']),
        "is_toc": (any(k in fn for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b|_toc|toc_', fn) and 'protocol' not in fn)),
        "is_color": any(k in combined for k in ['컬러', '칼라', 'color'])
    }
    return res

# --- [메인 시스템] ---
st.set_page_config(page_title="사내 견적 에이전트 V9.0", layout="wide")
st.title("📂 무결점 사내 견적 에이전트 팀 (V9.0 - 그룹비닐 및 USB 중복제거)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 선택하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {} 
    usb_done_folders = set() # USB 카운트 완료된 폴더 추적
    vinyl_groups_done = {} # {폴더: set(이미처리된그룹번호)}

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            # 1. 먼저 폴더별 텍스트 파일(지시서) 내용 수집
            all_paths = z.namelist()
            folder_notes = {}
            for p in all_paths:
                if p.endswith('.txt'):
                    with z.open(p) as tf:
                        folder_notes[os.path.dirname(p)] = tf.read().decode('utf-8', errors='ignore')

            # 2. 파일 분석 시작
            valid_files = [f for f in all_paths if not f.startswith('__MACOSX') and not f.endswith('/') and not f.lower().endswith(('.doc', '.docx'))]
            
            for f in valid_files:
                filename = os.path.basename(f)
                foldername = os.path.dirname(f)
                top_folder = f.split('/')[0] if '/' in f else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "특수":0, "총파일수":0}
                if top_folder not in vinyl_groups_done:
                    vinyl_groups_done[top_folder] = set()

                if "출력x" in filename.lower(): continue

                instr = folder_notes.get(foldername, "")
                info = analyze_file(filename, foldername, instr)
                f_div, f_mul = get_multiplier(filename)
                fold_div, fold_mul = get_multiplier(foldername)
                final_mul = f_mul if f_mul > 1 else fold_mul
                final_div = f_div if f_div < 1.0 else fold_div
                ext = os.path.splitext(f)[1].lower()
                
                p_bw, p_color, m_divider, m_vinyl, m_usb, m_special = 0, 0, 0, 0, 0, 0

                # [USB 정산 - 폴더당 1개만]
                if info["is_usb"] and foldername not in usb_done_folders:
                    m_usb = 1
                    usb_done_folders.add(foldername)

                # [비닐 정산 - 그룹핑 로직]
                if info["is_vinyl"]:
                    prefix = get_prefix(filename)
                    if info["is_group_vinyl"] and prefix:
                        group_key = f"{top_folder}_{prefix}"
                        if group_key not in vinyl_groups_done[top_folder]:
                            m_vinyl = 1 # 그룹당 처음 한 번만 비닐 1개 추가
                            vinyl_groups_done[top_folder].add(group_key)
                    else:
                        m_vinyl = final_mul if any(k in filename for k in ['각', '각각']) else f_mul

                # [색간지/특수 정산]
                if info["is_divider"]: m_divider = final_mul
                if info["is_special"]: m_special = final_mul

                # [페이지 계산 및 총파일수 집계]
                raw_p = 0
                is_counted_file = False
                if ext in ['.pdf', '.pptx'] and not info["is_binder"] and not info["is_toc"] and not info["is_divider"] and not info["is_usb"]:
                    try:
                        with z.open(f) as fd:
                            f_stream = io.BytesIO(fd.read())
                            if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                            elif ext == '.pptx' and HAS_PPTX: raw_p = len(Presentation(f_stream).slides)
                            
                            p_val = math.ceil(raw_p * final_div) * final_mul
                            if info["is_color"]: p_color = p_val
                            else: p_bw = p_val
                            if p_val > 0: is_counted_file = True
                    except: pass

                # 합산
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["USB or CD"] += m_usb
                summary[top_folder]["특수"] += m_special
                if is_counted_file: summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "원본P": raw_p, "배수": f"{final_div}x{final_mul}",
                    "흑백": p_bw, "컬러": p_color, "비닐": m_vinyl, "USB": m_usb, "체크여부": "출력물" if is_counted_file else "부속/제외"
                })

        # 화면 출력
        st.subheader("📊 1. 최상위 폴더별 견적 요약 리포트 (V9.0)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        cols = ["흑백", "컬러", "색간지", "비닐", "USB or CD", "특수", "총파일수"]
        st.dataframe(sum_df[cols], use_container_width=True)
        
        st.subheader("🔍 2. 상세 계산 근거 (검증용)")
        st.dataframe(pd.DataFrame(detailed_log), use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df[cols].to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V9.0 최종 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_리포트_V9.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
