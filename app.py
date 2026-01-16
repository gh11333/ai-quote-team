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

# --- [에이전트 지능: 고도화된 규칙 엔진] ---
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

def analyze_file(filename, foldername):
    fn = filename.lower()
    fd = foldername.lower()
    combined = fn + " " + fd
    
    res = {
        "is_usb": any(k in combined for k in ['usb', 'cd']),
        "is_vinyl": "비닐" in combined,
        "is_divider": any(k in combined for k in ['색지', '색간지', '간지', '탭지']),
        "is_special": any(k in combined for k in ['클립', '스테플러', '집게']),
        "is_binder": any(k in combined for k in ['cover', 'spine', 'face', '표지']),
        "is_toc": (any(k in fn for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b|_toc|toc_', fn) and 'protocol' not in fn)),
        "is_color": any(k in combined for k in ['컬러', '칼라', 'color'])
    }
    return res

# --- [메인 시스템] ---
st.set_page_config(page_title="사내 견적 에이전트 V8.0", layout="wide")
st.title("📂 무결점 사내 견적 에이전트 팀 (V8.0 - USB/간지 완벽대응)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    # 사용자 요청 컬럼 순서로 요약판 구성
    summary = {} 

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_files = [f for f in z.namelist() if not f.startswith('__MACOSX') and not f.endswith('/')]
            valid_files = [f for f in all_files if not f.lower().endswith(('.doc', '.docx'))]
            
            for f in valid_files:
                path_parts = f.split('/')
                top_folder = path_parts[0] if path_parts else "Root"
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "특수":0, "총파일수":0}
                
                filename = os.path.basename(f)
                foldername = os.path.dirname(f)
                if "출력x" in filename.lower(): continue

                # 지능형 분석
                info = analyze_file(filename, foldername)
                f_div, f_mul = get_multiplier(filename)
                fold_div, fold_mul = get_multiplier(foldername)
                
                final_mul = f_mul if f_mul > 1 else fold_mul
                final_div = f_div if f_div < 1.0 else fold_div
                
                ext = os.path.splitext(f).lower()[1]
                p_bw, p_color, m_divider, m_vinyl, m_usb, m_special = 0, 0, 0, 0, 0, 0

                # 1. 자재 및 부자재 정산 (X장 수량 반영)
                if info["is_usb"]: m_usb = 1 # USB는 일단 1개
                if info["is_vinyl"]: m_vinyl = final_mul if any(k in filename for k in ['각', '각각']) else f_mul
                if info["is_divider"]: m_divider = final_mul
                if info["is_special"]: m_special = final_mul

                # 2. 페이지 계산 (USB/CD 포함 시 출력 제외 핵심 로직)
                raw_p = 0
                if ext in ['.pdf', '.pptx'] and not info["is_binder"] and not info["is_toc"]:
                    if not info["is_usb"]: # USB/CD가 아닐 때만 인쇄 페이지 계산
                        try:
                            with z.open(f) as fd:
                                f_stream = io.BytesIO(fd.read())
                                if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                                elif ext == '.pptx' and HAS_PPTX: raw_p = len(Presentation(f_stream).slides)
                                
                                calc_p = math.ceil(raw_p * final_div) * final_mul
                                if info["is_color"]: p_color = calc_p
                                else: p_bw = calc_p
                        except: pass

                # 데이터 합산
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["USB or CD"] += m_usb
                summary[top_folder]["특수"] += m_special
                summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "원본P": raw_p, "배수": f"{final_div}x{final_mul}",
                    "흑백": p_bw, "컬러": p_color, "색간지": m_divider, "비닐": m_vinyl, "USB": m_usb, "총파일수": 1
                })

        # 화면 출력
        st.subheader("📊 1. 최상위 폴더별 견적 요약 (최종 양식)")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        st.dataframe(sum_df[["흑백", "컬러", "색간지", "비닐", "USB or CD", "특수", "총파일수"]])
        
        st.subheader("🔍 2. 상세 계산 근거")
        st.dataframe(pd.DataFrame(detailed_log))

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df.to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V8.0 무결점 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_리포트_V8.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
