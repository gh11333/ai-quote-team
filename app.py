import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# PPT 지원 부품 (Import 에러 방지)
try:
    from pptx import Presentation
    HAS_PPTX = True
except:
    HAS_PPTX = False

# --- [에이전트 지능: 고도화된 규칙 엔진] ---
def get_multiplier(text):
    text = text.lower().replace(" ", "")
    # 1. 분할 인쇄 (나누기)
    div_val = 1.0
    div_match = re.search(r'(\d+)(?:페이지|up|쪽모아|쪽)', text)
    if div_match:
        val = int(div_match.group(1))
        if val in [2, 4, 6, 8, 16]: div_val = 1 / val
    
    # 2. 부수/장수 (곱하기)
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
        "is_special": any(k in combined for k in ['클립', '스테플러', '집게', '핀', '고정']),
        "is_binder": any(k in combined for k in ['cover', 'spine', 'face', '표지']),
        "is_toc": (any(k in fn for k in ['tableofcontents', '목차']) or (re.search(r'\btoc\b|_toc|toc_', fn) and 'protocol' not in fn)),
        "is_color": any(k in combined for k in ['컬러', '칼라', 'color'])
    }
    return res

# --- [메인 시스템] ---
st.set_page_config(page_title="사내 견적 에이전트 V8.1", layout="wide")
st.title("📂 무결점 사내 견적 에이전트 팀 (V8.1 - 오류수정 및 양식최적화)")

uploaded_zip = st.file_uploader("작업 폴더(ZIP)를 선택하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {} 

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_files = [f for f in z.namelist() if not f.startswith('__MACOSX') and not f.endswith('/')]
            # 워드 중복 제거 (PDF가 있으면 워드는 무시)
            valid_files = [f for f in all_files if not f.lower().endswith(('.doc', '.docx'))]
            
            for f in valid_files:
                path_parts = f.split('/')
                top_folder = path_parts[0] if path_parts else "Root"
                if top_folder not in summary:
                    # 사용자 요청 순서로 초기화
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB or CD":0, "특수":0, "TOC":0, "바인더":0, "총파일수":0}
                
                filename = os.path.basename(f)
                foldername = os.path.dirname(f)
                if "출력x" in filename.lower(): continue

                # 지능형 분석
                info = analyze_file(filename, foldername)
                f_div, f_mul = get_multiplier(filename)
                fold_div, fold_mul = get_multiplier(foldername)
                
                # 규칙 우선순위 적용
                final_mul = f_mul if f_mul > 1 else fold_mul
                final_div = f_div if f_div < 1.0 else fold_div
                
                # 오류 수정 부분: ext 추출 방식 변경
                ext = os.path.splitext(f)[1].lower()
                
                p_bw, p_color, m_divider, m_vinyl, m_usb, m_special, m_toc, m_binder = 0, 0, 0, 0, 0, 0, 0, 0

                # 1. 부자재 및 물건 정산
                if info["is_usb"]: m_usb = 1
                if info["is_vinyl"]: m_vinyl = final_mul if any(k in filename for k in ['각', '각각']) else f_mul
                if info["is_divider"]: m_divider = final_mul
                if info["is_special"]: m_special = final_mul
                if info["is_toc"]: m_toc = final_mul
                if info["is_binder"]: m_binder = final_mul

                # 2. 페이지 계산 (USB/CD가 포함되면 인쇄 페이지는 무조건 0)
                raw_p = 0
                if ext in ['.pdf', '.pptx'] and not info["is_binder"] and not info["is_toc"] and not info["is_divider"]:
                    if not info["is_usb"]: 
                        try:
                            with z.open(f) as fd:
                                f_stream = io.BytesIO(fd.read())
                                if ext == '.pdf': raw_p = len(PdfReader(f_stream).pages)
                                elif ext == '.pptx' and HAS_PPTX: raw_p = len(Presentation(f_stream).slides)
                                
                                # 올림 계산 적용 (85/4 = 22)
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
                summary[top_folder]["TOC"] += m_toc
                summary[top_folder]["바인더"] += m_binder
                summary[top_folder]["총파일수"] += 1

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "원본P": raw_p, "배수": f"{final_div}x{final_mul}",
                    "흑백": p_bw, "컬러": p_color, "색간지": m_divider, "비닐": m_vinyl, "USB": m_usb, "카테고리": "인쇄물" if not info["is_usb"] else "USB용"
                })

        # 화면 출력
        st.subheader("📊 1. 최상위 폴더별 견적 요약 리포트")
        sum_df = pd.DataFrame.from_dict(summary, orient='index')
        # 요청하신 컬럼 순서대로 정렬하여 출력
        display_cols = ["흑백", "컬러", "색간지", "비닐", "USB or CD", "특수", "TOC", "바인더", "총파일수"]
        st.dataframe(sum_df[display_cols], use_container_width=True)
        
        st.subheader("🔍 2. 상세 계산 근거 (검증용)")
        st.dataframe(pd.DataFrame(detailed_log), use_container_width=True)

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sum_df[display_cols].to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        st.download_button("📂 V8.1 무결점 견적서 다운로드", data=output.getvalue(), file_name="최종_견적_리포트_V8.1.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
