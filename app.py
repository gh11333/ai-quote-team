import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader

# PPTX 라이브러리 체크
try:
    from pptx import Presentation
except ImportError:
    Presentation = None

# --- [정밀 추출 엔진] ---
def get_rules(text):
    """배수(부) 및 N-up 규칙 추출"""
    text = text.lower().replace(" ", "")
    div, mul = 1.0, None
    
    # 1. 배수(부수) 추출: 숫자 + 부/장
    if not any(k in text for k in ['비닐', '간지', '색지', '탭지']):
        m_mul = re.search(r'(\d+)(?:부|장)', text)
        if m_mul: mul = int(m_mul.group(1))
    
    # 2. N-up 추출: 2, 4, 6, 9, 16 대응
    m_div = re.search(r'(\d+)(?:up|페이지|쪽|면|쪽모아)', text)
    if m_div:
        val = int(m_div.group(1))
        if val in [2, 4, 6, 9, 16]: div = 1 / val
    return div, mul

def get_special_count(text, keywords, default_mul=1):
    """특수 자재(라벨, 스티커, 카드, 비닐 등) 수량 및 항목명 추출"""
    text = text.lower().replace(" ", "")
    found_item = ""
    count = 0
    for k in keywords:
        if k in text:
            found_item = k
            # 숫자 패턴 검색 (키워드 앞뒤)
            m = re.search(rf'{k}.*?(\d+)(?:장|개|매)|(\d+)(?:장|개|매).*?{k}', text)
            if m:
                # 두 그룹 중 매칭된 숫자 선택
                g1 = m.group(1)
                g2 = m.group(2)
                count = int(g1 if g1 else g2)
            else:
                # 숫자 지시가 없는데 '각각' 키워드가 있으면 부수만큼, 아니면 1개
                count = default_mul if any(x in text for x in ['각', '각각', '하나씩']) else 1
            break
    return count, found_item

# --- [메인 시스템] ---
st.set_page_config(page_title="무결점 에이전트 V28.1", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V28.1 - 무오류 정밀 버전)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_folders = set() # 폴더당 자재/특수항목 1회 합산용
    usb_counted_top = set()   # 최상위 폴더당 USB 1회 합산용

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            # __MACOSX 제외 전체 경로 리스트
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시서 및 구조 데이터베이스 구축
            db = {}
            for p in all_paths:
                d = os.path.dirname(p)
                if d not in db: db[d] = {"txt": "", "name": os.path.basename(d)}
                if p.lower().endswith('.txt'):
                    try:
                        with z.open(p) as f:
                            db[d]["txt"] += f.read().decode('utf-8', errors='ignore')
                    except: pass

            # 2. 정산 엔진 가동
            for p in all_paths:
                # 파일만 필터링 (폴더 제외 및 특정 확장자 무시)
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']):
                    continue
                
                # [변수 초기화]
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_special, special_note = 0, ""
                m_vinyl, m_divider = 0, 0
                
                clean_p = p.replace('\\', '/')
                filename = os.path.basename(clean_p)
                foldername = os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                # 요약표 초기화
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}

                # 지시 상속 (파일명 + 폴더명 + 지시서)
                all_instr = filename + " " + db.get(foldername, {}).get("name", "") + " " + db.get(foldername, {}).get("txt", "")
                
                # 배수 및 N-up 확정
                f_div, f_mul = get_rules(filename)
                p_div, p_mul = get_rules(db.get(foldername, {}).get("name", "") + " " + db.get(foldername, {}).get("txt", ""))
                
                final_div = f_div if f_div < 1.0 else p_div
                final_mul = f_mul if f_mul is not None else (p_mul if p_mul is not None else 1)

                # [카테고리 판정 - Binder 우선]
                cat = "흑백"
                if any(k in filename.lower() for k in ['face', 'spine', 'cover', '표지', 'binder']):
                    cat = "바인더"
                elif any(k in filename.lower() for k in ['toc', '목차']):
                    cat = "TOC"
                elif any(k in all_instr.lower() for k in ['컬러', 'color', '칼라']):
                    cat = "컬러"

                # [인쇄 차단 - USB/CD (단어 경계 체크)]
                if re.search(r'\b(usb|cd)\b', all_instr.lower()) and 'cdms' not in all_instr.lower():
                    cat = "SKIP(USB)"
                    if top_folder not in usb_counted_top:
                        summary[top_folder]["USB"] = 1
                        usb_counted_top.add(top_folder)

                # [자재 및 특수 단어 정산 - 폴더당 1회 합산]
                if foldername not in processed_folders:
                    m_vinyl, _ = get_special_count(all_instr, ["비닐"], final_mul)
                    m_divider, _ = get_special_count(all_instr, ["간지", "색지", "탭지"], final_mul)
                    # 클립, 카드, 라벨, 스티커 정산
                    m_special, special_note = get_special_count(all_instr, ["클립", "카드", "라벨", "스티커"], final_mul)
                    processed_folders.add(foldername)

                # [페이지 계산]
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as f:
                            f_stream = io.BytesIO(f.read())
                            if p.lower().endswith('.pdf'):
                                raw_p = len(PdfReader(f_stream).pages)
                            elif p.lower().endswith('.pptx') and Presentation:
                                raw_p = len(Presentation(f_stream).slides)
                        
                        # 공식: math.ceil(원본P * 분할배수) * 출력부수
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                        summary[top_folder]["총파일수"] += 1
                    except:
                        pass

                # [결과 요약 업데이트]
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                # 상세 로그 기록
                detailed_log.append({
                    "폴더": top_folder, 
                    "파일명": filename, 
                    "분류": cat, 
                    "원본P": raw_p, 
                    "계산식": f"{final_div}up x {final_mul}부", 
                    "최종P": final_p, 
                    "비닐": m_vinyl, 
                    "특수항목": special_note, 
                    "특수수량": m_special
                })

        # --- [화면 출력 및 다운로드] ---
        st.subheader("📊 V28.1 최종 요약 리포트")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='최종요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세근거')
        
        st.download_button(
            label="📂 V28.1 정산서(Excel) 다운로드",
            data=output.getvalue(),
            file_name="사내_견적_정산_V28_1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
