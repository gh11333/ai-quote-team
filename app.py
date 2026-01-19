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

# --- [에이전트 A: 수치 추출 엔진 - 안전성 극대화] ---
def extract_value(text, pattern):
    if not text: return None
    # 공백 제거 후 검색하여 변종 지시에 대응
    m = re.search(pattern, text.lower().replace(" ", ""))
    if m:
        # 매칭된 그룹 중 None이 아닌 첫 번째 값을 찾아 숫자로 변환
        for g in m.groups():
            if g is not None:
                try:
                    return int(g)
                except:
                    continue
    return None

# --- [에이전트 B: 카테고리 판독기] ---
def get_category(filename, context_text):
    fn = filename.lower()
    # 1순위: 바인더 부속 (인쇄 페이지 제외)
    if any(k in fn for k in ['face', 'spine', 'cover', '표지', 'binder']): return "바인더"
    # 2순위: 목차
    if any(k in fn for k in ['toc', '목차']): return "TOC"
    # 3순위: 컬러 여부
    if any(k in fn or k in context_text.lower() for k in ['컬러', 'color', '칼라']): return "컬러"
    return "흑백"

# --- [에이전트 C: 인쇄 차단 판독기] ---
def is_usb_folder(text):
    t = text.lower().replace(" ", "")
    # 실무 키워드 및 CDMS 예외 처리
    if any(k in t for k in ['usb', 'cd제작', 'usb제작', 'usb담기']):
        return 'cdms' not in t
    return False

# --- [메인 시스템] ---
st.set_page_config(page_title="최종 병기 V35.1", layout="wide")
st.title("📂 2026 사내 견적 자동화 시스템 (V35.1 - 런타임 오류 해결)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    
    # 중복 합산 방지 장치
    processed_fixed_materials = set() 
    processed_usb_folders = set()

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시사항 전수 DB화 (폴더 단위 지시 수집)
            db = {}
            for p in all_paths:
                d = os.path.dirname(p).replace('\\', '/')
                if d not in db: db[d] = {"instrs": [os.path.basename(d)], "folder_name": os.path.basename(d)}
                if p.lower().endswith('.txt'):
                    fname = os.path.basename(p)
                    db[d]["instrs"].append(fname)
                    try:
                        with z.open(p) as f:
                            content = f.read().decode('utf-8', errors='ignore')
                            if content.strip(): db[d]["instrs"].append(content)
                    except: pass

            # 2. 메인 정산 루프
            for p in all_paths:
                # 불필요한 파일 필터링
                if p.endswith('/') or any(k in p.lower() for k in ['.doc', '.docx', '.msg', '출력x']): continue
                
                # 변수 초기화
                raw_p, p_bw, p_color, final_p = 0, 0, 0, 0
                m_vinyl, m_divider, m_special = 0, 0, 0
                
                clean_p = p.replace('\\', '/')
                filename = os.path.basename(clean_p)
                foldername = os.path.dirname(clean_p)
                top_folder = clean_p.split('/')[0] if '/' in clean_p else "Root"
                
                if top_folder not in summary:
                    summary[top_folder] = {"흑백":0, "컬러":0, "색간지":0, "비닐":0, "USB":0, "TOC":0, "바인더":0, "특수":0, "총파일수":0}

                # [계층 구조 분석 - 상속 경로 생성]
                path_nodes = []
                curr = foldername
                while True:
                    path_nodes.append(curr)
                    if not curr or curr == '.': break
                    curr = os.path.dirname(curr)

                # [규칙 결정 - 우선순위 스택 로직]
                final_div, final_mul = 1.0, 1
                div_found, mul_found = False, False
                
                # A. 파일명 규칙 우선 적용
                d = extract_value(filename, r'(\d+)(?:up|페이지|쪽|면|쪽모아)')
                m = extract_value(filename, r'(\d+)(?:부|장)')
                if d: final_div, div_found = 1/d, True
                if m: final_mul, mul_found = m, True
                
                # B. 상위 폴더로 올라가며 빈자리 상속
                for node in path_nodes:
                    node_texts = db.get(node, {}).get("instrs", [])
                    for text in node_texts:
                        if not div_found:
                            d = extract_value(text, r'(\d+)(?:up|페이지|쪽|면|쪽모아)')
                            if d: final_div, div_found = 1/d, True
                        if not mul_found:
                            m = extract_value(text, r'(\d+)(?:부|장)')
                            if m: final_mul, mul_found = m, True

                # [자재 및 인쇄 차단 판독 환경 조성]
                context_texts = []
                for node in path_nodes: context_texts.extend(db.get(node, {}).get("instrs", []))
                context_all = " ".join(context_texts) + " " + filename
                
                # [자재 정산 - EACH와 FIXED 엔진]
                for item, keys in {"비닐": ["비닐"], "색간지": ["간지", "색지"], "특수": ["라벨", "스티커", "카드", "클립"]}.items():
                    # 1. 고정 수량 (특정 폴더에 적힌 숫자는 딱 한번만 합산)
                    local_instrs = db.get(foldername, {}).get("instrs", [])
                    for instr in local_instrs:
                        val = extract_value(instr, rf'{keys[0]}.*?(\d+)|(\d+).*?{keys[0]}')
                        if val:
                            key_id = f"{foldername}_{item}_{val}"
                            if key_id not in processed_fixed_materials:
                                if item == "비닐": m_vinyl += val
                                elif item == "색간지": m_divider += val
                                else: m_special += val
                                processed_fixed_materials.add(key_id)
                    
                    # 2. 개별 수량 (상위 어디든 '각'이 있으면 파일당 합산)
                    if any(k in context_all.lower() for k in keys):
                        if any(x in context_all.lower() for x in ['각', '각각', '하나씩']):
                            if item == "비닐": m_vinyl += (1 * final_mul)
                            elif item == "색간지": m_divider += (1 * final_mul)
                            else: m_special += (1 * final_mul)

                # [카테고리 분류 및 인쇄 제외 처리]
                cat = get_category(filename, context_all)
                
                if is_usb_folder(context_all):
                    cat = "SKIP(USB)"
                    if foldername not in processed_usb_folders:
                        summary[top_folder]["USB"] += 1
                        processed_usb_folders.add(foldername)

                # [최종 페이지 연산]
                if cat in ["흑백", "컬러"]:
                    try:
                        with z.open(p) as f:
                            f_stream = io.BytesIO(f.read())
                            if p.lower().endswith('.pdf'): raw_p = len(PdfReader(f_stream).pages)
                            elif p.lower().endswith('.pptx') and Presentation: raw_p = len(Presentation(f_stream).slides)
                        
                        # math.ceil 적용하여 정확한 출력 장수 산출
                        final_p = math.ceil(raw_p * final_div) * final_mul
                        if cat == "컬러": p_color = final_p
                        else: p_bw = final_p
                        summary[top_folder]["총파일수"] += 1
                    except: pass

                # 요약표 반영
                summary[top_folder]["흑백"] += p_bw
                summary[top_folder]["컬러"] += p_color
                summary[top_folder]["비닐"] += m_vinyl
                summary[top_folder]["색간지"] += m_divider
                summary[top_folder]["특수"] += m_special
                if cat == "TOC": summary[top_folder]["TOC"] += final_mul
                if cat == "바인더": summary[top_folder]["바인더"] += final_mul

                detailed_log.append({
                    "폴더": top_folder, "파일명": filename, "분류": cat, "원본P": raw_p, 
                    "계산식": f"{final_div}up x {final_mul}부", "최종P": final_p, "비닐": m_vinyl
                })

        st.subheader("📊 V35.1 정밀 요약 리포트 (오류 수정 완료)")
        st.dataframe(pd.DataFrame.from_dict(summary, orient='index'), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame.from_dict(summary, orient='index').to_excel(writer, sheet_name='요약')
            pd.DataFrame(detailed_log).to_excel(writer, sheet_name='상세')
        st.download_button("📂 V35.1 정산서 다운로드", data=output.getvalue(), file_name="최종_견적_V35_1.xlsx")

    except Exception as e:
        st.error(f"시스템 오류 발생: {e}")
