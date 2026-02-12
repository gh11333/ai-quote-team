import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation

# ==========================================
# 1. 핵심 파싱 엔진 (파일명 해석기)
# ==========================================
class InstructionParser:
    @staticmethod
    def parse_filename(filename):
        name_lower = filename.lower().replace(" ", "")
        
        # 1. N-up (모아찍기) - 기본값 1
        n_up = 1
        n_up_match = re.search(r'(\d+)(?:up|쪽|분할|면|슬라이드)', name_lower)
        if n_up_match:
            n_up = int(n_up_match.group(1))

        # 2. 인쇄 부수 (Copies) - 기본값 1
        copies = 1
        copy_match = re.search(r'(\d+)(?:부|권|copy|copies|set)', name_lower)
        if copy_match:
            copies = int(copy_match.group(1))

        # 3. 양면/단면 여부 (표기용, 계산에서는 제외)
        is_duplex = True # 기본 양면
        if any(k in name_lower for k in ['단면', 'single', 'simplex']):
            is_duplex = False
        
        # 4. 컬러/흑백
        is_color = False
        if any(k in name_lower for k in ['컬러', '칼라', 'color', 'rgb']):
            is_color = True

        # --- 부자재(Material) 파싱 ---
        materials = {
            "비닐내지": 0,
            "색지": 0,
            "USB": 0
        }

        # 비닐내지/내지
        if '비닐' in name_lower or '내지' in name_lower:
            # "비닐10장" 처럼 숫자가 붙어있는 경우
            cnt_match = re.search(r'(?:비닐|내지).*?(\d+)(?:장|개|매)?', name_lower)
            if cnt_match:
                materials["비닐내지"] = int(cnt_match.group(1))
            else:
                materials["비닐내지"] = 1 # 언급만 있으면 1개

        # 색지/간지
        if '색지' in name_lower or '간지' in name_lower:
            cnt_match = re.search(r'(?:색지|간지).*?(\d+)(?:장|개|매)?', name_lower)
            if cnt_match:
                materials["색지"] = int(cnt_match.group(1))
            else:
                # "뒤에 색지" 같은 경우, 파일당 1장으로 처리
                materials["색지"] = 1

        # USB
        if 'usb' in name_lower:
            cnt_match = re.search(r'usb.*?(\d+)(?:개)?', name_lower)
            if cnt_match:
                materials["USB"] = int(cnt_match.group(1))
            else:
                materials["USB"] = 1

        return {
            "n_up": n_up,
            "copies": copies,
            "is_duplex": is_duplex,
            "is_color": is_color,
            "materials": materials
        }

# ==========================================
# 2. 페이지 측정 엔진
# ==========================================
def get_page_count(file_bytes, ext):
    try:
        if ext == '.pdf':
            reader = PdfReader(io.BytesIO(file_bytes))
            return len(reader.pages)
        elif ext in ['.pptx', '.ppt']:
            prs = Presentation(io.BytesIO(file_bytes))
            return len(prs.slides)
        return 0
    except:
        return 0

# ==========================================
# 3. 메인 앱 (Streamlit)
# ==========================================
st.set_page_config(page_title="인쇄 견적 1차 집계 시스템", layout="wide")

st.title("🖨️ 인쇄/제본 1차 물량 산출기 (V2.1)")
st.info("수정사항: 폴더별 자동 분류 기능 추가, 양면인쇄 시 페이지 나누기 로직 삭제 (페이지 수 그대로 계산)")

uploaded_file = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_file:
    # 데이터 구조: { "폴더명": [파일결과리스트], ... }
    grouped_data = {}
    
    # 전체 합계용
    grand_total = {"흑백": 0, "컬러": 0, "비닐": 0, "색지": 0, "USB": 0}

    with zipfile.ZipFile(uploaded_file, 'r') as z:
        file_list = [f for f in z.namelist() if not f.startswith('__') and not f.endswith('/')]
        
        for filepath in file_list:
            # 경로 분리 (최상위 폴더 추출)
            parts = filepath.split('/')
            top_folder = parts[0] if len(parts) > 1 else "최상위 경로(Root)"
            filename = os.path.basename(filepath)
            ext = os.path.splitext(filename)[1].lower()
            
            # 그룹 초기화
            if top_folder not in grouped_data:
                grouped_data[top_folder] = []

            # 1. 파일명 파싱
            spec = InstructionParser.parse_filename(filename)
            
            # 2. 페이지 계산
            raw_pages = 0
            final_print_pages = 0
            print_type = "-"

            # PDF/PPTX 처리
            if ext in ['.pdf', '.pptx', '.ppt']:
                file_bytes = z.read(filepath)
                raw_pages = get_page_count(file_bytes, ext)
                
                if raw_pages > 0:
                    # [수정된 로직]
                    # 원본 161p, 4up -> 40.25 -> 41페이지 (양면 여부 상관없이 41면 출력)
                    pages_n_up = math.ceil(raw_pages / spec['n_up'])
                    
                    # 부수 적용
                    final_print_pages = pages_n_up * spec['copies']
                    
                    if spec['is_color']:
                        print_type = "컬러"
                        grand_total["컬러"] += final_print_pages
                    else:
                        print_type = "흑백"
                        grand_total["흑백"] += final_print_pages

            elif ext == '.txt':
                print_type = "지시서"

            # 3. 부자재 집계
            mats = spec['materials']
            grand_total["비닐"] += mats["비닐내지"]
            grand_total["색지"] += mats["색지"]
            grand_total["USB"] += mats["USB"]

            # 4. 데이터 저장
            row_data = {
                "파일명": filename,
                "원본P": raw_pages,
                "설정": f"{spec['n_up']}up / {spec['copies']}부",
                "참고": "양면" if spec['is_duplex'] else "단면", # 참고용 텍스트
                "계산된페이지": final_print_pages,
                "분류": print_type,
                "비닐": mats["비닐내지"],
                "색지": mats["색지"],
                "USB": mats["USB"]
            }
            grouped_data[top_folder].append(row_data)

    # --- 결과 출력 ---
    
    # 1. 전체 요약 (맨 위)
    st.markdown("### 📊 전체 총괄 합계")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("총 흑백(면)", grand_total["흑백"])
    c2.metric("총 컬러(면)", grand_total["컬러"])
    c3.metric("총 비닐(매)", grand_total["비닐"])
    c4.metric("총 색지(매)", grand_total["색지"])
    c5.metric("총 USB(개)", grand_total["USB"])
    
    st.divider()

    # 2. 폴더별 상세 내역 (반복문)
    st.markdown("### 📂 폴더별 상세 명세서")
    
    # 폴더 이름을 정렬해서 출력
    for folder_name in sorted(grouped_data.keys()):
        rows = grouped_data[folder_name]
        df = pd.DataFrame(rows)
        
        # 해당 폴더의 소계 계산
        sub_bw = df[df['분류']=='흑백']['계산된페이지'].sum()
        sub_color = df[df['분류']=='컬러']['계산된페이지'].sum()
        sub_vinyl = df['비닐'].sum()
        
        with st.expander(f"📁 {folder_name} (흑백: {sub_bw} / 컬러: {sub_color} / 비닐: {sub_vinyl})", expanded=True):
            st.dataframe(
                df,
                column_config={
                    "계산된페이지": st.column_config.NumberColumn("인쇄수량(면)", format="%d"),
                    "비닐": st.column_config.NumberColumn("비닐", format="%d"),
                    "색지": st.column_config.NumberColumn("색지", format="%d"),
                },
                use_container_width=True,
                hide_index=True
            )

    # (선택사항) 전체 엑셀 다운로드 준비
    all_rows = []
    for folder, rows in grouped_data.items():
        for r in rows:
            r['상위폴더'] = folder # 엑셀에는 폴더명 포함
            all_rows.append(r)
    
    final_df = pd.DataFrame(all_rows)
    # 컬럼 순서 조정
    cols = ['상위폴더', '파일명', '원본P', '설정', '참고', '계산된페이지', '분류', '비닐', '색지', 'USB']
    final_df = final_df[cols]
