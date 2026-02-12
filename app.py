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
# [Agent 1] 해석가 (Instruction Parser)
# : 다양한 자연어 파일명을 표준 인쇄 옵션으로 번역
# ==========================================
class InstructionParser:
    @staticmethod
    def parse_n_up(text):
        """
        N-up(모아찍기) 정보를 다양한 한국어/영어 패턴에서 추출
        우선순위: 복합 표현(1면4쪽) > 명시적 표현(4up) > 관용구(4분할)
        """
        text = text.lower().replace(" ", "")
        
        # 패턴 1: 복합 표현 (예: "1면4쪽", "한면에2슬라이드", "한면두쪽")
        # '두'쪽 같은 한글 숫자도 처리하기 위해 매핑
        kor_num = {'한':1, '두':2, '세':3, '네':4, '1':1, '2':2, '3':3, '4':4, '6':6, '8':8, '9':9}
        
        # 정규식: (1|한)면(에)?(숫자|한글)쪽
        complex_match = re.search(r'(?:1|한)면(?:에)?(\d+|두|네)(?:쪽|슬라이드|페이지)', text)
        if complex_match:
            val = complex_match.group(1)
            return kor_num.get(val, int(val) if val.isdigit() else 1)

        # 패턴 2: 명시적 N-up (예: "4up", "2-up")
        up_match = re.search(r'(\d+)\s*-?up', text)
        if up_match:
            return int(up_match.group(1))

        # 패턴 3: 분할/쪽모아 (예: "4분할", "2쪽모아")
        split_match = re.search(r'(\d+)(?:분할|쪽모아)', text)
        if split_match:
            return int(split_match.group(1))
            
        # 패턴 4: 슬라이드 수만 적힌 경우 (예: "4슬라이드") - 4up으로 간주
        slide_match = re.search(r'(\d+)슬라이드', text)
        if slide_match:
            return int(slide_match.group(1))

        return 1 # 기본값 (1-up)

    @staticmethod
    def parse_filename(filename):
        name_lower = filename.lower().replace(" ", "")
        
        # 1. N-up 해석 (강화된 로직 적용)
        n_up = InstructionParser.parse_n_up(name_lower)

        # 2. 인쇄 부수 (Copies)
        copies = 1
        copy_match = re.search(r'(\d+)(?:부|권|copy|copies|set)', name_lower)
        if copy_match:
            copies = int(copy_match.group(1))

        # 3. 양면/단면 (표기용)
        is_duplex = True
        if any(k in name_lower for k in ['단면', 'single', 'simplex']):
            is_duplex = False
        
        # 4. 컬러/흑백
        is_color = False
        if any(k in name_lower for k in ['컬러', '칼라', 'color', 'rgb']):
            is_color = True

        # 5. 부자재(Materials) 파싱
        materials = {"비닐내지": 0, "색지": 0, "USB": 0}

        # 비닐내지
        if '비닐' in name_lower or '내지' in name_lower:
            cnt_match = re.search(r'(?:비닐|내지).*?(\d+)(?:장|개|매)?', name_lower)
            if cnt_match:
                materials["비닐내지"] = int(cnt_match.group(1))
            else:
                materials["비닐내지"] = 1

        # 색지 (파일명에 '뒤에색지' 포함 시 1장, '색지10장' 시 10장)
        if '색지' in name_lower or '간지' in name_lower:
            cnt_match = re.search(r'(?:색지|간지).*?(\d+)(?:장|개|매)?', name_lower)
            if cnt_match:
                materials["색지"] = int(cnt_match.group(1))
            else:
                materials["색지"] = 1

        # USB
        if 'usb' in name_lower:
            cnt_match = re.search(r'usb.*?(\d+)(?:개)?', name_lower)
            materials["USB"] = int(cnt_match.group(1)) if cnt_match else 1

        return {
            "n_up": n_up,
            "copies": copies,
            "is_duplex": is_duplex,
            "is_color": is_color,
            "materials": materials
        }

# ==========================================
# [Agent 2] 측정가 (Page Counter)
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
# [Main] 통합 관리자 (System)
# ==========================================
st.set_page_config(page_title="인쇄 견적 1차 집계 시스템", layout="wide")

st.title("🖨️ 인쇄/제본 1차 물량 산출기 (V3.0 - 해석 엔진 강화)")
st.info("업데이트: '1면4쪽', '한면에2슬라이드' 등 복합 인쇄 용어를 정확히 N-up으로 해석합니다.")

uploaded_file = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_file:
    grouped_data = {}
    grand_total = {"흑백": 0, "컬러": 0, "비닐": 0, "색지": 0, "USB": 0}

    with zipfile.ZipFile(uploaded_file, 'r') as z:
        file_list = [f for f in z.namelist() if not f.startswith('__') and not f.endswith('/')]
        
        for filepath in file_list:
            parts = filepath.split('/')
            top_folder = parts[0] if len(parts) > 1 else "Root"
            filename = os.path.basename(filepath)
            ext = os.path.splitext(filename)[1].lower()
            
            if top_folder not in grouped_data: grouped_data[top_folder] = []

            # 1. 해석 (Agent 1)
            spec = InstructionParser.parse_filename(filename)
            
            # 2. 측정 및 계산 (Agent 2)
            raw_pages = 0
            final_print_pages = 0
            print_type = "-"

            if ext in ['.pdf', '.pptx', '.ppt']:
                file_bytes = z.read(filepath)
                raw_pages = get_page_count(file_bytes, ext)
                
                if raw_pages > 0:
                    # N-up 적용 (올림 처리)
                    # 예: 100페이지 / 4up = 25페이지
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

            # 4. 결과 기록
            row_data = {
                "파일명": filename,
                "원본P": raw_pages,
                "해석결과": f"{spec['n_up']}쪽 모아찍기" if spec['n_up'] > 1 else "1쪽(기본)",
                "부수": f"{spec['copies']}부",
                "계산된페이지": final_print_pages,
                "분류": print_type,
                "비닐": mats["비닐내지"],
                "색지": mats["색지"],
                "USB": mats["USB"]
            }
            grouped_data[top_folder].append(row_data)

    # --- 화면 출력 ---
    st.markdown("### 📊 전체 총괄 합계")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("총 흑백(면)", grand_total["흑백"])
    c2.metric("총 컬러(면)", grand_total["컬러"])
    c3.metric("총 비닐(매)", grand_total["비닐"])
    c4.metric("총 색지(매)", grand_total["색지"])
    c5.metric("총 USB(개)", grand_total["USB"])
    
    st.divider()

    st.markdown("### 📂 폴더별 상세 명세서")
    for folder_name in sorted(grouped_data.keys()):
        rows = grouped_data[folder_name]
        df = pd.DataFrame(rows)
        
        sub_bw = df[df['분류']=='흑백']['계산된페이지'].sum()
        sub_color = df[df['분류']=='컬러']['계산된페이지'].sum()
        
        with st.expander(f"📁 {folder_name} (흑백: {sub_bw} / 컬러: {sub_color})", expanded=True):
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
