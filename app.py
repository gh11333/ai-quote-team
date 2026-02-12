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
        """
        파일명에서 인쇄 옵션(N-up, 부수, 컬러여부)과
        부자재 수량(비닐, 색지, USB)을 추출합니다.
        """
        name_lower = filename.lower().replace(" ", "")
        
        # --- [A] 인쇄 옵션 파싱 ---
        
        # 1. N-up (모아찍기)
        # 예: 4up, 4쪽, 1면4쪽, 4슬라이드
        n_up = 1
        n_up_match = re.search(r'(\d+)(?:up|쪽|분할|면|슬라이드)', name_lower)
        if n_up_match:
            n_up = int(n_up_match.group(1))

        # 2. 인쇄 부수 (Copies)
        # 예: 3부, 5권, 10copy -> 인쇄물에만 적용되는 곱하기 인자
        copies = 1
        copy_match = re.search(r'(\d+)(?:부|권|copy|copies|set)', name_lower)
        if copy_match:
            copies = int(copy_match.group(1))

        # 3. 양면/단면 (기본값: 양면)
        is_duplex = True
        if any(k in name_lower for k in ['단면', 'single', 'simplex']):
            is_duplex = False
        # 파일명에 '양면'이 명시되면 확실히 양면
        if any(k in name_lower for k in ['양면', 'double', 'duplex']):
            is_duplex = True

        # 4. 컬러/흑백 (기본값: 흑백)
        is_color = False
        if any(k in name_lower for k in ['컬러', '칼라', 'color', 'rgb']):
            is_color = True

        # --- [B] 부자재(Material) 파싱 (독립적 수량) ---
        
        materials = {
            "비닐내지": 0,
            "색지": 0,
            "USB": 0,
            "바인더": 0
        }

        # 1. 비닐내지/내지
        if '비닐' in name_lower or '내지' in name_lower:
            # "비닐10장", "비닐내지3개" 처럼 숫자가 붙어있는 경우
            cnt_match = re.search(r'(?:비닐|내지).*?(\d+)(?:장|개|매)?', name_lower)
            if cnt_match:
                materials["비닐내지"] = int(cnt_match.group(1))
            else:
                # 숫자가 없으면 기본 1장 (사용자 피드백: 비닐내지 = 1)
                materials["비닐내지"] = 1

        # 2. 색지/간지
        if '색지' in name_lower or '간지' in name_lower:
            cnt_match = re.search(r'(?:색지|간지).*?(\d+)(?:장|개|매)?', name_lower)
            if cnt_match:
                materials["색지"] = int(cnt_match.group(1))
            else:
                materials["색지"] = 1 # 언급은 있는데 수량 없으면 1장

        # 3. USB
        if 'usb' in name_lower:
            cnt_match = re.search(r'usb.*?(\d+)(?:개)?', name_lower)
            if cnt_match:
                materials["USB"] = int(cnt_match.group(1))
            else:
                materials["USB"] = 1

        # 4. 바인더 (폴더나 파일명에 바인더 언급 시)
        if '바인더' in name_lower:
             # 보통 바인더는 파일 자체가 아니라 결과물이므로 여기서는 카운트가 모호하나,
             # 3대 바인더(제안서 등)인 경우를 위해 로직 추가
             materials["바인더"] = 1

        return {
            "n_up": n_up,
            "copies": copies,
            "is_duplex": is_duplex,
            "is_color": is_color,
            "materials": materials
        }

# ==========================================
# 2. 페이지 측정 엔진 (Reader)
# ==========================================
def get_page_count(file_bytes, ext):
    try:
        if ext == '.pdf':
            reader = PdfReader(io.BytesIO(file_bytes))
            return len(reader.pages)
        elif ext in ['.pptx', '.ppt']:
            prs = Presentation(io.BytesIO(file_bytes))
            return len(prs.slides)
        # 워드(.docx)나 한글(.hwp)은 라이브러리 의존성이 커서 일단 0 처리하거나 추후 추가
        return 0
    except:
        return 0

# ==========================================
# 3. 메인 앱 (Streamlit)
# ==========================================
st.set_page_config(page_title="인쇄 견적 1차 집계 시스템", layout="wide")

st.title("🖨️ 인쇄/제본 1차 물량 산출기 (Logic V2.0)")
st.markdown("""
- **.txt 파일**: 인쇄 매수에서 제외, **부자재(비닐, 색지 등) 수량만 카운트**합니다.
- **.pdf/.pptx 파일**: 페이지 수를 읽어 인쇄 매수를 계산하고, 파일명에 적힌 부자재를 추가합니다.
- **수량 계산**: `파일명의 부수(Copies)`는 인쇄 매수에만 곱해지며, **부자재 수량에는 곱해지지 않습니다.** (1:1 합산)
""")

uploaded_file = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_file:
    results = []
    
    # 전체 집계용 변수
    total_summary = {
        "흑백_페이지(면)": 0,
        "컬러_페이지(면)": 0,
        "비닐내지(매)": 0,
        "색지(매)": 0,
        "USB(개)": 0
    }

    with zipfile.ZipFile(uploaded_file, 'r') as z:
        # __MACOSX 등 불필요한 시스템 파일 제외
        file_list = [f for f in z.namelist() if not f.startswith('__') and not f.endswith('/')]
        
        for filepath in file_list:
            filename = os.path.basename(filepath)
            folder = os.path.dirname(filepath)
            ext = os.path.splitext(filename)[1].lower()
            
            # 1. 파일명 파싱 (인쇄 옵션 & 부자재 추출)
            # 폴더명은 참고용으로 표기만 하고, 로직은 파일명 기준 (사용자 요청 2번)
            spec = InstructionParser.parse_filename(filename)
            
            # 2. 파일 타입별 처리 로직
            raw_pages = 0
            calc_sheets = 0
            print_category = "-"
            
            # [Case A] 인쇄용 파일 (.pdf, .pptx)
            if ext in ['.pdf', '.pptx', '.ppt']:
                file_bytes = z.read(filepath)
                raw_pages = get_page_count(file_bytes, ext)
                
                if raw_pages > 0:
                    # N-up 적용 (올림 처리)
                    pages_n_up = math.ceil(raw_pages / spec['n_up'])
                    
                    # 양면/단면 적용
                    # 양면이면 2로 나누고 올림, 단면이면 그대로
                    sheets_per_copy = math.ceil(pages_n_up / 2) if spec['is_duplex'] else pages_n_up
                    
                    # 부수 적용 (최종 인쇄 매수)
                    calc_sheets = sheets_per_copy * spec['copies']
                    
                    # 컬러/흑백 분류
                    if spec['is_color']:
                        print_category = "컬러"
                        total_summary["컬러_페이지(면)"] += calc_sheets
                    else:
                        print_category = "흑백"
                        total_summary["흑백_페이지(면)"] += calc_sheets

            # [Case B] 지시서 파일 (.txt) -> 인쇄 매수는 0, 부자재만 체크
            elif ext == '.txt':
                raw_pages = 0
                calc_sheets = 0
                print_category = "지시서(Skip)"
                # txt 파일은 인쇄하지 않으므로 copies가 있어도 인쇄매수에 영향 없음

            # 3. 부자재 집계 (파일 종류 상관없이 파일명에 있으면 무조건 합산)
            # 사용자 요청: "3부 비닐내지는 인쇄매수*3 + 비닐내지1" -> 부자재는 copies 곱하지 않음
            mats = spec['materials']
            total_summary["비닐내지(매)"] += mats["비닐내지"]
            total_summary["색지(매)"] += mats["색지"]
            total_summary["USB(개)"] += mats["USB"]
            
            # 결과 리스트에 추가
            results.append({
                "폴더 경로": folder,
                "파일명": filename,
                "타입": ext,
                "원본P": raw_pages,
                "옵션": f"{spec['n_up']}up/{'양면' if spec['is_duplex'] else '단면'}/{spec['copies']}부",
                "부자재 추출": str([k for k, v in mats.items() if v > 0]),
                "인쇄매수": calc_sheets,
                "분류": print_category,
                "비닐": mats["비닐내지"],
                "색지": mats["색지"],
                "USB": mats["USB"]
            })

    # --- 결과 출력 ---
    st.subheader("📊 전체 집계 요약")
    
    # 보기 좋게 컬럼으로 나누기
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("🖨️ 흑백 인쇄(장)", total_summary["흑백_페이지(면)"])
    c2.metric("🎨 컬러 인쇄(장)", total_summary["컬러_페이지(면)"])
    c3.metric("📂 비닐내지(매)", total_summary["비닐내지(매)"])
    c4.metric("📄 색지/간지(매)", total_summary["색지(매)"])
    c5.metric("💾 USB(개)", total_summary["USB(개)"])

    st.divider()

    st.subheader("📑 상세 파일별 분석 로그")
    df = pd.DataFrame(results)
    
    # 데이터프레임 스타일링 (가독성 향상)
    st.dataframe(
        df, 
        column_config={
            "인쇄매수": st.column_config.NumberColumn("최종 인쇄(장)"),
            "비닐": st.column_config.NumberColumn("비닐(매)"),
            "색지": st.column_config.NumberColumn("색지(매)"),
        },
        use_container_width=True
    )
    
    # 엑셀 다운로드 버튼
    # (실제 배포 시 pandas의 to_excel 사용을 위해 openpyxl 필요할 수 있음)
    # output = io.BytesIO()
    # with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    #     df.to_excel(writer, sheet_name='Sheet1', index=False)
    # st.download_button(label="📥 엑셀로 결과 다운로드", data=output.getvalue(), file_name="quotation_result.xlsx")
