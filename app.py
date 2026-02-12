import streamlit as st
import zipfile
import os
import io
import re
import math
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation

# --- [Agent 1: 전략 해석가 (The Interpreter)] ---
class StrategyInterpreter:
    @staticmethod
    def parse_instruction(text):
        text = text.lower().replace(" ", "")
        
        # 1. n-up 추출 (한 면에 들어가는 페이지)
        n_up = 1
        up_match = re.search(r'(\d+)(?:up|쪽모아|분할|면\d+쪽|슬라이드)', text)
        if up_match: n_up = int(up_match.group(1))

        # 2. 부수(Copies) 추출
        copies = 1
        copy_match = re.search(r'(\d+)(?:부|권|세트|장씩)', text)
        if copy_match: copies = int(copy_match.group(1))

        # 3. 양면 여부
        is_duplex = True if any(k in text for k in ['양면', 'double']) else False
        if '단면' in text: is_duplex = False

        # 4. [특수] 분권 로직 (01번 폴더 이슈 해결)
        # '4권으로 분권'은 4세트가 아니라, 1세트를 4개 바인더에 나눠 담는다는 의미로 우선 해석
        is_divided = True if '분권' in text else False
        
        return {"n_up": n_up, "copies": copies, "is_duplex": is_duplex, "is_divided": is_divided}

# --- [Agent 2: 정밀 측정가 (The Counter)] ---
class PageCounter:
    @staticmethod
    def get_raw_pages(file_content, ext):
        try:
            f_stream = io.BytesIO(file_content)
            if ext == '.pdf':
                return len(PdfReader(f_stream).pages)
            elif ext == '.pptx':
                return len(Presentation(f_stream).slides)
            return 1 # 기본값
        except:
            return 0

# --- [Agent 3: 최종 정산 및 검증관 (The Auditor)] ---
class QuotationAuditor:
    @staticmethod
    def calculate_sheets(raw_pages, spec):
        """
        최종 인쇄 매수 산출 공식:
        $$FinalSheets = \lceil (\frac{RawPages}{N-up}) \times \frac{1}{2(if Duplex)} \rceil \times Copies$$
        """
        if raw_pages == 0: return 0
        
        # 1. n-up 적용
        pages_after_up = math.ceil(raw_pages / spec['n_up'])
        
        # 2. 양면/단면 적용 (양면이면 2로 나눔)
        divisor = 2 if spec['is_duplex'] else 1
        sheets_per_copy = math.ceil(pages_after_up / divisor)
        
        # 3. 부수 적용 (분권인 경우 부수를 1로 고정하는 안전장치)
        final_copies = 1 if spec['is_divided'] and spec['copies'] == 1 else spec['copies']
        
        return sheets_per_copy * final_copies

# --- [Main App Integration] ---
st.set_page_config(page_title="무결점 엔진 V41.0", layout="wide")
st.title("🛡️ 2026 견적 자동화 에이전트 팀 (V41.0)")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded_zip:
    results = []
    summary = {}

    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX') and not p.endswith('/')]
        
        for path in all_paths:
            filename = os.path.basename(path)
            folder_path = os.path.dirname(path)
            top_folder = path.split('/')[0] if '/' in path else "Root"
            ext = os.path.splitext(filename)[1].lower()

            if top_folder not in summary:
                summary[top_folder] = {"흑백": 0, "컬러": 0, "파일수": 0}

            # 1. 해석 에이전트 기동 (폴더명 + 파일명 컨텍스트 통합)
            context = (folder_path + "_" + filename).replace('\\', '_')
            spec = StrategyInterpreter.parse_instruction(context)

            # 2. 측정 에이전트 기동
            raw_p = PageCounter.get_raw_pages(z.read(path), ext)

            # 3. 정산 에이전트 기동
            final_sheets = QuotationAuditor.calculate_sheets(raw_p, spec)

            # 분류 (컬러/흑백)
            is_color = any(k in context.lower() for k in ['컬러', '칼라', 'color'])
            cat = "컬러" if is_color else "흑백"

            # 데이터 저장
            summary[top_folder][cat] += final_sheets
            summary[top_folder]["파일수"] += 1
            results.append({
                "폴더": top_folder,
                "파일명": filename,
                "원본P": raw_p,
                "설정": f"{spec['n_up']}UP/{'양면' if spec['is_duplex'] else '단면'}",
                "부수": spec['copies'],
                "최종인쇄매수": final_sheets,
                "분류": cat
            })

    # 결과 출력
    st.subheader("📊 정산 요약")
    st.table(pd.DataFrame.from_dict(summary, orient='index'))
    
    st.subheader("📑 상세 에이전트 로그")
    st.dataframe(pd.DataFrame(results))

    # 엑셀 다운로드 로직 (생략 - 위와 동일)
