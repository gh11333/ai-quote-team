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

# --- [에이전트 1: 정밀 수치 및 규칙 추출기] ---
def extract_rules(text, is_filename=False):
    t = " " + text.lower().replace(" ", " ") + " "
    div, mul = None, None
    # N-up 추출 (2, 4, 6, 9, 16)
    m_div = re.search(r'(\d+)\s*(?:up|페이지|쪽|면|쪽모아)', t)
    if m_div and int(m_div.group(1)) in [2, 4, 6, 9, 16]:
        div = 1 / int(m_div.group(1))
    # 부수 추출 (자재 키워드가 없을 때만)
    if not any(k in t for k in ['비닐', '간지', '색지', '탭지', '특수', '라벨', '스티커', '카드', '클립']):
        m_mul = re.search(r'(\d+)\s*(?:부|장)', t)
        if m_mul: mul = int(m_mul.group(1))
    return div, mul

def analyze_accessories(text_list, keyword):
    """지시 뭉치에서 EACH(각)와 FIXED(고정)를 분리 판독"""
    is_each, fixed_val, found = False, 0, False
    for txt in text_list:
        t = txt.lower().replace(" ", "")
        if keyword not in t: continue
        found = True
        if any(x in t for x in ['각', '각각', '하나씩']): is_each = True
        m = re.search(rf'{keyword}.*?(\d+)|(\d+).*?{keyword}', t)
        if m: fixed_val += int(m.group(1) or m.group(2))
    return is_each, fixed_val, found

# --- [에이전트 2: 엄격한 분류기] ---
def get_file_category(filename):
    """분류는 오직 파일명 독립 단어로만 결정 (폴더 상속 배제)"""
    fn = " " + filename.lower().replace("_", " ").replace("-", " ") + " "
    if any(re.search(rf'\b{k}\b', fn) for k in ['face', 'spine', 'cover', '표지', 'binder']): return "바인더"
    if any(re.search(rf'\b{k}\b', fn) for k in ['toc', '목차']): return "TOC"
    return "인쇄"

# --- [메인 시스템] ---
st.set_page_config(page_title="최종 병기 V37.1", layout="wide")
st.title("📂 2026 사내 견적 자동화 (V37.1 - 무오류 완결판)")

uploaded_zip = st.file_uploader("ZIP 파일을 업로드하세요", type="zip")

if uploaded_zip:
    detailed_log = []
    summary = {}
    processed_fixed_items = set() # 중복 합산 방지 장치

    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as z:
            all_paths = [p for p in z.namelist() if not p.startswith('__MACOSX')]
            
            # 1. 지시사항 전수 DB화
            db = {}
            for p in all_paths:
                d = os.path.dirname(p).replace('\\', '/')
                if d not in db: db[d] = {"instrs": [os.path.basename(d)], "folder_name": os.path.basename(d)}
                if p.lower().
