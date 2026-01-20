import streamlit as st
import os
import re

st.set_page_config(page_title="출력물 페이지 계산기", layout="wide")
st.write("✅ 앱 시작됨 (import 단계 통과)")

# -----------------------------
# 유틸 함수
# -----------------------------

def safe_int(x, default=0):
    try:
        return int(x)
    except:
        return default


def extract_vinyl_count(text):
    """
    비닐내지 숫자 추출 규칙
    - '비닐내지(3공) 5장' → 5
    - '비닐내지 10장' → 10
    - '비닐내지 안에 넣어주세요' → 1
    - 연도(2024, 2025 등) 절대 숫자로 인식 ❌
    """
    if
