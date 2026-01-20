import streamlit as st
import zipfile
import os
import io
import re
from collections import defaultdict
from pypdf import PdfReader

st.set_page_config(page_title="STEP1 흑백·비닐 안정화", layout="wide")
st.title("📦 STEP 1 · 흑백 / 비닐 계산 엔진")

uploaded_zip = st.file_uploader("ZIP 파일 업로드", type="zip")

# ---------- 유틸 ----------
def is_txt(name):
    return name.lower().endswith(".txt")

def is_pdf(name):
    return name.lower().endswith(".pdf")

def has_vinyl_keyword(text):
    return any(k in text for k in ["비닐", "비닐내지"])

def has_usb(text):
    return "usb" in text or "cd" in text

def extract_vinyl_from_txt(text):
    """
    TXT에서만 숫자 허용
    '비닐내지 5장' → 5
    숫자 없으면 1
    """
    m = re.search(r'비닐.*?(\d+)\s*장', text)
    if m:
        return int(m.group(1))
    if "비닐" in text:
        return 1
    return 0

def count_pdf_pages(file_bytes):
    reader = PdfReader(io.BytesIO(file_bytes))
    return len(reader.pages)

#
