import streamlit as st
import zipfile
import io
from pypdf import PdfReader

st.set_page_config(page_title="PDF 페이지 테스트", layout="wide")
st.title("📄 PDF 페이지 수 테스트")

uploaded = st.file_uploader("ZIP 파일 업로드", type="zip")

if uploaded:
    st.write("ZIP 분석 중...")
    with zipfile.ZipFile(uploaded) as z:
        for name in z.namelist():
            if name.lower().endswith(".pdf"):
                with z.open(name) as f:
                    data = io.BytesIO(f.read())
                    try:
                        reader = PdfReader(data)
                        st.write(f"📄 {name} → {len(reader.pages)} 페이지")
                    except Exception as e:
                        st.error(f"{name} 읽기 실패: {e}")
