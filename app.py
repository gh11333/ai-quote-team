import os
import re
from PyPDF2 import PdfReader

ROOT_DIR = r"C:\Users\안건희\Desktop\새 폴더"  # ★ 여기만 수정 ★

total_bw_pages = 0
total_vinyl = 0

debug_log = []

def extract_vinyl_count(text):
    """
    비닐내지 숫자 추출
    - '비닐내지 2장' → 2
    - '비닐내지(3공)' → 1
    """
    matches = re.findall(r"비닐내지[^0-9]*([0-9]+)?", text)
    count = 0
    for m in matches:
        count += int(m) if m else 1
    return count

for root, dirs, files in os.walk(ROOT_DIR):
    for file in files:
        path = os.path.join(root, file)

        # -------------------------
        # PDF → 페이지 계산
        # -------------------------
        if file.lower().endswith(".pdf"):
            try:
                reader = PdfReader(path)
                pages = len(reader.pages)
                total_bw_pages += pages
                debug_log.append(f"[PDF] {file} → {pages} pages")
            except Exception as e:
                debug_log.append(f"[PDF ERROR] {file} → {e}")

        # -------------------------
        # TXT → 비닐만 계산
        # -------------------------
        elif file.lower().endswith(".txt"):
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()

                vinyl_from_txt = extract_vinyl_count(content)
                vinyl_from_name = extract_vinyl_count(file)

                vinyl = vinyl_from_txt + vinyl_from_name
                total_vinyl += vinyl

                debug_log.append(
                    f"[TXT] {file} → 비닐 {vinyl} (본문 {vinyl_from_txt} / 파일명 {vinyl_from_name})"
                )
            except Exception as e:
                debug_log.append(f"[TXT ERROR] {file} → {e}")

# -------------------------
# 결과 출력
# -------------------------
print("====== 최종 집계 ======")
print(f"흑백 페이지 수: {total_bw_pages}")
print(f"비닐 내지 수: {total_vinyl}")

print("\n====== 디버그 로그 (상위 50개) ======")
for line in debug_log[:50]:
    print(line)
