import os
import re
from PyPDF2 import PdfReader

# =========================
# 공통 유틸
# =========================

def safe_int(value, default=0):
    try:
        return int(value)
    except:
        return default


# =========================
# 비닐 계산 (최종 안정화)
# =========================

def extract_vinyl_count(text: str) -> int:
    """
    비닐 계산 규칙 (최종본)
    1. '비닐' / '비닐내지'가 없으면 0
    2. 키워드 기준 ±10글자 내 숫자만 인정
    3. 1~200 범위만 유효
    4. 숫자 없으면 비닐 = 1
    """

    if not text:
        return 0

    text = text.lower()

    if "비닐" not in text:
        return 0

    pattern = r"(비닐내지|비닐).{0,10}?(\d{1,3})"
    matches = re.findall(pattern, text)

    valid = []
    for _, num in matches:
        n = safe_int(num)
        if 1 <= n <= 200:
            valid.append(n)

    if valid:
        return max(valid)

    return 1


# =========================
# 페이지 계산
# =========================

def calculate_pdf_pages(pdf_path: str) -> int:
    try:
        reader = PdfReader(pdf_path)
        return len(reader.pages)
    except:
        return 0


def extract_page_rule_from_text(text: str):
    """
    출력 규칙 파싱
    단면 / 양면
    1면에 2페이지 / 4페이지
    """
    text = text.lower()

    is_duplex = "양면" in text
    per_side = 1

    if "1면에2페이지" in text or "한면2페이지" in text:
        per_side = 2
    elif "1면에4페이지" in text or "한면4페이지" in text:
        per_side = 4

    return is_duplex, per_side


def calculate_printed_pages(original_pages, is_duplex, per_side):
    if per_side <= 0:
        return original_pages

    logical_pages = original_pages / per_side

    if is_duplex:
        return int((logical_pages + 1) // 2)
    else:
        return int(logical_pages)


# =========================
# 단일 파일 처리
# =========================

def process_file(file_path: str):
    filename = os.path.basename(file_path)
    ext = os.path.splitext(filename)[1].lower()

    vinyl = 0
    bw_pages = 0

    # TXT 먼저 읽기
    txt_content = ""
    if ext == ".txt":
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                txt_content = f.read()
        except:
            pass

    # 비닐 계산 (TXT + 파일명)
    vinyl += extract_vinyl_count(txt_content)
    vinyl += extract_vinyl_count(filename)

    # PDF 페이지 계산
    if ext == ".pdf":
        original_pages = calculate_pdf_pages(file_path)
        rule_text = filename + " " + txt_content
        is_duplex, per_side = extract_page_rule_from_text(rule_text)
        bw_pages = calculate_printed_pages(original_pages, is_duplex, per_side)

    return bw_pages, vinyl


# =========================
# 폴더 단위 집계
# =========================

def process_folder(root_folder: str):
    result = {}

    for root, dirs, files in os.walk(root_folder):
        folder_name = os.path.basename(root)
        if folder_name not in result:
            result[folder_name] = {"흑백": 0, "비닐": 0}

        for file in files:
            file_path = os.path.join(root, file)
            bw, vinyl = process_file(file_path)
            result[folder_name]["흑백"] += bw
            result[folder_name]["비닐"] += vinyl

    return result


# =========================
# 실행부
# =========================

if __name__ == "__main__":
    TARGET_FOLDER = "./data"  # ← 여기만 네 폴더 경로로 수정

    summary = process_folder(TARGET_FOLDER)

    print("\n📊 정산 결과")
    print("-" * 40)
    for folder, values in summary.items():
        print(
            f"{folder}\t흑백 {values['흑백']}\t비닐 {values['비닐']}"
        )
