import math

def calculate_pages(raw_pages, pages_per_sheet, copies):
    sheets = math.ceil(raw_pages / pages_per_sheet)
    final = sheets * copies
    return final, f"({raw_pages} ÷ {pages_per_sheet}) → {sheets} × {copies}"
