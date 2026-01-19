def detect_context(text):
    text = text.lower()

    if any(k in text for k in ["인쇄x", "출력x", "인쇄없음"]):
        return {"ignore": True}

    for k in ["비닐", "usb", "cd", "바인더"]:
        if k in text:
            return {"material_only": True, "material": k}

    if any(k in text for k in ["컬러", "칼라", "color"]):
        return {"print_type": "컬러"}

    return {"print_type": "흑백"}
