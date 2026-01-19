import re

def extract_folder_materials(folder, filenames, instructions):
    text = " ".join([folder] + filenames + instructions).lower()

    file_count = len([
        f for f in filenames
        if f.lower().endswith((".pdf", ".pptx"))
    ])

    materials = {
        "비닐": 0,
        "USB": 0,
        "CD": 0,
        "바인더": 0
    }

    # ✅ 바인더: 무조건 폴더당 1
    if any(k in text for k in ["바인더", "binder"]):
        materials["바인더"] = 1

    # ✅ USB
    if re.search(r"\busb\b", text):
        # 예외: 각 USB
        if any(k in text for k in ["각usb", "usb각", "각 usb"]):
            materials["USB"] = file_count
        else:
            materials["USB"] = 1

    # ✅ CD (USB와 동일 로직)
    if re.search(r"\bcd\b", text):
        if any(k in text for k in ["각cd", "cd각", "각 cd"]):
            materials["CD"] = file_count
        else:
            materials["CD"] = 1

    # ✅ 비닐 (현재는 폴더당 1)
    if "비닐" in text:
        materials["비닐"] = 1

    return materials
