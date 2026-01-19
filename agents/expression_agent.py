import re
import yaml

with open("rules.yaml", encoding="utf-8") as f:
    RULES = yaml.safe_load(f)

NUM_MAP = RULES["numbers"]

def korean_to_num(text):
    return NUM_MAP.get(text)

def extract_pages_per_sheet(text):
    for p in RULES["pages_per_sheet"]:
        m = re.search(p, text)
        if m:
            g = m.group(1)
            return int(g) if g.isdigit() else korean_to_num(g)
    return 1

def extract_copies(text):
    for p in RULES["copies"]:
        m = re.search(p, text)
        if m:
            return int(m.group(1))
    return 1

def extract_materials(text):
    found = {}
    for mat, keys in RULES["materials"].items():
        for k in keys:
            if k in text:
                found[mat] = found.get(mat, 0) + 1
    return found
