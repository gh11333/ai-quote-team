def aggregate(results, folder_materials):
    summary = {}

    for r in results:
        folder = r["folder"]

        if folder not in summary:
            summary[folder] = {
                "흑백": 0,
                "컬러": 0,
                "비닐": 0,
                "USB": 0,
                "CD": 0,
                "바인더": 0
            }

        if r.get("print_type") == "흑백":
            summary[folder]["흑백"] += r.get("pages", 0)

        if r.get("print_type") == "컬러":
            summary[folder]["컬러"] += r.get("pages", 0)

    # 🔥 자재는 폴더 기준으로 1회 세팅
    for folder, mats in folder_materials.items():
        for k, v in mats.items():
            summary[folder][k] = v

    return summary
