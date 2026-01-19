def aggregate(results):
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

        # 인쇄 페이지 합산
        if r.get("print_type") == "흑백":
            summary[folder]["흑백"] += r.get("pages", 0)

        if r.get("print_type") == "컬러":
            summary[folder]["컬러"] += r.get("pages", 0)

        # 자재 합산 (안전 처리)
        for m, v in r.get("materials", {}).items():
            key = m.upper() if m.lower() in ["usb", "cd"] else m

            if key not in summary[folder]:
                # 신규 자재는 무시하거나 로그용으로만 남김
                continue

            summary[folder][key] += v

    return summary
