def aggregate(results):
    summary = {}
    for r in results:
        folder = r["folder"]
        if folder not in summary:
            summary[folder] = {"흑백":0,"컬러":0,"비닐":0,"USB":0,"바인더":0}

        if r.get("print_type") == "흑백":
            summary[folder]["흑백"] += r["pages"]
        if r.get("print_type") == "컬러":
            summary[folder]["컬러"] += r["pages"]

        for m,v in r.get("materials", {}).items():
            summary[folder][m] += v

    return summary
