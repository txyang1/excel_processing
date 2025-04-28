if ok == "Involved I-Step:":
    s = str(val)
    # 1) 如果以 G070 或 U006 开头，沿用原逻辑
    if s.startswith("G070") or s.startswith("U006"):
        val = "NA05" + s[4:]
    else:
        # 2) 否则尝试提取全角或半角括号内的编号，如 “（25-07-452 ATS+3...）”
        m = re.search(r'[（(]([\d-]+)', s)
        if m:
            code = m.group(1)               # e.g. "25-07-452"
            val = f"NA05-{code}"           # 结果 "NA05-25-07-452"
