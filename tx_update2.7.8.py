# …（前面代码不变）

# -------------------------------
# 4. 遍历新表，更新或追加，只高亮更新的单元格
# -------------------------------
for _, new_row in df_new.iterrows():
    new_id = new_row.get("ID", "")
    if new_id in id2row:
        # —— 更新原有行，只高亮那些被写入新值的单元格 —— 
        r = id2row[new_id]
        # 先可选更新 ID 的超链接
        if new_id in id2url_new:
            cell = ws.cell(r, id_col)
            cell.hyperlink = id2url_new[new_id]
            cell.style     = "Hyperlink"
        # 对映射字段逐个检查
        for new_key, orig_key in mapping.items():
            v = new_row.get(new_key, "")
            if pd.notna(v) and v != "":
                c = header2col.get(orig_key)
                if c:
                    cell = ws.cell(r, c)
                    cell.value = v
                    cell.fill  = green_fill    # 仅对更新的单元格涂绿
    else:
        # —— 追加新行，只高亮那些有实际值的单元格 —— 
        row_vals = []
        for hdr in header2col:
            # 找到 new_key 对应的 orig_key==hdr
            for new_key, orig_key in mapping.items():
                if orig_key == hdr:
                    val = new_row.get(new_key, "")
                    if pd.isna(val): val = ""
                    row_vals.append(val)
                    break
            else:
                row_vals.append("")  # 没映射的列留空

        ws.append(row_vals)
        new_r = ws.max_row

        # 给 ID 单元格加超链接（如果有）
        if new_id in id2url_new:
            cell = ws.cell(new_r, id_col)
            cell.hyperlink = id2url_new[new_id]
            cell.style     = "Hyperlink"

        # 只高亮那些写入了非空新值的单元格
        for new_key, orig_key in mapping.items():
            v = new_row.get(new_key, "")
            if pd.notna(v) and v != "":
                c = header2col.get(orig_key)
                if c:
                    ws.cell(new_r, c).fill = yellow_fill

# …（后面公式、高级列填充等保持不变） …
