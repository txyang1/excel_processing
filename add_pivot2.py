def _refresh_all_pivot_tables(xlsx_path):
    """
    遍历整本工作簿，更新每张表上所有 PivotTable 的 SourceData 区域到
    各自表格的当前数据范围（从 A1 到最右、最下一行列），并刷新。
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

    for ws in wb.Worksheets:
        # 跳过空表或不需要刷新的表（可根据名称过滤）
        # if ws.Name not in ("Octane and jira", ...): continue

        # 找到该表数据区末尾
        last_row = ws.Cells(ws.Rows.Count, 1).End(win32com.client.constants.xlUp).Row
        last_col = ws.Cells(1, ws.Columns.Count).End(win32com.client.constants.xlToLeft).Column

        top_left     = ws.Cells(1,1).Address(True,True,1)               # $A$1
        bottom_right = ws.Cells(last_row, last_col).Address(True,True,1)
        source_data  = f"{ws.Name}!{top_left}:{bottom_right}"

        for pt in ws.PivotTables():
            cache = wb.PivotCaches().Create(
                SourceType=1,     # xlDatabase
                SourceData=source_data
            )
            pt.ChangePivotCache(cache)
            pt.RefreshTable()

    wb.Save()
    wb.Close(False)
    excel.Quit()
    print(f"→ 所有 PivotTable 已更新并刷新。")
