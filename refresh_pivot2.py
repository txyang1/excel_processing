# refresh_octane_pivots.py

import os
import sys
import win32com.client

# Excel VBA 常量（直接用数值，避免常量加载问题）
XL_UP       = -4162   # xlUp
XL_TOLEFT   = -4159   # xlToLeft
XL_DATABASE = 1       # xlDatabase

def refresh_pivots(workbook_path):
    if not os.path.isfile(workbook_path):
        print(f"文件不存在：{workbook_path}")
        return

    abs_path = os.path.abspath(workbook_path)
    print(f"🔄 打开并更新 PivotTable：{abs_path}")

    # 启动 Excel COM
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(abs_path)
    try:
        # 定位数据源工作表
        data_ws = wb.Worksheets("Octane and jira")
    except Exception:
        print("❌ 找不到工作表 “Octane and jira”")
        wb.Close(False)
        excel.Quit()
        return

    # 计算该表的有效数据区域 A1 → 最后一行/最后一列
    last_row = data_ws.Cells(data_ws.Rows.Count, 1).End(XL_UP).Row
    last_col = data_ws.Cells(1, data_ws.Columns.Count).End(XL_TOLEFT).Column

    top_left     = data_ws.Cells(1, 1).Address    # e.g. "$A$1"
    bottom_right = data_ws.Cells(last_row, last_col).Address
    source_ref   = f"'{data_ws.Name}'!{top_left}:{bottom_right}"

    print(f"  新的数据源范围：{source_ref}")

    # 遍历所有工作表中的所有 PivotTable，更新并刷新
    for ws in wb.Worksheets:
        pts = ws.PivotTables()
        if pts.Count == 0:
            continue
        print(f"  ▶ 工作表 “{ws.Name}” 有 {pts.Count} 个 PivotTable，将更新它们……")
        for pt in pts:
            cache = wb.PivotCaches().Create(
                SourceType=XL_DATABASE,
                SourceData=source_ref
            )
            pt.ChangePivotCache(cache)
            pt.RefreshTable()

    # 保存并退出
    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("✅ 所有 PivotTable 已更新并刷新。")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("用法: python refresh_octane_pivots.py <Excel文件路径>")
        sys.exit(1)
    refresh_pivots(sys.argv[1])
