# refresh_octane_pivots.py

import os
import sys
import win32com.client

# Excel VBA 常量
XL_UP       = -4162   # xlUp
XL_TOLEFT   = -4159   # xlToLeft
XL_DATABASE = 1       # xlDatabase

def refresh_pivots(workbook_path):
    if not os.path.isfile(workbook_path):
        print(f"❌ 文件不存在：{workbook_path}")
        return

    abs_path = os.path.abspath(workbook_path)
    print(f"🔄 打开并更新 PivotTable：{abs_path}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(abs_path)

    try:
        data_ws = wb.Worksheets("Octane and jira")
    except Exception:
        print("❌ 找不到工作表 “Octane and jira”")
        wb.Close(False)
        excel.Quit()
        return

    # 计算数据区域
    last_row = data_ws.Cells(data_ws.Rows.Count, 1).End(XL_UP).Row
    last_col = data_ws.Cells(1, data_ws.Columns.Count).End(XL_TOLEFT).Column

    if last_row < 2:
        print("⚠️ “Octane and jira” 表只有标题行，无数据可供 PivotTable 使用，跳过更新。")
        wb.Close(False)
        excel.Quit()
        return

    top_left     = data_ws.Cells(1, 1).Address    # "$A$1"
    bottom_right = data_ws.Cells(last_row, last_col).Address
    source_ref   = f"'{data_ws.Name}'!{top_left}:{bottom_right}"
    print(f"  新的数据源范围：{source_ref}")

    # 遍历所有工作表和它们的 PivotTable
    for ws in wb.Worksheets:
        pts = ws.PivotTables()
        if pts.Count == 0:
            continue
        print(f"  ▶ 表 “{ws.Name}” 上有 {pts.Count} 个 PivotTable，开始更新：")
        for idx in range(1, pts.Count+1):
            pt = pts.Item(idx)
            try:
                cache = wb.PivotCaches().Create(
                    SourceType=XL_DATABASE,
                    SourceData=source_ref
                )
                pt.ChangePivotCache(cache)
                pt.RefreshTable()
                print(f"    ✔️ 已更新 PivotTable “{pt.Name}”")
            except Exception as e:
                print(f"    ❌ 更新 PivotTable “{pt.Name}” 失败：{e}")

    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("✅ 所有可更新的 PivotTable 已处理完成。\n")

def main():
    if len(sys.argv) != 2:
        print("用法: python refresh_octane_pivots.py <Excel文件路径>")
        sys.exit(1)
    refresh_pivots(sys.argv[1])

if __name__ == "__main__":
    main()
