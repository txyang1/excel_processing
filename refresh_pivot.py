# refresh_pivots.py

import os
import sys
import win32com.client

# Excel VBA 常量数值
XL_UP      = -4162   # xlUp
XL_TOLEFT  = -4159   # xlToLeft
XL_DATABASE = 1      # xlDatabase

def refresh_pivots_in_workbook(xlsx_path):
    """
    打开 xlsx_path 指定的 Excel 文件，遍历所有工作表上的所有 PivotTable，
    将它们的数据源范围 (SourceData) 更新到该表的当前数据区域 (A1:末尾单元格)，
    并刷新 PivotTable。
    """
    if not os.path.isfile(xlsx_path):
        print(f"文件不存在：{xlsx_path}")
        return

    abs_path = os.path.abspath(xlsx_path)
    print(f"开始更新透视表：{abs_path}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(abs_path)

    for ws in wb.Worksheets:
        # 跳过空表
        try:
            # 找到当前表中第一列最后一行（数据区行末）
            last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
            # 找到当前表中第一行最后一列（数据区列末）
            last_col = ws.Cells(1, ws.Columns.Count).End(XL_TOLEFT).Column
        except Exception:
            continue

        # 构造数据源地址，例如 "Sheet1!$A$1:$F$200"
        top_left     = ws.Cells(1, 1).Address      # "$A$1"
        bottom_right = ws.Cells(last_row, last_col).Address
        source_data  = f"'{ws.Name}'!{top_left}:{bottom_right}"

        pts = ws.PivotTables()
        if pts.Count == 0:
            # 该表无 PivotTable，跳过
            continue

        print(f"  更新工作表 '{ws.Name}' 上 {pts.Count} 个 PivotTable，范围：{source_data}")
        for pt in pts:
            # 创建新的 PivotCache 并绑定
            cache = wb.PivotCaches().Create(
                SourceType=XL_DATABASE,
                SourceData=source_data
            )
            pt.ChangePivotCache(cache)
            pt.RefreshTable()

    # 保存并关闭
    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("所有 PivotTable 已更新并刷新。")

def main():
    if len(sys.argv) < 2:
        print("用法: python refresh_pivots.py <Excel文件1> [<Excel文件2> ...]")
        sys.exit(1)

    for arg in sys.argv[1:]:
        refresh_pivots_in_workbook(arg)

if __name__ == "__main__":
    main()
