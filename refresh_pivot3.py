# refresh_octane_pivots.py

import os
import win32com.client  # pip install pywin32

# =========== 内部配置 ===========
CONFIG = {
    # 要刷新的 Excel 文件路径（相对于脚本所在目录）
    "workbook_path": os.path.join("Orig_files", "summary.xlsx"),
    # 数据源所在的工作表名称
    "data_sheet": "Octane and jira",
    # 如果你的 PivotTable 分布在特定的 sheet 列表里，可以在这里列出；
    # 若要更新所有 sheet 上的 PivotTable，请留空或为 None
    "pivot_sheets": None  # e.g. ["PivotSheet1", "PivotSheet2"]
}
# ================================

# Excel VBA 常量（数字形式）
XL_UP       = -4162   # xlUp
XL_TOLEFT   = -4159   # xlToLeft
XL_DATABASE = 1       # xlDatabase

def refresh_octane_pivots():
    wb_path = CONFIG["workbook_path"]
    data_sheet = CONFIG["data_sheet"]
    pivot_sheets = CONFIG["pivot_sheets"]

    if not os.path.isfile(wb_path):
        print(f"❌ Excel 文件不存在：{wb_path}")
        return

    abs_path = os.path.abspath(wb_path)
    print(f"🔄 打开并更新 PivotTable：{abs_path}")

    # 启动 Excel COM
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(abs_path)

    # 定位数据表
    try:
        ds = wb.Worksheets(data_sheet)
    except Exception:
        print(f"❌ 找不到工作表 “{data_sheet}”")
        wb.Close(False)
        excel.Quit()
        return

    # 计算数据表的有效区域
    last_row = ds.Cells(ds.Rows.Count, 1).End(XL_UP).Row
    last_col = ds.Cells(1, ds.Columns.Count).End(XL_TOLEFT).Column

    if last_row < 2:
        print(f"⚠️ “{data_sheet}” 表只有标题行，无数据可用，跳过更新。")
        wb.Close(False)
        excel.Quit()
        return

    tl = ds.Cells(1,1).Address      # "$A$1"
    br = ds.Cells(last_row, last_col).Address
    source_ref = f"'{data_sheet}'!{tl}:{br}"
    print(f"  新的数据源范围：{source_ref}")

    # 决定要更新哪些工作表的 PivotTable
    sheets_to_scan = (
        pivot_sheets if pivot_sheets
        else [ws.Name for ws in wb.Worksheets]
    )

    for name in sheets_to_scan:
        try:
            ws = wb.Worksheets(name)
        except Exception:
            print(f"⚠️ 工作表 “{name}” 不存在，跳过。")
            continue

        pts = ws.PivotTables()
        if pts.Count == 0:
            continue

        print(f"  ▶ 更新工作表 “{name}” 上的 {pts.Count} 个 PivotTable：")
        for idx in range(1, pts.Count + 1):
            pt = pts.Item(idx)
            try:
                cache = wb.PivotCaches().Create(
                    SourceType=XL_DATABASE,
                    SourceData=source_ref
                )
                pt.ChangePivotCache(cache)
                pt.RefreshTable()
                print(f"    ✔️ PivotTable “{pt.Name}” 更新成功")
            except Exception as e:
                print(f"    ❌ PivotTable “{pt.Name}” 更新失败：{e}")

    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("✅ 所有 PivotTable 已更新并刷新。")

if __name__ == "__main__":
    refresh_octane_pivots()
