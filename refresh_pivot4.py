# refresh_octane_pivots.py

import os
import win32com.client  # pip install pywin32

# =========== 内部配置 ===========
WORKBOOK_PATH = os.path.join("Orig_files", "summary.xlsx")
DATA_SHEET    = "Octane and jira"
# 若只需更新特定 sheet 上的 PivotTable，请列出它们；否则设为 None
PIVOT_SHEETS  = None
# ================================

# Excel VBA 常量（直接用数字避免常量加载问题）
XL_UP       = -4162   # xlUp
XL_TOLEFT   = -4159   # xlToLeft
XL_DATABASE = 1       # xlDatabase

def refresh_octane_pivots():
    if not os.path.isfile(WORKBOOK_PATH):
        print(f"❌ Excel 文件不存在：{WORKBOOK_PATH}")
        return

    abs_path = os.path.abspath(WORKBOOK_PATH)
    print(f"🔄 打开并更新 PivotTable：{abs_path}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(abs_path)

    # 定位数据源表
    try:
        ds = wb.Worksheets(DATA_SHEET)
    except Exception:
        print(f"❌ 未找到工作表：{DATA_SHEET}")
        wb.Close(False)
        excel.Quit()
        return

    # 计算数据区域末尾
    last_row = ds.Cells(ds.Rows.Count, 1).End(XL_UP).Row
    last_col = ds.Cells(1, ds.Columns.Count).End(XL_TOLEFT).Column

    if last_row < 2:
        print(f"⚠️ “{DATA_SHEET}” 只有标题，无数据，跳过更新。")
        wb.Close(False)
        excel.Quit()
        return

    top_left     = ds.Cells(1, 1).Address    # "$A$1"
    bottom_right = ds.Cells(last_row, last_col).Address
    source_ref   = f"'{DATA_SHEET}'!{top_left}:{bottom_right}"
    print(f"  新的数据源范围：{source_ref}")

    # 确定要扫描哪些工作表
    sheets = PIVOT_SHEETS if PIVOT_SHEETS else [ws.Name for ws in wb.Worksheets]

    for name in sheets:
        try:
            ws = wb.Worksheets(name)
        except Exception:
            print(f"⚠️ 未找到工作表：{name}，跳过。")
            continue

        pts = ws.PivotTables()   # 注意：调用方法获取集合
        count = pts.Count
        print(f"  ▶ 表 “{name}” 上检测到 {count} 个 PivotTable")
        if count == 0:
            continue

        for i in range(1, count + 1):
            pt = pts.Item(i)
            try:
                cache = wb.PivotCaches().Create(
                    SourceType=XL_DATABASE,
                    SourceData=source_ref
                )
                pt.ChangePivotCache(cache)
                pt.RefreshTable()
                print(f"    ✔️ 已刷新 PivotTable：{pt.Name}")
            except Exception as e:
                print(f"    ❌ 刷新 PivotTable “{pt.Name}” 失败：{e}")

    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("✅ 所有 PivotTable 已更新并刷新。")

if __name__ == "__main__":
    refresh_octane_pivots()
