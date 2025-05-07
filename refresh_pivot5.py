# refresh_octane_pivots.py

import os
import win32com.client  # pip install pywin32

# =========== 内部配置 ===========
WORKBOOK_PATH = os.path.join("Orig_files", "summary.xlsx")
DATA_SHEET    = "Octane and jira"
# 若只需更新指定 sheet 上的 PivotTable，列在这里；否则设为 None
PIVOT_SHEETS  = None
# ================================

# Excel VBA 常量（数字形式）
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

    # 定位数据工作表
    try:
        ds = wb.Worksheets(DATA_SHEET)
    except Exception:
        print(f"❌ 未找到工作表：{DATA_SHEET}")
        wb.Close(False)
        excel.Quit()
        return

    # 1) 找到表头最后一列（用于确定列范围）
    header_row = 1
    last_header_col = ds.Cells(header_row, ds.Columns.Count).End(XL_TOLEFT).Column

    # 2) 在表头这一行扫描 “ID” 列号
    id_col_idx = None
    for c in range(1, last_header_col + 1):
        if str(ds.Cells(header_row, c).Value).strip() == "ID":
            id_col_idx = c
            break

    if not id_col_idx:
        print("⚠️ 在表头未找到 “ID” 列，无法确定数据行范围，跳过更新。")
        wb.Close(False)
        excel.Quit()
        return

    # 3) 用 ID 列 End(xlUp) 找到最后一行
    last_row = ds.Cells(ds.Rows.Count, id_col_idx).End(XL_UP).Row

    if last_row <= header_row:
        print(f"⚠️ “{DATA_SHEET}” 表中 “ID” 列无数据，跳过更新。")
        wb.Close(False)
        excel.Quit()
        return

    # 4) 用表头最后一列作为数据区域的列末
    last_col = last_header_col

    # 5) 构造 SourceData 字符串
    tl = ds.Cells(header_row, 1).Address    # 通常 "$A$1"
    br = ds.Cells(last_row, last_col).Address
    source_ref = f"'{DATA_SHEET}'!{tl}:{br}"
    print(f"  新的数据源范围：{source_ref}")

    # 决定要扫描哪些 sheet
    sheets = PIVOT_SHEETS if PIVOT_SHEETS else [ws.Name for ws in wb.Worksheets]

    # 遍历并更新每个 PivotTable
    for name in sheets:
        try:
            ws = wb.Worksheets(name)
        except Exception:
            print(f"⚠️ 未找到工作表：{name}，跳过。")
            continue

        pts = ws.PivotTables()      # 注意：必须调用 PivotTables()
        count = pts.Count
        if count == 0:
            continue
        print(f"  ▶ 表 “{name}” 上有 {count} 个 PivotTable，开始更新：")

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
                print(f"    ❌ 刷新 {pt.Name} 失败：{e}")

    # 保存并退出
    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("✅ 所有 PivotTable 已更新并刷新。")

if __name__ == "__main__":
    refresh_octane_pivots()
