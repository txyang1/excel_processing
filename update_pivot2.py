import os
from pathlib import Path
import win32com.client  # pip install pywin32

# =========== 内部配置 ===========
DATA_SHEET   = "Octane and jira"
# 若只需更新指定 sheet 上的 PivotTable，列在这里；否则设为 None
PIVOT_SHEETS = None
# ================================

def refresh_octane_pivots():
    # 读取并规范化用户输入的文件路径
    raw_path = input("Please enter the path to the Excel file: ")
    file_path = Path(raw_path).expanduser().resolve()

    if not file_path.is_file():
        print(f"❌ Excel 文件不存在：{file_path}")
        return

    abs_path = str(file_path)
    print(f"🔄 打开并更新 PivotTable：{abs_path}")

    # 初始化 COM 对象和工作簿引用
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = None
    try:
        # 尝试打开工作簿
        wb = excel.Workbooks.Open(abs_path)

        # 定位数据工作表
        try:
            ds = wb.Worksheets(DATA_SHEET)
        except Exception:
            print(f"❌ 未找到工作表：{DATA_SHEET}")
            return

        # Excel VBA 常量
        XL_UP       = -4162   # xlUp
        XL_TOLEFT   = -4159   # xlToLeft
        XL_DATABASE = 1       # xlDatabase

        # 1) 找到表头最后一列
        header_row = 1
        last_header_col = ds.Cells(header_row, ds.Columns.Count).End(XL_TOLEFT).Column

        # 2) 扫描“ID”列号
        id_col_idx = None
        for c in range(1, last_header_col + 1):
            if str(ds.Cells(header_row, c).Value).strip() == "ID":
                id_col_idx = c
                break

        if not id_col_idx:
            print("⚠️ 在表头未找到 “ID” 列，跳过更新。")
            return

        # 3) 根据 ID 列找到最后一行
        last_row = ds.Cells(ds.Rows.Count, id_col_idx).End(XL_UP).Row
        if last_row <= header_row:
            print(f"⚠️ “{DATA_SHEET}” 中 “ID” 列无数据，跳过更新。")
            return

        # 4) 数据区域右下角列
        last_col = last_header_col

        # 5) 构造 SourceData
        tl = ds.Cells(header_row, 1).Address
        br = ds.Cells(last_row, last_col).Address
        source_ref = f"'{DATA_SHEET}'!{tl}:{br}"
        print(f"  新的数据源范围：{source_ref}")

        # 读取需要更新的所有工作表
        sheets = PIVOT_SHEETS if PIVOT_SHEETS else [ws.Name for ws in wb.Worksheets]

        # 遍历并更新每个 PivotTable
        for name in sheets:
            try:
                ws = wb.Worksheets(name)
            except Exception:
                print(f"⚠️ 未找到工作表：{name}，跳过。")
                continue

            pts = ws.PivotTables()
            if pts.Count == 0:
                continue
            print(f"  ▶ 表 “{name}” 上有 {pts.Count} 个 PivotTable：")

            for i in range(1, pts.Count + 1):
                pt = pts.Item(i)
                try:
                    cache = wb.PivotCaches().Create(
                        SourceType=XL_DATABASE,
                        SourceData=source_ref
                    )
                    pt.ChangePivotCache(cache)
                    pt.RefreshTable()
                    print(f"    ✔️ 刷新：{pt.Name}")
                except Exception as e:
                    print(f"    ❌ 刷新 {pt.Name} 失败：{e}")

        # 保存更改
        wb.Save()
        print("✅ 所有 PivotTable 已更新并刷新。")

    except Exception as exc:
        print(f"❌ 运行时出错：{exc}")

    finally:
        # 仅在成功打开并存在 wb 时才关闭
        if wb is not None:
            wb.Close(False)
        excel.Quit()

if __name__ == "__main__":
    refresh_octane_pivots()
