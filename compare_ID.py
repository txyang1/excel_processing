import pandas as pd

# 1. 参数设置
input_file = 'input.xlsx'      # 源文件名
orig_sheet = 'Sheet1'          # 原数据所在 sheet 名称
new_sheet  = 'compared ID'     # 对齐结果要写入的 sheet 名称

# 2. 读取要对比的两列
df = pd.read_excel(input_file, sheet_name=orig_sheet, usecols=['ID', 'ID2'])
df1 = pd.DataFrame({'ID': df['ID'].dropna().astype(str)})
df2 = pd.DataFrame({'ID2': df['ID2'].dropna().astype(str)})

# 3. 外连接合并，按值对齐
result = pd.merge(df1, df2, left_on='ID', right_on='ID2', how='outer')

# 4. 先读取原文件所有 sheet，以便后面一起写回
all_sheets = pd.read_excel(input_file, sheet_name=None)

# 5. 写回原文件并新增对比 sheet
with pd.ExcelWriter(input_file, engine='xlsxwriter') as writer:
    # 5.1 重写所有原有 sheet
    for name, sheet_df in all_sheets.items():
        # 如果原文件里已存在‘compared ID’，先跳过避免冲突
        if name == new_sheet:
            continue
        sheet_df.to_excel(writer, sheet_name=name, index=False)
    # 5.2 写入对齐结果到新 sheet
    result.to_excel(writer, sheet_name=new_sheet, index=False)

    # 6. 对空白单元格（未匹配项）染成淡蓝色
    workbook  = writer.book
    worksheet = writer.sheets[new_sheet]
    fmt_blue = workbook.add_format({'bg_color': '#DCE6F1'})

    max_row = len(result) + 1  # +1 因为 header 在第 1 行
    worksheet.conditional_format(f'A2:A{max_row}', {
        'type':     'blanks',
        'format':   fmt_blue
    })
    worksheet.conditional_format(f'B2:B{max_row}', {
        'type':     'blanks',
        'format':   fmt_blue
    })

print(f"已在『{input_file}』中新建 sheet『{new_sheet}』并完成对齐与高亮。")
