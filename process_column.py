import pandas as pd

# 1. 读取原始 Excel
input_file = 'input.xlsx'      # 源文件名
sheet_name = 'Sheet1'          # 工作表名
df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=['ID', 'ID2'])

# 2. 拆分两列并分别构造 DataFrame
df1 = pd.DataFrame({'ID': df['ID'].dropna().astype(str)})
df2 = pd.DataFrame({'ID2': df['ID2'].dropna().astype(str)})

# 3. 外连接合并，按值对齐
result = pd.merge(df1, df2, left_on='ID', right_on='ID2', how='outer')

# 4. 输出到新的 Excel，并对空白（未匹配）单元格设置淡蓝填充
output_file = 'aligned_output.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    result.to_excel(writer, index=False, sheet_name=sheet_name)
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    # 定义淡蓝色格式
    fmt_blue = workbook.add_format({'bg_color': '#DCE6F1'})

    # 数据区范围（假设表头在第1行），从 A2/B2 开始到末行
    max_row = len(result) + 1
    worksheet.conditional_format(f'A2:A{max_row}', {
        'type':     'blanks',
        'format':   fmt_blue
    })
    worksheet.conditional_format(f'B2:B{max_row}', {
        'type':     'blanks',
        'format':   fmt_blue
    })

print(f"对齐并高亮差异后的文件已保存为：{output_file}")
