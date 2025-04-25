import os
import sys
import pandas as pd

# 1. 获取用户输入
directory = input("Please enter the path to the excel folder: ")
timestamp = input("Please enter the timestamp (e.g. 20250421_0905) ")
istep_and_cw = input("Please enter istep and CW  (e.g. 490CW16) ")

# 2. 初始化数据结构
data_dict = {}
file_names = []

# 3. 检查是否已有汇总 CSV（非必须，可选）
csv_file_path = os.path.join(directory, f"summary_{timestamp}_{istep_and_cw}.csv")
existing_file_names = []
if os.path.exists(csv_file_path):
    # 如果已存在，就读取已汇总过的文件名，用于去重
    existing_df = pd.read_csv(csv_file_path)
    if 'File Name' in existing_df.columns:
        existing_file_names = existing_df['File Name'].unique().tolist()

# 4. 遍历目录，读取每个 .xlsm 的 TRIGGER 表
for filename in os.listdir(directory):
    if filename.endswith(".xlsm") and not filename.startswith("summary") and filename not in existing_file_names:
        file_path = os.path.join(directory, filename)
        xls = pd.ExcelFile(file_path)
        if 'TRIGGER' in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name='TRIGGER', skiprows=6)
            # 删除跳过行后多余的首行
            df = df.drop(df.index[0])
            # 删除第 34 列（索引 33）及以后的冗余列
            df = df.drop(columns=df.columns[33:], errors='ignore')

            # 构造带时间戳和 CW 的文件名
            appended_filename = f"{filename}_{timestamp}_{istep_and_cw}"
            # 每行都记录来源文件名
            file_names.extend([appended_filename] * len(df))

            # 累积所有列的数据
            for col in df.columns:
                data_dict.setdefault(col, []).extend(df[col].tolist())
        else:
            print(f"Ignored (missing TRIGGER sheet): {filename}")

# 5. 合并为 DataFrame
try:
    new_df = pd.DataFrame(data_dict)
except Exception as e:
    print(f"格式合并出错：{e}")
    print("请检查各 Excel 的第 7 行表头是否一致、格式是否正确。")
    sys.exit(1)

# 插入“File Name”列
new_df.insert(0, 'File Name', file_names)

# 如果有“Date”列，则格式化为 “Apr 21” 样式
if 'Date' in new_df.columns:
    new_df['Date'] = pd.to_datetime(new_df['Date'], errors='coerce').dt.strftime('%b %d')

# 添加空的“Comment”列，并将“File Name”移到最后
new_df['Comment'] = ''
cols = new_df.columns.tolist()
cols.append(cols.pop(cols.index('File Name')))
new_df = new_df[cols]

# 6. 如果已有旧的 CSV，则合并
if os.path.exists(csv_file_path):
    old_df = pd.read_csv(csv_file_path)
    summary_df = pd.concat([old_df, new_df], ignore_index=True)
else:
    summary_df = new_df

# 7. 将汇总结果输出为 CSV
summary_df.to_csv(csv_file_path, index=False)
print(f"汇总已保存为 CSV 文件：{csv_file_path}")
