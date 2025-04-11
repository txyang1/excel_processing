import pandas as pd

# 1. 生成 original.xlsx 数据
original_data = {
    'ID': [1, 2, 3],
    'Name': ['Item A', 'Item B', 'Item C'],
    'Assigned ECU Group': ['Group1', 'Group2', 'Group3'],
    'Phase Group': ['PhaseGroup1', 'PhaseGroup2', 'PhaseGroup3'],
    'Ticket no.supplier': ['T001', 'T002', 'T003'],
    'Days in phase': [10, 20, 30],
    'Defect finder': ['FinderA', 'FinderB', 'FinderC']
}

original_df = pd.DataFrame(original_data)
original_df.to_excel('original.xlsx', index=False)

# 2. 生成 new.xlsx 数据，包含部分更新或异常数据：
new_data = {
    'ID': [1, 2, 3],
    'Name': ['Item A', 'Item B Updated', 'Item C'],
    'Assigned ECU Group': ['Group1', 'Group2', 'Group3'],
    'Phase Group': ['PhaseGroup1', 'PhaseGroup2', 'PhaseGroup3'],  # 对应的映射将根据 Defect finder 更新
    'Ticket no.supplier': ['T001', 'T002', 'T003'],
    'Days in phase': [15, 20, 'invalid'],  # 第1行更新了天数，第3行数据异常
    'Defect finder': ['FinderA', 'FinderB', 'FinderD']  # 第3行更改了 Defect finder 以触发映射更新示例
}

new_df = pd.DataFrame(new_data)
new_df.to_excel('new.xlsx', index=False)

print("已生成 original.xlsx 和 new.xlsx")
