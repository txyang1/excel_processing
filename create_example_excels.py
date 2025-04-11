import pandas as pd

# -------------------------------
# 1. 生成虚拟的 Defect_data 表格
# -------------------------------
data_defect = {
    "Assigned ECU Group": ["Group1", "Group2_TXTXTX", "Group3"],
    "Phase Group": ["Phase1", "Phase2", "Phase3"],
    "Name": ["Defect A", "Defect B", "Defect C"],
    "ID": [101, 102, 103],
    "Ticket no.supplier": ["T001", "T002", "T003"],
    "First use/SoP of function": ["2021-01-01", "2021-02-01", "2021-03-01"],
    "Days in phase": [15, 25, 10],
    "Owner": ["Alice", "Bob", "Charlie"],
    "Tags": ["tag1", "tag2", "tag3"],
    "Reporting relevance": ["High", "Medium", "Low"],
    "Closed in version": ["v1.0", "v1.1", "v1.2"],
    "Invoved I-step": ["I1", "I2", "I3"],
    "Found in function": ["Func1", "Func2_TXTXTX", "Func3"],
    "Defect finder": ["FinderA", "FinderB", "FinderC"],
    "Error occurrence": ["E1", "E2", "E3"],
    "Phase": ["P1", "P2", "P3"],
    "Planned closing version": ["v2.0", "v2.1", "v2.2"],
    "Creation time": ["2020-12-31", "2021-01-31", "2021-02-28"],
    "Assigned ECU": ["ECU1", "ECU2_TXTXTX", "ECU3"],
    "Target I-Step": ["T-I1", "T-I2", "T-I3"],
    "Target SET": ["Set1", "Set2", "Set3"],
    "Target Week": ["W1", "W2", "W3"],
    "Defect severity": ["Sev1", "Sev2", "Sev3"],
    "Defect category": ["Cat1", "Cat2", "Cat3"],
    "Display severity": ["Disp1", "Disp2", "Disp3"],
    "Target release": ["R1", "R2", "R3"],
    "Last modified": ["2021-01-10", "2021-02-10", "2021-03-10"]
}
df_defect = pd.DataFrame(data_defect)
# 保存到 Excel 文件
df_defect.to_excel("excel_code/data/Defect_data2.xlsx", index=False)


# -------------------------------
# 2. 生成虚拟的 Jira 表格
# -------------------------------
data_jira = {
    "Created": ["2021-01-02", "2021-02-02", "2021-03-02"],
    "issue Type": ["Bug", "Task_TXTXTX", "Bug"],
    "priority": ["High", "Medium", "Low"],
    "Issue key": ["JIRA-101", "JIRA-102", "JIRA-103"],
    "issue id": [101, 102, 103],  # 与 Defect_data 的 ID 对应，用于合并
    "Affects Version/s": ["v1.0", "v1.1", "v1.2"],
    "Status": ["Open", "Closed", "Open"],
    "Assignee": ["Dave", "Eve", "Frank"],
    "Reporter": ["Grace", "Heidi_TXTXTXTX", "Ivan"],
    "Updated": ["2021-01-15", "2021-02-15", "2021-03-15"],
    "Summary": ["Issue A summary", "Issue B summary", "Issue C summary"],
    "Fix Version/s": ["v1.0.1", "v1.1.1", "v1.2.1"],
    "Custom field(project Specific Labels)": ["Label1", "Label2", "Label3"],
    "Custom field(Epic Link)": ["Epic1", "Epic2", "Epic3"],
    "Custom field(External ID(Single Line))": ["Ext1", "Ext2", "Ext3"],
    "Inward issue link(Block)": ["Link1", "Link2", "Link3"],
    "Inward issue link(Causes)": ["Cause1", "Cause2", "Cause3"],
    "Inward issue link(Cloners)": ["Clone1", "Clone2", "Clone3"],
    "Inward issue link(Duplicate)": ["Dup1", "Dup2", "Dup3"],
    "Outward issue link(Duplicate)": ["ODup1", "ODup2", "ODup3"],
    "Inward issue link(Implement)": ["Implement1", "Implement2", "Implement3"],
    "Inward issue link(Relates)": ["Relates1", "Relates2", "Relates3"],
    "Outward issue link(Relates)": ["ORelates1", "ORelates2", "ORelates3"],
    "Custom field(Parent Link)": ["Parent1", "Parent2", "Parent3"],
    "Custom field(Other Text)": ["Other Text A", "Other Text B", "Other Text C"]
}
df_jira = pd.DataFrame(data_jira)
# 保存到 Excel 文件
df_jira.to_excel("excel_code/data/Jira2.xlsx", index=False)

print("已生成 Defect_data2.xlsx 和 Jira2.xlsx 文件")

'''# -----------------------------------------------------
# 3. 融合 Defect_data 与 Jira 生成总表 updated_excel
# -----------------------------------------------------
# 读取刚才生成的两个 Excel 文件
df_defect = pd.read_excel("Defect_data.xlsx")
df_jira = pd.read_excel("Jira.xlsx")

# 通过 Defect_data 的 'ID' 与 Jira 的 'issue id' 进行合并（假设是一对一匹配）
df_merged = pd.merge(df_defect, df_jira, left_on="ID", right_on="issue id", how="left")

# 构建总表 updated_excel，提取并映射需要的字段
# 注意部分列名称在原始表与目标总表之间名称存在不一致，下面做了适当映射

df_updated = pd.DataFrame({
    "Function": df_merged["Assigned ECU Group"],
    "ID": df_merged["ID"],
    "Ticket no. supplier": df_merged["Ticket no.supplier"],
    "Name": df_merged["Name"],
    "Closed in Version": df_merged["Closed in version"],
    "Involved I-Step": df_merged["Invoved I-step"],
    "First use/Sop of function": df_merged["First use/SoP of function"],
    "Top issue Candidate": df_merged["Issue key"],
    "Reporting relevance": df_merged["Reporting relevance"],
    "Creation time": df_merged["Creation time"],
    "Error occurrence": df_merged["Error occurrence"],
    "duplicated": df_merged["Inward issue link(Duplicate)"],
    "Phase": df_merged["Phase"],
    "Found in function": df_merged["Found in function"],
    "Defect finder": df_merged["Defect finder"],
    "Owner": df_merged["Owner"],
    "Comment": df_merged["Summary"],
    "Target I-Step:": df_merged["Target I-Step"],
    "Follow up": df_merged["Updated"],
    "Planned closing version": df_merged["Planned closing version"],
    "Root cause": "",  # 暂无数据，默认为空
    "Days": df_merged["Days in phase"],
    "open >20 days": df_merged["Days in phase"].apply(lambda x: "Yes" if x > 20 else "No"),
    "No TIS, Octane or Jira": "",  # 暂无数据，默认为空
    "Days in the phase": df_merged["Days in phase"]
})

# 将总表保存为 updated_excel.xlsx
df_updated.to_excel("updated_excel.xlsx", index=False)

print("已生成融合后的总表文件： updated_excel.xlsx")'''
