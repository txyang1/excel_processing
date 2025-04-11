import pandas as pd

# -------------------------------
# 1. 读取源文件数据
# -------------------------------
df_defect = pd.read_excel("excel_code/data/Defect_data.xlsx")
df_jira = pd.read_excel("excel_code/data/Jira.xlsx")

# -------------------------------
# 2. 合并数据
# -------------------------------
# 假设 Defect_data 的 "ID" 与 Jira 的 "issue id" 进行一对一匹配
df_merged = pd.merge(df_defect, df_jira, left_on="ID", right_on="issue id", how="left")

# -------------------------------
# 3. 定义目标字段映射关系
# -------------------------------
# mapping_dict 中的 key 为总表 updated_excel 中的字段名称，
# value 为来源表中的字段名称（如在 Defect_data 或 Jira 中），
# 若目标字段的值需要自定义处理（如计算或没有对应来源字段），则置为空字符串或单独处理。
mapping_dict = {
    "Function": "Assigned ECU Group",             # 来源：Defect_data
    "ID": "ID",                                    # 同名字段
    "Ticket no. supplier": "Ticket no.supplier",   # 来源：Defect_data
    "Name": "Name",                                # 来源：Defect_data
    "Closed in Version": "Closed in version",      # 来源：Defect_data
    "Involved I-Step": "Invoved I-step",           # 来源：Defect_data
    "First use/Sop of function": "First use/SoP of function",  # 来源：Defect_data
    "Top issue Candidate": "Issue key",            # 来源：Jira
    "Reporting relevance": "Reporting relevance",  # 来源：Defect_data
    "Creation time": "Creation time",              # 来源：Defect_data
    "Error occurrence": "Error occurrence",        # 来源：Defect_data
    "duplicated": "Inward issue link(Duplicate)",    # 来源：Jira
    "Phase": "Phase",                              # 来源：Defect_data
    "Found in function": "Found in function",      # 来源：Defect_data
    "Defect finder": "Defect finder",              # 来源：Defect_data
    "Owner": "Owner",                              # 来源：Defect_data
    "Comment": "Summary",                          # 来源：Jira
    "Target I-Step:": "Target I-Step",              # 来源：Defect_data
    "Follow up": "Updated",                        # 来源：Jira
    "Planned closing version": "Planned closing version",  # 来源：Defect_data
    "Root cause": "",                              # 无来源，默认空
    "Days": "Days in phase",                       # 来源：Defect_data
    "open >20 days": "Days in phase",              # 根据“Days in phase”计算（大于20为“Yes”，否则“No”）
    "No TIS, Octane or Jira": "",                  # 无来源，默认空
    "Days in the phase": "Days in phase"           # 来源：Defect_data
}

# -------------------------------
# 4. 根据映射构造总表 DataFrame
# -------------------------------
data_updated = {}
for target_field, source_field in mapping_dict.items():
    if target_field == "open >20 days":  
        # 根据"Days in phase"计算，如果该列存在则进行条件判断
        if source_field in df_merged.columns:
            data_updated[target_field] = df_merged[source_field].apply(lambda x: "Yes" if x > 20 else "No")
        else:
            data_updated[target_field] = ["No"] * len(df_merged)
    elif source_field == "":
        # 如果没有对应来源字段，填充空值或默认值
        data_updated[target_field] = ""
    else:
        # 如果来源字段存在，则直接赋值；否则填空字符串
        if source_field in df_merged.columns:
            data_updated[target_field] = df_merged[source_field]
        else:
            data_updated[target_field] = ""
            
# 构造总表 DataFrame
df_updated = pd.DataFrame(data_updated)

# -------------------------------
# 5. 保存总表
# -------------------------------
df_updated.to_excel("excel_code/results/updated_excel.xlsx", index=False)
print("已生成 updated_excel.xlsx")
