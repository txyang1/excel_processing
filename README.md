# excel_processing
# auto\_excel\_update README

## 概述

`auto_excel_update.py` 是一个自动化工具，用于监控两个文件夹（JIRA 导出与 Octane 导出），当有新文件或更新文件时，将其与原始汇总表进行比对并更新，然后将更新后的文件保存到指定目录，并且每天定时执行全量更新。

主要功能：

* 监控 JIRA 与 Octane 导出目录，新增或修改文件时自动触发更新
* 按配置读取新导出文件（Excel 或 CSV），比对原始汇总表中的工单 ID
* 对已有工单行进行字段更新并标绿色高亮，对新增工单行添加并标粉色高亮
* 更新完成后，将结果保存到 `Orig_files` 目录，并在文件名中加入时间戳、来源标识（Jira/Octane）
* 每日 08:00 定时全量扫描更新

## 目录结构

```
项目根目录/
├─ auto_excel_update.py         # 主脚本
├─ unified_config_auto2.json    # 配置文件
├─ Orig_files/                  # 更新后文件输出目录（与 ORIG_DIR 对应）
├─ JIRA_exports/                # JIRA 导出文件夹
└─ OCTANE_exports/              # Octane 导出文件夹
```

## 环境依赖

* Python 3.6+
* pandas
* openpyxl
* schedule
* watchdog

安装依赖：

```bash
pip install pandas openpyxl schedule watchdog
```

## 配置文件 `unified_config_auto2.json`

示例结构：

```json
{
  "folders": {
    "orig_dir": "Orig_files",
    "jira_dir": "JIRA_exports",
    "octane_dir": "OCTANE_exports"
  },
  "paths": {
    "original_file": "汇总表.xlsx",
    "updated_file": "tx_auto_updated_{time}_jira_{jira}_octane_{octane}.xlsx"
  },
  "sheet": { "target_sheet": "Data" },
  "sources": {
    "Jira": {
      "read_method": "excel",
      "date_col": "Updated Date",
      "mapping": { /* 字段映射 */ }
    },
    "Octane": {
      "read_method": "csv",
      "date_col": null,
      "mapping": { /* 字段映射 */ }
    }
  },
  "settings": {
    "clear_old_highlight": "Y"
  },
  "fund_function_patterns": {},
  "owner_root_cause_patterns": {}
}
```

* **orig\_dir**：更新后文件保存目录
* **jira\_dir**/**octane\_dir**：分别为 JIRA/Octane 导出文件夹
* **original\_file**：原始汇总表文件名
* **updated\_file**：输出文件名模板，可使用 `{time}`、`{jira}`、`{octane}` 占位
* **sources**：分别配置 JIRA 与 Octane 数据源的读取方式、日期列与字段映射
* **clear\_old\_highlight**：是否清除旧高亮 (`Y`/`N`)
* **fund\_function\_patterns**、**owner\_root\_cause\_patterns**：可选的补充规则

## 使用方法

1. **配置**

   * 修改 `unified_config_auto2.json` 中的目录与映射
   * 将 JIRA 导出的 Excel 放入 `JIRA_exports`，Octane 导出的 CSV 放入 `OCTANE_exports`

2. **运行脚本**

   ```bash
   python tx_auto_update2.3.1.2.py
   ```

3. **结果**

   * 监控期间新文件加入或修改时，自动更新并输出到 `Orig_files`
   * 每天 08:00 自动执行全量更新

## 自定义扩展

* 可以在 `sources` 中添加更多数据源类型
* 在 `settings` 中调整定时任务时间或是否清除旧高亮
* 如需更多字段映射规则，可在 `fund_function_patterns` 与 `owner_root_cause_patterns` 中配置

---

如有问题或建议，欢迎反馈！
