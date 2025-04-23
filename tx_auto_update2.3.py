# auto_excel_update.py

import os
import time
import json
import pandas as pd
import schedule
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# === 读取配置 ===
BASE_DIR  = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(BASE_DIR, "unified_config2.7.1.json"), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

# 文件夹配置
folders   = cfg['folders']
ORIG_DIR   = os.path.join(BASE_DIR, folders['orig_dir'])
JIRA_DIR   = os.path.join(BASE_DIR, folders['jira_dir'])
OCTANE_DIR = os.path.join(BASE_DIR, folders['octane_dir'])

# JSON 中的 paths 和其他设置
paths     = cfg['paths']
sheet     = cfg['sheet']['target_sheet']
sources   = cfg['sources']
settings  = cfg.get('settings', {})
clear_old = settings.get('clear_old_highlight','N').upper() == 'Y'
fund_pats = cfg.get('fund_function_patterns', {})
owner_pats= cfg.get('owner_root_cause_patterns', {})

def is_row_blank(ws, row):
    for c in range(1, ws.max_column + 1):
        if ws.cell(row, c).value not in (None, ""):
            return False
    return True

def trim_trailing_blank_rows(ws):
    for r in range(ws.max_row, 1, -1):
        if is_row_blank(ws, r):
            ws.delete_rows(r)
        else:
            break

def find_last_data_row(ws, key_col):
    for r in range(ws.max_row, 1, -1):
        if ws.cell(row=r, column=key_col).value not in (None, ""):
            return r
    return 1

def update_excel(new_fp):
    """针对单个新文件 new_fp，读取 Orig_files/summary.xlsx，更新后保存到同一文件夹下的 updated_file 名称。"""
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 开始更新: {new_fp}")
    folder = os.path.basename(os.path.dirname(new_fp))
    source_key = None
    if folder == os.path.basename(JIRA_DIR):   source_key = "Jira"
    if folder == os.path.basename(OCTANE_DIR): source_key = "Octane"
    if not source_key:
        print(f"  → 未识别来源 ({folder})，跳过。")
        return

    # 加载原表与目标路径
    orig_fp = os.path.join(ORIG_DIR, paths['original_file'])
    upd_name= paths['updated_file']
    dest_dir= os.path.dirname(new_fp)
    upd_fp  = os.path.join(dest_dir, upd_name)

    src_cfg   = sources[source_key]
    read_meth = src_cfg['read_method']
    date_col  = src_cfg.get('date_col')
    mapping   = src_cfg['mapping']

    # 读取新表
    df_new = pd.read_excel(new_fp) if read_meth=="excel" else pd.read_csv(new_fp)

    # 打开原表
    wb = load_workbook(orig_fp)
    ws = wb[sheet]

    # 清除旧高亮
    if clear_old:
        pink = {'FFC0CB','00FFC0CB','FFFFC0CB'}
        blue = {'FFADD8E6','00ADD8E6','ADD8E6'}
        gray = {'FFC0C0C0','00C0C0C0','C0C0C0'}
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                pf = cell.fill
                if getattr(pf,'patternType',None)=="solid":
                    argb = pf.fgColor.rgb or pf.fgColor.value
                    if argb in pink|blue|gray:
                        cell.fill = PatternFill()

    trim_trailing_blank_rows(ws)

    # 表头与已有 ID→行
    header2col = {ws.cell(1,c).value:c for c in range(1,ws.max_column+1)}
    id_col     = header2col['ID']
    id2row     = { ws.cell(r,id_col).value:r
                   for r in range(2, ws.max_row+1)
                   if ws.cell(r,id_col).value }

    # 收集新表 ID 超链接（若 excel 且有超链）
    id2url = {}
    if read_meth=="excel":
        tmp = load_workbook(new_fp).active
        cid = next((c for c in range(1,tmp.max_column+1)
                    if tmp.cell(1,c).value=="ID"), None)
        if cid:
            for r in range(2, tmp.max_row+1):
                cell = tmp.cell(r,cid)
                if cell.hyperlink:
                    id2url[cell.value] = cell.hyperlink.target

    last_row    = find_last_data_row(ws, id_col)
    pink_fill   = PatternFill('solid', fgColor='FFC0CB')
    blue_fill   = PatternFill('solid', fgColor='ADD8E6')

    # 更新/追加
    id_key = next(k for k,v in mapping.items() if v=="ID")
    for _, new_row in df_new.iterrows():
        nid = new_row.get(id_key)
        if pd.isna(nid) or not nid: continue

        if nid in id2row:
            r = id2row[nid]
            phase = ws.cell(r, header2col.get('Phase')).value or ""
            if any(x in phase for x in ("Concluded","Closed","Resolved")):
                continue
            # 更新字段（略去细节，可按 mapping 逻辑更新并填 blue_fill）
            # ……

        else:
            # 新增整行
            last_row += 1
            for col_name, cidx in header2col.items():
                val = ""
                for nk, ok in mapping.items():
                    if ok == col_name:
                        tmp = new_row.get(nk)
                        if pd.notna(tmp): val = tmp
                        break
                cell = ws.cell(last_row, cidx)
                cell.value = val
                if val: cell.fill = pink_fill

            # 超链接
            if nid in id2url:
                c = ws.cell(last_row, id_col)
                c.hyperlink, c.style = id2url[nid], "Hyperlink"

    # 填充公式示例：Days 列
    if "Days" in header2col and "Creation time" in header2col:
        colL = get_column_letter(header2col["Creation time"])
        for r in range(2, ws.max_row+1):
            ws.cell(r, header2col["Days"]).value = f'=DATEDIF(${colL}{r},TODAY(),"D")'

    # 保存
    wb.save(upd_fp)
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 更新完成，文件已保存到: {upd_fp}")

class FolderHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            folder = os.path.basename(os.path.dirname(event.src_path))
            if folder in (folders['jira_dir'], folders['octane_dir']):
                update_excel(event.src_path)
    on_modified = on_created

def main():
    # 确保目录存在
    for d in (ORIG_DIR, JIRA_DIR, OCTANE_DIR):
        os.makedirs(d, exist_ok=True)

    # 启动 Watchdog
    observer = Observer()
    observer.schedule(FolderHandler(), path=JIRA_DIR, recursive=False)
    observer.schedule(FolderHandler(), path=OCTANE_DIR, recursive=False)
    observer.start()
    print(f"监控目录：{JIRA_DIR}，{OCTANE_DIR}")

    # 每日 08:00 全量扫描
    schedule.every().day.at("08:00").do(lambda: [
        update_excel(os.path.join(JIRA_DIR, fn)) for fn in os.listdir(JIRA_DIR)
        if fn.lower().endswith((".xlsx",".csv"))
    ] + [
        update_excel(os.path.join(OCTANE_DIR, fn)) for fn in os.listdir(OCTANE_DIR)
        if fn.lower().endswith((".xlsx",".csv"))
    ])
    print("已添加每日 08:00 全量扫描更新任务。")

    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
