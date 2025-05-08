import os
import re
import time
import json
import pandas as pd
import schedule
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime
import win32com.client  # pip install pywin32

# === 加载配置 ===
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(BASE_DIR, "unified_config_auto2.json"), 'r', encoding='utf-8') as f:
    cfg = json.load(f)

folders         = cfg['folders']
ORIG_DIR        = os.path.join(BASE_DIR, folders['orig_dir'])
JIRA_DIR        = os.path.join(BASE_DIR, folders['jira_dir'])
OCTANE_DIR      = os.path.join(BASE_DIR, folders['octane_dir'])
paths           = cfg['paths']
sheet           = cfg['sheet']['target_sheet']
sources         = cfg['sources']
settings        = cfg.get('settings', {})
clear_old_flag  = settings.get('clear_old_highlight','N').upper() == 'Y'
fund_patterns   = cfg.get('fund_function_patterns', {})
owner_patterns  = cfg.get('owner_root_cause_patterns', {})

# Excel VBA 常量（数字形式，避免依赖 constants）
XL_UP       = -4162  # xlUp
XL_TOLEFT   = -4159  # xlToLeft
XL_DATABASE = 1      # xlDatabase

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

def _refresh_pivots_in_workbook(xlsx_path, data_sheet_name):
    """
    用 COM 打开 xlsx_path，基于 data_sheet_name 的 ID 列范围
    更新并刷新整本工作簿中所有 PivotTable。
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))

    try:
        ds = wb.Worksheets(data_sheet_name)
    except Exception:
        print(f"❌ COM: 找不到数据表 “{data_sheet_name}”，跳过 Pivot 刷新")
        wb.Close(False)
        excel.Quit()
        return

    # 1) 找到“ID”列在表头的列号
    header_row = 1
    last_header_col = ds.Cells(header_row, ds.Columns.Count).End(XL_TOLEFT).Column
    id_col_idx = None
    for c in range(1, last_header_col + 1):
        if str(ds.Cells(header_row, c).Value).strip() == "ID":
            id_col_idx = c
            break
    if not id_col_idx:
        print("❌ COM: 未找到 ‘ID’ 列，无法定位数据区域，跳过 Pivot 刷新")
        wb.Close(False)
        excel.Quit()
        return

    # 2) 用 ID 列 End(xlUp) 定位最后一行
    last_row = ds.Cells(ds.Rows.Count, id_col_idx).End(XL_UP).Row
    if last_row <= header_row:
        print("⚠️ COM: ‘ID’ 列无数据，跳过 Pivot 刷新")
        wb.Close(False)
        excel.Quit()
        return

    # 3) 用表头最后一列作为列末
    last_col = last_header_col

    # 4) 构造 SourceData 引用
    top_left     = ds.Cells(header_row, 1).Address    # "$A$1"
    bottom_right = ds.Cells(last_row, last_col).Address
    source_ref   = f"'{data_sheet_name}'!{top_left}:{bottom_right}"
    print(f"  COM: 设置 Pivot 数据源为 {source_ref}")

    # 5) 遍历所有工作表的 PivotTable
    for ws in wb.Worksheets:
        try:
            pts = ws.PivotTables()  # 必须调用
            cnt = pts.Count
        except Exception:
            continue
        if cnt == 0:
            continue
        print(f"  ▶ COM: 工作表 [{ws.Name}] 有 {cnt} 个 PivotTable，开始更新…")
        for i in range(1, cnt+1):
            pt = pts.Item(i)
            try:
                cache = wb.PivotCaches().Create(
                    SourceType=XL_DATABASE,
                    SourceData=source_ref
                )
                pt.ChangePivotCache(cache)
                pt.RefreshTable()
                print(f"    ✔️ 已刷新 PivotTable [{pt.Name}]")
            except Exception as e:
                print(f"    ❌ 刷新 [{pt.Name}] 失败：{e}")

    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("✅ COM: 所有 PivotTable 已更新并刷新\n")

def update_excel(new_fp):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    print(f"[{ts}] 开始更新: {new_fp}")

    folder = os.path.basename(os.path.dirname(new_fp))
    if folder == os.path.basename(JIRA_DIR):
        source_key  = "Jira"; jira_flag='Y'; octane_flag='N'
    elif folder == os.path.basename(OCTANE_DIR):
        source_key  = "Octane"; jira_flag='N'; octane_flag='Y'
    else:
        print(f" → 未识别来源 ({folder})，跳过。")
        return

    orig_fp  = os.path.join(ORIG_DIR, paths['original_file'])
    template = paths['updated_file']
    filename = template.format(time=ts, jira=jira_flag, octane=octane_flag)
    upd_fp   = os.path.join(os.path.dirname(orig_fp), filename)

    src_cfg   = sources[source_key]
    read_meth = src_cfg['read_method']
    date_col  = src_cfg.get('date_col')
    mapping   = src_cfg['mapping']

    df_new = pd.read_excel(new_fp) if read_meth=="excel" else pd.read_csv(new_fp)

    # 如果是 Octane 来源则清色
    clear_old = clear_old_flag if source_key=="Jira" else True
    print("  clear_old =", clear_old)

    wb = load_workbook(orig_fp)
    ws = wb[sheet]

    if clear_old:
        green = {"8ED973","008ED973","FF8ED973"}
        blue  = {'FFADD8E6','00ADD8E6','ADD8E6'}
        gray  = {'FFC0C0C0','00C0C0C0','C0C0C0'}
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                pf = cell.fill
                if getattr(pf,'patternType',None)=="solid":
                    argb = pf.fgColor.rgb or pf.fgColor.value
                    if argb in green|blue|gray:
                        cell.fill = PatternFill()

    trim_trailing_blank_rows(ws)
    # 重置行高
    for r in range(1, ws.max_row+1):
        ws.row_dimensions[r].height = None

    header2col = { ws.cell(1,c).value:c for c in range(1, ws.max_column+1) }
    id_col     = header2col['ID']

    # 已有 ID→行
    id2row = { ws.cell(r,id_col).value:r
               for r in range(2, ws.max_row+1)
               if ws.cell(r,id_col).value }

    # 提取超链接（若 Excel）
    id2url = {}
    if read_meth=="excel":
        tmpwb = load_workbook(new_fp); nws=tmpwb.active
        cid   = next((c for c in range(1,nws.max_column+1) if nws.cell(1,c).value=="ID"), None)
        if cid:
            for r in range(2,nws.max_row+1):
                cell=nws.cell(r,cid)
                if cell.hyperlink:
                    id2url[cell.value] = cell.hyperlink.target

    last_row      = find_last_data_row(ws, id_col)
    update_fill   = PatternFill("solid", fgColor="ADD8E6")
    new_fill      = PatternFill("solid", fgColor="8ED973")
    id_key        = next(k for k,v in mapping.items() if v=="ID")

    # 更新 or 追加
    for _, new_row in df_new.iterrows():
        nid = new_row.get(id_key)
        if pd.isna(nid) or not nid: continue

        if nid in id2row:
            r = id2row[nid]
            phase = ws.cell(r, header2col["Phase"]).value or ""
            if any(s in phase for s in ("Concluded","Closed","Resolved")):
                continue
            # 更新 Creation time
            if date_col:
                raw = new_row.get(date_col)
                if pd.notna(raw):
                    try:
                        dt = pd.to_datetime(raw).to_pydatetime()
                        c  = header2col[mapping[date_col]]
                        cell=ws.cell(r,c)
                        if cell.value!=dt:
                            cell.value=dt
                            cell.number_format="m/d/yyyy h:mm:ss AM/PM"
                            cell.fill=update_fill
                    except: pass
            # 更新其他
            for nk, ok in mapping.items():
                if nk==date_col or ok=="Function": continue
                val=new_row.get(nk)
                if pd.isna(val) or val=="": continue
                if ok=="Involved I-Step":
                    s=str(val)
                    if s.startswith(("G070","U006")):
                        val="NA05"+s[4:]
                    else:
                        m=re.search(r'[（(]([\d-]+)',s)
                        if m: val=f"NA05-{m.group(1)}"
                c=header2col.get(ok)
                if c and ws.cell(r,c).value!=val:
                    ws.cell(r,c).value=val
                    ws.cell(r,c).fill=update_fill
            # Root cause
            if "Owner" in header2col and "Root cause" in header2col:
                ow=str(ws.cell(r,header2col["Owner"]).value or "").lower()
                for cause,names in owner_patterns.items():
                    if any(n.lower() in ow for n in names):
                        c=ws.cell(r,header2col["Root cause"])
                        if c.value!=cause:
                            c.value=cause; c.fill=update_fill
                        break
        else:
            last_row+=1
            for idx,hdr in enumerate(header2col.keys(),start=1):
                val=""
                for nk,ok in mapping.items():
                    if ok==hdr:
                        tmp=new_row.get(nk)
                        if pd.notna(tmp):
                            val=tmp
                        break
                c=ws.cell(last_row,idx)
                c.value=val
                if val:
                    c.fill=new_fill
                    if hdr==mapping.get(date_col):
                        c.number_format="m/d/yyyy h:mm:ss AM/PM"
            if nid in id2url:
                lc=ws.cell(last_row,id_col)
                lc.hyperlink,lc.style=id2url[nid],"Hyperlink"

    # 公式 & 标记略（同之前）

    # 保存回原文件
    wb.save(orig_fp)
    print(f"[{ts}] 写回：{orig_fp}")

    # ---- PivotTable 自动刷新 ----
    _refresh_pivots_in_workbook(orig_fp, sheet)

class FolderHandler(FileSystemEventHandler):
    def __init__(self, folders): self.folders=folders
    def on_created(self,e): self._maybe(e.src_path)
    on_modified = on_created
    def on_moved(self,e): self._maybe(e.dest_path)
    def _maybe(self,p):
        if os.path.isdir(p): return
        f=os.path.basename(os.path.dirname(p))
        if f in (self.folders['jira_dir'],self.folders['octane_dir']):
            update_excel(p)

def main():
    for d in (ORIG_DIR,JIRA_DIR,OCTANE_DIR): os.makedirs(d,exist_ok=True)
    obs=Observer(); h=FolderHandler(folders)
    obs.schedule(h,path=JIRA_DIR,recursive=False)
    obs.schedule(h,path=OCTANE_DIR,recursive=False)
    obs.start()
    print(f"监控：{JIRA_DIR},{OCTANE_DIR}")
    try:
        while True:
            schedule.run_pending(); time.sleep(1)
    except KeyboardInterrupt:
        obs.stop(); obs.join()

if __name__=="__main__":
    main()
