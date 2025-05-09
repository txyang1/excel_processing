import time
from watchdog.events import FileSystemEventHandler

class FolderHandler(FileSystemEventHandler):
    def __init__(self, folders, debounce_seconds=5):
        self.folders = folders
        # 用来记录上次对同一路径触发更新的时间戳
        self._last_run = {}
        self._debounce = debounce_seconds

    def on_created(self, event):
        self._maybe_update(event.src_path)

    on_modified = on_created

    def on_moved(self, event):
        self._maybe_update(event.dest_path)

    def _maybe_update(self, path):
        # 只处理文件
        if os.path.isdir(path):
            return
        dir_name = os.path.basename(os.path.dirname(path))
        if dir_name not in (self.folders['jira_dir'], self.folders['octane_dir']):
            return

        now = time.time()
        last = self._last_run.get(path, 0)
        # 如果上次更新离现在小于阈值，就跳过
        if now - last < self._debounce:
            print(f"⚠️ 去抖：忽略短时间内重复触发 {path}")
            return

        # 记录本次时间，执行更新
        self._last_run[path] = now
        update_excel(path)
