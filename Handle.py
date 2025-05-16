class FolderHandler(FileSystemEventHandler):
    def __init__(self, folders, debounce_seconds=5):
        self.folders   = folders
        self._last_run = {}
        self._debounce = debounce_seconds

    # 只实现这一个方法
    def on_created(self, event):
        if event.is_directory:
            return

        path = os.path.normcase(os.path.abspath(event.src_path))
        # 如果是 .tmp 临时文件，直接跳过
        if path.lower().endswith('.tmp'):
            return

        folder = os.path.basename(os.path.dirname(path))
        if folder not in (self.folders['jira_dir'], self.folders['octane_dir']):
            return

        now = time.time()
        if now - self._last_run.get(path, 0) < self._debounce:
            return  # 去抖

        self._last_run[path] = now
        update_excel(path)

    # 不实现也可以，下面这几行可删可留
    def on_moved(self, event):   pass
    def on_deleted(self, event): pass
    def on_modified(self, event):pass
