import os
from PyQt6.QtCore import QObject, pyqtSignal, QTimer
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, callback, target_file):
        super().__init__()
        self.callback = callback
        self.target_file = os.path.abspath(target_file)
        
    def on_modified(self, event):
        if not event.is_directory and os.path.abspath(event.src_path) == self.target_file:
            self.callback()

class BackgroundManager(QObject):
    sync_requested = pyqtSignal()
    
    def __init__(self, config_data):
        super().__init__()
        self.config = config_data
        self.observer = None
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.sync_requested.emit)
        
        self.apply_config(self.config)

    def apply_config(self, config_data):
        self.config = config_data
        self._setup_watcher()
        self._setup_timer()

    def _setup_watcher(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
            
        src_file_raw = self.config.get("source_file", "")
        import urllib.parse
        if src_file_raw.startswith("file:///"):
            src_file_raw = urllib.parse.unquote(src_file_raw[8:])
            
        if self.config.get("watch_enabled", False) and os.path.exists(src_file_raw):
            target_dir = os.path.dirname(src_file_raw)
            handler = FileChangeHandler(self.sync_requested.emit, src_file_raw)
            self.observer = Observer()
            self.observer.schedule(handler, target_dir, recursive=False)
            self.observer.start()

    def pause_watch(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
            
    def resume_watch(self):
        if self.config.get("watch_enabled", False):
            self._setup_watcher()

    def _setup_timer(self):
        interval = self.config.get("sync_interval_minutes", 0)
        if interval > 0:
            self.timer.start(interval * 60 * 1000)
        else:
            self.timer.stop()

    def stop_all(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
        self.timer.stop()
