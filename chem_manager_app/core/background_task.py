import os
from PyQt6.QtCore import QObject, pyqtSignal, QTimer

class BackgroundManager(QObject):
    sync_requested = pyqtSignal()
    
    def __init__(self, config_data):
        super().__init__()
        self.config = config_data
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.sync_requested.emit)
        self._pause_depth = 0
        
        self.apply_config(self.config)

    def apply_config(self, config_data):
        self.config = config_data
        self._setup_timer()

    def pause_watch(self):
        self._pause_depth += 1
        self.timer.stop()
            
    def resume_watch(self):
        if self._pause_depth > 0:
            self._pause_depth -= 1
        if self._pause_depth == 0:
            self._setup_timer()

    def _setup_timer(self):
        if self._pause_depth > 0:
            self.timer.stop()
            return
        interval = self.config.get("sync_interval_minutes", 0)
        if interval > 0:
            self.timer.start(interval * 60 * 1000)
        else:
            self.timer.stop()

    def stop_all(self):
        self.timer.stop()
