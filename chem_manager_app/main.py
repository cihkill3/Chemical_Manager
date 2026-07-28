import sys
import os
import msvcrt
from PyQt6.QtWidgets import QApplication, QStyle
from PyQt6.QtGui import QIcon

from core.config_manager import load_config, save_config
from core.background_task import BackgroundManager
from gui.main_window import MainWindow, SyncWorker
from gui.tray_icon import TrayIcon

class AppController:
    def __init__(self):
        self.app = QApplication(sys.argv)
        
        # Single Instance Check
        self.lock_file_path = os.path.join(os.environ.get("TEMP", "."), "chem_manager.lock")
        self.lock_file = open(self.lock_file_path, "w")
        try:
            msvcrt.locking(self.lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        except IOError:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, "프로그램이 이미 실행 중입니다.", "알림", 0x30)
            sys.exit(0)
            
        self.app.setQuitOnLastWindowClosed(False)
        
        self.config = load_config()
        
        # Setup Background Manager
        self.bg_manager = BackgroundManager(self.config)
        self.bg_manager.sync_requested.connect(self.on_background_sync)
        
        # Setup Main Window
        self.main_window = MainWindow(self.config)
        self.main_window.config_updated.connect(self.on_config_updated)
        self.main_window.tray_requested.connect(self.hide_window)
        self.main_window.close_requested.connect(self.quit_app)
        self.main_window.manual_sync_started.connect(self.bg_manager.pause_watch)
        self.main_window.manual_sync_finished.connect(self.bg_manager.resume_watch)
        
        # Setup Tray Icon
        icon = self.app.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.tray_icon = TrayIcon(icon)
        self.tray_icon.show_window_requested.connect(self.show_window)
        self.tray_icon.exit_requested.connect(self.quit_app)
        self.tray_icon.show()
        
        # Show Window initially
        self.main_window.show()
        
        self.bg_worker = None

    def on_config_updated(self, new_config):
        self.config = new_config
        save_config(self.config)
        self.bg_manager.apply_config(self.config)

    def on_background_sync(self):
        if self.bg_worker is not None and self.bg_worker.isRunning():
            return # Already syncing
        
        self.bg_manager.pause_watch()
        self.main_window.log("[백그라운드] 자동 동기화 시작...")
        self.bg_worker = SyncWorker(self.config)
        self.bg_worker.progress.connect(lambda msg: self.main_window.log(f"[백그라운드] {msg}"))
        self.bg_worker.finished.connect(self.on_background_finished)
        self.bg_worker.start()

    def on_background_finished(self, result):
        self.bg_manager.resume_watch()
        if result.get("success"):
            new = result.get('new', 0)
            upd = result.get('updated', 0)
            if new > 0 or upd > 0:
                self.tray_icon.show_message("동기화 완료", f"신규 {new}건, 업데이트 {upd}건이 반영되었습니다.")
            self.main_window.log("[백그라운드] 동기화 성공")
        else:
            self.tray_icon.show_message("동기화 오류", "자동 동기화 중 오류가 발생했습니다.")
            self.main_window.log(f"[백그라운드] 오류: {result.get('error')}")

    def hide_window(self):
        self.main_window.hide()
        self.tray_icon.show_message("백그라운드 실행 중", "앱이 트레이로 최소화되었습니다. 자동 동기화가 설정에 따라 계속 실행됩니다.")

    def show_window(self):
        self.main_window.show()
        self.main_window.activateWindow()

    def quit_app(self):
        self.bg_manager.stop_all()
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec())

if __name__ == "__main__":
    controller = AppController()
    controller.run()
