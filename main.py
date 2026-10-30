"""源码入口与 PyInstaller 入口共用。"""
import ctypes
import sys
from pathlib import Path

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication, QMessageBox

from ppt_search.settings import configure_logging
from ppt_search.theme import STYLESHEET
from ppt_search.window import MainWindow


def main():
    program_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    app = QApplication(sys.argv)
    app.setApplicationName("PPT Search")
    app.setOrganizationName("PPTSearch")
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setStyleSheet(STYLESHEET)
    if sys.platform == "win32":
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PPTSearch.Desktop.1")
    try:
        configure_logging(program_dir / "data")
    except OSError as error:
        QMessageBox.critical(None, "无法写入数据目录", f"请把程序解压到可写文件夹后再启动。\n{error}")
        return 1
    initial_root = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 and Path(sys.argv[1]).is_dir() else None
    window = MainWindow(program_dir, program_dir / "data", initial_root)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
