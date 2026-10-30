"""原生 Qt Widgets 的统一浅色主题。"""
STYLESHEET = """
QWidget { font-family: 'Microsoft YaHei UI', 'Segoe UI'; font-size: 13px; color: #334155; }
QMainWindow { background: #f4f7fc; }
QMenuBar, QMenu { background: white; }
QLabel#brand { font-size: 25px; font-weight: 700; color: #0f172a; }
QLabel#subtitle { color: #64748b; font-size: 12px; }
QLabel#sectionTitle { color: #0f172a; font-size: 14px; font-weight: 600; }
QLabel#detailName { color: #0f172a; font-size: 16px; font-weight: 600; }
QLabel#muted { color: #64748b; }
QLabel#badge { color: #2563eb; background: #eff6ff; border-radius: 5px; padding: 5px 8px; }
QFrame#panel { background: white; border: 1px solid #e2e8f0; border-radius: 10px; }
QLineEdit { background: white; border: 1px solid #cbd5e1; border-radius: 7px; padding: 9px 12px; selection-background-color: #bfdbfe; selection-color: #0f172a; }
QLineEdit:focus { border: 1px solid #2563eb; }
QLineEdit#searchInput { font-size: 16px; padding: 12px 14px; }
QPushButton { background: white; border: 1px solid #cbd5e1; border-radius: 7px; padding: 9px 13px; }
QPushButton:hover { background: #eff6ff; border-color: #93c5fd; }
QPushButton:pressed { background: #dbeafe; }
QPushButton:disabled { color: #94a3b8; border-color: #e2e8f0; background: #f8fafc; }
QPushButton#primary { background: #2563eb; color: white; border: 1px solid #2563eb; font-weight: 600; }
QPushButton#primary:hover { background: #1d4ed8; }
QPushButton#primary:disabled { background: #93c5fd; border-color: #93c5fd; }
QCheckBox { spacing: 6px; padding: 4px 0; }
QListWidget { border: none; background: white; outline: none; }
QListWidget::item { border-radius: 5px; padding: 11px 8px; margin: 2px 0; }
QListWidget::item:selected { background: #eff6ff; color: #1d4ed8; }
QListWidget::item:hover { background: #f1f5f9; }
QTableView { border: none; background: white; alternate-background-color: #f8fafc; gridline-color: #f1f5f9; selection-background-color: #dbeafe; selection-color: #0f172a; outline: none; }
QHeaderView::section { border: none; border-bottom: 1px solid #e2e8f0; background: #f8fafc; color: #64748b; font-size: 12px; padding: 10px 7px; }
QTextBrowser { border: none; background: white; padding: 4px; }
QSplitter::handle { background: transparent; }
QProgressBar { border: none; background: #e2e8f0; height: 4px; border-radius: 2px; }
QProgressBar::chunk { background: #2563eb; border-radius: 2px; }
QToolTip { background: #0f172a; color: white; border: none; padding: 7px; }
"""
