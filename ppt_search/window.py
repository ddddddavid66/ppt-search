"""主窗口负责交互编排，提取、搜索和自动化分别由独立模块完成。"""
import logging
from pathlib import Path

from PySide6.QtCore import QTimer, Qt, QUrl
from PySide6.QtGui import QAction, QDesktopServices, QKeySequence, QTextOption
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
    QProgressBar, QPushButton, QSplitter, QTableView, QTextBrowser, QVBoxLayout, QWidget,
)

from .cache import TextCache
from .models import SearchOptions
from .powerpoint import OpenSlideWorker
from .presentation import FileDelegate, HighlightDelegate, ResultsFilter, ResultsModel, detail_html, export_csv
from .settings import load_settings, save_settings
from .workers import SearchWorker

LOGGER = logging.getLogger(__name__)

HELP_TEXT = """<h3>PPT Search · 使用说明</h3>
<p>选择或拖入文件夹，输入关键词后按 Enter。单击结果查看文字，双击在 PowerPoint 中打开原页。</p>
<p><b>搜索范围：</b>PPTX 标题、文本框、表格、嵌套组合形状和讲者备注。
按字面包含匹配，默认不区分大小写；完整短语可直接输入，无须加引号。</p>
<p><b>实际页码：</b>从第 1 张开始，按左侧缩略图顺序计数，包含隐藏页；
不使用页脚中自定义的编号。读取磁盘上已保存的版本。</p>
<p><b>首版边界：</b>不搜索图片文字、SmartArt 内部文字、图表内部数据、嵌入附件和旧格式 .ppt。
无原页图片预览。不支持解密加密文件。</p>
<p>搜索无需安装 PowerPoint；准确跳转需 Microsoft PowerPoint。
若无法跳转，可打开源文件或所在目录。不会保存或关闭您的文稿。</p>
<p>支持：Enter 搜索、Ctrl+F 聚焦搜索、Ctrl+O 选择目录、F5 刷新、Esc 取消。
导出 CSV 为当前筛选的结果。缓存和日志位于程序目录的 data，可在工具菜单清除缓存。</p>
<p>程序完全离线，不下载模型，不上传文件。版本 1.0.0。</p>"""


def panel(title: str):
    frame = QFrame()
    frame.setObjectName("panel")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(14, 15, 14, 12)
    heading = QLabel(title)
    heading.setObjectName("sectionTitle")
    layout.addWidget(heading)
    return frame, layout


class MainWindow(QMainWindow):
    def __init__(self, program_dir: Path, data_dir: Path, initial_root: Path | None = None):
        super().__init__()
        self.program_dir = program_dir.resolve()
        self.data_dir = data_dir
        self.settings_path = data_dir / "settings.json"
        self.settings = load_settings(self.settings_path)
        self.root = initial_root or Path(self.settings.get("directory", str(program_dir)))
        if not self.root.is_dir():
            self.root = program_dir
        self.root = self.root.resolve()
        self.job_id = 0
        self.search_workers = {}
        self.open_workers = {}
        self.documents = {}
        self.file_items = {}
        self.current_options = SearchOptions("linux")
        self.current_hit = None
        self.failures = []
        self.searching = False
        self.closing = False
        self._pending_open = None
        self.setWindowTitle("PPT Search · 幻灯片全文搜索")
        self.resize(1280, 800)
        self.setMinimumSize(1000, 650)
        self.setAcceptDrops(True)
        self._build_ui()
        self._build_menu()
        self._connect_events()
        self._set_detail_enabled(False)
        QTimer.singleShot(0, lambda: self.start_task(catalog=True))

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(13)
        header = QHBoxLayout()
        brand = QLabel("PPT Search")
        brand.setObjectName("brand")
        header.addWidget(brand)
        subtitle = QLabel("在课件中，快速找到答案")
        subtitle.setObjectName("subtitle")
        header.addWidget(subtitle)
        header.addStretch()
        badge = QLabel("本地离线  ·  精确到页")
        badge.setObjectName("badge")
        header.addWidget(badge)
        layout.addLayout(header)

        directory_row = QHBoxLayout()
        directory_row.addWidget(QLabel("搜索目录"))
        self.directory_input = QLineEdit(str(self.root))
        self.directory_input.setObjectName("directoryInput")
        directory_row.addWidget(self.directory_input, 1)
        self.browse_button = QPushButton("选择文件夹")
        self.refresh_button = QPushButton("刷新")
        directory_row.addWidget(self.browse_button)
        directory_row.addWidget(self.refresh_button)
        layout.addLayout(directory_row)

        search_row = QHBoxLayout()
        self.keyword_input = QLineEdit()
        self.keyword_input.setObjectName("searchInput")
        self.keyword_input.setPlaceholderText("搜索幻灯片内容，例如 linux、网络配置……")
        self.keyword_input.setClearButtonEnabled(True)
        self.search_button = QPushButton("搜索")
        self.search_button.setObjectName("primary")
        self.search_button.setMinimumWidth(108)
        self.search_button.setMinimumHeight(48)
        search_row.addWidget(self.keyword_input, 1)
        search_row.addWidget(self.search_button)
        layout.addLayout(search_row)
        options_row = QHBoxLayout()
        self.recursive_check = QCheckBox("包含子目录")
        self.notes_check = QCheckBox("搜索备注")
        self.case_check = QCheckBox("区分大小写")
        for name, checkbox, default in (("recursive", self.recursive_check, True),
                                        ("notes", self.notes_check, True),
                                        ("case_sensitive", self.case_check, False)):
            checkbox.setChecked(bool(self.settings.get(name, default)))
            options_row.addWidget(checkbox)
        options_row.addStretch()
        hint = QLabel("支持正文、表格、组合形状和备注")
        hint.setObjectName("muted")
        options_row.addWidget(hint)
        layout.addLayout(options_row)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        files_panel, files_layout = panel("文件列表")
        self.file_count = QLabel("正在读取目录…")
        self.file_count.setObjectName("muted")
        files_layout.addWidget(self.file_count)
        self.file_list = QListWidget()
        self.file_list.setObjectName("fileList")
        self.file_list.setItemDelegate(FileDelegate(self.file_list))
        self.file_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        files_layout.addWidget(self.file_list)
        files_panel.setMinimumWidth(170)
        splitter.addWidget(files_panel)

        results_panel, results_layout = panel("搜索结果")
        self.result_count = QLabel("输入关键词，跨文件查找幻灯片内容")
        self.result_count.setObjectName("muted")
        results_layout.addWidget(self.result_count)
        self.results_model = ResultsModel(self)
        self.results_filter = ResultsFilter(self)
        self.results_filter.setSourceModel(self.results_model)
        self.table = QTableView()
        self.table.setObjectName("resultsTable")
        self.table.setModel(self.results_filter)
        self.table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(46)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 145)
        self.table.setColumnWidth(1, 70)
        self.table.setColumnWidth(2, 55)
        self.delegate = HighlightDelegate(self.table)
        self.table.setItemDelegate(self.delegate)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(0, Qt.SortOrder.AscendingOrder)
        results_layout.addWidget(self.table, 1)
        self.empty_label = QLabel("选择课件目录，输入关键词开始搜索。\n也可以把文件夹拖到这个窗口。")
        self.empty_label.setWordWrap(True)
        self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_label.setObjectName("muted")
        results_layout.addWidget(self.empty_label)
        results_panel.setMinimumWidth(400)
        splitter.addWidget(results_panel)

        details_panel, details_layout = panel("内容详情")
        self.detail_name = QLabel("尚未选择结果")
        self.detail_name.setObjectName("detailName")
        self.detail_name.setWordWrap(True)
        self.detail_meta = QLabel("单击结果查看命中文本及上下文")
        self.detail_meta.setObjectName("muted")
        self.detail_meta.setWordWrap(True)
        self.detail_path = QTextBrowser()
        self.detail_path.setObjectName("muted")
        self.detail_path.setMaximumHeight(58)
        self.detail_path.setMinimumHeight(42)
        self.detail_path.setWordWrapMode(QTextOption.WrapMode.WrapAnywhere)
        self.detail_path.setStyleSheet("color:#64748b;font-size:11px;padding:0;")
        details_layout.addWidget(self.detail_name)
        details_layout.addWidget(self.detail_meta)
        details_layout.addWidget(self.detail_path)
        self.details_browser = QTextBrowser()
        self.details_browser.setObjectName("detailsBrowser")
        self.details_browser.setOpenLinks(False)
        self.details_browser.setOpenExternalLinks(False)
        details_layout.addWidget(self.details_browser, 1)
        self.open_button = QPushButton("在 PowerPoint 中打开此页")
        self.open_button.setObjectName("primary")
        details_layout.addWidget(self.open_button)
        actions = QHBoxLayout()
        self.folder_button = QPushButton("打开文件位置")
        self.copy_button = QPushButton("复制文本")
        actions.addWidget(self.folder_button)
        actions.addWidget(self.copy_button)
        details_layout.addLayout(actions)
        self.default_open_button = QPushButton("使用默认应用打开文件")
        details_layout.addWidget(self.default_open_button)
        details_panel.setMinimumWidth(265)
        splitter.addWidget(details_panel)
        splitter.setSizes([200, 650, 330])
        splitter.setChildrenCollapsible(False)
        layout.addWidget(splitter, 1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setMaximumHeight(4)
        layout.addWidget(self.progress_bar)
        footer = QHBoxLayout()
        self.status_label = QLabel("准备就绪")
        self.status_label.setObjectName("muted")
        footer.addWidget(self.status_label, 1)
        self.errors_button = QPushButton("查看失败文件")
        self.errors_button.hide()
        footer.addWidget(self.errors_button)
        self.export_button = QPushButton("导出 CSV")
        self.export_button.setEnabled(False)
        footer.addWidget(self.export_button)
        layout.addLayout(footer)

    def _build_menu(self):
        file_menu = self.menuBar().addMenu("文件")
        browse_action = QAction("选择文件夹", self)
        browse_action.setShortcut(QKeySequence.StandardKey.Open)
        browse_action.triggered.connect(self.choose_directory)
        file_menu.addAction(browse_action)
        focus_action = QAction("搜索内容", self)
        focus_action.setShortcut(QKeySequence.StandardKey.Find)
        focus_action.triggered.connect(self.keyword_input.setFocus)
        file_menu.addAction(focus_action)
        refresh_action = QAction("刷新", self)
        refresh_action.setShortcut(QKeySequence("F5"))
        refresh_action.triggered.connect(self.refresh_search)
        file_menu.addAction(refresh_action)
        cancel_action = QAction("取消搜索", self)
        cancel_action.setShortcut(QKeySequence("Escape"))
        cancel_action.triggered.connect(self.cancel_search)
        file_menu.addAction(cancel_action)
        tools_menu = self.menuBar().addMenu("工具")
        self.clear_cache_action = QAction("清除文本缓存", self)
        self.clear_cache_action.triggered.connect(self.clear_cache)
        tools_menu.addAction(self.clear_cache_action)
        logs_action = QAction("打开数据与日志目录", self)
        logs_action.triggered.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.data_dir))))
        tools_menu.addAction(logs_action)
        help_menu = self.menuBar().addMenu("帮助")
        help_action = QAction("使用说明与搜索范围", self)
        help_action.triggered.connect(lambda: QMessageBox.information(self, "使用说明", HELP_TEXT))
        help_menu.addAction(help_action)

    def _connect_events(self):
        self.browse_button.clicked.connect(self.choose_directory)
        self.refresh_button.clicked.connect(self.refresh_search)
        self.directory_input.returnPressed.connect(self.refresh_search)
        self.keyword_input.returnPressed.connect(self.start_search)
        self.search_button.clicked.connect(lambda: self.cancel_search() if self.searching else self.start_search())
        self.file_list.currentItemChanged.connect(self.filter_file)
        self.table.selectionModel().currentRowChanged.connect(self.show_detail)
        self.table.doubleClicked.connect(lambda _index: self.open_current_slide())
        self.open_button.clicked.connect(self.open_current_slide)
        self.folder_button.clicked.connect(self.open_file_location)
        self.copy_button.clicked.connect(self.copy_text)
        self.default_open_button.clicked.connect(self.open_default)
        self.export_button.clicked.connect(self.export_results)
        self.errors_button.clicked.connect(self.show_errors)
        self.recursive_check.toggled.connect(lambda _checked: self.refresh_search())

    def choose_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "选择存放 PPTX 的文件夹", str(self.root))
        if directory:
            self.directory_input.setText(directory)
            self.refresh_search()

    def refresh_search(self):
        self.start_task(catalog=not bool(self.keyword_input.text().strip()))

    def start_search(self):
        if not self.keyword_input.text().strip():
            self.status_label.setText("请输入搜索关键词")
            self.keyword_input.setFocus()
            return
        self.start_task()

    def start_task(self, catalog=False, pending_slide=None):
        root = Path(self.directory_input.text().strip()).expanduser()
        if not root.is_dir():
            self.status_label.setText("目录不存在或无法访问，请重新选择")
            return
        for worker in self.search_workers.values():
            worker.cancel()
        self.job_id += 1
        self.root = root.resolve()
        self.current_options = SearchOptions("" if catalog else self.keyword_input.text().strip(),
                                             self.recursive_check.isChecked(),
                                             self.notes_check.isChecked(), self.case_check.isChecked())
        self._pending_open = pending_slide
        self.documents.clear()
        self.failures.clear()
        self.file_items.clear()
        self.file_list.clear()
        self.results_model.reset_hits()
        self.results_filter.set_path(None)
        self.delegate.options = self.current_options
        self.clear_detail()
        self.errors_button.hide()
        self.export_button.setEnabled(False)
        self.searching = True
        self.search_button.setText("取消搜索")
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("正在读取目录…" if catalog else "正在扫描与搜索…")
        self.result_count.setText("输入关键词开始搜索" if catalog else "正在查找匹配内容…")
        self.empty_label.setText("正在读取目录…")
        self.empty_label.show()
        self.clear_cache_action.setEnabled(False)
        exclusions = tuple(self.program_dir / name for name in (".venv", "build", "dist", "outputs", "work", "data", "_internal"))
        worker = SearchWorker(self.job_id, self.root, self.current_options,
                              self.data_dir / "text-cache.sqlite3", exclusions, self)
        self.search_workers[self.job_id] = worker
        worker.discovered.connect(self.on_discovered)
        worker.file_ready.connect(self.on_file_ready)
        worker.progress.connect(self.on_progress)
        worker.completed.connect(self.on_completed)
        worker.finished.connect(lambda job=self.job_id: self.search_finished(job))
        worker.start()
        self.persist_settings()

    def on_discovered(self, job_id, files):
        if job_id != self.job_id:
            return
        self.file_count.setText(f"找到 {len(files)} 个 PPTX")
        all_item = QListWidgetItem(f"全部文件（{len(files)}）")
        all_item.setData(Qt.ItemDataRole.UserRole, None)
        self.file_list.addItem(all_item)
        for path in files:
            relative = str(path.relative_to(self.root))
            item = QListWidgetItem(relative)
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            item.setData(Qt.ItemDataRole.UserRole + 2, str(path.parent.relative_to(self.root)))
            item.setToolTip(str(path))
            self.file_items[str(path)] = (item, relative)
            self.file_list.addItem(item)
        self.file_list.setCurrentRow(0)
        self.progress_bar.setRange(0, max(1, len(files)))
        self.progress_bar.setValue(0)

    def on_file_ready(self, job_id, path, document, hits, error):
        if job_id != self.job_id:
            return
        if document is not None:
            self.documents[path] = document
        self.results_model.append_hits(hits)
        entry = self.file_items.get(path)
        if entry:
            item, relative = entry
            count = sum(len(hit.spans) for hit in hits)
            item.setData(Qt.ItemDataRole.UserRole + 1, -1 if error else count)
            item.setText(f"{relative}   {'读取失败' if error else str(count) + ' 处'}")
            if error:
                item.setToolTip(f"{path}\n{error}")
        self.empty_label.setVisible(not bool(self.results_model.hits))
        if hits:
            self.update_counts()

    def on_progress(self, job_id, done, total, filename):
        if job_id == self.job_id:
            self.progress_bar.setValue(done)
            self.status_label.setText(f"正在搜索 {done} / {total}：{filename}")

    def on_completed(self, job_id, summary):
        if job_id != self.job_id:
            return
        self.searching = False
        self.search_button.setText("搜索")
        self.progress_bar.setRange(0, max(1, summary.total))
        self.progress_bar.setValue(summary.processed if summary.cancelled else max(1, summary.total))
        self.failures = summary.failures
        self.errors_button.setVisible(bool(self.failures))
        self.errors_button.setText(f"查看失败项（{len(self.failures)}）")
        self.update_counts()
        if not self.current_options.keyword:
            self.empty_label.setText("输入关键词，搜索这些课件中的内容。" if summary.total else "目录中没有 PPTX 文件，请选择其他文件夹。")
            self.result_count.setText("输入关键词开始搜索")
            self.status_label.setText(f"目录扫描完成：{summary.total} 个文件" + (f" · {len(self.failures)} 项失败" if self.failures else ""))
        else:
            prefix = "已取消" if summary.cancelled else "扫描完成"
            self.status_label.setText(f"{prefix}：{summary.processed} / {summary.total} 个文件 · "
                                      f"缓存 {summary.cached} 个 · {summary.elapsed:.2f} 秒"
                                      + (f" · {len(self.failures)} 项失败" if self.failures else ""))
            if not self.results_model.hits:
                message = "搜索已取消，当前没有结果。" if summary.cancelled else "没有找到匹配内容。试试其他关键词，或检查是否勾选搜索备注。"
                if not summary.total:
                    message = "目录中没有 PPTX 文件，请选择其他文件夹。"
                self.empty_label.setText(message)
        self.empty_label.setVisible(not bool(self.results_model.hits))
        if self.results_filter.rowCount():
            self.table.selectRow(0)
        self._resume_pending_open(summary)

    def search_finished(self, job_id):
        worker = self.search_workers.pop(job_id, None)
        if worker:
            worker.deleteLater()
        self.clear_cache_action.setEnabled(not bool(self.search_workers))

    def cancel_search(self):
        worker = self.search_workers.get(self.job_id)
        if worker and worker.isRunning():
            worker.cancel()
            self.status_label.setText("正在取消；将在当前读取步骤结束后停止…")
            self._pending_open = None

    def update_counts(self):
        hits = self.visible_hits()
        files = len({hit.path for hit in hits})
        pages = len({(hit.path, hit.slide_id) for hit in hits})
        occurrences = sum(len(hit.spans) for hit in hits)
        self.result_count.setText(f"{files} 个文件  ·  {pages} 页  ·  {occurrences} 处命中")
        self.export_button.setEnabled(bool(hits))

    def visible_hits(self):
        return [self.results_filter.index(row, 0).data(Qt.ItemDataRole.UserRole)
                for row in range(self.results_filter.rowCount())]

    def filter_file(self, current, _previous=None):
        path = current.data(Qt.ItemDataRole.UserRole) if current else None
        self.results_filter.set_path(path)
        self.clear_detail()
        self.update_counts()
        if self.results_filter.rowCount():
            self.table.selectRow(0)

    def show_detail(self, current, _previous=None):
        hit = current.data(Qt.ItemDataRole.UserRole) if current.isValid() else None
        if hit is None:
            self.clear_detail()
            return
        self.current_hit = hit
        self.detail_name.setText(hit.filename)
        self.detail_meta.setText(f"第 {hit.slide_number} / {hit.slide_total} 页  ·  来源：{hit.block.source}"
                                 + ("  ·  隐藏页" if hit.hidden else ""))
        self.detail_path.setPlainText(hit.path)
        self.detail_path.setToolTip(hit.path)
        document = self.documents[hit.path]
        slide = next(slide for slide in document.slides if slide.slide_id == hit.slide_id)
        self.details_browser.setHtml(detail_html(hit, slide, self.current_options))
        self.details_browser.scrollToAnchor("selected")
        self._set_detail_enabled(True)

    def clear_detail(self):
        self.current_hit = None
        self.detail_name.setText("尚未选择结果")
        self.detail_meta.setText("单击结果查看命中文本及上下文")
        self.detail_path.clear()
        self.details_browser.clear()
        self._set_detail_enabled(False)

    def _set_detail_enabled(self, enabled):
        self.open_button.setEnabled(enabled and not bool(self.open_workers))
        for button in (self.folder_button, self.copy_button, self.default_open_button):
            button.setEnabled(enabled)

    def open_current_slide(self):
        hit = self.current_hit
        if hit is None or self.open_workers or self.searching:
            return
        path = Path(hit.path)
        try:
            stat = path.stat()
        except OSError:
            self.status_label.setText("源文件不存在或无法访问，请刷新搜索结果")
            return
        document = self.documents[hit.path]
        if (document.size, document.mtime_ns) != (stat.st_size, stat.st_mtime_ns):
            self.status_label.setText("源文件已变化，正在刷新后定位…")
            self.start_task(pending_slide=(hit.path, hit.slide_id))
            return
        self.launch_slide(path, hit.slide_id)

    def _resume_pending_open(self, summary):
        pending, self._pending_open = self._pending_open, None
        if pending is None or summary.cancelled:
            return
        path, slide_id = pending
        document = self.documents.get(path)
        if document is None or not any(slide.slide_id == slide_id for slide in document.slides):
            self.status_label.setText("目标幻灯片已删除或文件无法读取，请查看刷新后的结果")
            return
        self.launch_slide(Path(path), slide_id)

    def launch_slide(self, path, slide_id):
        self.status_label.setText("正在 PowerPoint 中打开目标幻灯片…")
        worker = OpenSlideWorker(self.job_id, path, slide_id, self)
        self.open_workers[id(worker)] = worker
        self.open_button.setEnabled(False)
        worker.opened.connect(self.slide_opened)
        worker.finished.connect(lambda token=id(worker): self.open_finished(token))
        worker.start()

    def slide_opened(self, job_id, number, error):
        if job_id != self.job_id:
            return
        if error:
            LOGGER.warning("PowerPoint 跳转失败：%s", error)
            self.status_label.setText("无法跳转，可使用默认应用打开文件或打开文件位置")
            QMessageBox.warning(self, "无法打开原页", f"{error}\n\n搜索结果仍然可用。可使用右侧的默认应用打开文件或打开文件位置。")
        else:
            self.status_label.setText(f"已在 PowerPoint 中定位到第 {number} 页")

    def open_finished(self, token):
        worker = self.open_workers.pop(token, None)
        if worker:
            worker.deleteLater()
        self.open_button.setEnabled(self.current_hit is not None and not bool(self.open_workers))

    def open_file_location(self):
        if self.current_hit:
            import subprocess
            path = Path(self.current_hit.path)
            if path.is_file():
                subprocess.Popen(["explorer.exe", "/select,", str(path)], shell=False)
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.parent)))

    def open_default(self):
        if self.current_hit:
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(self.current_hit.path)):
                self.status_label.setText("默认应用无法打开文件，请检查文件关联")

    def copy_text(self):
        if self.current_hit:
            hit = self.current_hit
            QApplication.clipboard().setText(f"{hit.filename} · 第 {hit.slide_number} 页 · {hit.block.source}\n{hit.block.text}")
            self.status_label.setText("已复制命中文本及来源")

    def export_results(self):
        filename, _ = QFileDialog.getSaveFileName(self, "导出当前筛选结果", str(self.root / "PPT搜索结果.csv"), "CSV 文件 (*.csv)")
        if not filename:
            return
        try:
            export_csv(Path(filename), self.visible_hits())
            self.status_label.setText(f"已导出：{filename}")
        except OSError as error:
            QMessageBox.warning(self, "无法导出", str(error))

    def clear_cache(self):
        if self.search_workers:
            return
        try:
            with TextCache(self.data_dir / "text-cache.sqlite3") as cache:
                cache.clear()
            self.status_label.setText("文本缓存已清除，下次搜索将重新读取源文件")
        except Exception as error:
            QMessageBox.warning(self, "无法清除缓存", str(error))

    def show_errors(self):
        box = QMessageBox(self)
        box.setWindowTitle("无法读取的文件或目录")
        box.setIcon(QMessageBox.Icon.Warning)
        box.setText(f"有 {len(self.failures)} 项无法读取，其余文件已继续处理。")
        box.setDetailedText("\n\n".join(f"{path}\n{message}" for path, message in self.failures))
        box.exec()

    def persist_settings(self):
        values = {"directory": str(self.root), "recursive": self.recursive_check.isChecked(),
                  "notes": self.notes_check.isChecked(), "case_sensitive": self.case_check.isChecked()}
        try:
            save_settings(self.settings_path, values)
        except OSError:
            LOGGER.exception("保存设置失败")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(Path(url.toLocalFile()).is_dir() for url in event.mimeData().urls() if url.isLocalFile()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            if url.isLocalFile() and Path(url.toLocalFile()).is_dir():
                self.directory_input.setText(url.toLocalFile())
                self.refresh_search()
                event.acceptProposedAction()
                break

    def closeEvent(self, event):
        self.persist_settings()
        active = [worker for worker in [*self.search_workers.values(), *self.open_workers.values()] if worker.isRunning()]
        if active:
            for worker in self.search_workers.values():
                worker.cancel()
            self.closing = True
            self.hide()
            # 不强制 terminate 正在读 ZIP/COM 的线程，避免 Qt 崩溃或损坏缓存。
            QTimer.singleShot(150, self.close)
            event.ignore()
            return
        event.accept()
