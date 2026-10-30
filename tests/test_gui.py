from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from ppt_search.models import SearchOptions
from ppt_search.theme import STYLESHEET
from ppt_search.window import MainWindow
from ppt_search.workers import SearchWorker


def wait_idle(qtbot, window):
    qtbot.waitUntil(lambda: not window.search_workers, timeout=20_000)


def test_gui_search_details_filter_clipboard_and_stale_task(qtbot, deck, tmp_path):
    root = deck[0].parent
    copy_root = root / "子目录"
    copy_root.mkdir()
    (copy_root / deck[0].name).write_bytes(deck[0].read_bytes())
    data = tmp_path / "data"
    data.mkdir()
    window = MainWindow(tmp_path, data, root)
    window.setStyleSheet(STYLESHEET)
    qtbot.addWidget(window)
    window.show()
    qtbot.waitUntil(lambda: window.file_list.count() == 3, timeout=20_000)
    wait_idle(qtbot, window)
    assert window.file_list.count() == 3
    window.keyword_input.setText("linux")
    qtbot.keyClick(window.keyword_input, Qt.Key.Key_Return)
    wait_idle(qtbot, window)
    assert window.results_model.rowCount() == 14
    assert "2 个文件" in window.result_count.text()
    assert "第 2 / 3 页" in window.detail_meta.text()
    assert "Linux" in window.details_browser.toPlainText()
    window.copy_text()
    assert "第 2 页" in QApplication.clipboard().text()
    window.file_list.setCurrentRow(1)
    assert window.results_filter.rowCount() == 7
    previous_count = len(window.results_model.hits)
    window.on_file_ready(window.job_id - 1, "ignored", None, [None], "")
    assert len(window.results_model.hits) == previous_count
    window.keyword_input.setText("不存在的关键词")
    window.start_search()
    wait_idle(qtbot, window)
    assert window.results_model.rowCount() == 0
    assert "没有找到" in window.empty_label.text()
    assert not window.open_button.isEnabled()


def test_worker_partial_failure_cache_and_cancel(qtbot, deck, tmp_path):
    (deck[0].parent / "损坏.pptx").write_bytes(b"broken zip")
    worker = SearchWorker(42, deck[0].parent, SearchOptions("linux"), tmp_path / "cache.sqlite3")
    with qtbot.waitSignal(worker.completed, timeout=20_000) as result:
        worker.start()
    worker.wait(5_000)
    summary = result.args[1]
    assert summary.total == 2 and summary.processed == 2
    assert len(summary.failures) == 1
    cached = SearchWorker(43, deck[0].parent, SearchOptions("linux"), tmp_path / "cache.sqlite3")
    with qtbot.waitSignal(cached.completed, timeout=20_000) as result:
        cached.start()
    cached.wait(5_000)
    assert result.args[1].cached == 1
    cancelled = SearchWorker(44, deck[0].parent, SearchOptions("linux"), tmp_path / "cache.sqlite3")
    cancelled.cancel()
    with qtbot.waitSignal(cancelled.completed, timeout=20_000) as result:
        cancelled.start()
    cancelled.wait(5_000)
    assert result.args[1].cancelled


def test_program_runtime_templates_are_excluded(qtbot, deck, tmp_path):
    program = tmp_path / "portable"
    runtime = program / "_internal" / "pptx" / "templates"
    runtime.mkdir(parents=True)
    (runtime / "default.pptx").write_bytes(deck[0].read_bytes())
    real_deck = program / "用户课件.pptx"
    real_deck.write_bytes(deck[0].read_bytes())
    data = program / "data"
    data.mkdir()
    window = MainWindow(program, data, program)
    qtbot.addWidget(window)
    qtbot.waitUntil(lambda: window.file_list.count() > 0, timeout=20_000)
    wait_idle(qtbot, window)
    assert set(window.file_items) == {str(real_deck)}


def test_replacing_running_search_keeps_latest_query(qtbot, deck, tmp_path, monkeypatch):
    import time
    import ppt_search.workers as workers
    original = workers.extract_document

    def slower(path, cancel):
        time.sleep(.04)
        return original(path, cancel)

    monkeypatch.setattr(workers, "extract_document", slower)
    data = tmp_path / "data"
    data.mkdir()
    window = MainWindow(tmp_path, data, deck[0].parent)
    qtbot.addWidget(window)
    qtbot.waitUntil(lambda: window.file_list.count() > 0, timeout=20_000)
    wait_idle(qtbot, window)
    window.keyword_input.setText("linux")
    window.start_search()
    window.keyword_input.setText("普通文字")
    window.start_search()
    wait_idle(qtbot, window)
    assert window.current_options.keyword == "普通文字"
    assert window.results_model.rowCount() == 1
    assert "普通文字" in window.results_model.hits[0].block.text
