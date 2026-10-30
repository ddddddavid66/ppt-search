"""执行固定样本冷/热搜索、原生 GUI 截图和可选 PowerPoint 真机验收。"""
import argparse
import hashlib
import json
import sys
from pathlib import Path
from time import perf_counter

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from ppt_search.cache import TextCache
from ppt_search.extract import extract_document
from ppt_search.models import SearchOptions
from ppt_search.powerpoint import navigate_application
from ppt_search.search import discover_files, search_document
from ppt_search.theme import STYLESHEET
from ppt_search.window import MainWindow


def benchmark(root: Path, cache_path: Path):
    manifest = json.loads((root / "expected-linux.json").read_text(encoding="utf-8"))
    output = {}
    with TextCache(cache_path) as cache:
        cache.clear()
        for mode in ("cold", "warm"):
            started = perf_counter()
            hits = []
            cache_count = 0
            for path in discover_files(root, True):
                document = cache.get(path)
                if document is None:
                    document = extract_document(path)
                    cache.put(document)
                else:
                    cache_count += 1
                hits.extend(search_document(document, SearchOptions("linux")))
            output[f"{mode}_seconds"] = round(perf_counter() - started, 4)
            output[f"{mode}_cached_files"] = cache_count
            observed = sorted([str(Path(hit.path).relative_to(root)), hit.slide_number,
                               hit.block.source, hit.block.text] for hit in hits)
            assert observed == sorted(manifest["expected"]), "固定样本搜索结果与独立清单不一致"
        output["files"] = manifest["files"]
        output["slides"] = manifest["slides"]
        output["hit_rows"] = len(hits)
        output["occurrences"] = sum(len(hit.spans) for hit in hits)
    return output


def office_smoke(path: Path):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    application = None
    owned_document = None
    before_count = None
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    try:
        application = win32com.client.Dispatch("PowerPoint.Application")
        before_count = application.Presentations.Count
        target = extract_document(path).slides[6]
        number = navigate_application(application, path, target.slide_id)
        for index in range(1, application.Presentations.Count + 1):
            candidate = application.Presentations.Item(index)
            if Path(candidate.FullName).resolve() == path.resolve():
                owned_document = candidate
                break
        assert owned_document is not None
        assert number == 7
        assert owned_document.ReadOnly != 0
        assert owned_document.Windows.Item(1).View.Slide.SlideID == target.slide_id
        count = application.Presentations.Count
        navigate_application(application, path, target.slide_id)
        assert application.Presentations.Count == count, "重复跳转不应重新打开文件"
        assert hashlib.sha256(path.read_bytes()).hexdigest() == digest
        return {"target_page": number, "readonly": True, "reused_document": True, "source_unchanged": True}
    finally:
        # 验收只关闭本次生成并打开的测试课件，不关闭任何用户文稿。
        if owned_document is not None:
            owned_document.Close()
        if application is not None and before_count == 0 and application.Presentations.Count == 0:
            application.Quit()
        pythoncom.CoUninitialize()


def gui_smoke(app, program_dir: Path, root: Path, data_dir: Path, screenshot: Path):
    window = MainWindow(program_dir, data_dir, root)
    window.show()
    output = {}
    failed = []
    phase = "catalog"
    started = 0.0
    timer = QTimer()

    def poll():
        nonlocal phase, started
        if window.search_workers:
            return
        if phase == "catalog":
            phase = "search"
            window.keyword_input.setText("linux")
            started = perf_counter()
            window.start_search()
            return
        if phase == "search":
            output["gui_first_search_seconds"] = round(perf_counter() - started, 4)
            phase = "warm"
            started = perf_counter()
            window.start_search()
            return
        try:
            output["gui_warm_search_seconds"] = round(perf_counter() - started, 4)
            assert window.results_model.rowCount() == 280
            assert window.table.model().rowCount() == 280
            assert "20 个文件" in window.result_count.text()
            window.table.selectRow(0)
            assert window.current_hit is not None
            assert "Linux" in window.details_browser.toPlainText()
            window.copy_text()
            assert "第 7 页" in app.clipboard().text()
            screenshot.parent.mkdir(parents=True, exist_ok=True)
            window.grab().save(str(screenshot))
            output["screenshot"] = str(screenshot)
        except Exception as error:
            failed.append(error)
        finally:
            timer.stop()
            window.close()
            app.quit()

    timer.timeout.connect(poll)
    timer.start(20)
    QTimer.singleShot(60_000, app.quit)
    app.exec()
    if failed:
        raise failed[0]
    assert output, "GUI 验收超时"
    return output


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path("work/demo"))
    parser.add_argument("--office", action="store_true")
    arguments = parser.parse_args()
    root = arguments.root.resolve()
    program_dir = Path(__file__).resolve().parent.parent
    work = program_dir / "work"
    data = work / "acceptance-data"
    data.mkdir(parents=True, exist_ok=True)
    output = benchmark(root, data / "text-cache.sqlite3")
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)
    output.update(gui_smoke(app, program_dir, root, data, work / "gui-preview.png"))
    if arguments.office:
        # 独立文件名避免复用用户之前打开的演示课件。
        smoke_path = work / "office-smoke-generated.pptx"
        smoke_path.write_bytes(discover_files(root, True)[0].read_bytes())
        output["powerpoint"] = office_smoke(smoke_path)
    (work / "acceptance.json").write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
