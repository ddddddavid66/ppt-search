"""任务通过信号跨线程传递结果，不从后台操作控件。"""
import logging
from pathlib import Path
from threading import Event
from time import perf_counter

from PySide6.QtCore import QThread, Signal

from .cache import TextCache
from .extract import Cancelled, check_cancel, extract_document
from .models import SearchOptions, SearchSummary
from .search import discover_files, search_document

LOGGER = logging.getLogger(__name__)


class SearchWorker(QThread):
    discovered = Signal(int, object)
    file_ready = Signal(int, str, object, object, str)
    progress = Signal(int, int, int, str)
    completed = Signal(int, object)

    def __init__(self, job_id: int, root: Path, options: SearchOptions,
                 cache_path: Path, excluded: tuple[Path, ...] = (), parent=None):
        super().__init__(parent)
        self.job_id = job_id
        self.root = root
        self.options = options
        self.cache_path = cache_path
        self.excluded = excluded
        self.cancel_event = Event()

    def cancel(self):
        self.cancel_event.set()

    def run(self):
        summary = SearchSummary()
        started = perf_counter()
        try:
            files = discover_files(self.root, self.options.recursive, self.cancel_event,
                                   summary.failures, self.excluded)
            summary.total = len(files)
            self.discovered.emit(self.job_id, files)
            if not self.options.keyword.strip():
                summary.elapsed = perf_counter() - started
                self.completed.emit(self.job_id, summary)
                return
            with TextCache(self.cache_path) as cache:
                cache.prune_missing()
                for path in files:
                    check_cancel(self.cancel_event)
                    self._process_file(path, cache, summary)
                    summary.processed += 1
                    self.progress.emit(self.job_id, summary.processed, summary.total, path.name)
        except Cancelled:
            summary.cancelled = True
        except Exception as error:
            LOGGER.exception("搜索任务失败")
            summary.failures.append((str(self.root), str(error)))
        summary.cancelled = summary.cancelled or self.cancel_event.is_set()
        summary.elapsed = perf_counter() - started
        self.completed.emit(self.job_id, summary)

    def _process_file(self, path: Path, cache: TextCache, summary: SearchSummary):
        try:
            document = cache.get(path)
            if document is None:
                document = extract_document(path, self.cancel_event)
                check_cancel(self.cancel_event)
                cache.put(document)
            else:
                summary.cached += 1
            hits = search_document(document, self.options, self.cancel_event)
            self.file_ready.emit(self.job_id, str(path), document, hits, "")
        except Cancelled:
            raise
        except Exception as error:
            LOGGER.warning("无法读取 %s：%s", path, error)
            cache.remove(path)
            message = str(error) or type(error).__name__
            summary.failures.append((str(path), message))
            self.file_ready.emit(self.job_id, str(path), None, [], message)
