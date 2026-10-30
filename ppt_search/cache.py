"""每个任务独占 SQLite 连接，缓存完整文本而非特定关键词的结果。"""
import json
import logging
import sqlite3
from pathlib import Path

from .models import Document

CACHE_VERSION = 2
LOGGER = logging.getLogger(__name__)


class TextCache:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(path, timeout=10)
        self.connection.execute("""CREATE TABLE IF NOT EXISTS documents (
            path TEXT PRIMARY KEY, size INTEGER NOT NULL, mtime_ns INTEGER NOT NULL,
            version INTEGER NOT NULL, content TEXT NOT NULL)""")
        self.connection.commit()

    @staticmethod
    def key(path: Path) -> str:
        return str(path.resolve()).casefold()

    def get(self, path: Path) -> Document | None:
        stat = path.stat()
        row = self.connection.execute(
            "SELECT content FROM documents WHERE path=? AND size=? AND mtime_ns=? AND version=?",
            (self.key(path), stat.st_size, stat.st_mtime_ns, CACHE_VERSION)).fetchone()
        if row is None:
            return None
        try:
            return Document.from_dict(json.loads(row[0]))
        except (ValueError, TypeError, KeyError):
            LOGGER.warning("无效的文本缓存，重新提取：%s", path)
            self.remove(path)
            return None

    def put(self, document: Document) -> None:
        self.connection.execute("INSERT OR REPLACE INTO documents VALUES (?, ?, ?, ?, ?)",
                                (self.key(Path(document.path)), document.size, document.mtime_ns,
                                 CACHE_VERSION, json.dumps(document.to_dict(), ensure_ascii=False)))
        self.connection.commit()

    def remove(self, path: Path) -> None:
        self.connection.execute("DELETE FROM documents WHERE path=?", (self.key(path),))
        self.connection.commit()

    def prune_missing(self) -> None:
        paths = [row[0] for row in self.connection.execute("SELECT path FROM documents")]
        missing = [(path,) for path in paths if not Path(path).is_file()]
        self.connection.executemany("DELETE FROM documents WHERE path=?", missing)
        self.connection.commit()

    def clear(self) -> None:
        self.connection.execute("DELETE FROM documents")
        self.connection.commit()
        self.connection.execute("VACUUM")

    def close(self) -> None:
        self.connection.close()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()
