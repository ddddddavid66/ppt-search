"""与 GUI 无关的目录发现及字面检索逻辑。"""
import os
import re
from pathlib import Path
from threading import Event

from .extract import check_cancel
from .models import Document, SearchHit, SearchOptions


def discover_files(root: Path, recursive: bool, cancel: Event | None = None,
                   errors: list | None = None, excluded: tuple[Path, ...] = ()) -> list[Path]:
    root = root.resolve()
    if not root.is_dir():
        raise NotADirectoryError(f"目录不存在：{root}")
    exclusions = {str(path.resolve()).casefold() for path in excluded}
    files = []

    def on_error(error):
        if errors is not None:
            errors.append((str(error.filename), f"无法读取目录：{error.strerror}"))

    for directory, folders, names in os.walk(root, followlinks=False, onerror=on_error):
        check_cancel(cancel)
        folders[:] = [name for name in folders
                      if str((Path(directory) / name).resolve()).casefold() not in exclusions
                      and not (Path(directory) / name).is_symlink()]
        for name in names:
            if name.lower().endswith(".pptx") and not name.startswith("~$"):
                files.append(Path(directory) / name)
        if not recursive:
            break
    return sorted(files, key=lambda path: str(path.relative_to(root)).casefold())


def make_pattern(options: SearchOptions) -> re.Pattern:
    if not options.keyword.strip():
        raise ValueError("请输入搜索关键词")
    # 对用户输入全部转义，正则仅用于获得准确的原文高亮范围。
    return re.compile(re.escape(options.keyword), 0 if options.case_sensitive else re.IGNORECASE)


def search_document(document: Document, options: SearchOptions,
                    cancel: Event | None = None) -> list[SearchHit]:
    pattern = make_pattern(options)
    hits = []
    for slide in document.slides:
        check_cancel(cancel)
        for block in slide.blocks:
            if block.source == "备注" and not options.include_notes:
                continue
            spans = tuple((match.start(), match.end()) for match in pattern.finditer(block.text))
            if spans:
                hits.append(SearchHit(document.path, slide.number, slide.slide_id,
                                      len(document.slides), slide.title, slide.hidden, block, spans))
    return hits
