"""高亮、结果模型和导出：不涉及文件扫描或 Office 自动化。"""
import csv
import html
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt, QSize
from PySide6.QtGui import QAbstractTextDocumentLayout, QPalette, QTextDocument, QColor
from PySide6.QtWidgets import QStyle, QStyledItemDelegate

from .models import SearchHit, SearchOptions
from .search import make_pattern


def highlight(text: str, options: SearchOptions) -> str:
    parts = []
    previous = 0
    for match in make_pattern(options).finditer(text):
        parts.append(html.escape(text[previous:match.start()]))
        parts.append('<span style="background-color:#fef08a;color:#713f12;">'
                     + html.escape(match.group()) + '</span>')
        previous = match.end()
    parts.append(html.escape(text[previous:]))
    return "".join(parts).replace("\n", "<br>").replace("\v", "<br>")


class ResultsModel(QAbstractTableModel):
    HEADERS = ("文件名", "页码", "来源", "命中文本")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.hits: list[SearchHit] = []

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.hits)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self.HEADERS[section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        hit = self.hits[index.row()]
        if role == Qt.ItemDataRole.UserRole:
            return hit
        if role == Qt.ItemDataRole.ToolTipRole:
            return f"{hit.path}\n第 {hit.slide_number} 页 · {hit.block.location}\n{hit.block.text}"
        if role == Qt.ItemDataRole.TextAlignmentRole and index.column() in (1, 2):
            return Qt.AlignmentFlag.AlignCenter
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        return (hit.filename, f"{hit.slide_number}" + (" · 隐藏" if hit.hidden else ""),
                hit.block.source, hit.snippet)[index.column()]

    def reset_hits(self):
        self.beginResetModel()
        self.hits.clear()
        self.endResetModel()

    def append_hits(self, hits: list[SearchHit]):
        if not hits:
            return
        self.beginInsertRows(QModelIndex(), len(self.hits), len(self.hits) + len(hits) - 1)
        self.hits.extend(hits)
        self.endInsertRows()


class ResultsFilter(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.path: str | None = None

    def set_path(self, path: str | None):
        self.path = path
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        return self.path is None or self.sourceModel().hits[row].path == self.path

    def lessThan(self, left, right):
        first = self.sourceModel().hits[left.row()]
        second = self.sourceModel().hits[right.row()]
        if left.column() == 1:
            return (first.slide_number, first.path.casefold()) < (second.slide_number, second.path.casefold())
        if left.column() == 0:
            return (first.path.casefold(), first.slide_number) < (second.path.casefold(), second.slide_number)
        return super().lessThan(left, right)


class HighlightDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.options = SearchOptions("linux")

    def paint(self, painter, option, index):
        if index.column() != 3:
            super().paint(painter, option, index)
            return
        self.initStyleOption(option, index)
        text = option.text
        option.text = ""
        style = option.widget.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, option, painter, option.widget)
        document = QTextDocument()
        document.setDefaultFont(option.font)
        document.setDocumentMargin(0)
        document.setHtml(highlight(text, self.options))
        document.setTextWidth(option.rect.width() - 20)
        context = QAbstractTextDocumentLayout.PaintContext()
        context.palette = option.palette
        if option.state & QStyle.StateFlag.State_Selected:
            context.palette.setColor(QPalette.ColorRole.Text, option.palette.color(QPalette.ColorRole.HighlightedText))
        painter.save()
        painter.setClipRect(option.rect)
        painter.translate(option.rect.left() + 10, option.rect.top() + 12)
        document.documentLayout().draw(painter, context)
        painter.restore()


class FileDelegate(QStyledItemDelegate):
    """文件名自动省略，命中数始终可见，目录单独占一行。"""
    def sizeHint(self, option, index):
        return QSize(180, 58)

    def paint(self, painter, option, index):
        path = index.data(Qt.ItemDataRole.UserRole)
        if path is None:
            super().paint(painter, option, index)
            return
        self.initStyleOption(option, index)
        option.text = ""
        option.widget.style().drawControl(QStyle.ControlElement.CE_ItemViewItem, option, painter, option.widget)
        rect = option.rect.adjusted(10, 6, -10, -4)
        count = index.data(Qt.ItemDataRole.UserRole + 1)
        caption = "—" if count is None else ("失败" if count == -1 else f"{count} 处")
        painter.save()
        font = option.font
        font.setPointSizeF(max(9, font.pointSizeF()))
        painter.setFont(font)
        metrics = painter.fontMetrics()
        count_width = metrics.horizontalAdvance(caption) + 10
        painter.setPen(QColor("#2563eb" if count != -1 else "#b45309"))
        painter.drawText(rect.adjusted(rect.width() - count_width, 0, 0, -rect.height() // 2),
                         Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, caption)
        painter.setPen(QColor("#334155"))
        name = metrics.elidedText(Path(path).name, Qt.TextElideMode.ElideMiddle, rect.width() - count_width - 8)
        painter.drawText(rect.adjusted(0, 0, -count_width - 8, -rect.height() // 2),
                         Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, name)
        font.setPointSizeF(max(8, font.pointSizeF() - 1))
        painter.setFont(font)
        painter.setPen(QColor("#94a3b8"))
        parent = index.data(Qt.ItemDataRole.UserRole + 2) or "当前目录"
        parent = painter.fontMetrics().elidedText(parent, Qt.TextElideMode.ElideMiddle, rect.width())
        painter.drawText(rect.adjusted(0, rect.height() // 2, 0, 0),
                         Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, parent)
        painter.restore()


def detail_html(hit: SearchHit, slide, options: SearchOptions) -> str:
    sections = []
    for source_name, heading in (("body", "幻灯片正文"), ("notes", "讲者备注")):
        if source_name == "notes" and not options.include_notes:
            continue
        blocks = [block for block in slide.blocks if (block.source == "备注") == (source_name == "notes")]
        if not blocks:
            continue
        sections.append(f'<h3 style="color:#64748b;font-size:12px;">{heading}</h3>')
        for block in blocks:
            location = html.escape(block.location)
            selected = block == hit.block
            decoration = "border-left:3px solid #2563eb;padding-left:10px;" if selected else ""
            anchor = '<a name="selected"></a>' if selected else ""
            sections.append(f'{anchor}<div style="{decoration}"><p style="color:#94a3b8;font-size:11px;">'
                            f'{location}</p><p>{highlight(block.text, options)}</p></div>')
    title = html.escape(hit.title)
    return f'<html><body style="color:#334155;font-size:13px;line-height:1.6;"><h2>{title}</h2>' + "".join(sections) + '</body></html>'


def export_csv(path: Path, hits: list[SearchHit]):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["文件名", "完整路径", "实际页码", "总页数", "标题", "来源", "位置", "命中文本", "命中次数", "隐藏页"])
        for hit in hits:
            values = [hit.filename, hit.path, hit.slide_number, hit.slide_total, hit.title,
                      hit.block.source, hit.block.location, hit.block.text, len(hit.spans), "是" if hit.hidden else "否"]
            # 导出的 PPT 文本按文本打开，避免 Excel 把以等号开头的原文当公式执行。
            writer.writerow(["'" + value if isinstance(value, str) and value.startswith(("=", "+", "-", "@"))
                             else value for value in values])
