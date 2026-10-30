"""PowerPoint 自动化只用于用户主动打开结果，检索不依赖 Office。"""
import os
from pathlib import Path

from PySide6.QtCore import QThread, Signal


class SlideUnavailableError(RuntimeError):
    pass


def same_file(left: str, right: str) -> bool:
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


def navigate_application(application, path: Path, slide_id: int) -> int:
    """独立于 COM 初始化，便于使用假对象验证文稿复用和稳定 ID 定位。"""
    application.Visible = -1
    presentation = None
    for index in range(1, application.Presentations.Count + 1):
        candidate = application.Presentations.Item(index)
        if same_file(candidate.FullName, str(path)):
            presentation = candidate
            break
    if presentation is None:
        presentation = application.Presentations.Open(str(path), -1, 0, -1)
    try:
        slide = presentation.Slides.FindBySlideID(slide_id)
    except Exception as error:
        raise SlideUnavailableError("目标幻灯片已删除或被替换，请刷新搜索结果") from error
    if presentation.Windows.Count == 0:
        presentation.NewWindow()
    window = presentation.Windows.Item(1)
    window.Activate()
    window.ViewType = 9  # ppViewNormal：在编辑窗口定位，不进入放映。
    window.View.GotoSlide(slide.SlideIndex)
    return int(slide.SlideIndex)


def open_slide(path: Path, slide_id: int) -> int:
    if not path.is_file():
        raise FileNotFoundError("源文件不存在，请刷新搜索结果")
    import pythoncom
    import win32com.client

    # COM 对象只能在初始化它的 STA 线程使用，不把对象传回 GUI 线程。
    pythoncom.CoInitialize()
    try:
        application = win32com.client.Dispatch("PowerPoint.Application")
        return navigate_application(application, path.resolve(), slide_id)
    finally:
        pythoncom.CoUninitialize()


class OpenSlideWorker(QThread):
    opened = Signal(int, int, str)

    def __init__(self, job_id: int, path: Path, slide_id: int, parent=None):
        super().__init__(parent)
        self.job_id = job_id
        self.path = path
        self.slide_id = slide_id

    def run(self):
        try:
            number = open_slide(self.path, self.slide_id)
            self.opened.emit(self.job_id, number, "")
        except Exception as error:
            self.opened.emit(self.job_id, 0, str(error))
