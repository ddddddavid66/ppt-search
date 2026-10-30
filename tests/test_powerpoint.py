from pathlib import Path
from types import SimpleNamespace

import pytest

from ppt_search.powerpoint import SlideUnavailableError, navigate_application


class Presentations:
    def __init__(self, document, already_open=True):
        self.document = document
        self.Count = int(already_open)
        self.open_arguments = None

    def Item(self, _index):
        return self.document

    def Open(self, *arguments):
        self.open_arguments = arguments
        return self.document


def fake_application(path, already_open=True, deleted=False):
    calls = []

    def find(slide_id):
        calls.append(("find", slide_id))
        if deleted:
            raise RuntimeError("missing")
        return SimpleNamespace(SlideIndex=7)

    window = SimpleNamespace(Activate=lambda: calls.append("activate"),
                             View=SimpleNamespace(GotoSlide=lambda page: calls.append(("goto", page))))
    document = SimpleNamespace(FullName=str(path), Slides=SimpleNamespace(FindBySlideID=find),
                               Windows=SimpleNamespace(Count=1, Item=lambda _: window))
    application = SimpleNamespace(Presentations=Presentations(document, already_open))
    return application, calls


def test_reuse_open_document_and_find_stable_id(tmp_path):
    path = tmp_path / "测试.pptx"
    application, calls = fake_application(path)
    assert navigate_application(application, path, 256) == 7
    assert application.Presentations.open_arguments is None
    assert ("find", 256) in calls and ("goto", 7) in calls


def test_new_document_readonly_and_deleted_slide(tmp_path):
    path = tmp_path / "测试.pptx"
    application, _ = fake_application(path, already_open=False)
    navigate_application(application, path, 256)
    assert application.Presentations.open_arguments == (str(path), -1, 0, -1)
    application, _ = fake_application(path, deleted=True)
    with pytest.raises(SlideUnavailableError):
        navigate_application(application, path, 256)
