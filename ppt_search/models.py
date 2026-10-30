"""提取、缓存、搜索和界面共享的轻量数据结构。"""
from dataclasses import asdict, dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class TextBlock:
    text: str
    source: str
    location: str


@dataclass
class SlideText:
    number: int
    slide_id: int
    title: str
    hidden: bool
    blocks: list[TextBlock] = field(default_factory=list)


@dataclass
class Document:
    path: str
    size: int
    mtime_ns: int
    slides: list[SlideText]

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict) -> "Document":
        slides = [SlideText(**{**slide, "blocks": [TextBlock(**block) for block in slide["blocks"]]})
                  for slide in value["slides"]]
        return cls(value["path"], value["size"], value["mtime_ns"], slides)


@dataclass(frozen=True)
class SearchOptions:
    keyword: str
    recursive: bool = True
    include_notes: bool = True
    case_sensitive: bool = False


@dataclass(frozen=True)
class SearchHit:
    path: str
    slide_number: int
    slide_id: int
    slide_total: int
    title: str
    hidden: bool
    block: TextBlock
    spans: tuple[tuple[int, int], ...]

    @property
    def filename(self) -> str:
        return Path(self.path).name

    @property
    def snippet(self) -> str:
        start = max(0, self.spans[0][0] - 32)
        end = min(len(self.block.text), self.spans[0][1] + 85)
        text = self.block.text[start:end].replace("\n", " ").replace("\v", " ")
        return ("…" if start else "") + text + ("…" if end < len(self.block.text) else "")


@dataclass
class SearchSummary:
    total: int = 0
    processed: int = 0
    cached: int = 0
    failures: list[tuple[str, str]] = field(default_factory=list)
    cancelled: bool = False
    elapsed: float = 0.0
