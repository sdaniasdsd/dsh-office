"""`docling.document_converter` 的桩实现。

行为由环境变量 `DOCX_PARSE_STUB_DOCLING` 描述的 JSON 决定：

    {
      "items": [
        {"label": "section_header", "text": "Intro", "level": 0, "textVia": "attr"}
      ],
      "bare": false,          # true 时 iterate_items() 直接产出 item（不裹元组）
      "raise": "boom"         # 非空时 convert() 抛错，用于验证失败降级
    }

真实 Docling 的关键形状（桩必须与之对齐，否则测的是假象）：
  - `DocumentConverter().convert(source)` → result.document
  - `document.iterate_items()` → 依次产出 `(item, level)` 元组
  - `item.label` 是枚举，稳定字符串在 `.value` 上（`title` / `section_header`）
  - `item.text` 属性与 `item.get_text()` 方法取到同一段文本
"""
from __future__ import annotations

import json
import os
from typing import Any

SPEC_ENV = "DOCX_PARSE_STUB_DOCLING"


class _StubLabel:
    """模拟 `DocItemLabel`：真实实现是枚举，稳定字符串挂在 `.value` 上。"""

    def __init__(self, value: str) -> None:
        self.value = value


class _StubItem:
    def __init__(self, label: str, text: str, text_via: str) -> None:
        self.label = _StubLabel(label)
        # textVia == "get_text" 时故意不挂 `.text`，用于验证调用方的回退路径。
        if text_via == "attr":
            self.text = text
        self._text = text

    def get_text(self) -> str:
        return self._text


class _StubDocument:
    def __init__(self, entries: list[Any], bare: bool) -> None:
        self._entries = entries
        self._bare = bare

    def iterate_items(self):
        if self._bare:
            return iter([entry[0] for entry in self._entries])
        return iter(self._entries)


class _StubResult:
    def __init__(self, document: _StubDocument) -> None:
        self.document = document


def _load_spec() -> dict[str, Any]:
    raw = os.environ.get(SPEC_ENV)
    if not raw:
        raise RuntimeError(f"{SPEC_ENV} is not set")
    return json.loads(raw)


class DocumentConverter:
    """只实现桥接脚本用到的那一小部分接口。"""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        pass

    def convert(self, source: Any) -> _StubResult:
        spec = _load_spec()
        if spec.get("raise"):
            raise RuntimeError(str(spec["raise"]))

        entries = [
            (
                _StubItem(
                    entry["label"],
                    entry.get("text", ""),
                    entry.get("textVia", "attr"),
                ),
                entry.get("level", 0),
            )
            for entry in spec.get("items", [])
        ]
        return _StubResult(_StubDocument(entries, bool(spec.get("bare"))))
