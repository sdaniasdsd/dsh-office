"""A deterministic layout test double; does not claim actual rendering accuracy."""
import os
from types import SimpleNamespace


class Document:
    def __init__(self, path):
        self.path = path

    def layout(self):
        mode = os.environ.get("DOCX_TEST_LAYOUT", "normal")
        if mode == "fail":
            raise ValueError("controlled layout failure")
        if mode == "empty":
            return []
        return [SimpleNamespace(body_index=body, physical_page=page, displayed_page=page,
                                bounds=SimpleNamespace(x=72, y=72, width=400, height=30))
                for body, page in [(0, 1), (1, 2), (2, 2), (2, 3), (2, 3)]]

    def layout_page(self, index):
        return SimpleNamespace(width=612, height=792)
