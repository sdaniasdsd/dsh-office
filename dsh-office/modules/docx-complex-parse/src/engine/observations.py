"""Read declared body pagination, sections and DrawingML without inventing layout.

Paths identify exact XML elements. Paragraph-internal breaks deliberately have no
beforeBodyIndex: a whole paragraph cannot represent a run-level boundary.
"""
from __future__ import annotations

from typing import Any

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def tag(name: str) -> str:
    prefix, local = name.split(":")
    return "{" + NS[prefix] + "}" + local


def number(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def on(value: str | None) -> bool:
    return value not in ("0", "false", "off")


def walk(node: Any, path: str):
    yield node, path
    # Deleted content and revision snapshots are not current document facts.
    if node.tag in {tag("w:del"), tag("w:sectPrChange"), tag("w:txbxContent")}:
        return
    chosen = None
    alternate = node.tag == tag("mc:AlternateContent")
    branches = list(node)
    if alternate:
        # Select one branch only; rendering both duplicates the same drawing.
        chosen = next((child for child in branches if child.tag == tag("mc:Choice")
                       and set((child.get("Requires") or "").split()) <= {"wps", "wpg"}), None)
        if chosen is None:
            chosen = node.find(tag("mc:Fallback"))
    counts: dict[str, int] = {}
    for child in branches:
        counts[child.tag] = counts.get(child.tag, 0) + 1
        if alternate and child is not chosen:
            continue
        uri, _, local = child.tag[1:].partition("}")
        prefix = next((key for key, value in NS.items() if uri == value), None)
        name = f"{prefix}:{local}" if prefix else f"*[local-name()='{local}']"
        yield from walk(child, f"{path}/{name}[{counts[child.tag]}]")


def geometry(sect: Any) -> dict[str, Any] | None:
    size, margin = sect.find(tag("w:pgSz")), sect.find(tag("w:pgMar"))
    if size is None or margin is None:
        return None
    width, height = number(size.get(tag("w:w"))), number(size.get(tag("w:h")))
    margins = {name: number(margin.get(tag(f"w:{name}")))
               for name in ("top", "right", "bottom", "left")}
    if not width or not height or width < 0 or height < 0 or None in margins.values():
        return None
    return {"width": width, "height": height, "unit": "twip",
            "orientation": "landscape" if width > height else "portrait", "margins": margins}


def drawing(node: Any, path: str, body_index: int) -> dict[str, Any]:
    inline = node.tag == tag("wp:inline")
    positions: list[tuple[str, str | None, int | None]] = []
    simple = node.find(tag("wp:simplePos"))
    for axis, attr in (("H", "x"), ("V", "y")):
        pos = node.find(tag("wp:position" + axis))
        if not inline and node.get("simplePos") in ("1", "true") and simple is not None:
            positions.append(("page", None, number(simple.get(attr))))
        elif pos is None:
            positions.append(("unknown", None, None))
        else:
            align, offset = pos.find(tag("wp:align")), pos.find(tag("wp:posOffset"))
            positions.append((pos.get("relativeFrom") or "unknown",
                              align.text if align is not None else None,
                              number(offset.text) if offset is not None else None))
    horizontal, vertical = positions
    behind = node.get("behindDoc") in ("1", "true")
    wrap = "inline" if inline else "none"
    for name, mode in (("wrapSquare", "square"), ("wrapTight", "tight"),
                       ("wrapThrough", "through"), ("wrapTopAndBottom", "topAndBottom"),
                       ("wrapNone", "behindText" if behind else "inFrontOfText")):
        if not inline and node.find(tag("wp:" + name)) is not None:
            wrap = mode
            break
    extent, prop = node.find(tag("wp:extent")), node.find(tag("wp:docPr"))
    blip, chart = node.find(".//" + tag("a:blip")), node.find(".//" + tag("c:chart"))
    textbox = node.find(".//" + tag("w:txbxContent"))
    kind = "image" if blip is not None else "chart" if chart is not None else "textbox" if textbox is not None else "shape"
    rel = chart.get(tag("r:id")) if chart is not None else (
        (blip.get(tag("r:embed")) or blip.get(tag("r:link"))) if blip is not None else None)
    return {
        "path": path, "bodyIndex": body_index, "kind": kind,
        "anchor": {"relativeFromHorizontal": horizontal[0], "relativeFromVertical": vertical[0],
                   "alignHorizontal": horizontal[1] if horizontal[1] in ("left", "center", "right") else None,
                   "alignVertical": vertical[1] if vertical[1] in ("top", "center", "bottom") else None,
                   "offsetX": horizontal[2], "offsetY": vertical[2], "unit": "emu",
                   "wrap": wrap, "behindText": behind},
        "width": number(extent.get("cx")) if extent is not None else None,
        "height": number(extent.get("cy")) if extent is not None else None,
        "unit": "emu", "relationshipId": rel,
        "name": prop.get("name") if prop is not None else None,
        "altText": prop.get("descr") if prop is not None else None,
    }


def read_observations(body: Any, flags: dict[str, Any], warnings: list[dict[str, Any]],
                      max_floats: int | None = None) -> dict[str, Any]:
    breaks: list[dict[str, Any]] = []
    sections: list[dict[str, Any]] = []
    floats: list[dict[str, Any]] = []
    body_nodes = [node for node in body if node.tag in (tag("w:p"), tag("w:tbl"))]
    start = 0
    counts = {tag("w:p"): 0, tag("w:tbl"): 0}
    order: dict[str, int] = {}
    block_paths: list[str] = []

    def section(sect: Any, path: str, end: int) -> None:
        nonlocal start
        kind = sect.find(tag("w:type"))
        start_type = (kind.get(tag("w:val")) if kind is not None else "nextPage")
        if start_type not in ("continuous", "nextPage", "evenPage", "oddPage"):
            start_type = "unknown"
        cols = sect.find(tag("w:cols"))
        column_count = 1 if cols is None else (
            len(cols.findall(tag("w:col"))) or number(cols.get(tag("w:num"))) or 1)
        sections.append({"startBodyIndex": start if start < len(body_nodes) else None,
                         "geometry": geometry(sect), "columnCount": max(1, column_count),
                         "startType": start_type})
        if start > 0 and start_type in ("nextPage", "evenPage", "oddPage"):
            breaks.append({"path": path, "kind": "section",
                           "beforeBodyIndex": start if start < len(body_nodes) else None})
        start = end + 1

    for index, block in enumerate(body_nodes):
        counts[block.tag] += 1
        name = "p" if block.tag == tag("w:p") else "tbl"
        path = f"/w:document/w:body/w:{name}[{counts[block.tag]}]"
        block_paths.append(path)
        for node, node_path in walk(block, path):
            order[node_path] = len(order)
            if flags.get("parsePageBreaks", True):
                if node.tag == tag("w:br") and node.get(tag("w:type")) == "page":
                    breaks.append({"path": node_path, "kind": "explicit", "beforeBodyIndex": None})
                elif node.tag == tag("w:lastRenderedPageBreak"):
                    breaks.append({"path": node_path, "kind": "rendered", "beforeBodyIndex": None})
                elif node.tag == tag("w:pageBreakBefore") and on(node.get(tag("w:val"))):
                    # Only a top-level paragraph can be addressed as a whole block.
                    before = index if node_path == path + "/w:pPr[1]/w:pageBreakBefore[1]" else None
                    breaks.append({"path": node_path, "kind": "explicit", "beforeBodyIndex": before})
            if flags.get("parseFloatingObjects", True):
                if node.tag in (tag("wp:anchor"), tag("wp:inline")):
                    floats.append(drawing(node, node_path, index))
                    if max_floats is not None and len(floats) > max_floats:
                        raise OverflowError("maxFloatingObjects")
                elif node.tag == tag("w:pict"):
                    warnings.append({"code": "UNSUPPORTED_CONTENT", "path": node_path,
                                     "message": "Legacy VML drawing is not modelled as a floating object."})
        ppr = block.find(tag("w:pPr")) if name == "p" else None
        sect = ppr.find(tag("w:sectPr")) if ppr is not None else None
        if sect is not None:
            section(sect, path + "/w:pPr[1]/w:sectPr[1]", index)
    final = body.find(tag("w:sectPr"))
    if final is not None:
        section(final, "/w:document/w:body/w:sectPr[1]", len(body_nodes) - 1)
    def break_order(item: dict[str, Any]) -> float:
        before = item["beforeBodyIndex"]
        if item["kind"] == "section" and before is not None:
            return order[block_paths[before]] - 0.5
        return order.get(item["path"], len(order))
    breaks.sort(key=break_order)
    # Omitted collections mean disabled/unsupported, distinct from observed empty.
    result: dict[str, Any] = {}
    if flags.get("parsePageBreaks", True):
        result["breaks"] = breaks
    if flags.get("parseSections", True):
        result["sections"] = sections
    if flags.get("parseFloatingObjects", True):
        result["floats"] = floats
    return result
