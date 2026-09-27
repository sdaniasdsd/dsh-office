#!/usr/bin/env python3
"""docx-parse 的 Docling 深度引擎桥接。

定位
----
本脚本是 `execute` / `verify` 的可选深度引擎；`inspect` 永远走 `docx_parse.py`。
它与轻量引擎共享**完全相同的协议**（stdout 上的同一份 JSON 载荷），
因此上层无需为换引擎改动任何契约、错误分类或 artifact 引用。

为什么不是「用 Docling 全量重写解析」
------------------------------------
Docling 把所有输入归一化成统一的 `DoclingDocument`（texts / tables / pictures），
这个模型**不携带** OPC 关系、批注、脚注，也不保留 `w:pStyle`、`w:outlineLvl`、
字符级 run 格式。若让 Docling 接管全流程，会静默丢掉本模块契约中的这些字段，
属于能力回退。

因此本桥接采取「OPC 层复用 + 正文增强」的分工：

  1. Content-Types、样式表、关系、批注、脚注、元数据 —— 复用 `docx_parse.py`
     的纯标准库实现（确定性、零依赖、与轻量引擎逐字节一致）；
  2. 正文块 —— 先由标准库产出**完整**记录，再用 Docling 做一次标题判定增强。

合并方向是单向的：Docling **只能补充**标准库未能判定为标题的段落，
永不覆盖 OOXML 已给出的权威结论（样式名 / `w:outlineLvl`）。
这保证了「开启深度引擎」在任何情况下都不会让结果比关闭时更差。

失败语义
--------
- Docling 不可导入 / 解析抛错 → 记录 issue 后回退标准库结果（文档级问题，
  不应让整次调用失败），上层会看到一条 `PARTIALLY_PARSED` 告警；
- Docling 正常运行但未发现任何标题 → 静默使用标准库结果（对无标题文档而言
  这是正常情况，报 warning 只会制造噪声）。
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any

import docx_parse as base

# Docling 中代表「标题」的标签值。
DOCLING_HEADING_LABELS = frozenset({"title", "section_header"})

# 标题层级上限：OOXML 只定义到 Heading 9，但实际语义上 6 级以上已无区分度。
MAX_HEADING_LEVEL = 6


def _normalize_key(text: str) -> str:
    """把段落文本归一化成可比较的键。

    只用「折叠空白 + 大小写无关」：Docling 与标准库提取同一段落的文本时，
    差异几乎总是出在空白折叠与软连字符上，而不是词序或标点。
    """
    return " ".join(text.split()).casefold()


def _item_text(item: Any) -> str:
    """取一个 Docling 条目的纯文本。

    `TextItem` 既有 `.text` 属性，也提供 `get_text()`；官方示例用后者。
    两条路径都尝试，避免因版本差异取到空串而静默漏掉标题。
    """
    text = getattr(item, "text", None)
    if isinstance(text, str) and text:
        return text
    getter = getattr(item, "get_text", None)
    if callable(getter):
        try:
            value = getter()
        except Exception:  # noqa: BLE001 - 取文本失败等同于「没有文本」
            return ""
        if isinstance(value, str):
            return value
    return ""


def _docling_heading_levels(path: str) -> dict[str, int]:
    """用 Docling 解析一次，返回 {归一化文本: 标题层级}。

    Docling 的 `iterate_items()` 产出 `(item, level)`，其中 level 是文档树深度。
    OOXML 并没有把「标题级别」这一语义交给视觉模型，因此这里只能把树深度
    映射为 1 基层级——这是**启发式**，也是本增强只做「补充」而不做「覆盖」的原因。
    """
    # 延迟导入：只有真正启用深度引擎时才付出加载模型的代价。
    from docling.document_converter import DocumentConverter

    converter = DocumentConverter()
    result = converter.convert(path)

    levels: dict[str, int] = {}
    for entry in result.document.iterate_items():
        # 官方签名产出 `(item, level)` 元组；非官方文档称只产出 item。
        # 两种形状都兼容，避免因版本差异直接崩掉。
        item, depth = entry if isinstance(entry, tuple) else (entry, 0)

        label = getattr(item, "label", None)
        # label 是 DocItemLabel 枚举，其 `.value` 才是稳定字符串（`title` / `section_header`）；
        # 枚举成员的 `str()` 在不同 Python 版本下形如 `DocItemLabel.TITLE`，不可直接用。
        label_value = getattr(label, "value", None) or str(label or "")
        if label_value not in DOCLING_HEADING_LABELS:
            continue

        text = _item_text(item)
        key = _normalize_key(text)
        if not key:
            continue

        depth_value = depth if isinstance(depth, int) else 0
        level = min(MAX_HEADING_LEVEL, max(1, depth_value + 1))
        # 同一文本多次出现时保留首个（即最靠前的判读）。
        levels.setdefault(key, level)
    return levels


def docling_body_parser(
    archive: Any,
    budget: base.Budget,
    flags: dict[str, Any],
    styles_by_id: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """正文解析器：标准库产出完整结构，Docling 单向补充标题判定。"""
    # 第 1 步：标准库给出权威且完整的正文记录（这是结果的下界保证）。
    blocks = base.parse_body_part(archive, budget, flags, styles_by_id)

    # 标题能力关闭时无需付出任何 Docling 成本。
    if not flags.get("parseHeadings", True):
        return blocks

    path = getattr(archive, "filename", None)
    if not path:
        budget.issue(
            "DOCLING_UNAVAILABLE",
            "cannot determine the artifact path, skipping Docling enhancement",
        )
        return blocks

    try:
        heading_levels = _docling_heading_levels(path)
    except ImportError as exc:
        budget.issue("DOCLING_UNAVAILABLE", f"docling is not importable: {exc}")
        return blocks
    except Exception as exc:  # noqa: BLE001 - 深度引擎的任何失败都必须降级而非中断
        budget.issue("DOCLING_FAILED", f"docling could not parse this artifact: {exc}")
        return blocks

    if not heading_levels:
        # 无标题文档属正常情况，不产生告警。
        return blocks

    # 第 2 步：只补充标准库判定为「非标题」的段落。
    for block in blocks:
        if block["kind"] != "paragraph":
            continue
        paragraph = block["paragraph"]
        if paragraph["headingLevel"] is not None:
            # OOXML 已给出权威判定，深度引擎无权改写。
            continue
        guessed = heading_levels.get(_normalize_key(paragraph["text"]))
        if guessed is None:
            continue
        paragraph["headingLevel"] = guessed
        if paragraph["outlineLevel"] is None:
            # outlineLvl 是 0 基，与 headingLevel 差 1。
            paragraph["outlineLevel"] = guessed - 1

    return blocks


# --------------------------------------------------------------------------- #
# CLI                                                                           #
# --------------------------------------------------------------------------- #


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Parse a DOCX via Docling into structural JSON")
    parser.add_argument("--path", required=True, help="Path to the artifact")
    parser.add_argument("--config", required=True, help="JSON config: {limits, featureFlags}")
    args = parser.parse_args(argv)

    try:
        config = json.loads(args.config)
        limits = config["limits"]
        flags = config["featureFlags"]
    except (KeyError, ValueError, TypeError) as exc:
        sys.stderr.write(f"invalid --config payload: {exc}\n")
        return 2

    try:
        payload = base.build_payload(args.path, limits, flags, body_parser=docling_body_parser)
    except Exception as exc:  # noqa: BLE001 - 兜底：任何未预期异常都不能污染协议
        payload = base.empty_payload()
        payload["error"] = {"kind": "READ_FAILED", "message": f"unexpected failure: {exc}"}

    # 强制 UTF-8 输出，避免 Windows 控制台默认编码把非 ASCII 内容写坏。
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except (AttributeError, ValueError):  # pragma: no cover
        pass
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
