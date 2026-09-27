# 逐 run 清单：包含空 run，用来解释「插件报了 55 个 run，而我的探测只看到 56 个字面 run」这类计数差。
import json
import sys

from pptx import Presentation


def main():
    rows = []
    for number, slide in enumerate(Presentation(sys.argv[1]).slides, 1):
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            kind = None
            if getattr(shape, "is_placeholder", False):
                try:
                    kind = shape.placeholder_format.type.name
                except (AttributeError, ValueError):
                    kind = "?"
            for p_index, paragraph in enumerate(shape.text_frame.paragraphs):
                for r_index, run in enumerate(paragraph.runs):
                    rows.append({"slide": number, "shapeId": int(shape.shape_id), "placeholder": kind,
                                 "p": p_index, "r": r_index,
                                 "size": None if run.font.size is None else run.font.size.pt,
                                 "chars": len(run.text), "empty": run.text.strip() == ""})
    sizes = {}
    for row in rows:
        sizes[str(row["size"])] = sizes.get(str(row["size"]), 0) + 1
    print(json.dumps({"totalRuns": len(rows), "emptyRuns": sum(1 for row in rows if row["empty"]),
                      "sizeHistogram": sizes,
                      "notAtTarget": [row for row in rows if row["size"] not in (30, 18)],
                      "emptyDetail": [row for row in rows if row["empty"]]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
