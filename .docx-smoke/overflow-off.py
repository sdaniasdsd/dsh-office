# 假设检验：两行的换行位置差异是不是「行末标点可不可以溢出页边距」造成的？
#
# 不修改任何模块源码，只在一份产物上把 `w:overflowPunct w:val="0"` 显式写进每个段落的
# pPr（ECMA-376 里它排在 kinsoku/wordWrap 之后、spacing/ind 之前），再渲染一次看换行。
import shutil
import sys
import zipfile

source, target = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}

document = entries["word/document.xml"].decode("utf-8")


def add_flag(match: "re.Match[str]") -> str:
    """在 <w:pPr> 之后插入 overflowPunct，位置符合 CT_PPr 的顺序。"""
    return match.group(0) + '<w:overflowPunct w:val="0"/>'


import re  # noqa: E402 - 放在函数之后只为让上面的注释先读到

patched = re.sub(r"<w:pPr>", add_flag, document)
assert patched != document
entries["word/document.xml"] = patched.encode("utf-8")

with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
    for name, data in entries.items():
        out.writestr(name, data)
print(f"wrote {target}: overflowPunct inserted into {patched.count('<w:overflowPunct')} paragraph(s)")
