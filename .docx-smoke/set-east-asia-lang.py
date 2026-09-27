# 假设检验 2：换行差异是不是「东亚语言标记」造成的？
#
# 标准样例的 styles.xml 把东亚语言标成 en-US（Word 默认模板留下的），我们的产物标成 zh-CN。
# 只在一份产物上把 docDefaults 的东亚语言改成 en-US，其它一个字节都不动，再渲染看换行。
import re
import sys
import zipfile

source, target, east_asia = sys.argv[1], sys.argv[2], sys.argv[3]
with zipfile.ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}

styles = entries["word/styles.xml"].decode("utf-8")
patched = re.sub(r'(<w:lang[^>]*w:eastAsia=")[^"]*(")', lambda m: m.group(1) + east_asia + m.group(2), styles)
assert patched != styles, "no w:lang w:eastAsia to patch"
entries["word/styles.xml"] = patched.encode("utf-8")

with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
    for name, data in entries.items():
        out.writestr(name, data)
print(f"wrote {target}: eastAsia language set to {east_asia}")
