"""Bounded ZIP/OOXML preflight used before ExcelJS inflates an XLSX package."""
import json
import sys
import zipfile


def main():
    path = sys.argv[1]
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > 4096:
            raise ValueError("XLSX archive entry count exceeds the Profile limit")
        total = 0
        names = set()
        for item in entries:
            if item.file_size > 64 * 1024 * 1024:
                raise ValueError("XLSX entry exceeds the Profile uncompressed-size limit")
            total += item.file_size
            if total > 256 * 1024 * 1024:
                raise ValueError("XLSX archive exceeds the Profile uncompressed-size limit")
            names.add(item.filename.lower())
        required = {"[content_types].xml", "xl/workbook.xml"}
        if not required.issubset(names):
            raise ValueError("ZIP package is not an XLSX workbook")
        content_types = archive.read("[Content_Types].xml")
        if b"spreadsheetml.sheet.main+xml" not in content_types:
            raise ValueError("Only macro-free .xlsx workbooks are supported")
        if any(name.endswith("vbaproject.bin") for name in names):
            raise ValueError("Macro-enabled Excel files are not supported")
        patterns = {
            "charts": ("xl/charts/",), "drawings": ("xl/drawings/",),
            "externalLinks": ("xl/externallinks/",), "embeddedObjects": ("xl/embeddings/", "xl/oleobjects/"),
            "activeX": ("xl/activex/", "xl/ctrlprops/"), "pivot": ("xl/pivot",),
            "slicers": ("xl/slicers/",), "queryTables": ("xl/querytables/", "xl/connections.xml"),
            "digitalSignatures": ("_xmlsignatures/", "xl/vbaprojectsignature.bin"),
            "threadedComments": ("xl/threadedcomments/", "xl/persons/"), "customXml": ("customxml/",),
            "dynamicArrayMetadata": ("xl/metadata.xml",),
        }
        unsupported = sorted(feature for feature, prefixes in patterns.items() if any(name.startswith(prefix) for name in names for prefix in prefixes))
    print(json.dumps({"ok": True, "unsupportedFeatures": unsupported}, separators=(",", ":")))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"{type(error).__name__}: {error}", file=sys.stderr)
        sys.exit(2)
