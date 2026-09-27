import json
import os
import sys
from pathlib import Path
from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageStat

root = Path(sys.argv[1]).resolve()
data = root / 'data' / 'objects'
out = root / 'visuals'
out.mkdir(parents=True, exist_ok=True)
report = json.loads((root / 'report.json').read_text(encoding='utf-8'))
pdf_records = [r for r in report['records'] if r.get('format') == 'pdf' and r.get('ok')]
by_source = {}
for record in pdf_records:
    by_source.setdefault(record['sourceDocx'], []).append(record)

def image_for(ref):
    key = ref['sha256']
    path = data / key / 'artifact.png'
    return Image.open(path).convert('RGB')

def thumb_for(page):
    ref = page.get('thumbnail') or page['image']
    return image_for(ref)

def make_sheet(source, chunk_index, chunk):
    margin, gap, label_h = 18, 18, 32
    cell_w = max(im.width for _, im in chunk)
    cell_h = max(im.height for _, im in chunk)
    cols, rows = 2, (len(chunk) + 1) // 2
    sheet = Image.new('RGB', (margin * 2 + cols * cell_w + gap, margin * 2 + rows * (cell_h + label_h) + (rows - 1) * gap), 'white')
    draw = ImageDraw.Draw(sheet)
    for idx, (number, im) in enumerate(chunk):
        x = margin + (idx % cols) * (cell_w + gap)
        y = margin + (idx // cols) * (cell_h + label_h + gap)
        sheet.paste(im, (x, y + label_h))
        draw.text((x, y), f'{source} — page {number}', fill='black')
    path = out / f'{Path(source).stem}-pages-{chunk_index:02d}.png'
    sheet.save(path)
    return path

visual = {'pageSheets': [], 'comparisons': []}
for source, records in by_source.items():
    first = records[0]['renderResult']
    last = records[-1]['renderResult']
    entries = [(p['pageNumber'], thumb_for(p)) for p in first['pages']]
    for chunk_no, offset in enumerate(range(0, len(entries), 4), 1):
        path = make_sheet(source, chunk_no, entries[offset:offset + 4])
        visual['pageSheets'].append(str(path))
    page_diffs = []
    for first_page, last_page in zip(first['pages'], last['pages']):
        a, b = image_for(first_page['image']), image_for(last_page['image'])
        if a.size != b.size:
            page_diffs.append({'page': first_page['pageNumber'], 'sizeMismatch': [a.size, b.size]})
            continue
        diff = ImageChops.difference(a, b).convert('L')
        stat = ImageStat.Stat(diff)
        hist = diff.histogram()
        changed = sum(hist[11:])
        total = a.width * a.height
        page_diffs.append({'page': first_page['pageNumber'], 'width': a.width, 'height': a.height,
                           'changedPixelsOver10': round(changed / total, 6), 'meanAbsDiff': round(stat.mean[0], 5)})
    visual['comparisons'].append({'sourceDocx': source, 'rounds': len(records), 'pageCount': first['pageCount'],
        'distinctPdfHashes': len({r['renderResult']['pdf']['sha256'] for r in records}),
        'distinctFirstPagePngHashes': len({r['renderResult']['pages'][0]['image']['sha256'] for r in records}),
        'firstVsTenthPageDiffs': page_diffs})
(out / 'visual-summary.json').write_text(json.dumps(visual, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'out': str(out), 'sheetCount': len(visual['pageSheets']), 'comparisons': visual['comparisons']}, ensure_ascii=False, indent=2))
