import json
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from PIL import Image, ImageDraw

root = Path(sys.argv[1]).resolve()
bundle = Path(r'C:\Users\AA\AppData\Roaming\com.yeagoo.dsh-desktop\harness\profiles\web\node_modules\@deepseek-ai\dsh-docx\runtime\win32-x64')
soffice = bundle / 'libreoffice' / 'program' / 'soffice.com'
pdftoppm = bundle / 'poppler' / 'poppler-26.09.0' / 'Library' / 'bin' / 'pdftoppm.exe'
source_root = root / 'workspace'
out = root / 'visuals' / 'pptx-xlsx-pages'
out.mkdir(parents=True, exist_ok=True)
temp_root = Path(tempfile.gettempdir()) / f'dsh-office-review-{uuid.uuid4().hex}'
temp_root.mkdir(parents=True, exist_ok=False)
specimens = [
    ('pptx', p.name) for p in sorted((source_root / 'pptx').glob('*.pptx'))
] + [
    ('xlsx', p.name) for p in sorted((source_root / 'xlsx').glob('*.xlsx'))
]
summary = {'renderer': 'bundled LibreOffice + Poppler, review-only (not the DSH Office call path)', 'dpi': 'scale-to 480 px long edge', 'cases': []}

def run(args, timeout):
    return subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=timeout,
                          creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0), check=False)

for fmt, filename in specimens:
    stem = Path(filename).stem
    source = source_root / fmt / filename
    case_dir = temp_root / fmt / stem
    inputs, pdfs, images = case_dir / 'in', case_dir / 'pdf', case_dir / 'images'
    for folder in (inputs, pdfs, images): folder.mkdir(parents=True, exist_ok=True)
    ascii_source = inputs / filename
    shutil.copyfile(source, ascii_source)
    profile = (case_dir / 'profile').as_posix()
    args = [str(soffice), f'-env:UserInstallation=file:///{profile}', '--headless', '--nologo', '--nodefault', '--norestore',
            '--convert-to', 'pdf', '--outdir', str(pdfs), str(ascii_source)]
    case = {'format': fmt, 'file': filename, 'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(), 'ok': False}
    try:
        result = run(args, 180)
        pdf = pdfs / f'{stem}.pdf'
        case['sofficeExitCode'] = result.returncode
        case['sofficeStderrTail'] = result.stderr[-1500:]
        if result.returncode != 0 or not pdf.is_file() or pdf.stat().st_size == 0:
            raise RuntimeError('LibreOffice did not emit a nonempty PDF.')
        convert = run([str(pdftoppm), '-scale-to', '480', '-png', str(pdf), str(images / 'page')], 180)
        case['popplerExitCode'] = convert.returncode
        if convert.returncode != 0: raise RuntimeError('Poppler failed to render pages.')
        page_files = sorted(images.glob('page-*.png'))
        if not page_files: raise RuntimeError('Poppler emitted no page images.')
        target_pdf = out / f'{fmt}-{stem}.pdf'
        shutil.copyfile(pdf, target_pdf)
        thumbs = []
        for idx, page in enumerate(page_files, 1):
            saved = out / f'{fmt}-{stem}-page-{idx:03d}.png'
            shutil.copyfile(page, saved)
            thumbs.append((idx, Image.open(saved).convert('RGB')))
        chunk_size = 12
        sheets = []
        for chunk_no, start in enumerate(range(0, len(thumbs), chunk_size), 1):
            chunk = thumbs[start:start + chunk_size]
            cols, rows = 3, (len(chunk) + 2) // 3
            cell_w, cell_h, label_h, gap, margin = 340, 480, 24, 12, 12
            sheet = Image.new('RGB', (margin*2 + cols*cell_w + gap*(cols-1), margin*2 + rows*(cell_h+label_h) + gap*(rows-1)), 'white')
            draw = ImageDraw.Draw(sheet)
            for pos, (page_no, im) in enumerate(chunk):
                im.thumbnail((cell_w, cell_h))
                x = margin + (pos % cols) * (cell_w + gap)
                y = margin + (pos // cols) * (cell_h + label_h + gap)
                sheet.paste(im, (x, y + label_h))
                draw.text((x, y), f'{fmt.upper()} {stem} — page {page_no}', fill='black')
            sheet_path = out / f'{fmt}-{stem}-contact-{chunk_no:02d}.png'
            sheet.save(sheet_path)
            sheets.append(str(sheet_path))
        case.update({'ok': True, 'pdfBytes': target_pdf.stat().st_size, 'pageCount': len(page_files), 'pageImages': [str(out / f'{fmt}-{stem}-page-{i:03d}.png') for i in range(1, len(page_files)+1)], 'contactSheets': sheets})
    except Exception as exc:
        case['error'] = str(exc)
    summary['cases'].append(case)
    print(json.dumps({'format': fmt, 'file': filename, 'ok': case['ok'], 'pages': case.get('pageCount'), 'error': case.get('error')}), flush=True)

(out / 'render-summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'summaryPath': str(out / 'render-summary.json'), 'tempRoot': str(temp_root), 'cases': len(summary['cases']), 'successes': sum(1 for case in summary['cases'] if case['ok'])}, ensure_ascii=False, indent=2))
