"""Produce an auditable local runtime/model inventory without publishing files."""
import hashlib
import importlib.metadata
import json
import platform
from pathlib import Path

root = Path(__file__).resolve().parent.parent
records = []
for path in sorted((root / 'models').rglob('*')):
    if path.is_file() and '.cache' not in path.parts:
        digest = hashlib.sha256()
        with path.open('rb') as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b''):
                digest.update(chunk)
        records.append({'path': path.relative_to(root).as_posix(), 'sizeBytes': path.stat().st_size, 'sha256': digest.hexdigest()})
output = {'python': platform.python_version(), 'platform': platform.platform(),
          'packages': sorted([{'name': d.metadata['Name'], 'version': d.version} for d in importlib.metadata.distributions()], key=lambda x: x['name'].lower()),
          'models': records}
target = root / 'runtime' / 'inventory.json'
target.parent.mkdir(parents=True, exist_ok=True)
target.write_text(json.dumps(output, indent=2, ensure_ascii=False), encoding='utf-8')
print(f'{target}: {len(records)} model files, {sum(r["sizeBytes"] for r in records)} bytes')
