"""Docling boundary: stdout is one UTF-8 JSON object, diagnostics go to stderr.
Models are provisioned separately. Execution is offline and never downloads weights.
"""
import argparse
import contextlib
import json
import os
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('pdf')
    parser.add_argument('--models', required=True)
    parser.add_argument('--ocr', choices=['true', 'false'], default='true')
    parser.add_argument('--tables', choices=['true', 'false'], default='true')
    parser.add_argument('--max-pages', type=int, required=True)
    parser.add_argument('--max-bytes', type=int, required=True)
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
    os.environ['HF_HUB_OFFLINE'] = '1'
    os.environ['TRANSFORMERS_OFFLINE'] = '1'
    models = Path(args.models).resolve()
    if not models.is_dir():
        raise RuntimeError('Pre-downloaded model directory is missing')
    with contextlib.redirect_stdout(sys.stderr):
        from importlib.metadata import version
        from docling.datamodel.base_models import InputFormat
        from docling.datamodel.pipeline_options import PdfPipelineOptions, RapidOcrOptions
        from docling.datamodel.accelerator_options import AcceleratorOptions, AcceleratorDevice
        from docling.document_converter import DocumentConverter, PdfFormatOption
        from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend

        options = PdfPipelineOptions(artifacts_path=models, do_ocr=args.ocr == 'true',
                                     do_table_structure=args.tables == 'true', enable_remote_services=False)
        options.ocr_options = RapidOcrOptions(backend='onnxruntime', lang=['ch'])
        options.accelerator_options = AcceleratorOptions(num_threads=4, device=AcceleratorDevice.CPU)
        converter = DocumentConverter(allowed_formats=[InputFormat.PDF], format_options={
            # Explicit supported backend, not an implicit failure fallback. The
            # native docling-parse resource loader fails on this Unicode Windows path.
            InputFormat.PDF: PdfFormatOption(pipeline_options=options, backend=PyPdfiumDocumentBackend)
        })
        result = converter.convert(Path(args.pdf), max_num_pages=args.max_pages,
                                   max_file_size=args.max_bytes, raises_on_error=True)
        status = result.status.value
        document = result.document
        items = []
        for item, _level in document.iterate_items():
            if not getattr(item, 'prov', None):
                continue
            items.append(item.model_dump(mode='json'))
        output = {'protocol': 'docling-observations/v1', 'version': version('docling'), 'backend': 'pypdfium2',
                  'status': status, 'pages': {str(n): p.model_dump(mode='json') for n, p in document.pages.items()},
                  'items': items}
    print(json.dumps(output, ensure_ascii=False))


if __name__ == '__main__':
    main()
