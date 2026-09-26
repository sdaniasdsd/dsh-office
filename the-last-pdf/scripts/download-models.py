"""Explicit provisioning step. This script needs network; conversion does not."""
from pathlib import Path
from docling.utils.model_downloader import download_models

download_models(output_dir=Path(__file__).resolve().parent.parent / 'models',
                with_layout=True, with_tableformer=True, with_code_formula=False,
                with_picture_classifier=False, with_rapidocr=True,
                rapidocr_models=['onnxruntime:ch'])
