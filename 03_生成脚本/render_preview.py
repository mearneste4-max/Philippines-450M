# -*- coding: utf-8 -*-
"""PPTX → 逐页 PNG 预览（LibreOffice 转 PDF，再用 PyMuPDF 栅格化）。

不直接用渲染 CLI 出图：它会把 13.33×7.5 in 的页面按 10.67×6.0 in 输出（裁掉 20%）。

用法：python render_preview.py <方案.pptx> [输出目录，缺省 02_页面预览] [宽度像素，缺省 1734]
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz  # PyMuPDF


def main(pptx, outdir, width=1734):
    pptx = Path(pptx).resolve()
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        sys.exit('找不到 LibreOffice（soffice）')
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, '--headless', '--convert-to', 'pdf', '--outdir', tmp, str(pptx)],
                       check=True, capture_output=True)
        pdf = Path(tmp) / (pptx.stem + '.pdf')
        doc = fitz.open(pdf)
        for i, page in enumerate(doc, 1):
            zoom = width / page.rect.width
            page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(outdir / ('P%02d.png' % i))
        print('rendered %d pages -> %s' % (len(doc), outdir))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    here = Path(__file__).resolve().parent
    main(sys.argv[1],
         sys.argv[2] if len(sys.argv) > 2 else here.parent / '02_页面预览',
         int(sys.argv[3]) if len(sys.argv) > 3 else 1734)
