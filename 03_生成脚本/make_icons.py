# -*- coding: utf-8 -*-
"""把 05_素材库/图标_lucide/svg 下的线性图标渲染成指定颜色的透明底 PNG。

构建脚本只读 PNG，不依赖本脚本；只有新增图标或新增颜色时才需要跑它。
依赖：pip install cairosvg（Windows 上 cairosvg 需要 Cairo 运行库，嫌麻烦就在 WSL 里跑）。

用法：python make_icons.py            渲染全部图标 × 全部颜色
图标来源：Lucide（https://lucide.dev），ISC 许可，可商用，许可文本见同目录 LICENSE。
"""
from pathlib import Path

import cairosvg

ROOT = Path(__file__).resolve().parent.parent / '05_素材库' / '图标_lucide'
COLORS = {'0B308E': 'navy', 'FFFFFF': 'white', '7F9B00': 'green', '55607A': 'steel'}
SIZE = 256        # 输出像素；PPT 里最大放到 ~60pt，256px 足够清晰
STROKE = '1.75'   # 比默认 2 略细，大尺寸下更精致

for svg in sorted((ROOT / 'svg').glob('*.svg')):
    src = svg.read_text(encoding='utf-8').replace('stroke-width="2"', 'stroke-width="%s"' % STROKE)
    for hexc, tag in COLORS.items():
        out = ROOT / 'png' / ('%s_%s.png' % (svg.stem, tag))
        cairosvg.svg2png(bytestring=src.replace('currentColor', '#' + hexc).encode(),
                         write_to=str(out), output_width=SIZE, output_height=SIZE)
print('rendered %d icons x %d colours -> %s' % (len(list((ROOT / 'svg').glob('*.svg'))), len(COLORS), ROOT / 'png'))
