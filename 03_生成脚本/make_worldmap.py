# -*- coding: utf-8 -*-
"""P5 世界地图 + 欧洲放大图 + P14 国家定位小图。

按 450 MHz Alliance《Annual Global Update》(Dec 2025) §3.2–3.5 逐国条目给国家按主要用途着色；
"试验 / 咨询中"的国家单列浅色。输出透明底 PNG 与 JSON（坐标范围、标注点、各类计数），
构建脚本据此在 PPT 里叠加国名标签（用 PPT 文字，保证字体一致）。

构建脚本只读输出文件，不依赖本脚本；改国家分类或标签位置时才需要重跑。
依赖：pip install "geopandas==0.14.4" matplotlib（自带 naturalearth_lowres 底图，离线可用）
用法：python make_worldmap.py
"""
import json
import math
import warnings
from pathlib import Path

import geopandas as gpd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

warnings.filterwarnings('ignore')
OUT = Path(__file__).resolve().parent.parent / '06_用到的图片' / '预处理成品'

# 出处：Global Update 2025 p6–p9；颜色与构建脚本的图例一致
CATS = {
    'utility': ('0B308E', ['Germany', 'Austria', 'Netherlands', 'Poland', 'Ireland', 'Denmark', 'Sweden',
                           'Malaysia', 'Saudi Arabia', 'Tunisia', 'Kenya', 'Brazil', 'Suriname']),
    'safety': ('7F9B00', ['Russia', 'Hungary', 'Spain', 'Slovenia', 'Ghana', 'Indonesia', 'Italy']),
    'rail': ('8CC63F', ['South Africa']),
    'broadband': ('5B86C5', ['Norway', 'Finland', 'Armenia', 'Cambodia', 'Sri Lanka', 'Namibia', 'Angola',
                             'Nigeria', 'Mexico', 'China', 'Argentina', 'Greece']),
    'consult': ('C3D2EA', ['United Kingdom', 'France', 'Portugal', 'Morocco', 'Botswana', 'Uganda', 'Senegal',
                           'Colombia', 'Bolivia', 'Canada', 'United States of America', 'India', 'Pakistan',
                           'Thailand', 'Vietnam', 'Iraq', 'Georgia']),
}
DOTS = {'Singapore': ('safety', 103.8, 1.35), 'Bahrain': ('utility', 50.5, 26.0)}   # 底图里太小，画点
PH = (122.0, 12.5)
BASE, EDGE = '#E6EAF1', 'white'

# 标签锚点（经度, 纬度）= 标签中心；手工放在不互相遮挡的位置
LABELS_MAIN = {
    'Russia': (95, 62), 'China': (103, 34), 'Mexico': (-103, 23), 'Brazil': (-51, -10),
    'Suriname': (-48, 7), 'Argentina': (-66, -36), 'Tunisia': (10, 33), 'Nigeria': (8, 9.5),
    'Ghana': (-10, 4), 'Kenya': (46, 0), 'Angola': (17, -11.5), 'Namibia': (15, -21.5),
    'South Africa': (25, -32), 'Saudi Arabia': (44, 23), 'Bahrain': (60, 28), 'Armenia': (45, 43),
    'Sri Lanka': (72, 5), 'Cambodia': (98, 17), 'Malaysia': (93, 9), 'Singapore': (94, -3),
    'Indonesia': (120, -10),
}
LABELS_EU = {
    'Ireland': (-8, 53.2), 'Spain': (-3.5, 40), 'Netherlands': (2, 54.6), 'Germany': (10.5, 50.6),
    'Denmark': (9.5, 56.4), 'Norway': (8.5, 62), 'Sweden': (15.5, 63.5), 'Finland': (26, 63.5),
    'Poland': (19.5, 52.2), 'Austria': (14, 47.8), 'Slovenia': (11, 45.3), 'Hungary': (20.5, 47.2),
    'Italy': (13, 42.2), 'Greece': (22, 39.3),
}
MAIN = {'lon': (-170, 180), 'lat': (-45, 72), 'w_in': 12}
EU = {'lon': (-12, 33), 'lat': (35, 71), 'w_in': 5}

w = gpd.read_file(gpd.datasets.get_path('naturalearth_lowres'))
w = w[w.name != 'Antarctica']
col = {n: c for c, names in CATS.values() for n in names}
missing = [n for n in col if n not in set(w.name)]
assert not missing, missing
w['fc'] = w.name.map(lambda n: '#' + col[n] if n in col else BASE)
w.loc[w.name == 'Philippines', 'fc'] = '#F2B705'


def render(spec, path, dot_size):
    lon, lat = spec['lon'], spec['lat']
    h_in = spec['w_in'] * (lat[1] - lat[0]) / (lon[1] - lon[0])
    fig = plt.figure(figsize=(spec['w_in'], h_in), dpi=220)
    ax = fig.add_axes([0, 0, 1, 1])
    w.plot(ax=ax, color=w.fc, edgecolor=EDGE, linewidth=0.4)
    for cat, x, y in DOTS.values():
        ax.scatter([x], [y], s=dot_size, color='#' + CATS[cat][0], edgecolors='white', linewidths=1.0, zorder=5)
    ax.set_xlim(*lon); ax.set_ylim(*lat); ax.set_aspect('auto'); ax.axis('off')
    fig.savefig(path, transparent=True, dpi=220)
    plt.close(fig)


render(MAIN, OUT / 'worldmap_450.png', 40)
render(EU, OUT / 'worldmap_eu.png', 40)
counts = {k: len(v[1]) + sum(1 for c, *_ in DOTS.values() if c == k) for k, v in CATS.items()}
meta = {'main': MAIN, 'eu': EU, 'ph': PH, 'counts': counts,
        'labels_main': LABELS_MAIN, 'labels_eu': LABELS_EU}
(OUT / 'worldmap_450.json').write_text(json.dumps(meta, indent=1), encoding='utf-8')
print('saved world + Europe maps', counts)

# ---------------------------------------------------------------- P14 定位小图（竖版，与案例照片同规格 104×150）
LOCS = {'germany': ('Germany', (10.4, 51.1), 13, None),
        'malaysia': ('Malaysia', (102.2, 4.2), 9.5, None),
        'singapore': (None, (103.8, 1.35), 9, (103.8, 1.35))}
RATIO = 104 / 150
for key, (name, (cx, cy), hdeg, dot) in LOCS.items():
    wdeg = hdeg * RATIO / math.cos(math.radians(cy))
    fig = plt.figure(figsize=(2.08, 3.0), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    fc = w.name.map(lambda n: '#0B308E' if n == name else '#D5DDEA')
    w.plot(ax=ax, color=fc, edgecolor='white', linewidth=0.6)
    if dot:
        ax.scatter([dot[0]], [dot[1]], s=420, color='#0B308E', alpha=0.18, zorder=5)
        ax.scatter([dot[0]], [dot[1]], s=90, color='#0B308E', edgecolors='white', linewidths=1.5, zorder=6)
    ax.set_xlim(cx - wdeg / 2, cx + wdeg / 2); ax.set_ylim(cy - hdeg / 2, cy + hdeg / 2)
    ax.set_aspect('auto'); ax.axis('off')
    fig.savefig(OUT / ('loc_%s.png' % key), transparent=True, dpi=300)
    plt.close(fig)
print('saved locator maps:', ', '.join(LOCS))


# ---------------------------------------------------------------- P18 德国覆盖示意：全境填色 + 稀疏基站图标（避免密集点阵的不适感）
import random

from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from shapely.geometry import Point

de = w[w.name == 'Germany'].geometry.iloc[0]
minx, miny, maxx, maxy = de.bounds
rnd = random.Random(450)
pts, step = [], 1.75
lat_step = step * 0.866 * math.cos(math.radians(51))
row, y = 0, miny + 0.45
while y <= maxy:
    x = minx + 0.5 + (step / 2 if row % 2 else 0)
    while x <= maxx:
        p = Point(x + rnd.uniform(-0.15, 0.15), y + rnd.uniform(-0.1, 0.1))
        if de.buffer(-0.25).contains(p):
            pts.append((p.x, p.y))
        x += step
    y += lat_step; row += 1
cx, cy, hdeg = 10.4, 51.1, 9.6
wdeg = hdeg * (208 / 300) / math.cos(math.radians(cy))
fig = plt.figure(figsize=(2.08, 3.0), dpi=300)
ax = fig.add_axes([0, 0, 1, 1])
w.plot(ax=ax, color=w.name.map(lambda n: '#0B308E' if n == 'Germany' else '#E6EAF1'), edgecolor='white', linewidth=0.6)
ico = plt.imread(str(Path(__file__).resolve().parent.parent / '05_素材库' / '图标_lucide' / 'png' / 'radio-tower_white.png'))
for x, y in pts:
    ax.add_artist(AnnotationBbox(OffsetImage(ico, zoom=0.048), (x, y), frameon=False, zorder=6))
ax.set_xlim(cx - wdeg / 2, cx + wdeg / 2); ax.set_ylim(cy - hdeg / 2, cy + hdeg / 2)
ax.set_aspect('auto'); ax.axis('off')
fig.savefig(OUT / 'loc_germany_sites.png', transparent=True, dpi=300)
plt.close(fig)
print('saved Germany coverage sketch: %d tower icons' % len(pts))
