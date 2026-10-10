# -*- coding: utf-8 -*-
"""菲律宾 450MHz 方案 v8（重构版，按 04_项目文档/v8_逐页大纲.md）

17 页全篇（第 3 步）。
基础图元、素材查找、模板准备与 v7 相同；v8 设计规则见下方注释。

用法：python build_ph450m_v8.py [模板.pptx] [输出.pptx]
  模板缺省为 01_交付物/菲律宾450M方案_v7.pptx（只复用母版与版式）；
  输出缺省为 01_交付物/菲律宾450M方案_v8.pptx。
"""
import copy
import hashlib
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR as ANCHOR
from pptx.enum.text import PP_ALIGN as A
from pptx.oxml.ns import qn
from pptx.util import Pt

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
TPL = sys.argv[1] if len(sys.argv) > 1 else str(REPO / '01_交付物' / '菲律宾450M方案_v7.pptx')
OUT = sys.argv[2] if len(sys.argv) > 2 else str(REPO / '01_交付物' / '菲律宾450M方案_v8.pptx')
G = HERE / '素材'
(G / '_hi').mkdir(parents=True, exist_ok=True)   # 预处理缓存，不入库

# ---------------------------------------------------------------- 设计令牌
C = {
    'WHITE': 'FFFFFF', 'PAPER': 'F7FAFD', 'CARD': 'EEF2F9', 'CARD2': 'E4EBF7',
    'LINE': 'D6E0F0', 'NAVY': '0B308E', 'NAVY_L': '3A5BA8', 'NAVY_D': '07225F',
    'BLUE_M': '8BB0CD', 'GREEN': '7F9B00', 'LIME': '8CC63F', 'GREEN_L': 'F1F6E2',
    'ORANGE': 'E8912A', 'INK': '16213E', 'STEEL': '55607A', 'SILVER': '93A0B4',
    'MIST': 'DCE5F3',
}
W, H = 960, 540
X0, X1 = 48, 912
CW = X1 - X0
EN = 'Arial'
L, CEN, R = A.LEFT, A.CENTER, A.RIGHT
TOP, MID = ANCHOR.TOP, ANCHOR.MIDDLE

# 仓库内素材排在最前，保证同名文件优先取仓库版本
ROOTS = [G / '_recovered', REPO / '06_用到的图片', REPO / '05_素材库']
ROOTS += [G / n for n in ('_orig', '_case', '_mcs', '_prod', '_cut', '_hi', '_ex', 'images')]
ROOTS += [G / '_casepool' / n / 'images' for n in
          ('PNG', 'ICCS', 'HYTALK', 'MCX', 'TETRA2', 'SDP', 'MBB', 'RAN', 'QOS')]
ROOTS += [Path(r'F:\Workspace\03_sources\references\2026-10_资源中心图库\Approved_可公开使用'),
          Path(r'F:\Workspace\04_templates\素材库'),
          Path(r'F:\Workspace\03_sources\references\2026-09_应急快速部署\素材图片'),
          Path(r'F:\Workspace\01_projects\2026-10_应急产品内部分享\working\_gen\素材\images_v3')]
_IX = {}
for r in ROOTS:
    if r.is_dir():
        for p in r.rglob('*'):
            if p.suffix.lower() in ('.png', '.jpg', '.jpeg'):
                _IX.setdefault(p.name, p)


def asset(name):
    p = Path(name)
    if p.is_absolute() and p.exists():
        return p
    if Path(name).name in _IX:
        return _IX[Path(name).name]
    raise FileNotFoundError(name)


_IMG = {}


def slim(src, maxw=1600):
    """嵌入前统一缩放并转 JPEG，避免 PPTX 体积失控。"""
    src = Path(src)
    key = hashlib.md5(str(src).encode()).hexdigest()[:10]
    out = G / '_hi' / ('v6_%s.jpg' % key)
    if not out.exists():
        from PIL import Image
        im = Image.open(src)
        if im.mode in ('RGBA', 'LA', 'P'):
            im = im.convert('RGBA')
            bg = Image.new('RGB', im.size, (255, 255, 255))
            bg.paste(im, mask=im.split()[-1])
            im = bg
        else:
            im = im.convert('RGB')
        if im.width > maxw:
            im = im.resize((maxw, max(1, int(im.height * maxw / im.width))), Image.LANCZOS)
        im.save(out, quality=88, optimize=True)
    return out


def cropcache(name, src, box):
    """按比例裁掉不要的区域（如报告页边框），缓存复用。"""
    out = G / '_hi' / name
    if not out.exists():
        from PIL import Image
        im = Image.open(src).convert('RGB')
        w, h = im.size
        im.crop((int(w * box[0]), int(h * box[1]), int(w * box[2]), int(h * box[3]))).save(out)
    return out


def illu(name, maxw=1200):
    """等距插画：保留透明通道，缩到 maxw 以内后存 PNG。"""
    p = asset(name)
    out = G / '_hi' / ('i_%s.png' % p.stem)
    if not out.exists():
        from PIL import Image
        im = Image.open(p).convert('RGBA')
        if im.width > maxw:
            im = im.resize((maxw, max(1, int(im.height * maxw / im.width))), Image.LANCZOS)
        im.save(out)
    return out


def enhance(name, src, minw=1500):
    """放大 + 锐化。src 可传文件名（延迟解析）：仓库里已有同名成品时直接用成品，
    原始素材不在本机也能重建。"""
    from PIL import Image, ImageFilter
    out = G / '_hi' / name
    if not out.exists() and name in _IX:
        return _IX[name]
    if not out.exists():
        im = Image.open(asset(src) if isinstance(src, str) else src).convert('RGB')
        if im.width < minw:
            im = im.resize((minw, int(im.height * minw / im.width)), Image.LANCZOS)
        im.filter(ImageFilter.UnsharpMask(radius=2, percent=110, threshold=3)).save(out)
    return out


# ---------------------------------------------------------------- 基础图元
def box(s, x, y, w, h, fill=None, outline=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE,
        shadow=False):
    sh = s.shapes.add_shape(shape, Pt(x), Pt(y), Pt(w), Pt(h))
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sh.fill.background()
    if outline:
        sh.line.color.rgb = RGBColor.from_string(outline); sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    if not shadow:
        sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def ln(s, x1, y1, x2, y2, color, w=1.0, dash=False):
    c = s.shapes.add_connector(1, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    c.line.color.rgb = RGBColor.from_string(color); c.line.width = Pt(w)
    if dash:
        c.line.dash_style = 4
    return c


def tx(s, x, y, w, h, text, size=11, bold=False, color='INK', align=L, anchor=TOP,
       line=1.25, after=3, font=EN):
    tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h)); tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, seg in enumerate(str(text).split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line
        p.space_after = Pt(after)
        r = p.add_run(); r.text = seg
        r.font.size = Pt(size); r.font.bold = bold; r.font.name = font
        r.font.color.rgb = RGBColor.from_string(C.get(color, color))
    return tb


def rich(s, x, y, w, h, parts, size=11, anchor=TOP, line=1.3, align=L):
    tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h)); tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = line
    for t, b, col in parts:
        r = p.add_run(); r.text = t
        r.font.size = Pt(size); r.font.bold = b; r.font.name = EN
        r.font.color.rgb = RGBColor.from_string(C.get(col, col))
    return tb


def keepalpha(src, maxw=900):
    """PNG（图标／插画）只做缩放，保留透明通道——绝不能走 slim() 的白底合成。"""
    src = Path(src)
    if src.suffix.lower() != '.png':
        return slim(src)
    key = hashlib.md5(str(src).encode()).hexdigest()[:10]
    out = G / '_hi' / ('a_%s.png' % key)
    if not out.exists():
        from PIL import Image
        im = Image.open(src)
        if im.mode not in ('RGBA', 'LA', 'P'):
            im = im.convert('RGBA')
        if im.width > maxw:
            im = im.resize((maxw, max(1, int(im.height * maxw / im.width))), Image.LANCZOS)
        im.save(out, optimize=True)
    return out


def fit(s, src, x, y, w, h):
    from PIL import Image
    src = keepalpha(src)
    with Image.open(src) as im:
        iw, ih = im.size
    r = min(w / iw, h / ih)
    return s.shapes.add_picture(str(src), Pt(x + (w - iw * r) / 2),
                                Pt(y + (h - ih * r) / 2), Pt(iw * r), Pt(ih * r))


def photo(s, src, x, y, w, h, fy=0.5, fx=0.5):
    from PIL import Image
    src = slim(src)
    with Image.open(src) as im:
        iw, ih = im.size
    p = s.shapes.add_picture(str(src), Pt(x), Pt(y), Pt(w), Pt(h))
    tgt, sr = w / h, iw / ih
    if sr > tgt:
        keep = tgt / sr; left = min(max(fx - keep / 2, 0), 1 - keep)
        p.crop_left, p.crop_right = left, 1 - keep - left
    else:
        keep = sr / tgt; top = min(max(fy - keep / 2, 0), 1 - keep)
        p.crop_top, p.crop_bottom = top, 1 - keep - top
    return p


# ---------------------------------------------------------------- 页面家具
def page(s, n, kicker, title, point=None):
    """浅色页头：白底 + 深蓝标题 + 右侧章节标签与页码 + 绿色细线。"""
    box(s, 0, 0, W, H, C['WHITE'])
    box(s, X0, 30, 18, 5, C['LIME'])
    tx(s, X0 + 26, 22, 780, 30, title, 21, True, 'NAVY')
    tx(s, X0, 62, 700, 16, kicker.upper(), 8.5, True, 'GREEN')
    tx(s, X1 - 44, 22, 44, 26, '%02d' % n, 17, True, 'NAVY', R)
    ln(s, X0, 82, X1, 82, C['LINE'], 1.0)
    if point:
        tx(s, X0, 94, CW, 18, point, 11.5, False, 'STEEL')


def concl(s, parts, y=452):
    box(s, 0, y, W, 40, C['NAVY'])
    rich(s, X0, y, CW, 40, parts, 12, anchor=MID)


def src(s, text):
    tx(s, X0, 508, CW, 14, 'Source: ' + text, 8, False, 'SILVER')


def stat(s, x, y, w, h, value, label, accent='NAVY', vsize=22):
    box(s, x, y, w, h, C['CARD'])
    box(s, x, y, 3.5, h, C[accent])
    tx(s, x + 16, y + 12, w - 32, 28, value, vsize, True, accent)
    tx(s, x + 16, y + 44, w - 32, h - 50, label, 9.5, False, 'STEEL', line=1.25)


def klab(s, x, y, text, color='GREEN'):
    tx(s, x, y, min(420, X1 - x), 14, text.upper(), 8.5, True, color)


def bullets(s, x, y, w, items, size=11, gap=34, dot='LIME'):
    for i, t in enumerate(items):
        yy = y + i * gap
        box(s, x, yy + 5, 6, 6, C[dot])
        tx(s, x + 16, yy, w - 16, gap - 4, t, size, False, 'STEEL', line=1.28)


# ---------------------------------------------------------------- 模板准备
prs = Presentation(TPL)
while len(prs.slides._sldIdLst):
    r = prs.slides._sldIdLst[0]
    prs.part.drop_rel(r.rId)
    prs.slides._sldIdLst.remove(r)
BLANK = None
for lay in prs.slide_layouts:
    if lay.name.lower() in ('blank', '空白'):
        BLANK = lay
BLANK = BLANK or prs.slide_layouts[6]


def newslide():
    return prs.slides.add_slide(BLANK)


# ================================================================ v8 设计规则
# 1. 标题即结论（action title），一行为主；不用底部深蓝结论带。
# 2. 字体：Segoe UI 家族（Windows 自带）——标题 Semibold、正文 Regular、大数字 Light。
#    Mac 上缺 Segoe UI 会回退，对外演示请用 Windows 机器或导出 PDF。
# 3. 字号：大数字 30–40pt，标题 22pt，短语 12–16pt，标签 9pt 加字距；页脚来源 7.5pt。
# 4. 少字：每个要点 ≤ 6 个词，一个图标配一个短语；解释性的话放进讲稿，不上页面。
# 5. 图标：Lucide 线性图标（ISC 许可，05_素材库/图标_lucide）；圆底反白或单色线性两种用法。
# 6. 网格：版心 48–912；左列 48–432、右列 456–912；三栏宽 272、间距 24；卡片内边距 16。
#    同一行的图标与文字按中心线对齐（文字框与图标同 y、同 h，垂直居中）。
F_HEAD, F_BODY, F_NUM = 'Segoe UI Semibold', 'Segoe UI', 'Segoe UI Light'
SECTIONS = ['A', 'B', 'C', 'D']
ICON = REPO / '05_素材库' / '图标_lucide' / 'png'
PAD = 16
COL3 = [X0 + i * 296 for i in range(3)]   # 三栏 x，栏宽 272
C3W = 272


def t8(s, x, y, w, h, text, size=13, color='INK', font=F_BODY, align=L, anchor=TOP, line=1.2, spc=0):
    tb = tx(s, x, y, w, h, text, size, False, color, align, anchor, line=line, after=0, font=font)
    if spc:
        for p in tb.text_frame.paragraphs:
            for r in p.runs:
                r._r.get_or_add_rPr().set('spc', str(int(spc * 100)))
    return tb


def cap(s, x, y, w, text, color='GREEN', size=9, align=L, h=14, anchor=TOP):
    """字距拉开的全大写小标签。"""
    return t8(s, x, y, w, h, text.upper(), size, color, F_HEAD, align, anchor, spc=1.2)


def icon(s, name, x, y, size, color='navy', disc=None, pad=0.24):
    """线性图标；disc 给定时先画圆底，图标内缩 pad。"""
    if disc:
        box(s, x, y, size, size, C[disc], shape=MSO_SHAPE.OVAL)
        x, y, size = x + size * pad, y + size * pad, size * (1 - 2 * pad)
    return s.shapes.add_picture(str(ICON / ('%s_%s.png' % (name, color))), Pt(x), Pt(y), Pt(size), Pt(size))


def icap(s, name, x, y, text, color='navy', capcol='NAVY', size=24):
    """图标 + 小标签，两者按中心线对齐。"""
    icon(s, name, x, y, size, color)
    cap(s, x + size + 10, y, 300, text, capcol, h=size, anchor=MID)


def page8(s, n, sec, topic, title):
    box(s, 0, 0, W, H, C['WHITE'])
    cap(s, X0, 22, 600, '%s   %s' % (sec, topic), h=16, anchor=MID)
    for i, k in enumerate(SECTIONS):
        box(s, X1 - 154 + i * 28, 28.5, 24, 3, C['NAVY'] if k == sec else C['MIST'])
    t8(s, X1 - 30, 22, 30, 16, '%02d' % n, 11, 'NAVY', F_HEAD, R, MID)
    t8(s, X0, 40, CW, 70, title, 22, 'NAVY', F_HEAD, L, MID, line=1.05)


def foot(s, text):
    t8(s, X0, 512, CW, 12, 'Source: ' + text, 7.5, 'SILVER')


def dots(s, cx, cy, k, col='NAVY'):
    """三个评分圆点，以 (cx, cy) 为中心。"""
    for i in range(3):
        box(s, cx - 21 + i * 16, cy - 5, 10, 10, C[col] if i < k else C['MIST'], shape=MSO_SHAPE.OVAL)


def chip(s, x, y, w, h, text, fg='GREEN', bg='WHITE', size=12):
    box(s, x, y, w, h, C[bg], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    t8(s, x, y, w, h, text, size, fg, F_HEAD, CEN, MID)


def strip(s, y, ic, head, body, bg='PAPER', h=40, col='NAVY', icol='navy', bodycol='INK'):
    """一行要点：图标 + 粗体短语 + 常规短语，垂直居中。"""
    box(s, X0, y, CW, h, C[bg])
    icon(s, ic, X0 + PAD, y + (h - 20) / 2, 20, icol)
    rich(s, X0 + PAD + 32, y, CW - 2 * PAD - 32, h,
         [(head + '   ', True, col), (body, False, bodycol)], 12.5, anchor=MID)
    for p in s.shapes[-1].text_frame.paragraphs:
        for r in p.runs:
            r.font.name = F_HEAD if r.font.bold else F_BODY
            r.font.bold = False


def rings(s, cx, cy, radii, cols):
    for r, col in zip(radii, cols):
        box(s, cx - r, cy - r * 0.92, 2 * r, 2 * r * 0.92, col, shape=MSO_SHAPE.OVAL)


PH_MAP = '菲律宾解决方案包装_slide11_1.png'
HANDSET = '终端_多模手台_1.png'


def cover_frame(s):
    box(s, 0, 0, W, H, C['WHITE'])
    box(s, 0, 0, W, 6, C['NAVY'])
    box(s, 0, 0, 320, 6, C['LIME'])
    box(s, 0, 6, 6, H - 6, C['NAVY'])
    fit(s, illu('hytera_logo.png', 900), X1 - 128, 34, 110, 40)
    rings(s, 724, 290, (176, 132, 88), ('F3F6FB', 'EAF0F8', 'DFE7F4'))
    fit(s, illu(PH_MAP, 1400), 650, 112, 170, 340)
    fit(s, illu(HANDSET), 552, 190, 86, 250)
    ln(s, X0, 470, X1, 470, C['LINE'], 0.75)


# ================================================================ P1 封面
s = newslide()
cover_frame(s)
cap(s, 68, 126, 440, 'Hytera  ×  Globe Telecom')
t8(s, 68, 150, 470, 130, '450 MHz\nMission-Critical Network\nfor the Philippines', 32, 'NAVY', F_HEAD, line=1.06)
box(s, 68, 292, 64, 4, C['LIME'])
t8(s, 68, 312, 440, 50, 'Coverage and resilience for utilities,\npublic safety and transport', 14, 'STEEL', line=1.3)
t8(s, X0, 478, 420, 24, 'Initial technical exchange  ·  October 2026', 11, 'NAVY', F_HEAD, L, MID)
t8(s, 512, 478, 400, 24, 'Hytera Communications Corporation Limited', 9.5, 'SILVER', F_BODY, R, MID)

# ================================================================ P2 执行摘要
s = newslide()
page8(s, 2, 'A', 'The answer', '450 MHz: the coverage that stays up after the storm')
ANS = [('layers', 'What it is', 'A resilience layer', 'Alongside 4G and 5G'),
       ('globe', 'Why now', 'A mainstream band', '77 countries deployed  ·  39 on LTE'),
       ('hand-coins', 'Who pays first', 'The utilities', 'Power and water anchor the build'),
       ('flag', 'What we propose', 'One pilot province', 'Two anchor customers')]
for i, (ic, lab, head, sub) in enumerate(ANS):
    y = 122 + i * 82
    icon(s, ic, X0, y + 9, 54, 'white', disc='NAVY' if i < 3 else 'GREEN')   # 圆心 = 文字块中心 y+36
    cap(s, X0 + 74, y + 4, 300, lab, 'GREEN')
    t8(s, X0 + 74, y + 20, 340, 26, head, 19, 'NAVY', F_HEAD)
    t8(s, X0 + 74, y + 48, 360, 20, sub, 13, 'STEEL')
    if i < 3:
        ln(s, X0 + 74, y + 76, 474, y + 76, C['LINE'], 0.75)
box(s, 520, 122, 392, 326, C['PAPER'])
rings(s, 716, 276, (150, 112, 74), ('EEF2F9', 'E4EBF7', 'DCE5F3'))
fit(s, illu(PH_MAP, 1400), 640, 130, 152, 290)
icon(s, 'radio-tower', 696, 232, 40, 'white', disc='GREEN', pad=0.22)
t8(s, 520, 420, 392, 20, 'One low-band layer across 7,641 islands', 11, 'STEEL', F_BODY, CEN, MID)
box(s, X0, 462, CW, 38, C['NAVY'])
icon(s, 'arrow-right', X0 + PAD, 471, 20, 'white')
t8(s, X0 + 48, 462, 120, 38, 'Today’s ask', 13, 'LIME', F_HEAD, L, MID)
t8(s, X0 + 170, 462, 680, 38, 'Confirm the spectrum  ·  Nominate a pilot province', 13, 'WHITE', F_BODY, L, MID)
foot(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p11 Table 1; NAMRIA')

# ================================================================ P3 国情
s = newslide()
page8(s, 3, 'B', 'The gap', 'What is missing is coverage that survives the storm')
GAP = [('mountain', '7,641', 'islands, mostly mountainous'),
       ('activity', '#1', 'WorldRiskIndex 2025'),
       ('tornado', '~20', 'tropical cyclones a year'),
       ('zap', '1.4M', 'consumers cut off by Typhoon Tino')]
for i, (ic, v, lab) in enumerate(GAP):
    r, c = divmod(i, 2)
    x, y = X0 + c * 198, 122 + r * 162
    box(s, x, y, 186, 150, C['CARD'])
    icon(s, ic, x + PAD, y + PAD, 36, 'white', disc='NAVY')
    t8(s, x + PAD, y + 62, 154, 44, v, 34, 'NAVY', F_NUM)
    t8(s, x + PAD, y + 110, 154, 34, lab, 12, 'STEEL', line=1.15)
photo(s, asset('China Yunnan Fire Rescue_Flood Rescue Operation_Girls and Cat_1.jpg'), 456, 122, 456, 312, fy=0.4)
t8(s, 456, 438, 456, 14, 'Flood rescue. Illustrative imagery.', 8, 'SILVER')
for j, (ic, k, v, bg, kc, vc, icol) in enumerate([
        ('building-2', '4G / 5G', 'capacity in the cities', 'CARD', 'NAVY', 'STEEL', 'navy'),
        ('radio-tower', '450 MHz', 'coverage that stays up', 'NAVY', 'LIME', 'WHITE', 'white')]):
    x, w = (X0, 384) if j == 0 else (456, 456)
    box(s, x, 458, w, 42, C[bg])
    icon(s, ic, x + PAD, 469, 20, icol)
    t8(s, x + 48, 458, 100, 42, k, 14, kc, F_HEAD, L, MID)
    t8(s, x + 140, 458, w - 156, 42, v, 13, vc, F_BODY, L, MID)
foot(s, 'NAMRIA; WorldRiskReport 2025; PAGASA; Department of Energy (5 Nov 2025)')

# ================================================================ P4 物理覆盖
s = newslide()
page8(s, 4, 'B', 'Why 450 MHz', 'One low-band grid covers what mid-band leaves dark')
SITES = [(2, 2), (2, 7), (7, 2), (7, 7)]


def waffle(s, x, y, uncovered):
    """10×10 方格：同一站址网格下的覆盖比例。离站址最远的格子先失去覆盖；绿格为站址。"""
    cells = [(r, c) for r in range(10) for c in range(10)]
    dist = {rc: min((rc[0] - a) ** 2 + (rc[1] - b) ** 2 for a, b in SITES) for rc in cells}
    dark = set(sorted(cells, key=lambda rc: (-dist[rc], rc))[:uncovered])
    for r, c in cells:
        col = C['LIME'] if (r, c) in SITES else (C['MIST'] if (r, c) in dark else C['NAVY'])
        box(s, x + c * 14.5, y + r * 14.5, 12, 12, col)


PANELS = [('410 / 450 MHz', 0, '100%', '1×'),
          ('740 MHz', 9, '91%', '2.5–3×'),
          ('2350 MHz', 41, '59%', '12–15×')]
for i, (f, unc, pct, sites) in enumerate(PANELS):
    x = X0 + i * 172
    box(s, x, 122, 160, 300, C['CARD'] if i else C['GREEN_L'])
    cap(s, x, 134, 160, f, 'NAVY' if i else 'GREEN', 10, CEN)
    waffle(s, x + 8.5, 158, unc)
    t8(s, x, 306, 160, 40, pct, 30, 'NAVY' if i else 'GREEN', F_NUM, CEN, MID)
    t8(s, x, 344, 160, 16, 'covered', 11, 'STEEL', F_BODY, CEN, MID)
    ln(s, x + 24, 368, x + 136, 368, C['LINE'], 0.75)
    t8(s, x, 376, 160, 24, sites, 17, 'NAVY', F_HEAD, CEN, MID)
    t8(s, x, 400, 160, 14, 'sites to close the gap', 9.5, 'STEEL', F_BODY, CEN, MID)
for k, (col, lab) in enumerate([('NAVY', 'covered'), ('MIST', 'not covered'), ('LIME', 'site')]):
    x = X0 + k * 110
    box(s, x, 438, 10, 10, C[col])
    t8(s, x + 16, 435, 90, 16, lab, 10, 'STEEL', F_BODY, L, MID)
cap(s, 576, 122, 336, 'What that buys')
for i, (ic, t) in enumerate([('mountain', 'Reaches over ridges\nand into valleys'),
                             ('house', 'Deep indoor: meter boxes\nand basements'),
                             ('zap', 'Few sites, so backup\npower is affordable')]):
    y = 146 + i * 72
    icon(s, ic, 576, y + 6, 44, 'white', disc='NAVY')
    t8(s, 636, y, 276, 56, t, 14, 'NAVY', F_HEAD, L, MID, line=1.15)
box(s, 576, 368, 336, 80, C['PAPER'])
icon(s, 'info', 576 + PAD, 384, 22, 'steel')
t8(s, 576 + 48, 380, 270, 30, 'Not a capacity layer', 14, 'NAVY', F_HEAD, L, MID)
t8(s, 576 + 48, 410, 270, 24, '3–5 MHz: voice, telemetry, low-rate video', 12, 'STEEL', F_BODY, L, MID)
foot(s, '450 MHz Alliance, Brief (Jun 2022) p15 - same site grid, reference modelled at 410 MHz')

# ================================================================ P5 频段已是主流
s = newslide()
page8(s, 5, 'B', 'Why now', 'The band is mainstream, and the operator model is proven')
TREND = [('globe', '86%', 'of the world lives where it\nis allocated or consulted'),
         ('radio-tower', '77 / 39', 'countries with networks /\nalready on LTE'),
         ('trending-up', '14M', 'sensor connections by 2030,\nplus 1.8M broadband')]
for i, (ic, v, lab) in enumerate(TREND):
    y = 122 + i * 104
    box(s, X0, y, 384, 94, C['CARD'])
    icon(s, ic, X0 + PAD, y + 25, 44, 'white', disc='NAVY')
    t8(s, X0 + 76, y, 120, 94, v, 30, 'NAVY', F_NUM, L, MID)
    t8(s, X0 + 196, y, 176, 94, lab, 11.5, 'STEEL', F_BODY, L, MID, line=1.15)
cap(s, 456, 122, 400, 'Operators running it')
OPS = [('Germany', '450connect', 'B72', 'zap', 'Utilities  ·  >95% coverage'),
       ('Malaysia', 'TNB', 'B31', 'zap', 'Utility grid  ·  4 areas'),
       ('Poland', 'PGE', 'B31', 'zap', 'Smart grid, field service'),
       ('Singapore', 'Public safety', 'B31', 'shield', 'Public safety, in service'),
       ('Indonesia', 'SeaTelSat', 'B31', 'satellite-dish', 'Rural LTE via satellite')]
for i, (ctry, op, band, ic, role) in enumerate(OPS):
    y = 144 + i * 57
    box(s, 456, y, 456, 52, C['CARD'] if i % 2 == 0 else C['PAPER'])
    t8(s, 456 + PAD, y + 7, 140, 20, ctry, 13.5, 'NAVY', F_HEAD, L, MID)
    t8(s, 456 + PAD, y + 27, 140, 18, op, 11, 'STEEL', F_BODY, L, MID)
    chip(s, 612, y + 14, 52, 24, band, 'NAVY', 'WHITE', 11)
    icon(s, ic, 680, y + 16, 20, 'green')
    t8(s, 710, y, 196, 52, role, 12, 'INK', F_BODY, L, MID)
strip(s, 460, 'check', 'Same pattern everywhere', 'one licensed operator carries it; utilities anchor the demand')
foot(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p4, p8-p9, p11-p12; RFI for devices (2025) p5')

# ================================================================ P6 行业优先级
s = newslide()
page8(s, 6, 'C', 'Who pays first', 'Start with the utilities: they pay first and fit best')
RANK = [('zap', 'Power', 3, 3, 3, 'Proven anchor abroad'),
        ('droplets', 'Water', 2, 3, 3, 'Same sites as power'),
        ('siren', 'Emergency', 2, 3, 2, 'Typhoon readiness'),
        ('shield', 'Police', 3, 2, 1, 'Interoperate, not replace'),
        ('train-front', 'Rail', 2, 3, 1, 'Long project cycles'),
        ('ship', 'Port', 2, 2, 2, 'No 450 MHz precedent yet')]
COLS = [('hand-coins', 'Ability to pay', 376), ('target', 'Fit with 450', 496), ('clock', 'Speed to revenue', 616)]
for ic, t, cx in COLS:
    icon(s, ic, cx - 9, 114, 18, 'steel')
    cap(s, cx - 60, 136, 120, t, 'STEEL', 8, CEN)
cap(s, 704, 136, 200, 'Key point', 'STEEL', 8)
for i, (ic, name, pay, fit_, speed, key) in enumerate(RANK):
    y = 156 + i * 50
    anchor = i < 2
    box(s, X0, y, CW, 44, C['GREEN_L'] if anchor else (C['CARD'] if i % 2 == 0 else C['PAPER']))
    t8(s, X0 + 8, y, 30, 44, str(i + 1), 20, 'GREEN' if anchor else 'SILVER', F_NUM, CEN, MID)
    icon(s, ic, X0 + 48, y + 7, 30, 'white', disc='GREEN' if anchor else 'NAVY')
    t8(s, X0 + 92, y, 160, 44, name, 15, 'NAVY', F_HEAD, L, MID)
    for (_, _, cx), k in zip(COLS, (pay, fit_, speed)):
        dots(s, cx, y + 22, k, 'GREEN' if anchor else 'NAVY')
    t8(s, 704, y, 200, 44, key, 12, 'INK', F_BODY, L, MID)
icon(s, 'check', X0, 464, 16, 'green')
t8(s, X0 + 24, 460, 600, 24, 'Green rows: the anchor customers', 11, 'GREEN', F_HEAD, L, MID)
t8(s, X0 + 24, 480, 700, 18, 'Our reading, to be tested in the pilot', 11, 'STEEL', F_BODY, L, MID)
foot(s, 'Hytera assessment; 450 MHz Alliance, Annual Global Update (Dec 2025) p8-p9; LWUA NRW data (2025)')

# ================================================================ P7–P12 六个行业页（同一模板，篇幅一致）
IND = [
    dict(n=7, rank=1, name='Power', ic='zap',
         title='Power anchors the network, as it does in Germany and Malaysia',
         hero='20260908_海本二部培训_HyTalk-MC方案_课件_V6.1_slide41_1.png', fy=0.5,
         stats=[('121', 'electric cooperatives\nunder NEA'), ('1.4M', 'consumers cut off\nby Typhoon Tino')],
         does=[('cpu', 'SCADA &\nfeeder automation'), ('radio', 'Push-to-talk\nfor line crews'),
               ('gauge', 'Smart meters,\ndeep indoors')],
         proof=[('Germany  450connect', 'B72  ·  >95% coverage'), ('Malaysia  TNB', 'B31  ·  4 areas')],
         sells=['Capacity package', 'IoT per meter', 'SCADA O&M'],
         src='NEA; Department of Energy (5 Nov 2025); 450 MHz Alliance, RFI for devices (2025) p5, Annual Global Update (Dec 2025) p8'),
    dict(n=8, rank=2, name='Water', ic='droplets',
         title='Water: find the losses instead of estimating them',
         hero='water_pumps.png', fy=0.5,
         stats=[('532', 'water districts\nunder LWUA'), ('149', 'kept losses within\nthe 20% limit')],
         does=[('gauge', 'Meters that\nreport'), ('activity', 'Pressure sensing,\nleak location'),
               ('cpu', 'Pump & reservoir\ntelemetry')],
         proof=[('Bahrain  EWA', 'water-meter network'), ('Kenya', 'band reserved for utilities')],
         sells=['IoT per meter', 'Telemetry', 'Integration'],
         src='LWUA NRW data, end-2025 (reported Aug 2026); 450 MHz Alliance, Annual Global Update (Dec 2025) p6, p8'),
    dict(n=9, rank=3, name='Emergency', ic='siren',
         title='Emergency: the layer that has to survive the storm',
         hero='China Yunnan Fire Rescue_Flood Rescue Operation_Girls and Cat_1.jpg', fy=0.42,
         stats=[('#1', 'WorldRiskIndex\n2025'), ('~20', 'tropical cyclones\na year')],
         does=[('package', 'Deployed\nwithin hours'), ('satellite-dish', 'Satellite\nbackhaul'),
               ('users', 'Field command\npost')],
         proof=[('Indonesia  SeaTelSat', 'B31 over satellite'), ('Hytera  E-pack200', 'rapid-deployment kit')],
         sells=['Readiness retainer', 'Standby kits', 'MCX service'],
         src='WorldRiskReport 2025; PAGASA; 450 MHz Alliance, RFI for devices (2025) p5; Hytera product documentation'),
    dict(n=10, rank=4, name='Police', ic='shield',
         title='Police: one group call across the agencies',
         hero='20241104_资源中心_ICC方案_场景版_V2_slide10_1.png', fy=0.42,
         stats=[('Interop', 'with DMR and TETRA\nalready in service'), ('Priority', 'when the public\nnetwork congests')],
         does=[('users', 'Cross-agency\ngroup call'), ('video', 'Live video\nfrom the scene'),
               ('network', 'DMR / TETRA\ninterworking')],
         proof=[('Singapore', 'B31  ·  public safety'), ('Russia  Rostelecom', 'B31  ·  public safety')],
         sells=['MCX per user', 'Devices as service', 'Integration'],
         src='450 MHz Alliance, Annual Global Update (Dec 2025) p8-p9; Hytera public-safety references'),
    dict(n=11, rank=5, name='Rail', ic='train-front',
         title='Rail: coverage that follows the track, not the towns',
         hero='v6_metro.png', fy=0.45,
         stats=[('Corridor', 'coverage, including\ncuttings and tunnels'), ('Train', 'to ground voice,\ndata and video')],
         does=[('radio', 'Dispatch\nalong the line'), ('video', 'On-board\nvideo'),
               ('cable', 'One network,\nnot a private one')],
         proof=[('South Africa  Transnet', 'B31 for rail'), ('Hytera  Sri Lanka', 'MCS on 9 lines')],
         sells=['MCX per user', 'Devices as service', 'Capacity'],
         src='450 MHz Alliance, Annual Global Update (Dec 2025) p6; Hytera showroom briefing L211'),
    dict(n=12, rank=6, name='Port', ic='ship',
         title='Port: the case rests on coverage over open water',
         hero='case_hzmb_bridge.jpg', fy=0.5,
         stats=[('Quay', 'to open sea on\none group call'), ('Coastal', 'sites also serve\nisland communities')],
         does=[('radio', 'Quay\ndispatch'), ('ship', 'Ship-to-shore\nvoice & data'),
               ('waves', 'Coast guard\npatrols')],
         proof=[('Precedent', 'no 450 MHz port yet'), ('Basis', 'low-band reach over sea')],
         sells=['MCX per user', 'Capacity', 'Coastal coverage'],
         src='450 MHz Alliance reports (no port deployment listed); Hytera transport and maritime references'),
]
for cfg in IND:
    s = newslide()
    page8(s, cfg['n'], 'C', 'Priority %d of 6  ·  %s' % (cfg['rank'], cfg['name']), cfg['title'])
    photo(s, asset(cfg['hero']), X0, 122, 384, 250, fy=cfg['fy'])
    icon(s, cfg['ic'], X0 + 14, 136, 44, 'white', disc='GREEN' if cfg['rank'] <= 2 else 'NAVY')
    t8(s, X0, 376, 384, 12, 'Illustrative imagery', 8, 'SILVER')
    for i, (v, lab) in enumerate(cfg['stats']):
        x = 456 + i * 228
        word = not v[0].isdigit() and not v.startswith(('#', '~'))
        t8(s, x, 116, 220, 50, v, 30 if word else 40, 'NAVY', F_NUM, L, ANCHOR.BOTTOM)
        t8(s, x, 170, 220, 36, lab, 12, 'STEEL', line=1.15)
    ln(s, 456, 216, X1, 216, C['LINE'], 0.75)
    cap(s, 456, 228, 300, 'What 450 MHz does')
    for i, (ic, t) in enumerate(cfg['does']):
        x = 456 + i * 156
        box(s, x, 250, 144, 122, C['CARD'])
        icon(s, ic, x + PAD, 266, 34, 'navy')
        t8(s, x + PAD, 314, 120, 46, t, 12.5, 'NAVY', F_HEAD, L, TOP, line=1.15)
    box(s, X0, 396, 384, 100, C['CARD'])
    icap(s, 'badge-check', X0 + PAD, 410, 'Proof point', 'navy', 'NAVY')
    for i, (who, what) in enumerate(cfg['proof']):
        y = 444 + i * 24
        t8(s, X0 + PAD, y, 176, 22, who, 12, 'NAVY', F_HEAD, L, MID)
        t8(s, X0 + 196, y, 180, 22, what, 12, 'STEEL', F_BODY, L, MID)
    box(s, 456, 396, 456, 100, C['GREEN_L'])
    icap(s, 'hand-coins', 456 + PAD, 410, 'What Globe sells', 'green', 'GREEN')
    for i, t in enumerate(cfg['sells']):
        chip(s, 456 + PAD + i * 145, 446, 134, 34, t)
    foot(s, cfg['src'])

# ================================================================ P13 端到端架构
s = newslide()
page8(s, 13, 'D', 'How we deliver', 'One build serves all six industries; only terminals and apps change')
LAY = [('smartphone', 'Terminal', 'Hytera 450M Product Introduction PPT_202609_slide9_5.png',
        [('PNC660 450M', 'handheld, B31/B72'), ('MNC650 450M', 'vehicle radio'),
         ('P70 5G+450M', 'handheld, 2026 Q4'), ('PNC660Ex', 'ATEX, 2027 Q2')]),
       ('radio-tower', 'Network', 'iBS一体化基站_1.png',
        [('450 MHz macro eNB', 'wide area, few sites'), ('Core', 'private or Globe’s own'),
         ('Backhaul', 'fibre, microwave, satellite'), ('Rapid deployment', 'portable site kit')]),
       ('layers', 'Application', 'Hytera PNC660 450MHz&MCX_slide8_3.png',
        [('MCS', 'MCPTT · MCVideo · MCData'), ('Dispatch', 'voice, video, GIS'),
         ('Recording', 'capture and playback'), ('Device management', 'zero-touch')])]
for i, (ic, name, img, items) in enumerate(LAY):
    x = COL3[i]
    box(s, x, 122, C3W, 40, C['NAVY'])
    icon(s, ic, x + PAD, 131, 22, 'white')
    t8(s, x + 48, 122, 200, 40, name, 15, 'WHITE', F_HEAD, L, MID)
    box(s, x, 162, C3W, 222, C['CARD'])
    box(s, x + PAD, 172, C3W - 2 * PAD, 62, C['WHITE'])
    fit(s, illu(img), x + PAD + 8, 176, C3W - 2 * PAD - 16, 54)
    for j, (k, d) in enumerate(items):
        y = 244 + j * 34
        box(s, x + PAD, y + 7, 6, 6, C['LIME'])
        t8(s, x + PAD + 14, y, 230, 18, k, 12, 'NAVY', F_HEAD, L, MID)
        t8(s, x + PAD + 14, y + 17, 230, 14, d, 10.5, 'STEEL', F_BODY, L, MID)
    if i < 2:
        icon(s, 'arrow-right', x + C3W + 2, 260, 20, 'steel')
for j, (ic, k, v, bg, col) in enumerate([
        ('check', 'MCS', 'a 3GPP standard: runs on any LTE or 5G radio', 'GREEN_L', 'green'),
        ('badge-check', 'Hytera', 'Alliance-listed device and eNB supplier; 450connect whitelisted', 'PAPER', 'navy')]):
    x = X0 + j * 444
    box(s, x, 400, 420, 96, C[bg])
    icap(s, ic, x + PAD, 414, k, col, 'GREEN' if col == 'green' else 'NAVY')
    t8(s, x + PAD, 446, 388, 40, v, 12.5, 'INK', F_BODY, L, TOP, line=1.2)
foot(s, '450M Roadmap; HyTalk MC V6.1; 450 MHz Alliance, Equipment Report (May 2025) p5; 450connect whitelist certificate')

# ================================================================ P14 交付记录（三个 MCS 案例）
s = newslide()
page8(s, 14, 'D', 'Delivery record', 'We already run mission-critical services at national scale')
strip(s, 122, 'info', 'Read this right:', 'these run MCS on public LTE and 5G, not on 450 MHz. They prove the service layer.',
      'PAPER', 36)
CASES = [('train-front', 'Sri Lanka', 'Sri Lanka Railways', 'Hytera PNC660 450MHz&MCX_slide13_1.jpg', 0.5,
          '1,950+', 'MCS terminals', ['9 lines  ·  403 stations', 'Public LTE (Dialog)']),
         ('train-front', 'Hong Kong', 'Light Rail', 'v6_dispatch.png', 0.5,
          '810', 'terminals', ['First railway MCS over 5G', 'Replaced TETRA']),
         ('shield', 'Papua New Guinea', 'Police HQ', '20260903_海本二部培训_HyTalk与HyTalk-Pro方案_课件_20260905_slide59_3.jpg', 0.4,
          '3,000', 'officers, one centre', ['APEC summit security', '60 s dispatch'])]
for i, (ic, ctry, name, img, fy, v, lab, facts) in enumerate(CASES):
    x = COL3[i]
    photo(s, asset(img), x, 170, C3W, 130, fy=fy)
    icon(s, ic, x + PAD, 280, 40, 'white', disc='NAVY')
    cap(s, x + 66, 304, 200, ctry, 'GREEN')
    t8(s, x + PAD, 324, 240, 24, name, 15, 'NAVY', F_HEAD, L, MID)
    t8(s, x + PAD, 348, 240, 36, v, 30, 'NAVY', F_NUM, L, MID)
    t8(s, x + PAD, 384, 240, 16, lab, 12, 'STEEL', F_BODY, L, MID)
    ln(s, x + PAD, 408, x + C3W - PAD, 408, C['LINE'], 0.75)
    for j, f in enumerate(facts):
        y = 414 + j * 24
        icon(s, 'check', x + PAD, y + 3, 16, 'green')
        t8(s, x + PAD + 24, y, 230, 22, f, 12, 'INK', F_BODY, L, MID)
t8(s, X0, 488, CW, 14, 'Hong Kong and Papua New Guinea images are illustrative.', 8, 'SILVER')
foot(s, 'Hytera showroom briefing L170, L211; Hytera PNC660 450MHz&MCX case records; PNG Police HQ case records')

# ================================================================ P15 商业模式
s = newslide()
page8(s, 15, 'D', 'Business model', 'The utility contract funds the coverage; five more lines use it')
CHAIN = [('package', 'Hytera', 'Radio, terminals, platform', 'NAVY'),
         ('radio-tower', 'Globe Telecom', 'Spectrum, network, customers', 'GREEN'),
         ('building-2', 'Anchor customers', 'Contract capacity first', 'NAVY_L')]
for i, (ic, name, d, col) in enumerate(CHAIN):
    x = X0 + i * 300
    box(s, x, 122, 264, 76, C[col])
    icon(s, ic, x + PAD, 140, 40, 'white')
    t8(s, x + 68, 130, 186, 30, name, 15, 'WHITE', F_HEAD, L, MID)
    t8(s, x + 68, 158, 186, 28, d, 12, 'MIST' if col != 'GREEN' else 'WHITE', F_BODY, L, MID)
    if i < 2:
        icon(s, 'arrow-right', x + 270, 150, 24, 'steel')
cap(s, X0, 216, 400, 'Six revenue lines  ·  and who buys them')
REV = [('gauge', 'IoT connectivity', 'Power  ·  Water'),
       ('radio', 'MCX as a service', 'Police  ·  Rail  ·  Port'),
       ('layers', 'Capacity packages', 'Power  ·  Port'),
       ('smartphone', 'Devices as a service', 'Police  ·  Rail'),
       ('cpu', 'Integration & O&M', 'Power  ·  Water'),
       ('siren', 'Emergency readiness', 'Emergency')]
for i, (ic, name, who) in enumerate(REV):
    r, c = divmod(i, 3)
    x, y = COL3[c], 238 + r * 104
    box(s, x, y, C3W, 92, C['CARD'])
    icon(s, ic, x + PAD, y + 25, 42, 'white', disc='NAVY' if i else 'GREEN')
    t8(s, x + 72, y + 18, 190, 28, name, 14, 'NAVY', F_HEAD, L, MID)
    t8(s, x + 72, y + 46, 190, 26, who, 12, 'GREEN', F_BODY, L, MID)
strip(s, 458, 'info', 'Sign demand before the build:', 'Portugal’s 450 MHz trial stalled on commercial terms, not technology')
foot(s, 'Hytera proposal; 450 MHz Alliance, Annual Global Update (Dec 2025) p9')

# ================================================================ P16 试点路线图
s = newslide()
page8(s, 16, 'D', 'Next steps', 'Pilot first: test the case in Globe’s own terrain')
STEPS = [('radio-tower', 'Confirm the spectrum', 'band, bandwidth'),
         ('map-pin', 'Pick a province', 'typhoon belt'),
         ('package', 'Build the pilot', 'sites, terminals, meters'),
         ('activity', 'Measure', 'coverage, uptime, cost'),
         ('trending-up', 'Decide on scale', 'national roll-out')]
for i, (ic, t, d) in enumerate(STEPS):
    x = X0 + i * 176
    first = i == 0
    box(s, x, 122, 160, 150, C['NAVY'] if first else C['CARD'])
    icon(s, ic, x + PAD, 138, 40, 'navy' if first else 'white', disc='WHITE' if first else 'NAVY')
    cap(s, x + 64, 138, 90, 'Step %d' % (i + 1), 'LIME' if first else 'GREEN', h=40, anchor=MID)
    t8(s, x + PAD, 190, 132, 44, t, 14, 'WHITE' if first else 'NAVY', F_HEAD, L, TOP, line=1.1)
    t8(s, x + PAD, 238, 132, 22, d, 11.5, 'MIST' if first else 'STEEL', F_BODY, L, MID)
    if i < 4:
        icon(s, 'arrow-right', x + 160, 189, 16, 'steel')
cap(s, X0, 292, 400, 'What we need from Globe')
NEED = [('radio-tower', 'Spectrum parameters', 'B31 or B72  ·  3 or 5 MHz'),
        ('map-pin', 'One candidate province', 'ideally in the typhoon belt'),
        ('handshake', 'Two introductions', 'a cooperative, a water district')]
for i, (ic, t, d) in enumerate(NEED):
    x = COL3[i]
    box(s, x, 314, C3W, 96, C['GREEN_L'])
    icon(s, ic, x + PAD, 340, 44, 'white', disc='GREEN')
    t8(s, x + 74, 330, 186, 30, t, 14, 'NAVY', F_HEAD, L, MID)
    t8(s, x + 74, 358, 186, 26, d, 12, 'STEEL', F_BODY, L, MID)
strip(s, 428, 'handshake', 'Not a build.', 'The next step is a joint pilot design, in Globe’s own terrain.', 'NAVY', 44, 'LIME', 'white', 'WHITE')
foot(s, 'Hytera proposal')

# ================================================================ P17 尾页
s = newslide()
cover_frame(s)
cap(s, 68, 126, 440, 'Hytera  ×  Globe Telecom')
t8(s, 68, 150, 440, 60, 'Thank you', 40, 'NAVY', F_HEAD)
box(s, 68, 226, 64, 4, C['LIME'])
cap(s, 68, 254, 440, 'Next step')
for i, (ic, t) in enumerate([('radio-tower', 'Confirm the spectrum'), ('map-pin', 'Nominate a pilot province'),
                             ('handshake', 'Design the pilot together')]):
    y = 278 + i * 44
    icon(s, ic, 68, y, 32, 'white', disc='NAVY')
    t8(s, 112, y, 380, 32, t, 15, 'NAVY', F_HEAD, L, MID)
t8(s, X0, 478, 420, 24, 'Mission First  ·  Safety Always', 11, 'GREEN', F_HEAD, L, MID)
t8(s, 512, 478, 400, 24, 'Hytera Communications Corporation Limited', 9.5, 'SILVER', F_BODY, R, MID)

# ---------------------------------------------------------------- 备注占位符（同 v7）
from lxml import etree
from pptx.oxml.ns import nsdecls

_BODY_PH = (
    '<p:sp %s><p:nvSpPr><p:cNvPr id="99" name="Notes Placeholder"/>'
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
    '<p:nvPr><p:ph type="body" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr>'
    '<p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm></p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>' % nsdecls('p', 'a')
)
for _part in prs.part.package.iter_parts():
    if str(_part.partname).endswith('notesMaster1.xml'):
        _tree = _part._element.find(qn('p:cSld')).find(qn('p:spTree'))
        if _tree.find('.//' + qn('p:ph')) is None:
            _tree.append(etree.fromstring(_BODY_PH))

prs.save(OUT)
print('saved: %s | slides: %d' % (OUT, len(prs.slides._sldIdLst)))
