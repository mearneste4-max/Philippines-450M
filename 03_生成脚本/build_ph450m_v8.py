# -*- coding: utf-8 -*-
"""菲律宾 450MHz 方案 v8（重构版，按 04_项目文档/v8_逐页大纲.md）

当前进度：第 2 步样张——只构建 P2 执行摘要、P6 行业优先级、P7 电力。
基础图元、素材查找、模板准备与 v7 相同；v8 设计规则见下方注释。

用法：python build_ph450m_v8.py [模板.pptx] [输出.pptx]
  模板缺省为 01_交付物/菲律宾450M方案_v7.pptx（只复用母版与版式）；
  输出缺省为 01_交付物/菲律宾450M方案_v8_样张.pptx。
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
OUT = sys.argv[2] if len(sys.argv) > 2 else str(REPO / '01_交付物' / '菲律宾450M方案_v8_样张.pptx')
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
# 3. 字号：大数字 30–40pt，标题 22pt，短语 13–16pt，标签 9pt 加字距；页脚来源 8pt。
# 4. 少字：每个要点 ≤ 6 个词，一个图标配一个短语；解释性的话放进讲稿，不上页面。
# 5. 图标：Lucide 线性图标（ISC 许可，05_素材库/图标_lucide），同一页同一种画法；
#    等距插画不再与线性图标混用。
# 6. 页眉：左上段标识（A–D + 主题，字距拉开），右上四段进度条 + 页码，取代目录页。
F_HEAD, F_BODY, F_NUM = 'Segoe UI Semibold', 'Segoe UI', 'Segoe UI Light'
SECTIONS = ['A', 'B', 'C', 'D']
ICON = REPO / '05_素材库' / '图标_lucide' / 'png'


def t8(s, x, y, w, h, text, size=13, color='INK', font=F_BODY, align=L, anchor=TOP, line=1.2, spc=0):
    tb = tx(s, x, y, w, h, text, size, False, color, align, anchor, line=line, after=0, font=font)
    if spc:
        for p in tb.text_frame.paragraphs:
            for r in p.runs:
                r._r.get_or_add_rPr().set('spc', str(int(spc * 100)))
    return tb


def cap(s, x, y, w, text, color='GREEN', size=9, align=L):
    """字距拉开的全大写小标签。"""
    return t8(s, x, y, w, 14, text.upper(), size, color, F_HEAD, align, spc=1.2)


def icon(s, name, x, y, size, color='navy', disc=None, pad=0.24):
    """线性图标；disc 给定时先画圆底，图标内缩 pad。"""
    if disc:
        box(s, x, y, size, size, C[disc], shape=MSO_SHAPE.OVAL)
        x, y, size = x + size * pad, y + size * pad, size * (1 - 2 * pad)
    return s.shapes.add_picture(str(ICON / ('%s_%s.png' % (name, color))), Pt(x), Pt(y), Pt(size), Pt(size))


def page8(s, n, sec, topic, title):
    box(s, 0, 0, W, H, C['WHITE'])
    cap(s, X0, 24, 600, '%s   %s' % (sec, topic))
    for i, k in enumerate(SECTIONS):
        box(s, X1 - 150 + i * 28, 29, 24, 3, C['NAVY'] if k == sec else C['MIST'])
    t8(s, X1 - 30, 21, 30, 16, '%02d' % n, 11, 'NAVY', F_HEAD, R)
    t8(s, X0, 40, CW, 70, title, 22, 'NAVY', F_HEAD, L, MID, line=1.05)   # 标题区居中：一行、两行都不留大空


def foot(s, text):
    t8(s, X0, 512, CW, 12, 'Source: ' + text, 7.5, 'SILVER')


def dots(s, x, y, k, col='NAVY'):
    for i in range(3):
        box(s, x + i * 16, y, 10, 10, C[col] if i < k else C['MIST'], shape=MSO_SHAPE.OVAL)


# ================================================================ P2 执行摘要
s = newslide()
page8(s, 2, 'A', 'The answer', '450 MHz: the coverage that stays up after the storm')
ANS = [('layers', 'What it is', 'A resilience layer', 'Alongside 4G and 5G'),
       ('globe', 'Why now', 'A mainstream band', '77 countries deployed  ·  39 on LTE'),
       ('hand-coins', 'Who pays first', 'The utilities', 'Power and water anchor the build'),
       ('flag', 'What we propose', 'One pilot province', 'Two anchor customers, measured')]
for i, (ic, lab, head, sub) in enumerate(ANS):
    y = 122 + i * 82
    icon(s, ic, X0, y + 6, 54, 'white', disc='NAVY' if i < 3 else 'GREEN')
    cap(s, X0 + 74, y + 4, 300, lab, 'GREEN')
    t8(s, X0 + 74, y + 20, 340, 26, head, 19, 'NAVY', F_HEAD)
    t8(s, X0 + 74, y + 48, 360, 20, sub, 13, 'STEEL')
    if i < 3:
        ln(s, X0 + 74, y + 76, 470, y + 76, C['LINE'], 0.75)
# 右：菲律宾地图 + 覆盖圈（示意 450 MHz 广覆盖）
box(s, 520, 118, 392, 330, C['PAPER'])
for k, (r, col) in enumerate([(150, 'EEF2F9'), (112, 'E4EBF7'), (74, 'DCE5F3')]):
    box(s, 716 - r, 283 - r * 0.92, 2 * r, 2 * r * 0.92, col, shape=MSO_SHAPE.OVAL)
fit(s, illu('菲律宾解决方案包装_slide11_1.png', 1400), 640, 136, 152, 296)
icon(s, 'radio-tower', 690, 236, 40, 'white', disc='GREEN', pad=0.22)
t8(s, 532, 420, 370, 18, 'One low-band layer across 7,641 islands', 11, 'STEEL', F_BODY, CEN)
# 底：今天要的决定
box(s, X0, 462, CW, 38, C['NAVY'])
icon(s, 'arrow-right', X0 + 16, 471, 20, 'white')
t8(s, X0 + 46, 462, 300, 38, 'Today’s ask', 13, 'LIME', F_HEAD, L, MID)
t8(s, X0 + 170, 462, 680, 38, 'Confirm the spectrum  ·  Nominate a pilot province', 13, 'WHITE', F_BODY, L, MID)
foot(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p11 Table 1; NAMRIA')

# ================================================================ P6 行业优先级
s = newslide()
page8(s, 6, 'C', 'Who pays first', 'Start with the utilities: they pay first and fit best')
RANK = [('zap', 'Power', 3, 3, 3, 'Proven anchor abroad'),
        ('droplets', 'Water', 2, 3, 3, 'Same sites as power'),
        ('siren', 'Emergency', 2, 3, 2, 'Typhoon readiness'),
        ('shield', 'Police', 3, 2, 1, 'Interoperate, not replace'),
        ('train-front', 'Rail', 2, 3, 1, 'Long project cycles'),
        ('ship', 'Port', 2, 2, 2, 'No 450 MHz precedent yet')]
COLS = [('hand-coins', 'Ability to pay', 352), ('target', 'Fit with 450', 470), ('clock', 'Speed to revenue', 588)]
for ic, t, x in COLS:
    icon(s, ic, x + 12, 114, 18, 'steel')
    cap(s, x - 30, 136, 104, t, 'STEEL', 8, CEN)
cap(s, 700, 136, 200, 'Key point', 'STEEL', 8)
for i, (ic, name, pay, fit_, speed, key) in enumerate(RANK):
    y = 156 + i * 50
    anchor = i < 2
    box(s, X0, y, CW, 44, C['GREEN_L'] if anchor else (C['CARD'] if i % 2 == 0 else C['PAPER']))
    t8(s, X0 + 14, y, 24, 44, str(i + 1), 20, 'GREEN' if anchor else 'SILVER', F_NUM, CEN, MID)
    icon(s, ic, X0 + 50, y + 7, 30, 'white', disc='GREEN' if anchor else 'NAVY')
    t8(s, X0 + 94, y, 160, 44, name, 15, 'NAVY', F_HEAD, L, MID)
    for (_, _, x), k in zip(COLS, (pay, fit_, speed)):
        dots(s, x + 4, y + 17, k, 'GREEN' if anchor else 'NAVY')
    t8(s, 700, y, 206, 44, key, 12, 'INK', F_BODY, L, MID)
icon(s, 'check', X0, 462, 16, 'green')
t8(s, X0 + 22, 460, 600, 20, 'Green rows: the anchor customers', 11, 'GREEN', F_HEAD)
t8(s, X0 + 22, 478, 700, 18, 'Our reading, to be tested in the pilot', 11, 'STEEL')
foot(s, 'Hytera assessment; 450 MHz Alliance, Annual Global Update (Dec 2025) p8-p9; LWUA NRW data (2025)')

# ================================================================ P7 电力（行业页样板）
s = newslide()
page8(s, 7, 'C', 'Priority 1 of 6  ·  Power', 'Power anchors the network, as it does in Germany and Malaysia')
photo(s, asset('20260908_海本二部培训_HyTalk-MC方案_课件_V6.1_slide41_1.png'),
      X0, 122, 384, 250, fy=0.5)
icon(s, 'zap', X0 + 14, 136, 44, 'white', disc='NAVY')
t8(s, X0, 376, 384, 12, 'Illustrative imagery', 8, 'SILVER')
# 右上：两个菲律宾数字
for i, (v, lab) in enumerate([('121', 'electric cooperatives\nunder NEA'),
                              ('1.4M', 'consumers cut off\nby Typhoon Tino')]):
    x = 456 + i * 230
    t8(s, x, 116, 220, 50, v, 40, 'NAVY', F_NUM)
    t8(s, x, 168, 220, 36, lab, 12, 'STEEL', line=1.15)
ln(s, 456, 216, X1, 216, C['LINE'], 0.75)
# 右中：三个能力图块
cap(s, 456, 228, 300, 'What 450 MHz does')
for i, (ic, t) in enumerate([('cpu', 'SCADA &\nfeeder automation'),
                             ('radio', 'Push-to-talk\nfor line crews'),
                             ('gauge', 'Smart meters,\ndeep indoors')]):
    x = 456 + i * 152
    box(s, x, 250, 144, 122, C['CARD'])
    icon(s, ic, x + 16, 264, 34, 'navy')
    t8(s, x + 16, 312, 120, 50, t, 12.5, 'NAVY', F_HEAD, line=1.15)
# 底：先例 + Globe 卖什么
box(s, X0, 396, 418, 100, C['CARD'])
icon(s, 'badge-check', X0 + 16, 410, 26, 'navy')
cap(s, X0 + 52, 416, 200, 'Proof point', 'NAVY')
for i, (who, what) in enumerate([('Germany  450connect', 'Band 72  ·  >95% coverage'),
                                 ('Malaysia  TNB', 'Band 31  ·  4 areas')]):
    y = 444 + i * 24
    t8(s, X0 + 18, y, 180, 20, who, 12.5, 'NAVY', F_HEAD)
    t8(s, X0 + 196, y, 210, 20, what, 12.5, 'STEEL')
box(s, 494, 396, 418, 100, C['GREEN_L'])
icon(s, 'hand-coins', 510, 410, 26, 'green')
cap(s, 546, 416, 200, 'What Globe sells', 'GREEN')
for i, t in enumerate(['Capacity package', 'IoT per meter', 'SCADA O&M']):
    x = 510 + i * 132
    box(s, x, 446, 124, 34, C['WHITE'], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    t8(s, x, 446, 124, 34, t, 12, 'GREEN', F_HEAD, CEN, MID)
foot(s, 'NEA; Department of Energy (5 Nov 2025); 450 MHz Alliance, RFI for devices (2025) p5, Annual Global Update (Dec 2025) p8')

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
