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
SECTIONS = ['01', '02', '03', '04']
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
    cap(s, x + size + 10, y, 600, text, capcol, h=size, anchor=MID)


def page8(s, n, sec, topic, title):
    n = len(prs.slides._sldIdLst)          # 页码自动取当前页序，增删页不用手改
    box(s, 0, 0, W, H, C['WHITE'])
    cap(s, X0, 22, 600, ('%s   %s' % (sec, topic)) if sec else topic, h=16, anchor=MID)
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


def pick(stem, fallback):
    """替换位：06_用到的图片/待替换/ 下放 <stem>.jpg/.png 即自动启用，否则用 fallback 并提示。"""
    for ext in ('.jpg', '.jpeg', '.png'):
        if stem + ext in _IX:
            return stem + ext
    print('  [slot] %s 未提供，暂用 %s' % (stem, fallback))
    return fallback


PH_MAP = '菲律宾解决方案包装_slide11_1.png'
HANDSET = str(REPO / '06_用到的图片' / '预处理成品' / 'pnc660_450m_pair.png')   # PNC660 450M 正反面


def cover_frame(s):
    box(s, 0, 0, W, H, C['WHITE'])
    box(s, 0, 0, W, 6, C['NAVY'])
    box(s, 0, 0, 320, 6, C['LIME'])
    box(s, 0, 6, 6, H - 6, C['NAVY'])
    fit(s, illu('hytera_logo.png', 900), X1 - 128, 34, 110, 40)
    rings(s, 724, 290, (176, 132, 88), ('F3F6FB', 'EAF0F8', 'DFE7F4'))
    fit(s, illu(PH_MAP, 1400), 650, 112, 170, 340)
    fit(s, HANDSET, 520, 196, 170, 240)
    ln(s, X0, 470, X1, 470, C['LINE'], 0.75)


# ================================================================ P1 封面
s = newslide()
cover_frame(s)
cap(s, 68, 126, 440, 'Hytera  ×  Globe Telecom')
t8(s, 68, 150, 470, 130, 'Exploring Application\nScenarios for a 450 MHz\nMission-Critical Network', 30, 'NAVY', F_HEAD, line=1.06)
box(s, 68, 292, 64, 4, C['LIME'])
t8(s, 68, 312, 440, 50, 'Coverage and resilience for utilities,\npublic safety and transport', 14, 'STEEL', line=1.3)
t8(s, X0, 478, 420, 24, 'Initial technical exchange  ·  October 2026', 11, 'NAVY', F_HEAD, L, MID)
t8(s, 512, 478, 400, 24, 'Hytera Communications Corporation Limited', 9.5, 'SILVER', F_BODY, R, MID)

# ================================================================ 目录与章节过渡页
CHAPTERS = [('01', 'radio-tower', '450 MHz advantages', 'What can this band do that mid-band cannot?'),
            ('02', 'users', 'Target industries', 'Who needs it, and who would pay first?'),
            ('03', 'globe', 'Reference cases', 'How is it built, and where does it run today?'),
            ('04', 'coins', 'Business model', 'How could the network pay for itself?')]
PAGES = ['Pages 3 – 6', 'Pages 7 – 14', 'Pages 15 – 19', 'Pages 20 – 21']   # 增删页后同步改这里


def numicon(s, ic, num, x, y, size, disc='NAVY', badge='LIME', icol='white'):
    """带序号的图标：圆底图标 + 右上角序号圆章。"""
    icon(s, ic, x, y, size, icol, disc=disc)
    b = size * 0.42
    box(s, x + size - b * 0.75, y - b * 0.2, b, b, C[badge], 'FFFFFF', 1.5, shape=MSO_SHAPE.OVAL)
    t8(s, x + size - b * 0.75, y - b * 0.2, b, b, num, b * 0.5, 'NAVY' if badge == 'LIME' else 'WHITE', F_HEAD, CEN, MID)


def divider(s, k):
    """章节过渡页：左侧巨号章节序号与标题，右侧四章列表、当前章高亮。"""
    num, ic, title, q = CHAPTERS[k]
    box(s, 0, 0, W, H, C['WHITE'])
    box(s, 0, 0, W, 6, C['NAVY'])
    box(s, 0, 0, 320, 6, C['LIME'])
    cap(s, X0 + 20, 120, 300, 'Chapter %s of 04' % num)
    t8(s, X0 + 14, 136, 380, 150, num, 130, 'NAVY', F_NUM, L, MID)
    box(s, X0 + 20, 296, 64, 4, C['LIME'])
    t8(s, X0 + 20, 314, 380, 44, title, 30, 'NAVY', F_HEAD, L, MID)
    t8(s, X0 + 20, 360, 380, 44, q, 15, 'STEEL', F_BODY, L, TOP, line=1.25)
    for i, (n2, ic2, t2, _) in enumerate(CHAPTERS):
        y = 120 + i * 80
        on = i == k
        box(s, 488, y, 424, 68, C['GREEN_L'] if on else C['PAPER'])
        if on:
            box(s, 488, y, 4, 68, C['GREEN'])
        numicon(s, ic2, n2, 510, y + 14, 40, 'GREEN' if on else 'MIST', 'NAVY' if on else 'SILVER',
                'white' if on else 'steel')
        t8(s, 574, y, 320, 68, t2, 17 if on else 14, 'NAVY' if on else 'SILVER', F_HEAD, L, MID)
    t8(s, X1 - 30, 500, 30, 16, '%02d' % len(prs.slides._sldIdLst), 11, 'NAVY', F_HEAD, R, MID)


s = newslide()
page8(s, 0, '', 'Agenda', 'Four questions for today\u2019s discussion')
for i, (num, ic, title, q) in enumerate(CHAPTERS):
    x = X0 + i * 222
    box(s, x, 124, 198, 300, C['CARD'])
    box(s, x, 124, 198, 4, C['NAVY'])
    numicon(s, ic, num, x + PAD, 148, 56)
    cap(s, x + PAD, 222, 160, 'Chapter %s' % num, 'GREEN')
    t8(s, x + PAD, 240, 168, 50, title, 17, 'NAVY', F_HEAD, L, TOP, line=1.1)
    ln(s, x + PAD, 300, x + 198 - PAD, 300, C['LINE'], 0.75)
    t8(s, x + PAD, 310, 168, 70, q, 12.5, 'STEEL', F_BODY, L, TOP, line=1.25)
    cap(s, x + PAD, 396, 160, PAGES[i], 'SILVER', 8.5)
strip(s, 444, 'handshake', 'An open discussion:', 'the spectrum is Globe\u2019s; these are options to explore together.')

s = newslide()
divider(s, 0)

# ================================================================ 第 1 章：业务定位（领导意见 1：稀缺低频，B2B/B2G，不适合 2C——原因是带宽小）
s = newslide()
page8(s, 0, '01', 'Positioning', 'A scarce low band, best spent on B2B and B2G, not on consumer service')
cap(s, X0, 122, 400, 'Why not consumer: the channel is narrow')
BW = [('450 MHz (Band 31 / 72)', 5, '≈ 2×5 MHz', 'GREEN'),
      ('One typical LTE carrier', 20, '20 MHz', 'NAVY_L'),
      ('One typical 5G carrier (3.5 GHz)', 100, '100 MHz', 'NAVY')]
for i, (lab, mhz, txt, col) in enumerate(BW):
    y = 146 + i * 40
    t8(s, X0, y, 384, 14, lab, 10, 'STEEL', F_BODY, L, MID)
    box(s, X0, y + 17, max(6, mhz * 3.0), 14, C[col])
    t8(s, X0 + max(6, mhz * 3.0) + 8, y + 15, 90, 18, txt, 11, 'NAVY', F_HEAD, L, MID)
t8(s, X0, 268, 384, 34, 'One 450 MHz cell covers a very large area, so this narrow channel\nis shared by far more users than a mid-band cell.', 10.5, 'STEEL', F_BODY, L, TOP, line=1.2)
for j, (head, rows, bg, col, ic) in enumerate([
        ('Fits: B2B and B2G', ['Power, water, oil & gas, rail', 'Emergency and public safety', 'IoT, SCADA, MCPTT'], 'GREEN_L', 'GREEN', 'check'),
        ('Not for consumers', ['Smartphone internet', 'Video streaming', 'Home broadband'], 'CARD', 'STEEL', 'info')]):
    x = X0 + j * 198
    box(s, x, 314, 186, 130, C[bg])
    box(s, x, 314, 186, 4, C[col])
    icap(s, ic, x + PAD, 326, head, col.lower() if col != 'STEEL' else 'steel', col, 18)
    for k, r in enumerate(rows):
        t8(s, x + PAD, 356 + k * 26, 160, 22, r, 11.5, 'INK', F_BODY, L, MID)
cap(s, 456, 122, 456, 'Three more reasons')
WHY = [('smartphone', 'No consumer handsets', 'Mainstream smartphones do not support Band 31 or 72. Devices are routers, meters and rugged handhelds.'),
       ('badge-check', 'Licensed for critical infrastructure', 'Germany reserved the band for energy and water; Singapore for public safety; South Africa for rail.'),
       ('zap', 'Built for resilience, not consumer cost', 'Sites carry long backup power (72 h at 450connect) and priority for mission-critical traffic.')]
for i, (ic, head, body) in enumerate(WHY):
    y = 146 + i * 100
    box(s, 456, y, 456, 90, C['CARD'])
    icon(s, ic, 456 + PAD, y + 23, 44, 'white', disc='NAVY')
    t8(s, 456 + 76, y + 10, 364, 22, head, 13.5, 'NAVY', F_HEAD, L, MID)
    t8(s, 456 + 76, y + 34, 364, 48, body, 11, 'STEEL', F_BODY, L, TOP, line=1.2)
strip(s, 456, 'info', 'Market signal:', 'Norway\u2019s ice began moving its 450 MHz consumer users to industrial service in 2024.', h=40)
foot(s, '3GPP LTE / NR channel bandwidths; Bundesnetzagentur (Nov 2020, Mar 2021); 450 MHz Alliance, Equipment Report (May 2025) p4-p5; Annual Global Update (Dec 2025) p6-p9')

# ================================================================ 第 1 章：需求 + 450 MHz 的回答（原 The gap 与 Why 450 合并）
s = newslide()
page8(s, 0, '01', 'Why 450 MHz', 'Reach every island, cross the mountains, stay up through the storm')
SITES = [(2, 2), (2, 7), (7, 2), (7, 7)]


def waffle_png(frac, n=24):
    """n×n 方格（PIL 渲染，576 格）：同一站址网格下的覆盖比例。离站址最远的格子先失去覆盖；
    绿点为 4 个站址。返回缓存 PNG 路径。"""
    from PIL import Image, ImageDraw
    unc = round(n * n * frac)
    out = G / '_hi' / ('waffle_%d_%d.png' % (n, unc))
    if out.exists():
        return out
    cell, gap = 40, 8
    size = n * (cell + gap) - gap
    im = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    sites = [(n / 4, n / 4), (n / 4, 3 * n / 4), (3 * n / 4, n / 4), (3 * n / 4, 3 * n / 4)]
    cells = [(r, c) for r in range(n) for c in range(n)]
    dist = {(r, c): min((r + 0.5 - a) ** 2 + (c + 0.5 - b) ** 2 for a, b in sites) for r, c in cells}
    dark = set(sorted(cells, key=lambda rc: (-dist[rc], rc))[:unc])
    rgb = lambda h: tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    for r, c in cells:
        x0, y0 = c * (cell + gap), r * (cell + gap)
        d.rounded_rectangle([x0, y0, x0 + cell, y0 + cell], 7,
                            fill=rgb(C['MIST'] if (r, c) in dark else C['NAVY']))
    for a, b in sites:
        cx, cy = b * (cell + gap) - gap / 2, a * (cell + gap) - gap / 2
        d.ellipse([cx - 30, cy - 30, cx + 30, cy + 30], fill=(255, 255, 255))
        d.ellipse([cx - 22, cy - 22, cx + 22, cy + 22], fill=rgb(C['LIME']))
    im.save(out)
    return out


PANELS = [('450 MHz', 0.0, '100%', '1×'),
          ('740 MHz', 0.087, '91%', '2.5–3×'),
          ('2.35 GHz', 0.413, '59%', '12–15×')]
cap(s, X0, 122, 300, 'What the Philippines needs')
NEEDS = [('waves', 'Reach far', '7,641 islands', 'One site where mid-band needs 2.5–15×'),
         ('mountain', 'Cross rugged terrain', 'Mountainous island interiors', 'Bends over ridges and into valleys'),
         ('tornado', 'Survive the storm', '#1 WorldRiskIndex  ·  ~20 cyclones a year', 'Few sites, so backup power is affordable')]
for i, (ic, head, fact, ans) in enumerate(NEEDS):
    y = 144 + i * 92
    box(s, X0, y, 292, 84, C['CARD'])
    icon(s, ic, X0 + PAD, y + 20, 44, 'white', disc='NAVY')
    t8(s, X0 + 72, y + 8, 212, 22, head, 14, 'NAVY', F_HEAD, L, MID)
    t8(s, X0 + 72, y + 30, 212, 18, fact, 10.5, 'STEEL', F_BODY, L, MID)
    icon(s, 'arrow-right', X0 + 72, y + 55, 12, 'green')
    t8(s, X0 + 88, y + 50, 198, 24, ans, 10, 'GREEN', F_HEAD, L, MID, line=1.1)
cap(s, 360, 122, 400, 'What 450 MHz delivers  ·  same site grid')
for i, (f, unc, pct, sites) in enumerate(PANELS):
    x = 360 + i * 188
    box(s, x, 144, 176, 268, C['CARD'] if i else C['GREEN_L'])
    cap(s, x, 154, 176, f, 'NAVY' if i else 'GREEN', 10, CEN)
    s.shapes.add_picture(str(waffle_png(unc)), Pt(x + 16), Pt(174), Pt(144), Pt(144))
    t8(s, x, 322, 176, 36, pct, 28, 'NAVY' if i else 'GREEN', F_NUM, CEN, MID)
    t8(s, x, 356, 176, 14, 'covered', 10.5, 'STEEL', F_BODY, CEN, MID)
    ln(s, x + 30, 376, x + 146, 376, C['LINE'], 0.75)
    t8(s, x, 380, 176, 18, 'Baseline site grid' if i == 0 else sites + '  sites to close the gap', 10.5, 'NAVY', F_HEAD, CEN, MID)
for k, (col, lab) in enumerate([('NAVY', 'covered'), ('MIST', 'not covered'), ('LIME', 'site')]):
    x = 360 + k * 100
    box(s, x, 425, 10, 10, C[col])
    t8(s, x + 16, 422, 80, 16, lab, 10, 'STEEL', F_BODY, L, MID)
strip(s, 452, 'info', 'Not a capacity layer:', '3–5 MHz carries voice, telemetry and low-rate video, alongside 4G and 5G.', h=40)
foot(s, 'NAMRIA; WorldRiskReport 2025; PAGASA; 450 MHz Alliance Brief (Jun 2022) p15, grid modelled at 410 MHz. Vs 410 MHz, 450 MHz adds ~1 dB path loss, 740 MHz ~7 dB (Hata).')

# ================================================================ P5 频段已是主流（世界地图 + 欧洲放大）
import json
s = newslide()
page8(s, 0, '01', 'Why now', 'The band is mainstream, and utilities lead the way')
PRE = REPO / '06_用到的图片' / '预处理成品'
WM = json.loads((PRE / 'worldmap_450.json').read_text(encoding='utf-8'))


def geo(spec, x0, y0, wpt, lon, lat):
    """经纬度 → 页面坐标（等距圆柱投影，与 make_worldmap.py 一致）。"""
    hpt = wpt * (spec['lat'][1] - spec['lat'][0]) / (spec['lon'][1] - spec['lon'][0])
    return (x0 + (lon - spec['lon'][0]) / (spec['lon'][1] - spec['lon'][0]) * wpt,
            y0 + (spec['lat'][1] - lat) / (spec['lat'][1] - spec['lat'][0]) * hpt)


def label(s, cx, cy, text, size=6.5):
    wl = len(text) * size * 0.52 + 8
    box(s, cx - wl / 2, cy - 6, wl, 12, 'FFFFFF', shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    t8(s, cx - wl / 2, cy - 6, wl, 12, text, size, 'INK', F_BODY, CEN, MID)


MS, MX, MY, MW = WM['main'], X0, 120, 600
MH = MW * (MS['lat'][1] - MS['lat'][0]) / (MS['lon'][1] - MS['lon'][0])
s.shapes.add_picture(str(PRE / 'worldmap_450.png'), Pt(MX), Pt(MY), Pt(MW), Pt(MH))
for n, (lo, la) in WM['labels_main'].items():
    label(s, *geo(MS, MX, MY, MW, lo, la), n)
px, py = geo(MS, MX, MY, MW, *WM['ph'])
box(s, px - 8, py - 8, 16, 16, None, 'F2B705', 1.5, shape=MSO_SHAPE.OVAL)
ln(s, px + 7, py - 4, px + 22, py - 16, 'F2B705', 1.0)
chip(s, px + 18, py - 28, 88, 16, 'Philippines · Globe', 'INK', 'WHITE', 7.5)
box(s, px + 18, py - 28, 88, 16, None, 'F2B705', 1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# 右：欧洲放大
ES, EX, EY, EW = WM['eu'], 672, 120, 240
EH = EW * (ES['lat'][1] - ES['lat'][0]) / (ES['lon'][1] - ES['lon'][0])
box(s, EX, EY, EW, EH, C['PAPER'])
s.shapes.add_picture(str(PRE / 'worldmap_eu.png'), Pt(EX), Pt(EY), Pt(EW), Pt(EH))
for n, (lo, la) in WM['labels_eu'].items():
    label(s, *geo(ES, EX, EY, EW, lo, la), n)
cap(s, EX + 8, EY + 6, 120, 'Europe', 'STEEL', 8)
# 图例（一行）
LEG = [('0B308E', 'Utilities & industry', 'utility'), ('7F9B00', 'Public safety', 'safety'), ('8CC63F', 'Rail', 'rail'),
       ('5B86C5', 'Operator', 'broadband'), ('C3D2EA', 'Trial / consultation', 'consult')]
lx = X0
for col, lab, k in LEG:
    box(s, lx, EY + EH + 14, 10, 10, col)
    txt = '%s  %d' % (lab, WM['counts'][k])
    t8(s, lx + 16, EY + EH + 9, 150, 20, txt, 10.5, 'INK', F_BODY, L, MID)
    lx += 16 + len(txt) * 5.6 + 22
# 统计（三栏）
TREND = [('globe', '86%', 'of people: allocated\nor in consultation'),
         ('radio-tower', '77 / 39', 'countries with networks /\nalready on LTE'),
         ('trending-up', '14M', 'sensor links by 2030,\nplus 1.8M broadband')]
ty = EY + EH + 40
for i, (ic, v, lab) in enumerate(TREND):
    x = COL3[i]
    box(s, x, ty, C3W, 72, C['CARD'])
    icon(s, ic, x + PAD, ty + 18, 36, 'white', disc='NAVY')
    t8(s, x + 62, ty, 90, 72, v, 24, 'NAVY', F_NUM, L, MID)
    t8(s, x + 150, ty, 116, 72, lab, 10.5, 'STEEL', F_BODY, L, MID, line=1.15)
strip(s, ty + 86, 'check', 'The common pattern:', 'one licensed operator per country; utilities are the largest user group', h=34)
foot(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p4, p6-p9 country updates, p11-p12; countries coloured by main use')

# ================================================================ 第 2 章
s = newslide()
divider(s, 1)

# ================================================================ P6 行业优先级
s = newslide()
page8(s, 0, '02', 'Who pays first', 'Start with the utilities: they pay first and fit best')
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
t8(s, X0 + 24, 480, 700, 18, 'Our reading, to be tested with Globe', 11, 'STEEL', F_BODY, L, MID)
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
         hero=pick('port_hero', 'case_hzmb_bridge.jpg'), fy=0.5,
         stats=[('Quay', 'to open sea on\none group call'), ('Coastal', 'sites also serve\nisland communities')],
         does=[('radio', 'Quay\ndispatch'), ('ship', 'Ship-to-shore\nvoice & data'),
               ('waves', 'Coast guard\npatrols')],
         proof=[('Precedent', 'no 450 MHz port yet'), ('Basis', 'low-band reach over sea')],
         sells=['MCX per user', 'Capacity', 'Coastal coverage'],
         src='450 MHz Alliance reports (no port deployment listed); Hytera transport and maritime references'),
]
for cfg in IND:
    s = newslide()
    page8(s, 0, '02', 'Priority %d of 6  ·  %s' % (cfg['rank'], cfg['name']), cfg['title'])
    if cfg['hero']:
        photo(s, asset(cfg['hero']), X0, 122, 384, 250, fy=cfg['fy'])
    else:   # 占位：插画，待 06_用到的图片/待替换/ 放入实拍
        box(s, X0, 122, 384, 250, C['PAPER'])
        fit(s, illu(cfg['illu'], 1400), X0 + 70, 140, 300, 216)
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

# ================================================================ 第 3 章
s = newslide()
divider(s, 2)

# ================================================================ P13 端到端架构
s = newslide()
page8(s, 0, '03', 'How it is built', 'One build serves all six industries; only terminals and apps change')
LAY = [('smartphone', 'Terminal', 'Hytera 450M Product Introduction PPT_202609_slide9_5.png',
        [('PNC660 450M', 'handheld'), ('MNC650 450M', 'vehicle radio'),
         ('P70 5G+450M', '2026 Q4'), ('PNC660Ex', 'ATEX, 2027 Q2')]),
       ('radio-tower', 'Network', [pick('tower_site', str(REPO / '06_用到的图片' / '预处理成品' / 'tower_rru.png')), 'iBS一体化基站_1.png'],
        [('RRU5800M', 'on tower, Band 31'), ('BBU5801M', 'baseband in shelter'),
         ('iBS3800', 'mobile, Band 31'), ('Core', 'private or shared')]),
       ('layers', 'Application', 'Hytera PNC660 450MHz&MCX_slide8_3.png',
        [('MCS', 'PTT · video · data'), ('Dispatch', 'voice, video, GIS'),
         ('Recording', 'playback'), ('Device mgmt', 'zero-touch')])]
for i, (ic, name, img, items) in enumerate(LAY):
    x = COL3[i]
    box(s, x, 122, C3W, 40, C['NAVY'])
    icon(s, ic, x + PAD, 131, 22, 'white')
    t8(s, x + 48, 122, 200, 40, name, 15, 'WHITE', F_HEAD, L, MID)
    box(s, x, 162, C3W, 222, C['CARD'])
    box(s, x + PAD, 172, C3W - 2 * PAD, 114, C['WHITE'])
    if isinstance(img, list):   # 网络层：固定站 + 移动站并排
        # 网络层：左 = 铁塔挂 RRU5800M（王总意见：基站上塔），右 = 移动站 iBS3800；两图等宽、底边对齐
        from PIL import Image as _I
        BASE_Y, MAXH = 262, 86
        for im2, tag, xx, ww in zip(img, ['Fixed site  \u00b7  RRU5800M', 'Mobile site  \u00b7  iBS3800'],
                                     [x + PAD, x + PAD + 120], [120, 120]):
            src2 = keepalpha(asset(im2) if not str(im2).startswith('/') else im2)
            with _I.open(src2) as _im:
                iw, ih = _im.size
            r = min(ww / iw, MAXH / ih)
            s.shapes.add_picture(str(src2), Pt(xx + (ww - iw * r) / 2), Pt(BASE_Y - ih * r), Pt(iw * r), Pt(ih * r))
            t8(s, xx, 266, ww, 16, tag, 9, 'STEEL', F_HEAD, CEN, MID)
        ln(s, x + PAD + 120, 182, x + PAD + 120, 276, C['LINE'], 0.75)
    else:
        fit(s, illu(img, 1400), x + PAD + 6, 176, C3W - 2 * PAD - 12, 106)
    for j, (k, d) in enumerate(items):
        y = 294 + j * 22
        box(s, x + PAD, y + 8, 6, 6, C['LIME'])
        t8(s, x + PAD + 14, y, 112, 22, k, 12, 'NAVY', F_HEAD, L, MID)
        t8(s, x + PAD + 128, y, 112, 22, d, 11, 'STEEL', F_BODY, L, MID)
    if i < 2:
        icon(s, 'arrow-right', x + C3W + 2, 260, 20, 'steel')
for j, (ic, k, v, bg, col) in enumerate([
        ('check', 'MCS', 'a 3GPP standard: runs on any LTE or 5G radio', 'GREEN_L', 'green'),
        ('badge-check', 'Hytera', 'Alliance-listed device and eNB supplier; 450connect whitelisted', 'PAPER', 'navy')]):
    x = X0 + j * 444
    box(s, x, 400, 420, 96, C[bg])
    icap(s, ic, x + PAD, 414, k, col, 'GREEN' if col == 'green' else 'NAVY')
    t8(s, x + PAD, 446, 388, 40, v, 12.5, 'INK', F_BODY, L, TOP, line=1.2)
foot(s, '450M Roadmap; RRU5800M Product Description V3.5.01 \u00a76.5; iBS3800 Product Description V3.0; MNC650 EU DoC (B31/B72); HyTalk MC V6.1; Alliance Equipment Report (May 2025) p5')

# ================================================================ 生态与兼容性（王总意见：体现 450 生态成员身份 + 与主流厂商核心网兼容）
s = newslide()
page8(s, 0, '03', 'Ecosystem and compatibility', 'A 450 MHz Alliance member, built to plug into the network Globe runs')
cap(s, X0, 122, 380, 'In the 450 MHz ecosystem')
ECO = [('badge-check', '450 MHz Alliance member', 'Joined August 2024'),
       ('smartphone', 'Alliance device supplier list', 'Equipment Report 2025'),
       ('radio-tower', 'Alliance macro eNB supplier list', 'Equipment Report 2025'),
       ('check', 'Operator-accepted terminal', '450connect whitelist, Germany'),
       ('layers', 'Band 31 and Band 72 portfolio', 'Terminals, eNB and MCS platform')]
for i, (ic, head, sub) in enumerate(ECO):
    y = 144 + i * 58
    box(s, X0, y, 384, 52, C['GREEN_L'] if i == 0 else C['CARD'])
    icon(s, ic, X0 + PAD, y + 9, 34, 'white', disc='GREEN' if i == 0 else 'NAVY')
    t8(s, X0 + 64, y + 6, 310, 22, head, 13, 'NAVY', F_HEAD, L, MID)
    t8(s, X0 + 64, y + 27, 310, 18, sub, 10.5, 'STEEL', F_BODY, L, MID)
cap(s, 456, 122, 456, 'Compatible by design  ·  3GPP standard interfaces')
HY = [('smartphone', 'Hytera terminals', 'LTE Uu'), ('radio-tower', 'Hytera 450 MHz eNB', 'S1'),
      ('layers', 'Hytera MCS platform', 'SIP / MCX')]
for i, (ic, t, itf) in enumerate(HY):
    y = 144 + i * 72
    box(s, 456, y, 168, 60, C['NAVY'])
    icon(s, ic, 456 + 12, y + 18, 24, 'white')
    t8(s, 456 + 44, y, 118, 60, t, 11.5, 'WHITE', F_HEAD, L, MID, line=1.1)
    ln(s, 624, y + 30, 744, y + 30, C['NAVY_L'], 1.25)
    chip(s, 644, y + 19, 80, 22, itf, 'NAVY', 'MIST', 9.5)
box(s, 744, 144, 168, 204, C['GREEN_L'])
box(s, 744, 144, 168, 4, C['GREEN'])
icon(s, 'server', 744 + PAD, 162, 30, 'green')
t8(s, 744 + PAD, 198, 140, 40, 'Globe\u2019s existing core', 13, 'NAVY', F_HEAD, L, TOP, line=1.1)
t8(s, 744 + PAD, 238, 140, 18, 'EPC / 5GC from', 10.5, 'STEEL', F_BODY, L, MID)
for k, v in enumerate(['Huawei', 'ZTE', 'Nokia', 'Ericsson']):
    icon(s, 'check', 744 + PAD, 262 + k * 20, 12, 'green')
    t8(s, 744 + PAD + 18, 258 + k * 20, 120, 20, v, 11.5, 'INK', F_HEAD, L, MID)
box(s, 456, 362, 456, 66, C['PAPER'])
icon(s, 'network', 456 + PAD, 380, 30, 'navy')
t8(s, 456 + 58, 368, 384, 30, 'Multi-vendor by standard', 13, 'NAVY', F_HEAD, L, MID)
t8(s, 456 + 58, 394, 384, 30, 'The Alliance\u2019s eNB list itself spans Ericsson, Huawei, Hytera, Nokia and ZTE.', 10.5, 'STEEL', F_BODY, L, MID, line=1.15)
strip(s, 452, 'handshake', 'For Globe:', 'add Hytera radio and terminals without replacing the core you run today. Interoperability is verified per project.', h=40)
foot(s, 'Hytera news release (Aug 2024); 450 MHz Alliance, Equipment Report (May 2025) p5 \u00a74.1-4.2; 450connect whitelist; 3GPP TS 36.300 / 36.413 / 23.280')

# ================================================================ 第 3 章：德国 450connect 详解（领导意见 2）
s = newslide()
page8(s, 0, '03', 'Reference  \u00b7  Germany', '450connect: one nationwide 450 MHz network, shared by Germany\u2019s utilities')
PRE = REPO / '06_用到的图片' / '预处理成品'
box(s, X0, 122, 270, 320, C['PAPER'])
s.shapes.add_picture(str(PRE / 'loc_germany.png'), Pt(X0 + 24), Pt(132), Pt(208), Pt(300))
chip(s, X0 + PAD, 410, 110, 22, 'Band 72', 'WHITE', 'NAVY', 10.5)
chip(s, X0 + PAD + 118, 410, 120, 22, 'Licence to 2040', 'NAVY', 'WHITE', 10.5)
FACTS = [('radio-tower', '~1,600', 'sites for nationwide coverage'),
         ('zap', '72 h', 'battery backup at every site'),
         ('waves', '2\u00d74.74 MHz', '451.00\u2013455.74 / 461.00\u2013465.74 MHz'),
         ('users', '4 \u00d7 25%', 'owned by utility groups incl. E.ON and Alliander'),
         ('server', 'Nokia', 'radio, core, backhaul and power backup; O&M to 2040'),
         ('smartphone', 'PNC660', 'Hytera handheld on the 450connect whitelist')]
for i, (ic, v, lab) in enumerate(FACTS):
    r, c = divmod(i, 2)
    x, y = 342 + c * 290, 122 + r * 82
    box(s, x, y, 280, 74, C['GREEN_L'] if i == 5 else C['CARD'])
    icon(s, ic, x + PAD, y + 19, 36, 'white', disc='GREEN' if i == 5 else 'NAVY')
    t8(s, x + 64, y + 8, 206, 32, v, 22, 'GREEN' if i == 5 else 'NAVY', F_NUM, L, MID)
    t8(s, x + 64, y + 40, 206, 28, lab, 10, 'STEEL', F_BODY, L, MID, line=1.1)
cap(s, 342, 374, 400, 'From licence to network')
TL = [('Nov 2020', 'Band reserved for energy & water'), ('Mar 2021', 'Licence awarded'),
      ('Feb 2022', 'Nokia selected'), ('End 2024', 'Full operation planned')]
for i, (d, t) in enumerate(TL):
    x = 342 + i * 143
    box(s, x, 398, 135, 44, C['CARD'])
    t8(s, x + 10, 400, 120, 18, d, 11, 'GREEN', F_HEAD, L, MID)
    t8(s, x + 10, 418, 120, 22, t, 9.5, 'INK', F_BODY, L, MID, line=1.05)
strip(s, 456, 'check', 'For Globe:', 'one operator, one network, many utility customers \u2014 and no consumer service.', h=40)
foot(s, 'Bundesnetzagentur press releases (Mar 2021); Nokia newsroom (Feb 2022); 450 MHz Alliance, RFI (2025) p5, Global Update (Dec 2025) p9; 450connect; Hytera whitelist certificate')

# ================================================================ P14 参考案例：3 个 450 MHz 网络 + 3 个 MCS 部署（1:1，卡片同规格）
s = newslide()
page8(s, 0, '03', 'References', 'Three 450 MHz networks, three national mission-critical deployments')
PRE = REPO / '06_用到的图片' / '预处理成品'
ROWS = [('radio-tower', 'green', 'GREEN', '450 MHz networks  ·  operator networks', 'GREEN_L', [
            ('Germany', '450connect', 'loc_germany.png', 'B72', 'national utility network', ['>95% coverage', 'PNC660 whitelisted']),
            ('Malaysia', 'TNB', 'loc_malaysia.png', 'B31', 'utility network', ['Phase 1: 4 areas', 'Phase 2 tendering']),
            ('Singapore', 'Public safety', 'loc_singapore.png', 'B31', 'public safety network', ['In service', 'Same band as PH'])]),
        ('layers', 'navy', 'NAVY', 'Hytera mission-critical deployments at national scale', 'CARD', [
            ('Sri Lanka', 'Railways', 'Hytera PNC660 450MHz&MCX_slide13_1.jpg', '1,950+', 'MCS terminals', ['9 lines, 403 stations', 'Dialog: 450 licence']),
            ('Hong Kong', 'Light Rail', pick('hk_lightrail', 'PNC660_Snow_Train.jpg'), '810', 'MCS terminals', ['1st railway MCS on 5G', 'Replaced TETRA']),
            ('Papua New Guinea', 'Police HQ', '20260903_海本二部培训_HyTalk与HyTalk-Pro方案_课件_20260905_slide59_3.jpg', '3,000', 'officers served', ['APEC security', '60 s dispatch'])])]
IMW, CH = 104, 150
for r, (ric, icol, capcol, head, bg, cards) in enumerate(ROWS):
    y0 = 122 + r * 186
    icap(s, ric, X0, y0, head, icol, capcol, 18)
    for i, (ctry, name, img, v, lab, facts) in enumerate(cards):
        x, y = COL3[i], y0 + 24
        box(s, x, y, C3W, CH, C[bg])
        if r == 0:
            box(s, x, y, IMW, CH, C['WHITE'])
            s.shapes.add_picture(str(PRE / img), Pt(x), Pt(y), Pt(IMW), Pt(CH))
        else:
            photo(s, asset(img), x, y, IMW, CH, fy=0.5)
        tx0, tw = x + IMW + 14, C3W - IMW - 22
        cap(s, tx0, y + 12, tw, ctry, capcol)
        t8(s, tx0, y + 28, tw, 20, name, 13.5, 'NAVY', F_HEAD, L, MID)
        t8(s, tx0, y + 50, tw, 34, v, 26, 'GREEN' if r == 0 else 'NAVY', F_NUM, L, MID)
        t8(s, tx0, y + 84, tw, 16, lab, 10.5, 'STEEL', F_BODY, L, MID)
        for k, f in enumerate(facts):
            yy = y + 106 + k * 20
            icon(s, 'check', tx0, yy + 4, 12, 'green')
            t8(s, tx0 + 18, yy, tw - 18, 20, f, 10.5, 'INK', F_BODY, L, MID)
t8(s, X0, 486, CW, 14, 'Top row: operator networks; Hytera’s link is the 450connect whitelist. PNG photo is illustrative.', 8, 'SILVER')
foot(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p7-p9; RFI for devices (2025) p5; 450connect whitelist; Hytera showroom briefing L170, L211; PNG case records')

# ================================================================ 第 4 章
s = newslide()
divider(s, 3)

# ================================================================ P15 商业模式
s = newslide()
page8(s, 0, '04', 'Business model', 'Utilities fund the coverage; every other industry builds on it')
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

# ================================================================ P17 尾页
s = newslide()
cover_frame(s)
cap(s, 68, 126, 440, 'Hytera  ×  Globe Telecom')
t8(s, 68, 150, 440, 60, 'Thank you', 40, 'NAVY', F_HEAD)
box(s, 68, 226, 64, 4, C['LIME'])
cap(s, 68, 254, 440, 'For discussion')
for i, (ic, t) in enumerate([('radio-tower', 'Band and bandwidth'), ('users', 'Which industries matter most'),
                             ('map-pin', 'Where coverage is needed first')]):
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
