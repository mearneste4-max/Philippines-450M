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
# 1. 标题即结论（action title），22pt，最多两行；不再用底部深蓝结论带。
# 2. 字号下限：正文 12pt，标签 10pt，只有页脚来源允许 8pt。
# 3. 页型四种：摘要卡（P2）、矩阵（P6）、行业页（P7–P12）、数据页（其余）。
# 4. 一种图标体系：等距插画只做行业识别；项目符号一律用编号或绿色方块，不再拿手台线稿当符号。
# 5. 页眉左上是段标识（A–D）+ 主题，右上是四段进度条 + 页码，取代目录页。
SECTIONS = ['A', 'B', 'C', 'D']


def page8(s, n, sec, topic, title):
    box(s, 0, 0, W, H, C['WHITE'])
    tx(s, X0, 22, 600, 14, '%s  ·  %s' % (sec, topic.upper()), 9.5, True, 'GREEN')
    for i, k in enumerate(SECTIONS):
        box(s, X1 - 160 + i * 30, 27, 26, 4, C['NAVY'] if k == sec else C['MIST'])
    tx(s, X1 - 36, 18, 36, 18, '%02d' % n, 12, True, 'NAVY', R)
    tx(s, X0, 40, CW, 60, title, 22, True, 'NAVY', L, ANCHOR.BOTTOM, line=1.08, after=0)   # 底对齐：单行标题贴近分隔线
    ln(s, X0, 108, X1, 108, C['LINE'], 1.0)


def foot(s, text):
    tx(s, X0, 510, CW, 14, 'Source: ' + text, 8, False, 'SILVER')


def dots(s, x, y, k, col='NAVY'):
    for i in range(3):
        box(s, x + i * 15, y, 9, 9, C[col] if i < k else C['MIST'], shape=MSO_SHAPE.OVAL)


def numbered(s, x, y, w, items, gap=46, size=12):
    for i, t in enumerate(items):
        yy = y + i * gap
        box(s, x, yy, 22, 22, C['NAVY'], shape=MSO_SHAPE.OVAL)
        tx(s, x, yy, 22, 22, str(i + 1), 10, True, 'WHITE', CEN, MID)
        tx(s, x + 34, yy + 1, w - 34, gap - 6, t, size, False, 'INK', line=1.25)


# ================================================================ P2 执行摘要
s = newslide()
page8(s, 2, 'A', 'The answer', '450 MHz does one job well: coverage that is still there after the storm')
SUMMARY = [
    ('01', 'WHAT IT IS', 'A coverage and resilience layer',
     'Alongside Globe’s 4G and 5G, not instead of them. Voice, telemetry and low-rate video, not capacity.'),
    ('02', 'WHY NOW', 'The band is mainstream',
     'Networks in 77 countries, 39 of them already on LTE. Globe already holds the spectrum.'),
    ('03', 'WHO PAYS FIRST', 'The utilities anchor the build',
     'Electric cooperatives and water districts first; public safety and transport use the same sites.'),
    ('04', 'WHAT WE PROPOSE', 'A pilot, not a build',
     'One typhoon-belt province and two anchor customers, measured in Globe’s own terrain.'),
]
for i, (num, lab, head, body) in enumerate(SUMMARY):
    r, c = divmod(i, 2)
    x, y = X0 + c * 438, 124 + r * 164
    box(s, x, y, 426, 152, C['CARD'])
    box(s, x, y, 4, 152, C['NAVY'] if i < 3 else C['GREEN'])
    tx(s, x + 22, y + 14, 50, 32, num, 24, True, 'LIME', L, MID)
    tx(s, x + 74, y + 14, 320, 32, lab, 10, True, 'GREEN', L, MID)
    tx(s, x + 22, y + 56, 390, 26, head, 17, True, 'NAVY')
    tx(s, x + 22, y + 88, 386, 56, body, 12, False, 'STEEL', line=1.3)
box(s, X0, 460, CW, 36, C['GREEN_L'])
rich(s, X0 + 18, 460, CW - 36, 36,
     [('Today’s ask:  ', True, 'GREEN'),
      ('confirm the spectrum parameters and nominate a pilot province.', False, 'INK')], 12, anchor=MID)
foot(s, '450 MHz Alliance, Annual Global Update (Dec 2025) p11 Table 1; Brief (Jun 2022) p15')

# ================================================================ P6 行业优先级
s = newslide()
page8(s, 6, 'C', 'Who pays first', 'Start where the money and the fit are both highest: the utilities')
RANK = [
    ('Power', '电力抢修.png', 3, 3, 3,
     'The utilities anchor 450connect in Germany and TNB in Malaysia; scattered assets need a private data path.'),
    ('Water', '水库大坝.png', 2, 3, 3,
     'Same sites and same IoT as power. Budgets vary: only 149 of 532 districts meet the 20% loss limit.'),
    ('Emergency', '消防站火灾.png', 2, 3, 2,
     'The direct answer to the typhoon risk. Paid as standby readiness, not by usage.'),
    ('Police', '警察局.png', 3, 2, 1,
     'Large user base, but DMR and TETRA are already in service. The entry point is interoperability.'),
    ('Rail', '交通枢纽.png', 2, 3, 1,
     'Corridor coverage suits the band, but the rail network is limited and projects run long.'),
    ('Port', '船舶.png', 2, 2, 2,
     'Coverage over water is an advantage, but there is no 450 MHz port deployment to point to yet.'),
]
COLS = [('ABILITY TO PAY', 334), ('FIT WITH 450 MHZ', 424), ('TIME TO REVENUE', 514)]
tx(s, X0 + 44, 118, 120, 14, 'INDUSTRY', 10, True, 'STEEL')
for t, x in COLS:
    tx(s, x - 18, 118, 92, 14, t, 9, True, 'STEEL', CEN)
tx(s, 610, 118, 300, 14, 'WHY', 10, True, 'STEEL')
for i, (name, ic, pay, fit_, speed, why) in enumerate(RANK):
    y = 138 + i * 52
    anchor = i < 2
    box(s, X0, y, CW, 46, C['GREEN_L'] if anchor else (C['CARD'] if i % 2 == 0 else C['PAPER']))
    tx(s, X0 + 10, y, 26, 46, str(i + 1), 20, True, 'GREEN' if anchor else 'NAVY', CEN, MID)
    fit(s, illu(ic), X0 + 44, y + 4, 44, 38)
    tx(s, X0 + 98, y, 160, 46, name, 14, True, 'NAVY', L, MID)
    for (t, x), k in zip(COLS, (pay, fit_, speed)):
        dots(s, x + 6, y + 18, k, 'GREEN' if anchor else 'NAVY')
    tx(s, 610, y, 296, 46, why, 10.5, False, 'INK', L, MID, line=1.2)
box(s, X0 - 8, 138, 4, 98, C['GREEN'])
tx(s, X0, 456, 500, 16, 'GREEN ROWS = ANCHOR CUSTOMERS', 9.5, True, 'GREEN')
tx(s, X0, 474, CW, 30,
   'Our reading, to be tested in the pilot. Three dots = strongest. Police has the largest budget but the slowest path to revenue.',
   10.5, False, 'STEEL', line=1.25)
foot(s, 'Hytera assessment; 450 MHz Alliance, Annual Global Update (Dec 2025) p8-p9; LWUA NRW data (2025)')

# ================================================================ P7 电力（行业页样板）
s = newslide()
page8(s, 7, 'C', 'Priority 1 of 6  ·  Power',
      'Power is the anchor customer - as it is for 450connect in Germany and TNB in Malaysia')
photo(s, asset('20260908_海本二部培训_HyTalk-MC方案_课件_V6.1_slide41_1.png'),
      X0, 124, 400, 236, fy=0.5)
tx(s, X0, 364, 400, 14, 'Line crews on a repair job. Illustrative imagery.', 8.5, False, 'SILVER')
for i, (v, lab) in enumerate([('121', 'electric cooperatives under NEA serve the rural grid'),
                              ('1.4M', 'electricity consumers without power after Typhoon Tino, Nov 2025')]):
    y = 124 + i * 64
    box(s, 470, y, 442, 56, C['CARD'])
    box(s, 470, y, 4, 56, C['NAVY'])
    tx(s, 488, y, 96, 56, v, 24, True, 'NAVY', L, MID)
    tx(s, 590, y, 306, 56, lab, 12, False, 'INK', L, MID, line=1.25)
klab(s, 470, 262, 'What 450 MHz does')
numbered(s, 470, 282, 442, [
    'A private data path for SCADA, feeder automation and substations',
    'Push-to-talk, video and lone-worker alarms for line crews',
    'Smart meters on the same sites, reaching boxes deep indoors',
], gap=30, size=12)
for j, (lab, body, col, bg) in enumerate([
        ('PROOF POINT', 'Germany: 450connect, Band 72, more than 95% national coverage.  Malaysia: TNB, Band 31, initial deployment in four areas.', 'NAVY', 'CARD'),
        ('WHAT GLOBE SELLS', 'A capacity package for the cooperative, IoT connectivity per meter, and SCADA integration and O&M.', 'GREEN', 'GREEN_L')]):
    x = X0 + j * 438
    box(s, x, 392, 426, 96, C[bg])
    box(s, x, 392, 426, 4, C[col])
    tx(s, x + 18, 406, 390, 14, lab, 10, True, col)
    tx(s, x + 18, 426, 390, 64, body, 12, False, 'INK', line=1.3)
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
