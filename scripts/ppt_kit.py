# -*- coding: utf-8 -*-
r"""
bjfu-meetingppt 构建工具库
将组会汇报 Markdown 内容写成基于母版版式的 PPT。
配套默认模板（assets/PPT模版.pptx）：版式 0封面 / 1目录 / 2章节封面 / 3内容 / 4空白。

内容脚本模式：
    import sys, os
    SKILL = r"<skill根目录>"
    sys.path.insert(0, os.path.join(SKILL, "scripts"))
    from ppt_kit import Deck
    deck = Deck(os.path.join(SKILL, "assets", "PPT模版.pptx"))
    deck.cover("组会工作汇报", "汇报人：张佳林 ｜ 2026年11月5日")
    deck.toc(["项目一", "项目二", "其他工作"])
    deck.section(1, "项目一")
    s = deck.body("1.1 项目进展：xxx")
    s.bullets(["要点一", "要点二"], 0.45, 1.02, 12.43)
    s.picture(r"attachments\xx.png", 0.45, 2.5, 6.55, 3.47, caption="图注")
    deck.save(r"out.pptx")
"""
import os
import subprocess
import zipfile

from PIL import Image, ImageFilter
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# ---- 设计规范（取自默认模板主题）----
TEAL_CHIP = RGBColor(0x31, 0x85, 0x9C)   # 编号块/强调色（模板自带）
TITLE_CLR = RGBColor(0x46, 0x7E, 0x90)   # 标题青蓝（模板 tx2·lumMod75%）
NAVY = RGBColor(0x24, 0x38, 0x48)        # 深藏青（accent2）
BODY_CLR = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PANEL_BG = RGBColor(0xF1, 0xF6, 0xF8)    # 浅青灰面板
FONT = "微软雅黑"

# ---- 示例设计语法（references/example-design.md）----
LABEL_NAVY = RGBColor(0x1F, 0x38, 0x64)  # 二级小节标签深藏青
RED = RGBColor(0xC0, 0x00, 0x00)         # 红字强调（成果/关键数字）
TEAL_DARK = RGBColor(0x1B, 0x7A, 0x8E)   # 缎带折角阴影
BAND_BLUE = RGBColor(0xD6, 0xE2, 0xF0)   # 表格隔行淡蓝（示例网格风）
FLOW_COLORS = [RGBColor(0xF7, 0x96, 0x46), RGBColor(0x2E, 0x5F, 0xA3),
               RGBColor(0x70, 0x30, 0xA0), RGBColor(0x4C, 0xAF, 0x50),
               RGBColor(0x8E, 0xAA, 0x98)]  # 流程卡一步一色
ARROW_BLUE = RGBColor(0x44, 0x72, 0xC4)  # 流程块箭头

SW, SH = 13.333, 7.5                     # 16:9 画布（英寸）


# ---------- 底层工具 ----------

def set_ea(run, font=FONT):
    """中文字体必须同时设置 a:ea/a:cs，否则回退默认字体。"""
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", font)


def style_run(run, size=16, bold=False, color=BODY_CLR, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    set_ea(run, font)


def set_bullet(para, color="31859C"):
    """罗列式项目符号：buClr+buChar+悬挂缩进（顺序符合 DrawingML schema）。"""
    pPr = para._p.get_or_add_pPr()
    pPr.set("marL", "216000")
    pPr.set("indent", "-216000")
    for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar", "a:buNone"):
        e = pPr.find(qn(tag))
        if e is not None:
            pPr.remove(e)
    buClr = pPr.makeelement(qn("a:buClr"), {})
    srgb = pPr.makeelement(qn("a:srgbClr"), {"val": color})
    buClr.append(srgb)
    pPr.append(buClr)
    pPr.append(pPr.makeelement(qn("a:buSzPct"), {"val": "90000"}))
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))


def add_text(slide, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP, wrap=True):
    """lines: [dict(text=, size=, bold=, color=, align=, bullet=, space_after=, line_spacing=)]
    段内混色强调：spec 用 runs=[(文本, 颜色或None), ...] 代替 text（红字关键词等）。"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, spec in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        if spec.get("space_after") is not None:
            p.space_after = Pt(spec["space_after"])
        if spec.get("line_spacing") is not None:
            p.line_spacing = spec["line_spacing"]
        d_size, d_bold = spec.get("size", 16), spec.get("bold", False)
        d_color = spec.get("color", BODY_CLR)
        if "runs" in spec:
            for t, c in spec["runs"]:
                r = p.add_run()
                r.text = t
                style_run(r, d_size, d_bold, c if c is not None else d_color)
        else:
            r = p.add_run()
            r.text = spec["text"]
            style_run(r, d_size, d_bold, d_color)
        if spec.get("bullet"):
            set_bullet(p)
    return tb


def _rect(slide, x, y, w, h, fill):
    sp = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def prep_image(src, cache_dir, max_px=2200):
    """python-pptx 仅收 BMP/GIF/JPEG/PNG/TIFF/WMF；iPhone 实况照片(MPO)等需转存 JPEG，超大图降采样。"""
    with Image.open(src) as im:
        fmt = (im.format or "").upper()
        if fmt not in ("PNG", "JPEG", "BMP", "GIF", "TIFF", "WMF") or max(im.size) > max_px:
            os.makedirs(cache_dir, exist_ok=True)
            out = os.path.join(cache_dir, os.path.splitext(os.path.basename(src))[0] + ".jpg")
            im.convert("RGB").thumbnail((max_px, max_px))
            im.convert("RGB").save(out, "JPEG", quality=90)
            return out
    return src


def content_crop_169(src, cache_dir):
    """内容感知 16:9 居中裁剪（按边缘能量选窗口），用于照片组等大排布。"""
    os.makedirs(cache_dir, exist_ok=True)
    dst = os.path.join(cache_dir, "crop169_" + os.path.splitext(os.path.basename(src))[0] + ".jpg")
    if os.path.exists(dst):
        return dst
    im = Image.open(src).convert("RGB")
    w, h = im.size
    if abs(w / h - 16 / 9) < 0.01:
        im.save(dst, "JPEG", quality=92)
        return dst
    g = np.asarray(im.convert("L").filter(ImageFilter.FIND_EDGES), dtype=np.float32)
    if w / h > 16 / 9:  # 太宽，裁宽
        n = min(int(h * 16 / 9), w)
        scores = np.convolve(g.sum(axis=0), np.ones(n), mode="valid")
        x0 = int(np.argmax(scores))
        box = (x0, 0, x0 + n, h)
    else:               # 太高，裁高
        n = min(int(w * 9 / 16), h)
        scores = np.convolve(g.sum(axis=1), np.ones(n), mode="valid")
        y0 = int(np.argmax(scores))
        box = (0, y0, w, y0 + n)
    im.crop(box).save(dst, "JPEG", quality=92)
    return dst


# ---------- Deck / Slide ----------

class Slide:
    """对 add_slide 结果的薄封装，提供常用排版动作。"""

    def __init__(self, deck, slide, cache_dir):
        self.deck = deck
        self.slide = slide
        self.cache = cache_dir

    # -- 文字 --
    def text(self, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP):
        return add_text(self.slide, x, y, w, h, lines, anchor=anchor)

    def bullets(self, items, x, y, w, size=16, space_after=9, color=BODY_CLR):
        lines = [dict(text=t, size=size, bullet=True, space_after=space_after,
                      color=color) for t in items]
        return add_text(self.slide, x, y, w, 0.40 * len(items) + 0.12, lines)

    def caption(self, x, y, w, text, color=LABEL_NAVY):
        """图注：示例语法为深藏青粗体（红色用于成绩强调）。"""
        return add_text(self.slide, x, y, w, 0.35,
                        [dict(text=text, size=11, bold=True, color=color,
                              align=PP_ALIGN.CENTER)])

    # -- 图片 --
    def picture(self, src, fx, fy, fw, fh, border=False, caption=None):
        """在 (fx,fy,fw,fh) 框内等比放置并居中。同组图片给相同框即等大外框。
        图注贴实际图片下缘居中（框大于图时不会脱节）。"""
        src = prep_image(src, self.cache)
        with Image.open(src) as im:
            iw, ih = im.size
        scale = min(fw / (iw / 96), fh / (ih / 96))
        w, h = (iw / 96) * scale, (ih / 96) * scale
        pic = self.slide.shapes.add_picture(src, Inches(fx + (fw - w) / 2),
                                            Inches(fy + (fh - h) / 2),
                                            Inches(w), Inches(h))
        if border:
            pic.line.color.rgb = RGBColor(0xD5, 0xDD, 0xE1)
            pic.line.width = Pt(0.75)
        if caption:
            self.caption(fx, Emu(pic.top + pic.height).inches + 0.10, fw, caption)
        return pic

    def photo_grid(self, sources, cols, x, y, cw, ch, gap=0.25, crop_169=True):
        """照片网格。crop_169=True 时先统一 16:9 裁剪，使每张精确等大、边缘对齐。"""
        import math
        n = len(sources)
        rows = math.ceil(n / cols)
        for i, src in enumerate(sources):
            gx, gy = x + (cw + gap) * (i % cols), y + (ch + gap) * (i // cols)
            if crop_169:
                src = content_crop_169(src, self.cache)
            self.picture(src, gx, gy, cw, ch, border=True)
        return rows

    # -- 示例设计组件 --
    def pic_notes(self, items, x, y, w, row_h, **kw):
        """示例式图片拼贴（每行等高、自然比例、每图必注）。kw 同模块级 pic_notes。"""
        kw.setdefault("cache_dir", self.cache)
        return pic_notes(self.slide, items, x, y, w, row_h, **kw)

    def flow_cards(self, steps, x, y, w, h, **kw):
        """流程步骤卡（彩色标题条+白底细边卡+块箭头）。kw 同模块级 flow_cards。"""
        kw.setdefault("cache_dir", self.cache)
        return flow_cards(self.slide, steps, x, y, w, h, **kw)

    def frame(self, x, y, w, h, color=TEAL_CHIP, width=1.0):
        """细线容器框：大块图文区整体框住。"""
        return frame_box(self.slide, x, y, w, h, color, width)

    def footer_note(self, runs, y=6.88, x=0.45, w=12.43, size=13):
        """底部成果/总结条（支持红字关键词混排）。"""
        return footer_note(self.slide, runs, y=y, x=x, w=w, size=size)

    # -- 面板罗列 / 卡片 --
    def panel_list(self, items, x=0.45, y0=1.15, w=12.43, h=0.82, dy=1.02, size=16):
        for i, item in enumerate(items):
            y = y0 + i * dy
            _rect(self.slide, x, y, w, h, PANEL_BG)
            _rect(self.slide, x, y, 0.09, h, TEAL_CHIP)
            add_text(self.slide, x + 0.35, y, w - 0.6, h,
                     [dict(text=item, size=size)], anchor=MSO_ANCHOR.MIDDLE)

    def cards(self, items, x0=0.45, y=2.05, w=6.01, h=3.30, gap=0.42, num_size=40, txt_size=19):
        for i, txt in enumerate(items):
            cx = x0 + i * (w + gap)
            _rect(self.slide, cx, y, w, h, PANEL_BG)
            _rect(self.slide, cx, y, w, 0.12, TEAL_CHIP)
            add_text(self.slide, cx + 0.45, y + 0.70, 1.6, 1.0,
                     [dict(text=f"{i+1:02d}", size=num_size, bold=True, color=TEAL_CHIP)])
            add_text(self.slide, cx + 0.45, y + 2.10, w - 0.9, 1.2,
                     [dict(text=txt, size=txt_size, bold=True, color=NAVY, line_spacing=1.3)])

    # -- 表格 --
    def table(self, data, x, y, w, col_ws, header=True, row_h=0.62, header_h=0.44,
              size=12, header_size=13, style="teal"):
        """data[0] 为表头行；col_ws 为各列英寸宽列表。
        style: "teal"=青底白字表头+浅青灰隔行（默认）；"grid"=示例网格风
        （白底黑粗体表头、隔行淡蓝 #D6E2F0，边框由表格默认样式提供）。"""
        grid = (style == "grid")
        rows, cols = len(data), len(data[0])
        gfx = self.slide.shapes.add_table(rows, cols, Inches(x), Inches(y),
                                          Inches(w), Inches(header_h + row_h * (rows - 1)))
        tbl = gfx.table
        for ci, cw in enumerate(col_ws):
            tbl.columns[ci].width = Inches(cw)
        for ri in range(rows):
            tbl.rows[ri].height = Inches(header_h if (header and ri == 0) else row_h)
            for ci in range(cols):
                cell = tbl.cell(ri, ci)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.02)
                cell.fill.solid()
                if header and ri == 0:
                    cell.fill.fore_color.rgb = WHITE if grid else TEAL_CHIP
                else:
                    cell.fill.fore_color.rgb = (WHITE if ri % 2 else
                                                (BAND_BLUE if grid else PANEL_BG))
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                r = p.add_run()
                r.text = data[ri][ci]
                if ri == 0:
                    style_run(r, header_size, True, BODY_CLR if grid else WHITE)
                else:
                    style_run(r, size, grid, BODY_CLR)
        return tbl


class Deck:
    """默认模板五版式封装。换模板时改 layout 索引与占位符 idx（见 references/template-map.md）。"""

    def __init__(self, template, work_dir="_ppt_build"):
        self.prs = Presentation(template)
        self.master = self.prs.slide_masters[0]
        layouts = self.master.slide_layouts
        if len(layouts) < 5:
            raise ValueError("模板版式不足 5 个，请先做版式勘察（references/template-map.md）")
        self.LAY_COVER, self.LAY_TOC, self.LAY_SECT, self.LAY_BODY, self.LAY_BLANK = \
            [layouts[i] for i in range(5)]
        self.cache = os.path.join(work_dir, "assets")
        self._clear()

    def _clear(self):
        """删除模板示例页；必须同时 drop_rel，否则保存时产生重复 zip 条目。"""
        lst = self.prs.slides._sldIdLst
        for sld in list(lst):
            rId = sld.get(qn("r:id"))
            if rId:
                self.prs.part.drop_rel(rId)
            lst.remove(sld)

    def _new(self, layout):
        return Slide(self, self.prs.slides.add_slide(layout), self.cache)

    @staticmethod
    def _ph(slide, idx):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        return None

    @staticmethod
    def _fill_ph(ph, text, size, bold, color, align=PP_ALIGN.LEFT):
        tf = ph.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        style_run(r, size, bold, color)

    # ---- 页面 ----
    def cover(self, title, subtitle, title_size=44):
        """封面/结束页：中部青蓝横带内大标题 + 白色副标题。"""
        s = self._new(self.LAY_COVER)
        ph = self._ph(s.slide, 10)
        self._fill_ph(ph, title, title_size, True, WHITE, PP_ALIGN.CENTER)
        if subtitle:
            add_text(s.slide, 0.5, 3.80, SW - 1.0, 0.45,
                     [dict(text=subtitle, size=18, color=WHITE, align=PP_ALIGN.CENTER)])
        return s

    def toc(self, items, size=16):
        """目录：复用模板编号块；标题居中放在编号右侧胶囊条上（行位取自版式）。"""
        s = self._new(self.LAY_TOC)
        rows = []
        for sh in self.LAY_TOC.shapes:
            if sh.name.startswith("Google Shape;397") and sh.width and sh.width > Emu(1000000):
                t = sh.text_frame.text.strip()
                if t in [str(i) for i in range(1, 10)]:
                    rows.append((int(t), sh.top, sh.height))
        rows.sort()
        for i, item in enumerate(items):
            if i < len(rows):
                num, top, hgt = rows[i]
            else:
                num, top, hgt = i + 1, rows[-1][1] + (rows[-1][1] - rows[-2][1]) * (i - len(rows) + 1), rows[-1][2]
            if i == 0:
                box = self._ph(s.slide, 10)
                box.left, box.top = Inches(7.95), top
                box.width, box.height = Inches(4.25), hgt
                self._fill_ph(box, item, size, True, NAVY, PP_ALIGN.CENTER)
                box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            else:
                add_text(s.slide, 7.95, top / 914400, 4.25, hgt / 914400,
                         [dict(text=item, size=size, bold=True, color=NAVY,
                               align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
        return s

    def section(self, num, title, size=28):
        """章节封面：左侧深色标题 + 版式自带超大数字。"""
        s = self._new(self.LAY_SECT)
        ph_n = self._ph(s.slide, 11)
        if ph_n is not None:
            ph_n.text_frame.text = str(num)   # 继承版式大字号与配色
        ph_t = self._ph(s.slide, 10)
        ph_t.left, ph_t.top = Inches(1.35), Inches(3.90)
        ph_t.width, ph_t.height = Inches(6.20), Inches(0.95)
        self._fill_ph(ph_t, title, size, True, NAVY)
        return s

    def body(self, chapter, section=None):
        """内容页（示例双层标题语法）：一级青蓝缎带（章节名，每页重复强化语境）
        + 二级深藏青小节标签（如"1.1 项目进展"）。
        只传一个参数时缎带显示该文本、无小节标签（兼容旧用法）。
        内容区从返回值 s.content_y（约 1.45）起，底至 6.85，页码自动。"""
        s = self._new(self.LAY_BODY)
        s.content_y = ribbon_title(s.slide, chapter, section)
        return s

    def save(self, path):
        self.prs.save(path)
        # 保存后完整性自检
        z = zipfile.ZipFile(path)
        assert z.testzip() is None, "zip 损坏"
        names = z.namelist()
        assert len(names) == len(set(names)), "存在重复部件（清页未 drop_rel?）"
        return path


# ---------- 示例设计组件（references/example-design.md）----------

def ribbon_title(slide, chapter, section=None, y=0.27):
    """一级青蓝缎带标题（右端斜切折角）+ 可选二级深藏青小节标签。
    返回内容区起始 y（约 1.45）。"""
    n = max(len(chapter), 2)
    w = min(max(0.30 * n + 0.35, 1.6), 5.4)
    x0, h, cut = 0.13, 0.50, 0.12
    # 缎带主体：右下角斜切的多边形
    pts = [(x0 + w, y), (x0 + w, y + h - cut), (x0 + w - cut, y + h), (x0, y + h)]
    fb = slide.shapes.build_freeform(Inches(x0), Inches(y))
    fb.add_line_segments([(Inches(px), Inches(py)) for px, py in pts], close=True)
    sp = fb.convert_to_shape()
    sp.fill.solid(); sp.fill.fore_color.rgb = TEAL_CHIP
    sp.line.fill.background(); sp.shadow.inherit = False
    # 折角阴影小三角
    fb2 = slide.shapes.build_freeform(Inches(x0 + w - cut), Inches(y + h))
    fb2.add_line_segments([(Inches(x0 + w), Inches(y + h)),
                           (Inches(x0 + w), Inches(y + h - cut))], close=True)
    tri = fb2.convert_to_shape()
    tri.fill.solid(); tri.fill.fore_color.rgb = TEAL_DARK
    tri.line.fill.background(); tri.shadow.inherit = False
    # 缎带白字
    add_text(slide, x0 + 0.15, y, w - 0.22, h,
             [dict(text=chapter, size=18, bold=True, color=WHITE)],
             anchor=MSO_ANCHOR.MIDDLE)
    # 缎带下方小三角点缀（示例原样）
    dot = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                 Inches(0.13), Inches(y + h + 0.02), Inches(0.06), Inches(0.06))
    dot.fill.solid(); dot.fill.fore_color.rgb = TEAL_CHIP
    dot.line.fill.background(); dot.shadow.inherit = False
    if section:
        add_text(slide, 0.34, y + h + 0.02, 12.5, 0.50,
                 [dict(text=section, size=16, bold=True, color=LABEL_NAVY)])
    return y + h + 0.60


def frame_box(slide, x, y, w, h, color=TEAL_CHIP, width=1.0):
    """细线容器框：把大块图文区整体框住（示例 p12/p13/p18 语法）。"""
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.background()
    sp.line.color.rgb = color
    sp.line.width = Pt(width)
    sp.shadow.inherit = False
    return sp


def pic_notes(slide, items, x, y, w, row_h, cap_h=0.34, gap=0.06, vgap=0.12,
              cap_color=LABEL_NAVY, cache_dir="_ppt_build/assets", max_bottom=6.85):
    """示例式图片拼贴：每行等高、宽度随自然比例、行内紧排顶满 w；每图配粗体图注。
    items: [图路径] 或 [(图路径, 图注)] 或 [(图路径, 图注, 图注颜色)]。
    与 photo_grid（等大 16:9 网格）互补：整齐划一用网格，饱满现场感用拼贴。
    行高自动收缩，保证最后一行图注不超过 max_bottom。"""
    norm = []
    for it in items:
        if isinstance(it, str):
            norm.append((it, None, cap_color))
        elif len(it) == 2:
            norm.append((it[0], it[1], cap_color))
        else:
            norm.append((it[0], it[1], it[2]))
    prep = []
    for path, cap, cc in norm:
        p = prep_image(path, cache_dir)
        with Image.open(p) as im:
            prep.append((p, cap, cc, im.width / im.height))  # 宽高比
    # 按初始 row_h 打包行
    rows_idx, cur, cur_w = [], [], 0.0
    for i, (_, _, _, ratio) in enumerate(prep):
        iw = row_h * ratio
        if cur and cur_w + iw + gap * len(cur) > w + 0.02:
            rows_idx.append(cur)
            cur, cur_w = [], 0.0
        cur.append(i)
        cur_w += iw
    if cur:
        rows_idx.append(cur)
    n = len(rows_idx)
    avail = max_bottom - y - n * cap_h - (n - 1) * vgap
    if n and avail / n < row_h:
        row_h = avail / n
    row_y = y
    for r in rows_idx:
        row = [(prep[i][0], prep[i][1], row_h * prep[i][3], prep[i][2]) for i in r]
        _flush_row(slide, row, x, w, row_y, row_h, cap_h, gap, cache_dir)
        row_y += row_h + cap_h + vgap
    return row_y


def _flush_row(slide, row, x, w, row_y, row_h, cap_h, gap, cache_dir):
    total = sum(r[2] for r in row) + gap * (len(row) - 1)
    cx = x + max((w - total) / 2, 0)
    for p, cap, iw, cc in row:
        slide.shapes.add_picture(p, Inches(cx), Inches(row_y), Inches(iw), Inches(row_h))
        if cap:
            add_text(slide, cx, row_y + row_h + 0.03, iw, cap_h,
                     [dict(text=cap, size=10.5, bold=True, color=cc,
                           align=PP_ALIGN.CENTER, line_spacing=1.0)])
        cx += iw + gap


def flow_cards(slide, steps, x, y, w, h, gap=0.60, header_h=0.32, cache_dir="_ppt_build/assets"):
    """流程步骤卡（示例 p12 语法）：彩色标题条 + 白底细边主体卡 + 块箭头串联。
    steps: [dict(title=, pic=图路径|None, text=|None, arrow=箭头标注|None)]。
    箭头标注会折行（示例亦如此），保持 3–6 字为宜。"""
    n = len(steps)
    card_w = (w - gap * (n - 1)) / n
    for i, st in enumerate(steps):
        cx = x + i * (card_w + gap)
        color = FLOW_COLORS[i % len(FLOW_COLORS)]
        hd = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(cx), Inches(y), Inches(card_w), Inches(header_h))
        hd.fill.solid(); hd.fill.fore_color.rgb = color
        hd.line.fill.background(); hd.shadow.inherit = False
        add_text(slide, cx + 0.06, y, card_w - 0.12, header_h,
                 [dict(text=st["title"], size=12.5, bold=True, color=WHITE,
                       align=PP_ALIGN.CENTER, line_spacing=1.0)], anchor=MSO_ANCHOR.MIDDLE)
        body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(cx), Inches(y + header_h + 0.05),
                                      Inches(card_w), Inches(h - header_h - 0.05))
        body.fill.solid(); body.fill.fore_color.rgb = WHITE
        body.line.color.rgb = color; body.line.width = Pt(1.0)
        body.shadow.inherit = False
        pad = 0.10
        if st.get("pic"):
            p = prep_image(st["pic"], cache_dir)
            avail_h = (h - header_h - 0.05) - (0.45 if st.get("text") else 0) - 2 * pad
            avail_w = card_w - 2 * pad
            ratio = _ratio(p)  # h/w
            pw, ph = avail_w, avail_w * ratio
            if ph > avail_h:
                pw, ph = avail_h / ratio, avail_h
            slide.shapes.add_picture(p, Inches(cx + (card_w - pw) / 2),
                                     Inches(y + header_h + 0.05 + pad), Inches(pw), Inches(ph))
        if st.get("text"):
            if st.get("pic"):
                add_text(slide, cx + pad, y + h - 0.42, card_w - 2 * pad, 0.36,
                         [dict(text=st["text"], size=10.5, color=BODY_CLR,
                               align=PP_ALIGN.CENTER, line_spacing=1.05)])
            else:  # 纯文字卡：在卡体内垂直居中，避免卡体大面积空白
                add_text(slide, cx + pad, y + header_h + 0.05,
                         card_w - 2 * pad, h - header_h - 0.05,
                         [dict(text=st["text"], size=11.5, bold=True, color=BODY_CLR,
                               align=PP_ALIGN.CENTER, line_spacing=1.2)],
                         anchor=MSO_ANCHOR.MIDDLE)
        if i < n - 1:
            ax = cx + card_w + 0.07
            ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(ax),
                                        Inches(y + h / 2 - 0.10), Inches(gap - 0.14), Inches(0.20))
            ar.fill.solid(); ar.fill.fore_color.rgb = ARROW_BLUE
            ar.line.fill.background(); ar.shadow.inherit = False
            if st.get("arrow"):
                add_text(slide, ax - 0.05, y + h / 2 - 0.50, gap - 0.04, 0.38,
                         [dict(text=st["arrow"], size=9.5, bold=True, color=BODY_CLR,
                               align=PP_ALIGN.CENTER, line_spacing=1.0)])


def _ratio(path):
    with Image.open(path) as im:
        return im.height / im.width


def footer_note(slide, runs, y=6.88, x=0.45, w=12.43, size=13):
    """底部成果/总结条（示例仪式化收尾）：全宽粗体一行，支持红字关键词。
    runs: 纯文本 str 或 [(文本, 颜色或None), ...]。"""
    if isinstance(runs, str):
        spec = dict(text=runs, size=size, bold=True, color=BODY_CLR, line_spacing=1.1)
    else:
        spec = dict(runs=runs, size=size, bold=True, color=BODY_CLR, line_spacing=1.1)
    return add_text(slide, x, y, w, 0.5, [spec])


def collect_md_images(md_path):
    """以 md 中的实际图片链接为索引定位图片文件（附件夹名不固定，勿假设 attachments/）。

    支持 ![[x.png]]（Obsidian，可带 |尺寸 参数）与 ![](path/x.png)（标准 md）两种引用。
    定位顺序：链接内路径（相对 md 所在目录或绝对）→ 按 basename 在 md 目录递归匹配
    （跳过隐藏目录与 _ppt_build）。返回 (引用列表, 路径列表)，路径与 md 中出现顺序一致，
    找不到的为 None。
    """
    import re
    md_path = os.path.abspath(md_path)
    base = os.path.dirname(md_path)
    text = open(md_path, encoding="utf-8").read()
    pat = re.compile(r"!\[\[([^\]|]+?)(?:\|[^\]]*)?\]\]|!\[[^\]]*\]\(([^)]+)\)")
    refs = []
    for m in pat.finditer(text):
        name = (m.group(1) or m.group(2) or "").strip()
        if name:
            refs.append(name)
    index = {}
    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d != "_ppt_build"]
        for f in files:
            index.setdefault(f, os.path.join(root, f))
    paths = []
    for name in refs:
        p = name if os.path.isabs(name) else os.path.join(base, name)
        paths.append(os.path.abspath(p) if os.path.exists(p) else index.get(os.path.basename(name)))
    return refs, paths


def export_pngs(pptx_path, out_dir, width=1280, height=720):
    """Windows + PowerPoint COM 导出每页 PNG（质检用）。"""
    os.makedirs(out_dir, exist_ok=True)
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    ps = (
        "$pp = New-Object -ComObject PowerPoint.Application;"
        "$pres = $pp.Presentations.Open('%s', $true, $false, $false);"
        "$i = 1;"
        "foreach ($s in $pres.Slides) {"
        "  $s.Export(('%s\\' + $i + '.png'), 'PNG', %d, %d); $i++ };"
        "$pres.Close(); $pp.Quit(); Write-Output ('exported ' + ($i - 1))"
    ) % (pptx_path, out_dir, width, height)
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    return (r.stdout or "") + (r.stderr or "")
