"""スライド生成の共通部品（図形・文字・はみ出し判定）。

build_pptx.py と build_template.py の両方から使う。
座標と大きさはすべてインチ（16:9 = 13.333 x 7.5 インチ）。
"""
import copy
import json
import re
import sys
import unicodedata
from pathlib import Path

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SKILL_DIR = Path(__file__).resolve().parent.parent
STYLES_DIR = SKILL_DIR / "styles"

SLIDE_W = 13.333
SLIDE_H = 7.5

FONT_JA = "游ゴシック"
FONT_EN = "Arial"
# 文字の高さに対する1行の高さ。游ゴシックは行の高さが大きいので余裕を見ておく
LINE_HEIGHT = 1.38

# 文字の大きさの下限（digital-agency-design-system の Office 基準 S9：本文 18pt 以上、補助の文字も 12pt 以上）。
# 本文の種類（BODY_LABELS）は 18pt 未満にしない。それ以外の文字も 12pt 未満にしない。
# 枠に入りきらないときも、下限より小さくは縮めず「文字が多すぎて枠に収まりません」の警告を出す（スライドを分けるか文字を減らす）。
MIN_PT = 12
BODY_PT = 18
BODY_LABELS = {
    "カードの本文", "カードの箇条書き", "要点の補足", "補足の箇条書き", "比較の本文", "比較の箇条書き",
    "ステップの本文", "ステップの箇条書き", "根拠の本文", "根拠の箇条書き", "回答", "引用",
    "読み取りのポイント", "まとめの一文", "リード文", "サブタイトル", "図の要点",
}


def floor_pt(label=""):
    """その文字の種類の、大きさの下限。"""
    return BODY_PT if label in BODY_LABELS else MIN_PT

ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


# ---------------------------------------------------------------- 実行環境

def console_utf8():
    """日本語版 Windows でも表示が文字化けしないよう、出力を UTF-8 にする。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def as_text(v):
    """None は空文字、数字などは文字列にする（deck.json の書き間違いで止まらないように）。"""
    if v is None:
        return ""
    return v if isinstance(v, str) else str(v)


# ---------------------------------------------------------------- 設定の読み込み

def load_theme(style):
    """styles/<style>/theme.json を読み込む。"""
    path = STYLES_DIR / style / "theme.json"
    if not path.exists():
        raise SystemExit(f"テーマが見つかりません: {path}")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def rgb(hex_str):
    return RGBColor.from_string(hex_str.lstrip("#").upper())


# ---------------------------------------------------------------- 警告

class Warnings:
    """はみ出しや文字数超過の警告を集める。"""

    def __init__(self):
        self.items = []
        self.where = ""

    def add(self, msg):
        self.items.append(f"{self.where}{msg}")


WARN = Warnings()


# ---------------------------------------------------------------- 文字幅の見積もり

def char_em(ch, bold=False):
    """1文字の幅（em 単位）を見積もる。全角=1.0、半角英数≒0.6。"""
    if ch == "\t":
        return 2.0
    ea = unicodedata.east_asian_width(ch)
    if ea in ("F", "W", "A"):
        return 1.0
    if ch == " ":
        return 0.32
    if ch in "il.,:;|!'`":
        return 0.32
    if ch in "mwMW@%":
        return 0.9
    if ch.isupper() or ch.isdigit():
        return 0.66 if bold else 0.62
    return 0.58 if bold else 0.55


def text_em(text, bold=False, spacing_em=0.0):
    return sum(char_em(c, bold) + spacing_em for c in text)


def count_lines(text, width_in, size_pt, bold=False, spacing_em=0.0):
    """幅 width_in に size_pt で流し込んだときの行数を見積もる。"""
    if width_in <= 0:
        return 999
    width_em = width_in * 72.0 / size_pt
    lines = 0
    for para in text.split("\n"):
        if not para:
            lines += 1
            continue
        cur = 0.0
        n = 1
        for ch in para:
            w = char_em(ch, bold) + spacing_em
            if cur + w > width_em and cur > 0:
                # 行頭禁則（句読点・閉じ括弧は前の行に残す）は無視して単純に折り返す
                n += 1
                cur = w
            else:
                cur += w
        lines += n
    return lines


def text_height(text, width_in, size_pt, bold=False, line_spacing=1.0, spacing_em=0.0):
    n = count_lines(text, width_in, size_pt, bold, spacing_em)
    return n * size_pt * LINE_HEIGHT * line_spacing / 72.0


def fit_size(text, width_in, height_in, max_pt, min_pt, bold=False, line_spacing=1.0,
             spacing_em=0.0, label=""):
    """箱に収まる最大の文字サイズを返す。最小でも収まらなければ警告する。"""
    size = max_pt
    while size >= min_pt:
        if text_height(text, width_in, size, bold, line_spacing, spacing_em) <= height_in:
            return size
        size -= 0.5
    WARN.add(f"文字が多すぎて枠に収まりません（{label or text[:16]}…）。文を短くしてください")
    return min_pt


def max_chars_warn(text, limit, label):
    # 表示するときに取る空白（日本語と数字の間など）は数えない
    n = len(tidy_ja(as_text(text)).replace("\n", "").replace("**", ""))
    if n > limit:
        WARN.add(f"{label}が {n} 文字です（目安 {limit} 文字以内）")


# ---------------------------------------------------------------- 文字の装飾

EMPH_RE = re.compile(r"\*\*(.+?)\*\*")


def split_emphasis(text):
    """'**強調**' を [(文字, 強調か)] に分ける。"""
    parts = []
    pos = 0
    for m in EMPH_RE.finditer(text):
        if m.start() > pos:
            parts.append((text[pos:m.start()], False))
        parts.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        parts.append((text[pos:], False))
    return parts or [("", False)]


_CJK = r"\u3000-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uff00-\uffef"
_SP1 = re.compile(rf"(?<=[{_CJK}]) (?=[0-9A-Za-z%＋+\-−])")
_SP2 = re.compile(rf"(?<=[0-9A-Za-z%]) (?=[{_CJK}])")


_SP3 = re.compile(rf"(?<=[{_CJK}]) +(?=\*\*)|(?<=\*\*) +(?=[{_CJK}])")


def tidy_ja(text):
    """日本語と半角英数字の間の半角スペースを取る（「3,240 件」→「3,240件」）。
    強調（**…**）の前後で日本語と接するスペースも取る。"""
    return _SP2.sub("", _SP1.sub("", _SP3.sub("", text)))


def plain(text):
    return EMPH_RE.sub(r"\1", as_text(text))


def style_run(run, size=None, bold=None, color=None, font=None, latin=None, italic=None,
              spacing=None):
    f = run.font
    if size is not None:
        f.size = Pt(max(size, MIN_PT))
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color is not None:
        f.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    ja = font or FONT_JA
    lat = latin or ja
    for tag, face in (("a:latin", lat), ("a:ea", ja), ("a:cs", lat)):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", face)
    # a:latin / a:ea / a:cs は solidFill より後ろに置く必要がある
    fill = rPr.find(qn("a:solidFill"))
    if fill is not None:
        rPr.remove(fill)
        rPr.insert(0, fill)
    if spacing is not None:
        rPr.set("spc", str(int(spacing * 100)))
    rPr.set("lang", "ja-JP")
    rPr.set("altLang", "en-US")


def set_paragraph(p, align="l", line_spacing=None, space_before=None, space_after=None):
    p.alignment = ALIGN[align]
    if line_spacing is not None:
        p.line_spacing = line_spacing
    if space_before is not None:
        p.space_before = Pt(space_before)
    if space_after is not None:
        p.space_after = Pt(space_after)


# 強調の帯: (強調の文字色, 帯の色)。デザイン（style）が決める。強調語がこの文字色で書かれるとき、
# 文字の下に薄い色の太い下線（帯）を付ける。下線なので、折り返した文や後から直した文でもずれない。
EMPH_BAND = None


def underline_band(run, color):
    """文字に、薄い色の太い下線（強調の帯）を付ける。"""
    rPr = run._r.get_or_add_rPr()
    rPr.set("u", "heavy")
    for tag in ("a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill"):
        el = rPr.find(qn(tag))
        if el is not None:
            rPr.remove(el)
    uf = etree.Element(qn("a:uFill"))
    sf = etree.SubElement(uf, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val=color.lstrip("#").upper())
    # 並び順の決まり（塗り → 下線 → 書体）に合わせ、書体（latin など）の前に置く
    anchor = next((rPr.find(qn(t)) for t in ("a:latin", "a:ea", "a:cs", "a:sym")
                   if rPr.find(qn(t)) is not None), None)
    if anchor is not None:
        anchor.addprevious(uf)
    else:
        rPr.append(uf)


def write_runs(p, text, size, color, bold=False, emph_color=None, emph_bold=True, font=None,
               latin=None, spacing=None, italic=None, tidy=True, underline=True):
    """段落に文字を書く。'**強調**' は emph_color（なければ太字のみ）で表示。
    強調の文字色が EMPH_BAND の色なら、文字の下に薄い色の帯（太い下線）も付ける
    （underline=False のときは付けない。見出しは図形の帯 emphasis_bands を使う）。"""
    text = as_text(text)
    if tidy:
        text = tidy_ja(text)
    for chunk, emph in split_emphasis(text):
        if not chunk:
            continue
        r = p.add_run()
        r.text = chunk
        style_run(
            r,
            size=size,
            bold=(emph_bold if emph else bold),
            color=(emph_color or color) if emph else color,
            font=font,
            latin=latin,
            spacing=spacing,
            italic=italic,
        )
        ec = (emph_color or "").lstrip("#").upper()
        if underline and emph and EMPH_BAND and ec and ec == EMPH_BAND[0].lstrip("#").upper():
            underline_band(r, EMPH_BAND[1])


# ---------------------------------------------------------------- 図形

def _no_shadow(shape):
    spPr = shape._element.spPr
    if spPr.find(qn("a:effectLst")) is None:
        etree.SubElement(spPr, qn("a:effectLst"))


def add_box(shapes, x, y, w, h, fill=None, line=None, line_w=0.75, geom="rect", radius=None,
            shadow=False, dash=None, rotation=0, adjustments=None, name=None):
    """塗り・線つきの図形を置く。radius は角丸の半径（インチ）。"""
    shape_type = {
        "rect": MSO_SHAPE.RECTANGLE,
        "round": MSO_SHAPE.ROUNDED_RECTANGLE,
        "round_top": MSO_SHAPE.ROUND_2_SAME_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "pill": MSO_SHAPE.ROUNDED_RECTANGLE,
        "chevron": MSO_SHAPE.CHEVRON,
        "pentagon": MSO_SHAPE.PENTAGON,
        "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
        "rtriangle": MSO_SHAPE.RIGHT_TRIANGLE,
        "arc": MSO_SHAPE.BLOCK_ARC,
        "wave": MSO_SHAPE.WAVE,
        "diamond": MSO_SHAPE.DIAMOND,
        "callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
        "parallelogram": MSO_SHAPE.PARALLELOGRAM,
        "star": MSO_SHAPE.STAR_6_POINT,
        "plus": MSO_SHAPE.CROSS,
        "donut": MSO_SHAPE.DONUT,
        "tear": MSO_SHAPE.TEAR,
        "moon": MSO_SHAPE.MOON,
        "heart": MSO_SHAPE.HEART,
    }[geom]
    shp = shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shp.name = name
    if geom in ("round", "round_top") and radius is not None:
        shp.adjustments[0] = max(0.0, min(0.5, radius / max(min(w, h), 0.01)))
    if geom == "pill":
        shp.adjustments[0] = 0.5
    if adjustments:
        for i, v in enumerate(adjustments):
            shp.adjustments[i] = v
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
        if dash:
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            shp.line.dash_style = {"dash": MSO_LINE_DASH_STYLE.DASH,
                                   "dot": MSO_LINE_DASH_STYLE.ROUND_DOT}[dash]
    else:
        shp.line.fill.background()
    if rotation:
        shp.rotation = rotation
    if shadow:
        set_soft_shadow(shp)
    else:
        _no_shadow(shp)
    # 図形の中の文字の既定値
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    return shp


def set_soft_shadow(shape, alpha=14, blur=0.06, dist=0.02, color="1F2937"):
    spPr = shape._element.spPr
    old = spPr.find(qn("a:effectLst"))
    if old is not None:
        spPr.remove(old)
    eff = etree.SubElement(spPr, qn("a:effectLst"))
    sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad=str(int(Emu(Inches(blur)))),
                          dist=str(int(Emu(Inches(dist)))), dir="5400000", algn="t",
                          rotWithShape="0")
    clr = etree.SubElement(sh, qn("a:srgbClr"), val=color)
    etree.SubElement(clr, qn("a:alpha"), val=str(alpha * 1000))


def add_line(shapes, x1, y1, x2, y2, color, width=1.0, dash=None):
    ln = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    if dash:
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        ln.line.dash_style = {"dash": MSO_LINE_DASH_STYLE.DASH,
                              "dot": MSO_LINE_DASH_STYLE.ROUND_DOT}[dash]
    return ln


def add_text(shapes, x, y, w, h, text, size=14, color="000000", bold=False, align="l",
             anchor="t", font=None, latin=None, line_spacing=1.0, spacing=None,
             emph_color=None, fit_min=None, label="", italic=None, para_space=0,
             margin=0.0, shape=None, wrap=True, bullet=None, bullet_color=None, tidy=True,
             bands=False):
    """文字を置く。text は文字列（改行で段落）か段落のリスト。

    fit_min を指定すると、枠に収まるよう size から fit_min まで縮める。
    shape を渡すとその図形の中に書く（新しい箱は作らない）。
    """
    paras = [as_text(t) for t in text] if isinstance(text, list) else as_text(text).split("\n")
    raw = "\n".join(plain(t) for t in paras)
    if shape is not None:
        # 既存の図形の中に書くときは、その図形の大きさで判定する
        w = shape.width / 914400
        h = shape.height / 914400
    inner_w = w - 2 * margin - (0.22 if bullet else 0)
    inner_h = h - 2 * margin
    floor = floor_pt(label)
    size = max(size, floor)
    if fit_min is not None:
        fit_min = min(size, max(fit_min, floor))
    if fit_min is not None:
        size = fit_size(raw, inner_w, inner_h, size, fit_min, bold, line_spacing,
                        (spacing / size) if spacing else 0.0, label)
    elif text_height(raw, inner_w, size, bold, line_spacing) > inner_h + 0.02 and label:
        WARN.add(f"「{label}」が枠からはみ出す可能性があります")
    if shape is None:
        shape = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        if label:
            shape.name = label          # 読み上げ順を決めるときの手がかり（order_for_reading）
    tf = shape.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = ANCHOR[anchor]
    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    for child in list(bodyPr):
        bodyPr.remove(child)
    for i, t in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        set_paragraph(p, align, line_spacing, space_before=para_space if i else 0)
        if bullet:
            _set_bullet(p, bullet, bullet_color or color, size)
        write_runs(p, t, size, color, bold=bold, emph_color=emph_color, font=font,
                   latin=latin, spacing=spacing, italic=italic, tidy=tidy, underline=not bands)
    if bands and not bullet:
        # 見出しまわりの強調は、太い下線ではなく後ろの図形の帯（emphasis_bands）
        emphasis_bands(shapes, "\n".join(paras), shape.left / 914400 + margin,
                       shape.top / 914400 + margin, inner_w, inner_h, size, bold, align, anchor,
                       emph_color, line_spacing, tidy=tidy)
    return shape


def _set_bullet(p, char, color, size):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Emu(Inches(0.22)))))
    pPr.set("indent", str(int(-Emu(Inches(0.22)))))
    for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar", "a:buNone"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr"), val=color.lstrip("#").upper())
    etree.SubElement(pPr, qn("a:buSzPct"), val="100000")
    etree.SubElement(pPr, qn("a:buFont"), typeface="Arial")
    etree.SubElement(pPr, qn("a:buChar"), char=char)


def fill_placeholder(ph, text, size, color, bold=True, align="l", anchor=None, font=None,
                     latin=None, emph_color=None, fit_min=None, w=None, h=None, label="",
                     line_spacing=1.0, spacing=None, bands=False):
    """プレースホルダー（タイトル等）に文字を入れる。位置はレイアウトに従う。
    bands=True なら、強調語の後ろに図形の帯（emphasis_bands）を置く（太い下線は付けない）。"""
    w = w if w is not None else ph.width / 914400
    h = h if h is not None else ph.height / 914400
    text = as_text(text)
    raw = plain(text)
    floor = floor_pt(label)
    size = max(size, floor)
    if fit_min is not None:
        fit_min = min(size, max(fit_min, floor))
    if fit_min is not None:
        size = fit_size(raw, w - 0.2, h - 0.1, size, fit_min, bold, line_spacing, 0.0,
                        label or raw[:16])
    tf = ph.text_frame
    tf.clear()
    if anchor:
        tf.vertical_anchor = ANCHOR[anchor]
    lines = text.split("\n")
    for i, t in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        set_paragraph(p, align, line_spacing)
        write_runs(p, t, size, color, bold=bold, emph_color=emph_color, font=font, latin=latin,
                   spacing=spacing, underline=not bands)
    if bands:
        lins = _inset(ph, "lIns", 91440)
        tins = _inset(ph, "tIns", 45720)
        emphasis_bands(ph._parent, text, ph.left / 914400 + lins,
                       ph.top / 914400 + tins, w - 2 * lins, h - 2 * tins, size, bold, align,
                       anchor or "m", emph_color, line_spacing)
    return size


# 見出しの強調の帯（図形）: 強調語の後ろに置く長方形。高さは文字の 0.36 倍、上端は行の縦の中心のすぐ上（0.02 倍上）
# （文字の下半分に重なる。游ゴシックの PowerPoint の画面で確かめた値）。行の折り返しは文字幅の見積もり（char_em）でたどるので、2 行の見出しにも付く。
BAND_H = 0.36
BAND_TOP = -0.02
BAND_PAD = 0.06       # 左右に広げる幅（文字の大きさに対する割合）


def _wrap_chars(chars, width_em, bold):
    """[(文字, 強調か)] を width_em で折り返し、行のリストにする（文字幅は char_em の見積もり）。"""
    lines, cur, w = [], [], 0.0
    for ch, e in chars:
        cw = char_em(ch, bold)
        if cur and w + cw > width_em:
            lines.append(cur)
            cur, w = [], 0.0
            if ch == " ":
                continue
        cur.append((ch, e))
        w += cw
    lines.append(cur)
    return lines


def emphasis_bands(shapes, text, x, y, w, h, size, bold, align, anchor, color, line_spacing=1.0,
                   tidy=True):
    """枠 (x, y, w, h) に書いた文の **強調** の後ろに、図形の帯を置く。color は EMPH_BAND の帯の色。"""
    if not EMPH_BAND or shapes is None:
        return
    ec = (color or "").lstrip("#").upper()
    if ec != EMPH_BAND[0].lstrip("#").upper():
        return
    band = EMPH_BAND[1]
    text = as_text(text)
    if tidy:
        text = tidy_ja(text)
    em = size / 72.0
    line_h = em * LINE_HEIGHT * line_spacing
    lines = []
    for para in text.split("\n"):
        chars = [(ch, e) for chunk, e in split_emphasis(para) for ch in chunk]
        lines += _wrap_chars(chars, w / em * 0.98, bold)
    if not any(e for ln in lines for _, e in ln):
        return
    block = len(lines) * line_h
    top = y if anchor == "t" else (y + h - block if anchor == "b" else y + (h - block) / 2)
    for i, ln in enumerate(lines):
        total = sum(char_em(ch, bold) for ch, _ in ln) * em
        x0 = x + ((w - total) / 2 if align == "c" else (w - total if align == "r" else 0))
        cx, start = x0, None
        for ch, e in ln + [("", False)]:
            if e and start is None:
                start = cx
            if not e and start is not None:
                center = top + line_h * (i + 0.5)
                pad = BAND_PAD * em
                add_box(shapes, start - pad, center + BAND_TOP * em, cx - start + 2 * pad,
                        BAND_H * em, fill=band, name="マーカー")
                start = None
            cx += char_em(ch, bold) * em if ch else 0


def _inset(ph, attr, default):
    """プレースホルダーの内側の余白（インチ）。スライド → レイアウトの順に探す。"""
    shape = ph
    while shape is not None:
        bodyPr = shape._element.find(".//" + qn("a:bodyPr"))
        if bodyPr is not None and bodyPr.get(attr) is not None:
            return int(bodyPr.get(attr)) / 914400
        try:
            shape = shape._base_placeholder
        except AttributeError:
            shape = None
    return default / 914400


def place(shape, x=None, y=None, w=None, h=None):
    """プレースホルダーの位置を変える。継承している位置を明示してから変える
    （一部だけ指定すると、残りが 0 になってしまうため）。"""
    cur = (shape.left, shape.top, shape.width, shape.height)
    shape.left, shape.top, shape.width, shape.height = cur
    if x is not None:
        shape.left = Inches(x)
    if y is not None:
        shape.top = Inches(y)
    if w is not None:
        shape.width = Inches(w)
    if h is not None:
        shape.height = Inches(h)


def clone_element(el):
    return copy.deepcopy(el)


# ---------------------------------------------------------------- 自由な形（モチーフ用）

def add_path(shapes, x, y, w, h, paths, fill=None, line=None, line_w=1.0, rotation=0,
             name=None, size=1000):
    """自由な形を置く。paths は 0〜size の座標で書いた
    [('M', x, y), ('L', x, y), ('C', x1, y1, x2, y2, x3, y3), ('Z',)] のリスト。
    複数の輪郭を描くときは paths をリストのリストにする。"""
    shp = add_box(shapes, x, y, w, h, fill=fill, line=line, line_w=line_w, rotation=rotation,
                  name=name)
    spPr = shp._element.spPr
    prst = spPr.find(qn("a:prstGeom"))
    cust = etree.Element(qn("a:custGeom"))
    for tag in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
        etree.SubElement(cust, qn(tag))
    etree.SubElement(cust, qn("a:rect"), l="0", t="0", r="r", b="b")
    pl = etree.SubElement(cust, qn("a:pathLst"))
    groups = paths if paths and isinstance(paths[0], list) else [paths]
    for cmds in groups:
        pe = etree.SubElement(pl, qn("a:path"), w=str(size), h=str(size))
        if fill is None:
            pe.set("fill", "none")
        for cmd in cmds:
            op = cmd[0]
            if op == "Z":
                etree.SubElement(pe, qn("a:close"))
                continue
            tag = {"M": "a:moveTo", "L": "a:lnTo", "C": "a:cubicBezTo"}[op]
            el = etree.SubElement(pe, qn(tag))
            pts = cmd[1:]
            for i in range(0, len(pts), 2):
                etree.SubElement(el, qn("a:pt"), x=str(int(pts[i])), y=str(int(pts[i + 1])))
    spPr.replace(prst, cust)
    return shp


# ---------------------------------------------------------------- アイコン（デジタル庁の素材）
# assets/da-icons.json: デジタル庁「イラストレーション・アイコン素材」のアイコン 60 種（線 line・塗り fill）。
# SVG の輪郭（24×24 の座標）を、そのまま PowerPoint の図形（自由な形）にする。色は元の色（1A1A1C）のまま使う
# （色を変えると規約の「加工」に当たり、資料ごとに出典と加工の内容を書く必要が出るため）。

ICONS_PATH = Path(__file__).resolve().parent.parent / "assets" / "da-icons.json"
_ICONS = None
_SVG_TOKEN = re.compile(r"[MLHVCZ]|-?\d*\.?\d+(?:e-?\d+)?", re.I)


def icon_data():
    global _ICONS
    if _ICONS is None:
        with open(ICONS_PATH, encoding="utf-8") as f:
            _ICONS = json.load(f)
    return _ICONS


def icon_names():
    return sorted(icon_data()["icons"])


def _svg_cmds(d, scale):
    """SVG の d（絶対座標の M L H V C Z だけ）を add_path の形にする。"""
    out, cur, x, y = [], None, 0.0, 0.0
    nums = []
    toks = _SVG_TOKEN.findall(d)

    def emit(op, vals):
        nonlocal x, y
        if op == "M":
            x, y = vals
            out.append(("M", x * scale, y * scale))
        elif op == "L":
            x, y = vals
            out.append(("L", x * scale, y * scale))
        elif op == "H":
            x = vals[0]
            out.append(("L", x * scale, y * scale))
        elif op == "V":
            y = vals[0]
            out.append(("L", x * scale, y * scale))
        elif op == "C":
            out.append(("C",) + tuple(v * scale for v in vals))
            x, y = vals[4], vals[5]

    need = {"M": 2, "L": 2, "H": 1, "V": 1, "C": 6}
    for t in toks:
        if t.isalpha():
            if t.upper() != t:
                raise ValueError("相対座標のアイコンには対応していません")
            if t == "Z":
                out.append(("Z",))
                cur = None
            else:
                cur = t
            continue
        nums.append(float(t))
        if cur and len(nums) == need[cur]:
            emit(cur, nums)
            nums = []
            if cur == "M":
                cur = "L"            # M の後ろに続く数は L とみなす（SVG の決まり）
    return out


def add_icon(shapes, name, x, y, size, variant="line"):
    """アイコンを置く（x, y は左上、size は一辺のインチ）。name は「help」または「help_fill」。
    知らない名前なら None を返す。"""
    data = icon_data()
    base, kind = name, variant
    for suf in ("_line", "_fill"):
        if name.endswith(suf):
            base, kind = name[: -len(suf)], suf[1:]
    icon = data["icons"].get(base)
    if icon is None:
        return None
    vb = data["viewbox"]
    scale = 1000
    groups = [_svg_cmds(d, scale) for d in icon[kind]]
    return add_path(shapes, x, y, size, size, groups, fill=data["color"], line=None,
                    name="アイコン", size=vb * scale)


# assets/da-illustrations.zip: 同じ素材のイラスト 74 点（PNG。資料に使う大きさに縮めたもの）。
ILLUST_PATH = Path(__file__).resolve().parent.parent / "assets" / "da-illustrations.zip"


def illustration_names():
    import zipfile
    with zipfile.ZipFile(ILLUST_PATH) as z:
        return sorted(n[:-4] for n in z.namelist() if n.endswith(".png"))


def illustration_bytes(name):
    """同梱のイラストの PNG（知らない名前なら None）。"""
    import zipfile
    with zipfile.ZipFile(ILLUST_PATH) as z:
        try:
            return z.read(name + ".png")
        except KeyError:
            return None


def add_image(shapes, data, x, y, w, h, alt, name="画像", valign="m"):
    """画像（バイト列）を (x, y, w, h) の枠に縦横比を保って収め、左右の中央に置く。
    valign は縦の寄せ（"t" 上・"m" 中央）。置いた画像を返す。"""
    import io
    from PIL import Image  # python-pptx が使う Pillow
    iw, ih = Image.open(io.BytesIO(data)).size
    k = min(w / iw, h / ih)
    pw, ph = iw * k, ih * k
    top = y if valign == "t" else y + (h - ph) / 2
    pic = shapes.add_picture(io.BytesIO(data), Inches(x + (w - pw) / 2), Inches(top),
                             Inches(pw), Inches(ph))
    pic.name = name
    set_alt(pic, alt)
    return pic


def add_illustration(shapes, name, x, y, w, h, alt):
    """イラストを (x, y, w, h) の枠に縦横比を保って収め、中央に置く。知らない名前なら None。"""
    data = illustration_bytes(name)
    if data is None:
        return None
    return add_image(shapes, data, x, y, w, h, alt, name="イラスト")


# figure の画像の相対パスの基準（build_pptx.build が deck.json のフォルダにする）
DECK_DIR = Path.cwd()


# ---------------------------------------------------------------- 読み上げ順と代替テキスト
# スクリーンリーダーは、スライドの図形を spTree の並び順（＝重なりの順）に読む。
# 描く都合の順のままだと「見出し 3 つ → 本文 3 つ」「注記が本文より先」のように読まれるので、
# 1 枚描き終えたら order_for_reading で、見た目の順（上 → 下、左 → 右。カードはカードごと）に並べ替える。
# 文字の無い図形（カードの地・線・帯）は「装飾用」にして読み上げから外し、重なりの一番下に回す
# （文字の無い図形は文字の後ろにあるものなので、見た目は変わらない）。

HEADER_BOTTOM = 1.38   # これより上から始まる図形は、見出しまわり（ラベル・見出し・リード文）とみなす
DECORATIVE_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
ADEC_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"
_SHAPE_TAGS = {qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:grpSp"), qn("p:cxnSp")}


def _c_nv_pr(el):
    return el.find(".//" + qn("p:cNvPr"))


def set_alt(shape, text):
    """図形に代替テキスト（読み上げられる説明）を付ける。"""
    el = shape._element if hasattr(shape, "_element") else shape
    _c_nv_pr(el).set("descr", plain(as_text(text)))
    return shape


HIDDEN_TITLE_NAME = "読み上げ用の題名"


def add_hidden_title(slide, text):
    """スライドの外（上）に題名のプレースホルダーを置く。読み上げソフトと「タイトルが無い」の検査のため。"""
    sid = slide.shapes._next_shape_id
    xml = (
        '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="{HIDDEN_TITLE_NAME}"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(Inches(0.6))}" y="{int(Inches(-1.2))}"/>'
        f'<a:ext cx="{int(Inches(8))}" cy="{int(Inches(0.8))}"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2800"/>'
        f'<a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')
    slide.shapes._spTree.append(etree.fromstring(xml))


def set_decorative(el):
    """図形を「装飾用」にする（PowerPoint の［装飾用にする］と同じ。読み上げられない）。"""
    nv = _c_nv_pr(el)
    if nv is None:
        return
    ext_lst = nv.find(qn("a:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(nv, qn("a:extLst"))
    for ext in ext_lst.findall(qn("a:ext")):
        if ext.get("uri") == DECORATIVE_URI:
            return
    ext = etree.SubElement(ext_lst, qn("a:ext"), uri=DECORATIVE_URI)
    etree.SubElement(ext, "{%s}decorative" % ADEC_NS, nsmap={"adec": ADEC_NS}, val="1")


def _has_text(el):
    return any((t.text or "").strip() for t in el.iter(qn("a:t")))


def _is_content(el):
    """読み上げる図形か。文字・表・グラフ・画像・代替テキストのある図形。"""
    if el.tag in (qn("p:graphicFrame"), qn("p:pic")):
        return True
    nv = _c_nv_pr(el)
    if nv is not None and (nv.get("descr") or "").strip():
        return True
    if el.find(".//" + qn("p:ph")) is not None:
        return True
    return _has_text(el)


def _rows(items, key):
    """items を行に分けて、上の行から、行の中は左から並べる。
    key(item) → (left, top, right, bottom)。縦の重なりが小さい方の高さの半分以上なら同じ行。"""
    rows = []
    for it in sorted(items, key=lambda i: (key(i)[1], key(i)[0])):
        l, t, r, b = key(it)
        for row in rows:
            rt, rb = row["t"], row["b"]
            over = min(b, rb) - max(t, rt)
            if over >= 0.5 * max(1, min(b - t, rb - rt)):
                row["items"].append(it)
                row["t"], row["b"] = min(rt, t), max(rb, b)
                break
        else:
            rows.append({"t": t, "b": b, "items": [it]})
    rows.sort(key=lambda r: r["t"])
    return [it for row in rows for it in sorted(row["items"], key=lambda i: key(i)[0])]


def order_for_reading(slide, first=(), keep_drawn=False):
    """スライドの図形を見た目の順に並べ替え、文字の無い図形を装飾用にする。

    first: 先に読ませる図形の名前（add_text の label）。例: ツリーの根。
    keep_drawn: True なら、読む図形は描いた順のまま（型の描き方がすでに読む順のとき）。
    カード（文字の無い箱）の中の図形はカードごとにまとめ、カードのすぐ上に接して置いた
    図形（process の矢羽根など）もそのカードの組に入れる。"""
    tree = slide.shapes._spTree
    shapes = list(slide.shapes)
    els = [sh._element for sh in shapes if sh._element.tag in _SHAPE_TAGS]
    if not els:
        return
    box = {}
    for sh in shapes:
        el = sh._element
        try:
            l, t, w, h = sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0
        except Exception:
            l = t = w = h = 0
        box[el] = (l, t, l + w, t + h)
    deco = [el for el in els if not _is_content(el)]
    content = [el for el in els if _is_content(el)]
    for el in deco:
        set_decorative(el)
    # カード: 文字の無い、ある程度大きい箱（線は除く）
    min_area = int(0.3 * 914400 * 914400)
    cards = [el for el in deco if el.tag == qn("p:sp")
             and (box[el][2] - box[el][0]) * (box[el][3] - box[el][1]) >= min_area]

    def above(el):
        """el のすぐ下に接しているカード（el より十分に背の高いもの）。無ければ None。"""
        l, t, r, b = box[el]
        for cd in cards:
            cl, ct, cr, cb = box[cd]
            over = min(r, cr) - max(l, cl)
            if (cd is not el and r > l and over >= 0.7 * (r - l)
                    and -0.05 * 914400 <= ct - b <= 0.4 * 914400 and (b - t) <= 0.4 * (cb - ct)):
                return cd
        return None

    def owner(el):
        """el を読む組（カード）。カードの中ならそのカード、カードのすぐ上の小さな図形
        （process の矢羽根など）なら下のカード。"""
        own = _inside(el)
        up = above(own)
        return up if up is not None else own

    def _inside(el):
        l, t, r, b = box[el]
        cx, cy = (l + r) / 2, (t + b) / 2
        best = None
        for cd in cards:
            cl, ct, cr, cb = box[cd]
            if cl <= cx <= cr and ct <= cy <= cb:
                if best is None or (cr - cl) * (cb - ct) < (box[best][2] - box[best][0]) * (box[best][3] - box[best][1]):
                    best = cd
        return best if best is not None else el

    groups = {}
    for el in content:
        groups.setdefault(owner(el), []).append(el)

    def gbox(members, key_el):
        bs = [box[m] for m in members] + ([box[key_el]] if key_el in box else [])
        return (min(b[0] for b in bs), min(b[1] for b in bs), max(b[2] for b in bs), max(b[3] for b in bs))

    if keep_drawn:
        # 見出しまわり（上端が本文より上の図形）だけは見た目の順、残りは描いた順
        top_lim = int(HEADER_BOTTOM * 914400)
        header = _rows([el for el in content if box[el][1] < top_lim], key=lambda e: box[e])
        ordered = header + [el for el in content if el not in header]
    else:
        glist = [(k, v, gbox(v, k)) for k, v in groups.items()]
        ordered = []
        for k, v, _ in _rows(glist, key=lambda g: g[2]):
            ordered += _rows(v, key=lambda e: box[e])
    # 先に読ませる図形（名前で指定）
    names = {el: (_c_nv_pr(el).get("name") if _c_nv_pr(el) is not None else "") for el in ordered}
    head = [el for nm in first for el in ordered if names[el] == nm]
    if head:
        # 見出しまわりの直後に入れる
        top_lim = int(HEADER_BOTTOM * 914400)
        rest = [el for el in ordered if el not in head]
        k = 0
        while k < len(rest) and box[rest[k]][1] < top_lim:
            k += 1
        ordered = rest[:k] + head + rest[k:]
    # ページの題名（タイトルのプレースホルダー）は、上にラベルやバッジがあっても最初に読む
    def is_title(el):
        ph = el.find(".//" + qn("p:ph"))
        return ph is not None and ph.get("type") in ("title", "ctrTitle")
    ordered = [el for el in ordered if is_title(el)] + [el for el in ordered if not is_title(el)]
    # 並べ替え: 装飾（描いた順のまま）→ 読む図形（読む順）
    for el in els:
        tree.remove(el)
    for el in deco + ordered:
        tree.append(el)
