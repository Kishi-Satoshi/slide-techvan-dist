"""要約・案内まわりの型（目次・注記・サマリー・次のアクション・声・よくある質問）。

agenda / notes / summary / next-actions / quote / faq を登録する。
ここでは配置だけを決め、色・角丸・影などの見た目は style（styles_impl.py）の部品に任せる。
色はすべて style.c の役割名で指定する（白 FFFFFF だけは直接書く）。
型で共通の決まり（styles_impl.py の「型で共通に使う決まり」）に従う:
  本体は Area の上端 a.y から上詰めで置き、余りは下に残す／文字サイズは style.size の段階／
  番号の丸は style.marker／横に並べるときは a.x〜a.right を style.gutter() の間隔で分ける。
"""
import re
from contextlib import contextmanager

import slidekit as K
from styles_impl import Area
from renderers import register, as_items, check_count, card_needed_height


# ====================================================================== この型だけで使う部品

_SIZE = {"item_head": 16, "body": 14, "body_large": 16, "body_small": 12, "caption": 11, "quote": 18}


def _pt(style, key):
    """文字サイズの段階（テーマの metrics.sizes。なければ既定）。"""
    return style.size(key, _SIZE[key])


def _s(v):
    """値を文字列にそろえる（年や内線番号を数字で書いても扱えるように）。None は空文字。"""
    return "" if v is None else str(v)


def _list(v):
    """1つだけ（文字列・数字・辞書）で書かれた値もリストにそろえる。"""
    if v is None:
        return []
    if isinstance(v, (str, int, float, dict)):
        return [v]
    return list(v)


def _items(v, key, fields):
    """項目を辞書のリストにそろえる。文字列だけの項目は {key: …} にし、fields の値は文字列にする。"""
    out = as_items([it if isinstance(it, dict) else _s(it) for it in _list(v)])
    for it in out:
        if key not in it and "text" in it:
            it[key] = it.pop("text")
        for f in fields:
            if f in it:
                it[f] = _s(it[f]).strip()
    return out


def _clamp(v, lo, hi):
    return max(lo, min(hi, v))


def _fit_count(h, gap, min_row):
    """高さ h に、高さ min_row 以上の行がいくつ入るか（1 以上）。"""
    return max(1, int((h + gap) / (min_row + gap) + 1e-6))


def _limit(items, lo, hi, fit, what):
    """個数を確かめる。fit 個を超える分は表示しない（警告は1つにまとめる）。"""
    n = len(items)
    if n > fit:
        K.WARN.add(f"{what}は {lo}〜{hi} 個にしてください（いまは {n} 個。"
                   f"{fit + 1} 個目以降は表示していません）")
        return items[:fit]
    check_count(items, lo, hi, what)
    return items


def _default_label(style, sd, en, ja, always):
    """ラベル（英字のアイブロウと日本語のラベル）が指定されていないときの既定。

    英字（eyebrow）と日本語（eyebrow_ja・label）は別々に補う。
    always が False なら、直前に章扉があるときは章扉のラベルを使う。"""
    sec = style.current_section
    use = always or not (sec.get("en") or sec.get("ja"))
    out = dict(sd)
    if use and not sd.get("eyebrow"):
        out["eyebrow"] = en
    if use and not (sd.get("eyebrow_ja") or sd.get("label")):
        out["eyebrow_ja"] = ja
    return out


def _accent(style, i):
    """i 番目の項目の色（紺でそろえる）。"""
    return style.c["primary"]


def _tail(style, s, base_x, base_w, tip_x, y, tail):
    """吹き出しのしっぽ（カードの下辺 y から話し手へ向かう三角）。

    カード（plain_card）の上に重ね、付け根の枠線は塗りで隠す。枠の色はカードと同じ card_border。"""
    c = style.c
    x0 = min(base_x, tip_x)
    w = max(base_x + base_w, tip_x) - x0
    over = 0.03                      # 枠線を隠すため、付け根は少し上から塗る
    h = tail + over
    sx, sy = 1000 / w, 1000 / h
    bl, br, tp = (base_x - x0) * sx, (base_x + base_w - x0) * sx, (tip_x - x0) * sx
    edge = over * sy
    K.add_path(s.shapes, x0, y - over, w, h,
               [("M", bl, 0), ("L", br, 0), ("L", br, edge), ("L", tp, 1000), ("L", bl, edge), ("Z",)],
               fill=c["card_bg"])
    K.add_path(s.shapes, x0, y - over, w, h, [("M", bl, edge), ("L", tp, 1000), ("L", br, edge)],
               line=c["card_border"], line_w=1.0)


def _person(style, s, x, y, d, fill):
    """人のアイコン（丸の中に白い頭と肩）。"""
    K.add_box(s.shapes, x, y, d, d, fill=fill, geom="oval")
    k = 0.5523 * 170
    head = [("M", 500, 190), ("C", 500 + k, 190, 670, 360 - k, 670, 360),
            ("C", 670, 360 + k, 500 + k, 530, 500, 530),
            ("C", 500 - k, 530, 330, 360 + k, 330, 360),
            ("C", 330, 360 - k, 500 - k, 190, 500, 190), ("Z",)]
    body = [("M", 190, 900), ("C", 190, 700, 330, 590, 500, 590),
            ("C", 670, 590, 810, 700, 810, 900), ("Z",)]
    K.add_path(s.shapes, x + d * 0.15, y + d * 0.13, d * 0.7, d * 0.7, [head, body], fill="FFFFFF")


@contextmanager
def _same_header(style, items, w):
    """横に並ぶカードの見出し帯の高さを、一番高いものにそろえる（本文の書き出しをそろえる）。

    style.card は style.card_header_h で帯の高さを決めるので、この間だけ同じ値を返させる。"""
    hh = max(style.card_header_h(it, w) for it in items)
    style.card_header_h = lambda item, width: hh
    try:
        yield hh
    finally:
        del style.card_header_h


def _content_fits(it, inner, size):
    """カードの中身（body と points）が inner に size で収まるか。_card_content と同じ見積もり。"""
    body = K.plain(it.get("body", ""))
    points = it.get("points") or []
    bh = K.text_height(body, inner.w, size, line_spacing=1.08) if body else 0.0
    ph = (K.text_height("\n".join(K.plain(p) for p in points), inner.w - 0.22, size - 0.5)
          if points else 0.0)
    if body and points:
        return bh <= inner.h * 0.6 and ph <= inner.h - bh - 0.18
    return bh <= inner.h and ph <= inner.h


def _card_content(style, s, inner, it, size):
    """カードの中身（body と points）。renderers.cards と同じ組み方で、文字の大きさは size に固定する。"""
    c = style.c
    body = it.get("body", "")
    points = it.get("points") or []
    if body and points:
        body_h = min(inner.h * 0.6,
                     K.text_height(K.plain(body), inner.w, size, line_spacing=1.08) + 0.12)
        style.text(s, inner.x, inner.y, inner.w, body_h, body, size, c["text"], anchor="t",
                   fit_min=size, label="根拠の本文", line_spacing=1.08)
        py = inner.y + body_h + 0.06
        style.text(s, inner.x, py, inner.w, inner.bottom - py, points, size - 0.5, c["text"],
                   anchor="t", fit_min=size - 0.5, label="根拠の箇条書き", bullet="・",
                   bullet_color=c["primary"], para_space=2)
    elif body:
        style.text(s, inner.x, inner.y, inner.w, inner.h, body, size, c["text"], anchor="t",
                   fit_min=size, label="根拠の本文", line_spacing=1.08)
    elif points:
        style.text(s, inner.x, inner.y, inner.w, inner.h, points, size - 0.5, c["text"],
                   anchor="t", fit_min=size - 0.5, label="根拠の箇条書き", bullet="・",
                   bullet_color=c["primary"], para_space=2)


# ====================================================================== 目次

@register("agenda")
def agenda(style, sd):
    """目次。items: [{title, desc}]（2〜6個）。current で今の章（0 から数える）を強調する。

    ラベルを指定しなければ「AGENDA ／ 目次」にする（直前の章扉のラベルは引き継がない）。"""
    s, a = style.body(_default_label(style, sd, "AGENDA", "目次", True))
    c = style.c
    items = _items(sd.get("items"), "title", ("title", "desc"))
    gap = 0.2 if len(items) <= 4 else 0.14
    if items and (a.h - gap * (len(items) - 1)) / len(items) < 0.62:
        gap = 0.08
    items = _limit(items, 2, 6, _fit_count(a.h, gap, 0.52), "目次の項目")
    if not items:
        return
    n = len(items)
    cur = sd.get("current")
    if cur is not None:
        try:
            cur = int(cur)
        except (TypeError, ValueError):
            cur = -1
        if not 0 <= cur < n:
            K.WARN.add(f"current は 0〜{n - 1} の番号にしてください（いまは {sd.get('current')}）")
            cur = None
    has_desc = any(it.get("desc") for it in items)
    # 行は上から詰めて置き、余りは下に残す。番号の丸は大（l）、行が低いときは中（m）
    row_h = min(1.0 if n <= 4 else 0.86, (a.h - gap * (n - 1)) / n)
    mk = "l" if row_h >= style.MARKER["l"][0] + 0.24 else "m"
    d = style.MARKER[mk][0]
    pad = 0.24
    title_size = _pt(style, "item_head")
    desc_size = _pt(style, "body" if n <= 4 else "body_small")
    inset = min(0.24, row_h * 0.22)
    tx = a.x + pad + d + 0.26
    ph = min(max(0.34, _pt(style, "caption") * K.LINE_HEIGHT / 72 + 0.14), row_h - 0.1)   # ラベルの文字が縮まない高さ
    pill_w = 1.05 if cur is not None else 0.0
    right = a.right - pad - (pill_w + 0.25 if pill_w else 0.0)
    if has_desc:
        longest = max(K.text_em(K.plain(it.get("title", "")), True) for it in items)
        tw = min(max(longest * title_size / 72 + 0.2, 3.0), (right - tx) * 0.5)
    else:
        tw = right - tx
    y = a.y
    for i, it in enumerate(items):
        K.max_chars_warn(it.get("title", ""), 20, "目次の項目")
        is_cur = cur == i
        dim = cur is not None and not is_cur
        style.plain_card(s, Area(a.x, y, a.w, row_h), highlight=is_cur)
        # 番号は濃い丸。今の章を示すときは、ほかの章の番号を控えめにする
        style.marker(s, a.x + pad, y + (row_h - d) / 2, f"{i + 1:02d}", mk, solid=not dim, muted=dim)
        style.text(s, tx, y + 0.04, tw, row_h - 0.08, it.get("title", ""), title_size,
                   c["text_muted"] if dim else c["heading"], bold=True, anchor="m",
                   fit_min=min(12, title_size), label="目次の項目")
        if it.get("desc"):
            K.max_chars_warn(it["desc"], 40, "目次の補足")
            lx = tx + tw + 0.2
            K.add_line(s.shapes, lx, y + inset, lx, y + row_h - inset, c["rule"], 1.0)
            style.text(s, lx + 0.3, y + 0.04, right - lx - 0.3, row_h - 0.08, it["desc"],
                       desc_size, c["text_muted"] if dim else c["text"], anchor="m",
                       fit_min=min(10, desc_size), label="目次の補足")
        if is_cur:
            style.pill(s, a.right - pad - pill_w, y + (row_h - ph) / 2, pill_w, ph, "この章",
                       tone="primary", size=_pt(style, "caption"))
        y += row_h + gap


# ====================================================================== 注記・出典

_MARK_W = 0.55                        # 注記の番号（※1・[1]）の幅
_PAD, _TOP, _BOTTOM = 0.3, 0.22, 0.24  # 箱の内側の余白（左右・上・下）
_HEAD = 0.52                          # 箱の小見出しと、その下の間隔


def _note_texts(v):
    """注記・出典を文字列のリストにそろえる。辞書（{"title": …, "url": …}）は値を全角スペースでつなぐ。
    空の項目は除く。"""
    out = []
    for t in _list(v):
        if isinstance(t, dict):
            t = "　".join(_s(x).strip() for x in t.values() if _s(x).strip())
        t = _s(t).strip()
        if t:
            out.append(t)
    return out


_TOKEN = re.compile(r"[!-~]+|.")   # 半角の連続（URL・英単語）はひとかたまり、それ以外は1文字ずつ


def _note_lines(text, w, size):
    """注記の行数の見積もり。URL や英単語（半角の連続）は途中で折り返さないものとして数える
    （K.count_lines は1文字ずつ折り返すので、URL の多い出典では行数を少なく見積もってしまう）。"""
    width = w * 72 / size
    lines = 0
    for para in text.split("\n"):
        n, cur = 1, 0.0
        for tok in _TOKEN.findall(para):
            tw = K.text_em(tok) * (1.08 if len(tok) > 1 else 1.0)  # 語は少し広めに見積もる
            if cur + tw <= width:
                cur += tw
            elif tw <= width:
                n += 1
                cur = 0.0 if tok.isspace() else tw
            else:
                # 1行より長い語（長い URL）は新しい行から始め、行の幅で切って数える
                if cur > 0:
                    n += 1
                k = int(tw // width)
                n += k
                cur = tw - k * width
        lines += n
    return lines


def _note_heights(texts, w, size):
    """各項目の高さ（w は番号を含む幅）。"""
    return [_note_lines(K.plain(t), w - _MARK_W, size) * size * K.LINE_HEIGHT * 1.1 / 72
            for t in texts]


def _box_h(texts, w, size, gap):
    """注記の箱1つの高さ（w は箱の幅）。"""
    hs = _note_heights(texts, w - 2 * _PAD, size)
    return _TOP + _HEAD + sum(hs) + gap * (len(texts) - 1) + _BOTTOM


def _col_h(col, size, gap, box_gap):
    """上から積んだ箱の列の高さ。col は [(texts, w), …]。"""
    return sum(_box_h(t, w, size, gap) for t, w in col) + box_gap * (len(col) - 1)


def _note_fit(cols, h, top, box_gap):
    """高さ h に収まる文字の大きさと項目の間隔を探す（top から 9pt まで。最後は間隔も詰める）。

    戻り値: (文字の大きさ, 間隔, 収まるか, 一番高い列の高さ)。"""
    size = top
    while size >= 9:
        for gap in (0.24, 0.2, 0.17, 0.14):
            used = max(_col_h(col, size, gap, box_gap) for col in cols)
            if used <= h:
                return size, gap, True, used
        size -= 0.5
    used = max(_col_h(col, 9.0, 0.08, box_gap) for col in cols)
    return 9.0, 0.08, used <= h, used


def _note_lines_total(col, size):
    """列に積んだ箱の行数の合計（折り返しの多さをくらべる）。"""
    return sum(_note_lines(K.plain(t), w - 2 * _PAD - _MARK_W, size) for texts, w in col for t in texts)


def _note_split(notes_, sources, a, top, g):
    """左右の幅を変えるときの左（注記）の幅。文字が一番大きく、高い方の箱が一番低くなる分け方を選ぶ
    （片方が 34〜66%。同じくらいなら半々に近い方）。戻り値: (左の幅, _note_fit の結果)。"""
    best = None
    for k in range(9):
        r = 0.34 + 0.04 * k
        lw = (a.w - g) * r
        fit = _note_fit([[(notes_, lw)], [(sources, a.w - lw - g)]], a.h, top, g)
        key = (fit[2], fit[0], -round(fit[3], 1), -abs(r - 0.5))
        if best is None or key > best[0]:
            best = (key, lw, fit)
    return best[1], best[2]


def _note_list(style, s, a, texts, fmt, size, gap, label):
    """番号つきの注記・出典を上から並べる。枠に入りきらない項目は描かずに警告する。"""
    c = style.c
    lh = size * K.LINE_HEIGHT / 72
    y = a.y
    for k, (t, h) in enumerate(zip(texts, _note_heights(texts, a.w, size))):
        if y + h > a.bottom + 0.02:
            K.WARN.add(f"{label}の {k + 1} 個目以降（{len(texts) - k} 個）は枠に収まらないため"
                       "表示していません。2枚に分けてください")
            return
        if k:
            style.row_rule(s, a.x, a.right, y - gap / 2)
        style.text(s, a.x, y, _MARK_W, lh, fmt.format(k + 1), size, c["heading"], bold=True,
                   anchor="t", latin=style.font("number_latin"))
        style.text(s, a.x + _MARK_W, y, a.w - _MARK_W, h + 0.04, t, size, c["text"], anchor="t",
                   line_spacing=1.1, fit_min=9, label=label)
        y += h + gap


@register("notes")
def notes(style, sd):
    """注記・出典。notes: [文字列]、sources: [文字列]。

    両方あれば同じ幅で左右に並べる（文が長くて左右では窮屈なときは上下に積む）。
    2つの箱は同じ部品（plain_card）で、高さは中身に合わせる。"""
    s, a = style.body(sd)
    notes_ = _note_texts(sd.get("notes"))
    sources = _note_texts(sd.get("sources"))
    if not notes_ and not sources:
        K.WARN.add("注記ページには notes か sources が必要です")
        return
    if len(notes_) + len(sources) > 12:
        K.WARN.add(f"注記と出典が合わせて {len(notes_) + len(sources)} 個あります"
                   "（12 個以内を推奨。2枚に分けてください）")
    for t in notes_:
        K.max_chars_warn(t, 120, "注記")
    top = _pt(style, "body")
    head_pt = _pt(style, "item_head")
    g = style.gutter()
    blk_n = ("注記", notes_, "※{}")
    blk_s = ("出典", sources, "[{}]")
    if notes_ and sources:
        half = (a.w - g) / 2
        fs = _note_fit([[(notes_, half)], [(sources, half)]], a.h, top, g)
        fk = _note_fit([[(notes_, a.w), (sources, a.w)]], a.h, top, g)
        # 同じ幅で左右に並べるのを基本にする。上下に積むのは、積む方が 1pt 以上大きくできるとき・
        # 全幅ならどの項目も1行に収まるのに左右では折り返すとき（同じ大きさで収まる場合）・
        # 左右では収まらないとき。どちらでも収まらなければ、左右の幅を変えてみる
        one_line = len(notes_) + len(sources)
        wraps = (_note_lines_total([(notes_, a.w), (sources, a.w)], fs[0]) == one_line
                 < _note_lines_total([(notes_, half), (sources, half)], fs[0]))
        if ((fk[2] and (fk[0] >= fs[0] + 1 or (fk[0] >= fs[0] and wraps)))
                or (not fs[2] and (fk[2] or fk[3] < fs[3]))):
            layout, fit = [[(a.x, a.w, blk_n), (a.x, a.w, blk_s)]], fk
        else:
            layout, fit = [[(a.x, half, blk_n)], [(a.x + half + g, half, blk_s)]], fs
            if not fs[2]:
                lw, fu = _note_split(notes_, sources, a, top, g)
                if fu[2]:
                    layout, fit = [[(a.x, lw, blk_n)], [(a.x + lw + g, a.w - lw - g, blk_s)]], fu
    else:
        blk = blk_n if notes_ else blk_s
        fit = _note_fit([[(blk[1], a.w)]], a.h, top, g)
        layout = [[(a.x, a.w, blk)]]
    size, gap, ok, _ = fit
    if ok and size < 10:
        K.WARN.add(f"注記・出典が多いため文字を {size:g}pt まで小さくしました。項目を減らすか、2枚に分けてください")
    # 左右に並べるときは、2つの箱を同じ高さにそろえる
    same_h = (min(a.h, max(_box_h(col[0][2][1], col[0][1], size, gap) for col in layout))
              if len(layout) > 1 else None)
    for col in layout:
        hs = [same_h] if same_h else [_box_h(b[1], w, size, gap) for _, w, b in col]
        if sum(hs) + g * (len(hs) - 1) > a.h:
            # 入りきらないときは高さを按分する（入らない項目は _note_list が省いて警告する）
            k = (a.h - g * (len(hs) - 1)) / sum(hs)
            hs = [h * k for h in hs]
        y = a.y
        for (x, w, (head, texts, fmt)), h in zip(col, hs):
            box = Area(x, y, w, min(h, a.bottom - y))
            style.plain_card(s, box)
            hh = style.sub_heading(s, box.x + _PAD, box.y + _TOP, box.w - 2 * _PAD, head, head_pt)
            top_y = box.y + _TOP + hh + 0.14
            _note_list(style, s, Area(box.x + _PAD, top_y, box.w - 2 * _PAD,
                                      box.bottom - _BOTTOM - top_y), texts, fmt, size, gap, head)
            y = box.bottom + g


# ====================================================================== エグゼクティブサマリー

def _conclusion(text, w, top, low):
    """結論の文字の大きさ（2行に収まるまで top から low まで縮める）と、結論の面の高さ。"""
    size = top
    while size > low and K.count_lines(K.plain(text), w, size, True) > 2:
        size -= 0.5
    th = K.text_height(K.plain(text), w, size, True, line_spacing=1.1)
    return size, max(0.9, th + 0.44)


@register("summary")
def summary(style, sd):
    """エグゼクティブサマリー。conclusion（2行以内）＋ points: [{heading, body, points}]（2〜4個）。

    上から結論の面 → 根拠のカード（a.x〜a.right を等分）の順に詰めて置く。
    根拠に body か points があればカードで並べ、見出しだけなら番号つきの箇条書きにする。
    ラベルを指定せず、直前に章扉もないときは「SUMMARY ／ 要約」にする。"""
    s, a = style.body(_default_label(style, sd, "SUMMARY", "要約", False))
    c = style.c
    conclusion = _s(sd.get("conclusion")).strip()
    points = _items(sd.get("points"), "heading", ("heading", "body", "label", "mark"))
    for it in points:
        it["points"] = [_s(t) for t in _list(it.get("points")) if _s(t).strip()]
    check_count(points, 2, 4, "サマリーの根拠（points）")
    if not conclusion:
        K.WARN.add("サマリーには conclusion（結論）を入れてください")
        if not points:
            return
    g = style.gutter()
    head = _pt(style, "item_head")
    low = _pt(style, "body_small")
    pad, lab_w, lab_h = 0.3, 1.0, 0.36
    tx = a.x + pad + lab_w + 0.3
    tw = a.right - pad - tx

    def draw_conclusion(y, size, ch):
        style.panel(s, Area(a.x, y, a.w, ch))
        style.pill(s, a.x + pad, y + (ch - lab_h) / 2, lab_w, lab_h, "結論", tone="primary",
                   size=low)
        style.text(s, tx, y + 0.08, tw, ch - 0.16, conclusion, size, c["heading"], bold=True,
                   anchor="m", fit_min=min(12, size), label="結論", line_spacing=1.1)

    # 結論は見出し（title）より明らかに小さく、項目の見出し＋2pt まで
    csize, ch = (0, 0.0)
    if conclusion:
        K.max_chars_warn(conclusion, 70, "結論")
        csize, ch = _conclusion(conclusion, tw, head + 2, head)
    if not points:
        draw_conclusion(a.y, csize, min(a.h, ch))
        return
    space = g if conclusion else 0.0
    for it in points:
        K.max_chars_warn(it.get("heading", ""), 24, "根拠の見出し")
    n = len(points)
    if not any(it.get("body") or it["points"] for it in points):
        _summary_list(style, s, a, points, ch + space,
                      (lambda y: draw_conclusion(y, csize, ch)) if conclusion else None)
        return
    w = (a.w - g * (n - 1)) / n
    # 本文は cards 型と同じ段階（card_body）。結論の下に置くので大きい段階は使わない
    size = style.size("card_body", _pt(style, "body"))
    for it in points:
        if it.get("body"):
            K.max_chars_warn(it["body"], {1: 120, 2: 120, 3: 75}.get(n, 50), "根拠の本文")
    with _same_header(style, points, w):
        def need(sz):
            return max(card_needed_height(style, it, w, sz) for it in points)
        avail = a.h - ch - space
        while need(size) > avail and size > low:
            size -= 0.5
        if need(size) > avail and conclusion:
            # それでも入らなければ結論を1段小さくして、カードの場所を空ける
            csize, ch = _conclusion(conclusion, tw, head, low)
            avail = a.h - ch - space
        while need(size) > avail and size > 10:
            size -= 0.5
        # カードは中身に合わせた高さにし、結論のすぐ下に置く（余りは下に残す）
        card_h = min(avail, need(size))
        y = a.y
        if conclusion:
            draw_conclusion(y, csize, ch)
        y += ch + space
        inners = [style.card(s, Area(a.x + i * (w + g), y, w, card_h), it, i)
                  for i, it in enumerate(points)]
    # 本文の大きさは全カードでそろえる（一番窮屈なカードに合わせる）
    while size > 9.5 and not all(_content_fits(it, box, size) for it, box in zip(points, inners)):
        size -= 0.5
    if size < 10.5:
        K.WARN.add(f"根拠のカードの文字を {size:g}pt まで小さくしました。本文を短くするか、"
                   "lead・callout・note を減らしてください")
    for it, box in zip(points, inners):
        _card_content(style, s, box, it, size)


def _summary_list(style, s, a, points, head_h, draw_head):
    """見出しだけの根拠を、番号の丸＋見出しの一覧で並べる（bullets と同じ組み方）。"""
    c = style.c
    n = len(points)
    d = style.MARKER["m"][0]
    tx = a.x + d + 0.25
    tw = a.w - d - 0.25
    size = _pt(style, "item_head")
    hs = [max(d + 0.04, K.text_height(K.plain(it.get("heading", "")), tw, size, True) + 0.08)
          for it in points]
    avail = a.h - head_h
    pitch = min(max(hs) + 0.36, avail / n)
    y = a.y
    if draw_head:
        draw_head(y)
    y += head_h + 0.05
    for i, (it, hh) in enumerate(zip(points, hs)):
        style.marker(s, a.x, y + (hh - d) / 2, it.get("mark") or i + 1, "m")
        style.text(s, tx, y, tw, hh, it.get("heading", ""), size, c["heading"], bold=True,
                   anchor="m", fit_min=min(12, size), label="根拠の見出し")
        if i < n - 1:
            style.row_rule(s, a.x, a.right, y + (hh + pitch) / 2)
        y += pitch


# ====================================================================== 次のアクション

@register("next-actions")
def next_actions(style, sd):
    """まとめ・次のアクション。items: [{action, owner, due, note}]（1〜6個）。

    列の見出し（アクション・担当・期限）を a.y に置き、その下に行を上から詰めて並べる。"""
    s, a = style.body(sd)
    c = style.c
    items = _items(sd.get("items"), "action", ("action", "owner", "due", "note", "mark"))
    cap = _pt(style, "caption")
    head_h = max(0.3, cap * K.LINE_HEIGHT / 72 + 0.08)
    gap = 0.14 if len(items) >= 5 else 0.2
    items = _limit(items, 1, 6, _fit_count(a.h - head_h - 0.08, gap, 0.42), "次のアクション")
    if not items:
        return
    n = len(items)
    owners = [K.plain(it.get("owner", "")) for it in items]
    dues = [K.plain(it.get("due", "")) for it in items]
    has_owner = any(owners)
    has_due = any(dues)
    has_note = any(it.get("note") for it in items)
    if n <= 2:
        pref = 1.1 if has_note else 0.92
    elif n <= 3:
        pref = 1.0 if has_note else 0.84
    else:
        pref = 0.9 if has_note else 0.72
    max_row = (a.h - head_h - 0.08 - gap * (n - 1)) / n
    pad = 0.24
    # 番号の丸は中（m）。行がとても低いときだけ小（s）
    mk = "m" if min(pref, max_row) >= style.MARKER["m"][0] + 0.12 else "s"
    d = style.MARKER[mk][0]
    # アクションと期限は項目の見出しの大きさ、補足は小さめの本文、担当は注釈の大きさ
    asz = _pt(style, "item_head")
    nsz = _pt(style, "body_small")
    # 担当の列は一番長い名前に、期限の列は一番長い期限に合わせる
    owner_w = (_clamp(max(K.text_em(o, True) for o in owners) * cap / 72 + 0.75, 1.6, 3.0)
               if has_owner else 0.0)
    due_w = (_clamp(max(K.text_em(t, True) for t in dues) * asz / 72 + 0.5, 1.4, 2.6)
             if has_due else 0.0)
    ax = a.x + pad + d + 0.26
    due_x = a.right - due_w
    owner_x = due_x - owner_w
    aw = owner_x - ax - 0.25
    muted = c["text_muted"]

    def block(it, a_pt, n_pt):
        al = min(2, K.count_lines(K.plain(it.get("action", "")), aw, a_pt, True))
        nl = min(2, K.count_lines(K.plain(it.get("note", "")), aw, n_pt)) if it.get("note") else 0
        return al * a_pt * K.LINE_HEIGHT / 72, nl * n_pt * K.LINE_HEIGHT / 72

    def need(a_pt, n_pt):
        return max(sum(block(it, a_pt, n_pt)) + (0.05 if it.get("note") else 0) for it in items) + 0.2

    # 行の高さは中身に合わせて伸ばし、入りきらなければ文字を小さくする
    row_h = min(max(pref, need(asz, nsz)), max_row)
    # 行が低いときは、補足を2行目に置かずアクションの後ろに続ける
    inline = has_note and row_h < 0.62
    if inline:
        asz = min(asz, int((row_h - 0.14) * 72 / K.LINE_HEIGHT * 2) / 2)
        nsz = min(nsz, asz - 2)
        # 1行に収まるかは少し余裕を見て判定する（見積もりより実際の文字幅が広いことがある）
        while asz > 11 and any(
                (K.text_em(K.plain(it.get("action", "")), True) * asz
                 + K.text_em("　" + K.plain(it.get("note", ""))) * nsz) / 72 > aw * 0.93
                for it in items):
            asz -= 0.5
            nsz = max(9.0, nsz - 0.5)
    else:
        while need(asz, nsz) > row_h + 0.1 and (asz > 11 or nsz > 9):
            asz = max(11.0, asz - 0.5)
            nsz = max(9.0, nsz - 0.5)
    ph = min(max(0.34, cap * K.LINE_HEIGHT / 72 + 0.14), row_h - 0.1)   # ラベルの文字が縮まない高さ
    due_pt = min(asz, int((row_h - 0.1) * 72 / K.LINE_HEIGHT * 2) / 2)
    if has_due:
        # 期限は1行で読めるように、一番長い期限が列に収まる大きさにする
        due_pt = min(due_pt, int((due_w - 0.3) * 72 / max(K.text_em(t, True) for t in dues) * 2) / 2)
    due_pt = max(11.0, due_pt)
    inset = max(0.14, row_h * 0.22)
    y = a.y
    style.text(s, ax, y, aw, head_h, "アクション", cap, muted, bold=True, anchor="m")
    if has_owner:
        style.text(s, owner_x, y, owner_w, head_h, "担当", cap, muted, bold=True, align="c",
                   anchor="m")
    if has_due:
        style.text(s, due_x, y, due_w, head_h, "期限", cap, muted, bold=True, align="c",
                   anchor="m")
    y += head_h + 0.08
    for i, it in enumerate(items):
        K.max_chars_warn(it.get("action", ""), 30, "アクション")
        style.plain_card(s, Area(a.x, y, a.w, row_h))
        style.marker(s, a.x + pad, y + (row_h - d) / 2, it.get("mark") or i + 1, mk)
        note = it.get("note", "")
        if note:
            K.max_chars_warn(note, 40, "アクションの補足")
        if note and inline:
            box = style.text(s, ax, y + 0.04, aw, row_h - 0.08, it.get("action", ""), asz,
                             c["heading"], bold=True, anchor="m", fit_min=10, label="アクション")
            K.write_runs(box.text_frame.paragraphs[0], "　" + note, nsz, muted,
                         emph_color=c.get("emph"), font=style.font("ja"), latin=style.font("latin"))
            wide = (K.text_em(K.plain(it.get("action", "")), True) * asz
                    + K.text_em("　" + K.plain(note)) * nsz) / 72
            if wide > aw * (2 if row_h >= 2 * asz * K.LINE_HEIGHT / 72 + 0.08 else 1):
                K.WARN.add("「アクションの補足」が枠からはみ出す可能性があります。補足を短くしてください")
        elif note:
            ah, nh = block(it, asz, nsz)
            by = y + (row_h - ah - nh - 0.05) / 2
            style.text(s, ax, by, aw, ah + 0.02, it.get("action", ""), asz, c["heading"],
                       bold=True, anchor="b", fit_min=10, label="アクション")
            style.text(s, ax, by + ah + 0.05, aw, nh + 0.02, note, nsz, muted, anchor="t",
                       fit_min=9, label="アクションの補足")
        else:
            style.text(s, ax, y + 0.04, aw, row_h - 0.08, it.get("action", ""), asz,
                       c["heading"], bold=True, anchor="m", fit_min=10, label="アクション")
        if row_h >= 0.6:
            for lx in ([owner_x] if has_owner else []) + ([due_x] if has_due else []):
                K.add_line(s.shapes, lx, y + inset, lx, y + row_h - inset, c["rule"], 0.75)
        if owners[i]:
            K.max_chars_warn(owners[i], 12, "担当")
            pw = min(owner_w - 0.3, K.text_em(owners[i], True) * cap / 72 + 0.5)
            style.pill(s, owner_x + (owner_w - pw) / 2, y + (row_h - ph) / 2, pw, ph,
                       it["owner"], tone="soft", size=cap)
        if dues[i]:
            style.text(s, due_x + 0.1, y + 0.04, due_w - 0.2, row_h - 0.08, it["due"], due_pt,
                       c["emph"], bold=True, align="c", anchor="m",
                       latin=style.font("number_latin"), fit_min=min(11.0, due_pt), label="期限")
        y += row_h + gap


# ====================================================================== 声・引用

@register("quote")
def quote(style, sd):
    """声・引用。quotes: [{text, who, role}]（1〜3個）。吹き出しと話し手を並べる。

    吹き出しは a.y から置き、a.x〜a.right を style.gutter() の間隔で等分する。"""
    s, a = style.body(sd)
    c = style.c
    quotes = []
    for k, q in enumerate(_items(sd.get("quotes"), "text", ("text", "who", "role"))):
        if q.get("text"):
            quotes.append(q)
        else:
            K.WARN.add(f"引用の {k + 1} 個目に text（発言）がありません。表示していません")
    quotes = _limit(quotes, 1, 3, 3, "引用")
    if not quotes:
        return
    n = len(quotes)
    g = style.gutter()
    w = (a.w - g * (n - 1)) / n
    limit = {1: 120, 2: 90, 3: 70}[n]
    size = _pt(style, "quote")
    low = _pt(style, "body_small")
    who_pt = _pt(style, "body")
    role_pt = _pt(style, "caption")
    pad = 0.34
    mark_h = 0.5
    tail = 0.3
    person = 0.62
    max_bh = a.h - tail - person - 0.12
    tw = w - 2 * pad

    def need(sz):
        return max(K.text_height(K.plain(q["text"]), tw, sz, line_spacing=1.15) for q in quotes)

    # 一番長い文が収まる大きさにそろえる
    while size > low and need(size) + mark_h + pad + 0.1 > max_bh:
        size -= 0.5
    # 吹き出しは一番長い文に合わせた高さにし、上から置く（余りは下に残す）
    bh = min(max_bh, max(1.4, need(size) + mark_h + pad + 0.15))
    y = a.y
    for i, q in enumerate(quotes):
        x = a.x + i * (w + g)
        K.max_chars_warn(q["text"], limit, "引用")
        accent = _accent(style, i)
        px = x + 0.3
        style.plain_card(s, Area(x, y, w, bh))
        # しっぽは話し手の丸に向ける
        _tail(style, s, px + person * 0.55, 0.42, px + person * 0.42, y + bh, tail)
        # 引用符は飾りの記号なので、文字サイズの段階によらず大きく置く
        style.text(s, x + pad - 0.06, y + 0.08, 1.0, mark_h + 0.1, "“", 48, accent, bold=True,
                   anchor="t", latin=style.font("number_latin"), wrap=False)
        style.text(s, x + pad, y + mark_h + 0.06, tw, bh - mark_h - 0.06 - pad * 0.8,
                   q["text"], size, c["heading"], anchor="t",
                   fit_min=min(11, size), label="引用", line_spacing=1.15)
        py = y + bh + tail + 0.12
        _person(style, s, px, py, person, accent)
        nx = px + person + 0.16
        nw = x + w - nx
        who = q.get("who", "")
        role = q.get("role", "")
        if role:
            style.text(s, nx, py + 0.02, nw, 0.32, who, who_pt, c["heading"], bold=True,
                       anchor="m", fit_min=10, label="話し手")
            style.text(s, nx, py + 0.33, nw, 0.27, role, role_pt, c["text_muted"], anchor="m",
                       fit_min=9, label="話し手の役職")
        elif who:
            style.text(s, nx, py, nw, person, who, who_pt, c["heading"], bold=True, anchor="m",
                       fit_min=10, label="話し手")


# ====================================================================== よくある質問

def _faq_levels(style):
    """文字・印・間隔の段階（大きい順）。q は質問、a は回答、mk は印（Q・A の丸）の大きさ。"""
    ih, b, bs = (_pt(style, k) for k in ("item_head", "body", "body_small"))
    out = [{"mk": "m", "q": ih, "a": b, "sep": 0.3, "qa": 0.1},
           {"mk": "m", "q": b, "a": bs, "sep": 0.24, "qa": 0.08},
           {"mk": "s", "q": bs, "a": bs, "sep": 0.2, "qa": 0.06}]
    for p in out:
        p["d"] = style.MARKER[p["mk"]][0]
    return out


def _faq_heights(items, w, p):
    """各項目の（質問の高さ, 回答の高さ）と全体の高さ。p は文字・印・間隔の大きさ。回答がなければ 0。"""
    d = p["d"]
    tw = w - d - 0.24 - 0.1  # 見積もりは少し狭めにして余裕を持たせる
    hs = []
    for it in items:
        qh = K.text_height(K.plain(it.get("q", "")), tw, p["q"], True)
        ah = K.text_height(K.plain(it.get("a", "")), tw, p["a"], line_spacing=1.1)
        hs.append((max(d, qh + d - p["q"] * K.LINE_HEIGHT / 72),
                   max(d, ah + d - p["a"] * K.LINE_HEIGHT / 72) if it.get("a") else 0.0))
    total = sum(qh + (p["qa"] + ah if ah else 0.0) for qh, ah in hs) + p["sep"] * (len(items) - 1)
    return hs, total


def _faq_fit(style, items, w, h, levels=None):
    """高さ h に収まる段階を選ぶ。一番小さい段階でも入らなければ、回答を 10pt まで少しずつ縮める。"""
    levels = levels or _faq_levels(style)
    for p in levels:
        if _faq_heights(items, w, p)[1] <= h:
            return p
    p = dict(levels[-1])
    while _faq_heights(items, w, p)[1] > h and p["a"] > 10:
        p["a"] -= 0.5
        p["q"] = max(p["a"], p["q"] - 0.5)
        p["sep"] = max(0.14, p["sep"] - 0.02)
        p["qa"] = max(0.04, p["qa"] - 0.01)
    return p


def _faq_sep(items, a, p):
    """項目の間隔。余裕があれば広げる（最大 0.8 インチ。余白の 6 割まで）。"""
    hs, total = _faq_heights(items, a.w, p)
    if len(items) > 1 and total < a.h:
        return min(0.8, p["sep"] + (a.h - total) / (len(items) - 1) * 0.6)
    return p["sep"]


def _faq_split(items, w, p):
    """2列に分ける位置。左右の列の高さがなるべくそろうように選ぶ（同じくらいなら半分に近い位置）。"""
    half = (len(items) + 1) // 2

    def key(k):
        h = max(_faq_heights(items[:k], w, p)[1], _faq_heights(items[k:], w, p)[1])
        return round(h, 1), abs(k - half)
    return min(range(1, len(items)), key=key)


def _faq_column(style, s, a, items, p, sep, row_h=None):
    """1列分の質問と回答を描く。row_h（行ごとの高さ）を渡すと、2列の行の頭をそろえる。"""
    c = style.c
    hs, total = _faq_heights(items, a.w, p)
    n = len(items)
    d = p["d"]
    tx = a.x + d + 0.24
    tw = a.right - tx
    y = a.y
    for k, (it, (qh, ah)) in enumerate(zip(items, hs)):
        top = y
        K.max_chars_warn(it.get("q", ""), 40, "質問")
        # 質問の Q は濃い丸、回答の A は控えめな丸
        style.marker(s, a.x, y, "Q", p["mk"], solid=True)
        off = (d - p["q"] * K.LINE_HEIGHT / 72) / 2
        style.text(s, tx, y + off, tw, qh - off, it.get("q", ""), p["q"], c["heading"], bold=True,
                   anchor="t", fit_min=min(11, p["q"]), label="質問")
        y += qh
        if it.get("a"):
            y += p["qa"]
            style.marker(s, a.x, y, "A", p["mk"])
            off = (d - p["a"] * K.LINE_HEIGHT / 72) / 2
            style.text(s, tx, y + off, tw, ah - off + 0.02, it["a"], p["a"], c["text"],
                       anchor="t", fit_min=min(9.5, p["a"]), label="回答", line_spacing=1.1)
            y += ah
        if row_h and k < len(row_h):
            y = max(y, top + row_h[k])
        if k < n - 1:
            style.row_rule(s, a.x, a.right, y + sep / 2)
            y += sep


@register("faq")
def faq(style, sd):
    """よくある質問。items: [{q, a}]（2〜5個）。1列で収まらないとき（4個以上）は2列にする。"""
    s, a = style.body(sd)
    raw = _items(sd.get("items"), "q", ("q", "a"))
    check_count(raw, 2, 5, "よくある質問")
    items = []
    for k, it in enumerate(raw):
        if not it.get("q"):
            K.WARN.add(f"よくある質問の {k + 1} 個目に質問 q がありません。表示していません")
            continue
        if not it.get("a"):
            K.WARN.add(f"よくある質問の {k + 1} 個目に回答 a がありません")
        items.append(it)
    if not items:
        return
    # 回答の文字数の上限は、2列に分ける前の全体の数で決める
    for it in items:
        K.max_chars_warn(it.get("a", ""), 75 if len(items) <= 3 else 30, "回答")
    levels = _faq_levels(style)
    cols = [(a, items)]
    # 1列で、質問を本文の大きさまで縮めても入らないときは2列にする（4個以上のとき）
    if len(items) >= 4 and all(_faq_heights(items, a.w, p)[1] > a.h for p in levels[:2]):
        g = style.gutter()
        cw = (a.w - g) / 2
        k = _faq_split(items, cw, levels[0])
        cols = [(Area(a.x, a.y, cw, a.h), items[:k]),
                (Area(a.x + cw + g, a.y, cw, a.h), items[k:])]
    # 列が2つのときも文字の大きさはそろえる（小さい方に合わせる）
    p = min((_faq_fit(style, col, ar.w, ar.h, levels) for ar, col in cols),
            key=lambda q: (q["a"], q["q"]))
    if any(_faq_heights(col, ar.w, p)[1] > ar.h + 0.05 for ar, col in cols):
        K.WARN.add("よくある質問が枠に収まりません。項目を減らすか、回答を短くしてください")
    # 項目の間隔は、どの列も自分の高さに収まる一番狭いものにそろえる（2列の行の高さをそろえる）
    sep = min(_faq_sep(col, ar, p) for ar, col in cols)
    row_h = None
    if len(cols) == 2:
        # 2列のときは、行ごとに左右の高い方に合わせて、行の頭をそろえる
        per_col = []
        for ar, col in cols:
            hs, _ = _faq_heights(col, ar.w, p)
            per_col.append([qh + (p["qa"] + ah if ah else 0.0) for qh, ah in hs])
        rows = max(len(x) for x in per_col)
        row_h = [max(x[r] for x in per_col if r < len(x)) for r in range(rows)]
        need = sum(row_h) + sep * (rows - 1)
        if need > a.h + 0.05:
            sep = max(0.14, sep - (need - a.h) / max(1, rows - 1))
    for ar, col in cols:
        _faq_column(style, s, ar, col, p, sep, row_h)
