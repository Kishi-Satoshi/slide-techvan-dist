#!/usr/bin/env python3
"""テーマ JSON の検査（デザインを追加・変更したときに使う）。

使い方:
    python check_theme.py            # styles/*/theme.json をすべて検査
    python check_theme.py techvan    # 1つだけ

検査内容:
  - 必要な色の役割（キー）がそろっているか
  - 文字と背景の組み合わせが読みやすいか（WCAG のコントラスト比）
      本文・小さい文字 4.5 以上 / 大きい文字・太字の見出し 3.0 以上
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
STYLES = HERE.parent / "styles"



def _console_utf8():
    """日本語版 Windows でも表示が文字化けしないよう、出力を UTF-8 にする。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

def lum(hex_str):
    h = hex_str.lstrip("#")
    vals = []
    for i in (0, 2, 4):
        c = int(h[i:i + 2], 16) / 255
        vals.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    r, g, b = vals
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a, b):
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


TECHVAN_PAIRS = [
    ("text", "bg", 4.5, "本文 / 地"),
    ("text", "bg_alt", 4.5, "本文 / 淡いパネル"),
    ("text_muted", "bg", 4.5, "補足文 / 地"),
    ("heading", "bg", 4.5, "見出し / 地"),
    ("on_primary", "primary", 4.5, "白文字 / 紺"),
    ("on_primary", "primary_dark", 4.5, "白文字 / 濃紺"),
    ("number", "card_bg", 3.0, "大きな数字 / カード"),
    ("emph", "marker", 3.0, "強調語（太字） / マーカー"),
    ("heading", "marker", 4.5, "見出し / マーカー"),
    ("on_primary", "confidential", 3.0, "CONFIDENTIAL（厳秘・太字） / 赤地"),
    ("confidential", "bg", 3.0, "CONFIDENTIAL（社外秘・太字） / 白地"),
]


def check(path):
    theme = json.loads(Path(path).read_text(encoding="utf-8"))
    c = theme["colors"]
    style = Path(path).parent.name
    problems = []
    pairs = TECHVAN_PAIRS
    for group in ("status", "levels"):
        for key, pair in c.get(group, {}).items():
            r = contrast(pair[1], pair[0])
            if r < 4.5:
                problems.append(f"{group}.{key} の文字が読みにくい: {pair[1]} / {pair[0]} = {r:.2f}（4.5 以上）")
    if "status" not in c or "levels" not in c:
        problems.append("色の役割「status」「levels」がありません")
    for fg, bg, need, label in pairs:
        if fg not in c or bg not in c:
            continue
        r = contrast(c[fg], c[bg])
        if r < need:
            problems.append(f"{label}: {c[fg]} / {c[bg]} = {r:.2f}（{need} 以上が必要）")
    name = style
    if problems:
        print(f"NG {name}")
        for p in problems:
            print("   - " + p)
        return False
    print(f"OK {name}")
    return True


def main():
    _console_utf8()
    targets = []
    if len(sys.argv) > 1:
        targets = [STYLES / arg / "theme.json" for arg in sys.argv[1:]]
    else:
        targets = sorted(STYLES.glob("*/theme.json"))
    ok = all([check(p) for p in targets])
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
