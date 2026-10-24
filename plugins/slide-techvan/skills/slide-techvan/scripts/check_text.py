#!/usr/bin/env python3
"""deck.json の文面を調べ、「機械が書いたように見える」ところを指摘する（文面の確認）。

調べること（語の一覧とパターンの説明は references/ai-smell-ja.md）:
  1. 確度「高」の語（決まり文句・誇張・先送り・絵文字・全角ダッシュなど）
  2. ラベルのままの見出し（「課題」「スケジュール」など）
  3. 見出しの語尾のそろいすぎ（です・ます で終わる見出しが半分以上、同じ語尾が 3 枚続く）
  4. 列挙の数のそろいすぎ（ほとんどが 3 個）
  5. 文頭の接続詞の連発

指摘は直すきっかけで、作成は止めない。build_pptx.py が作成のあとに自動で呼ぶ。
単体でも使える:
  python3 check_text.py 資料.deck.json
標準ライブラリだけで動く（Python 3.8 以上）。
"""
import json
import re
import sys
from pathlib import Path

LEXICON = Path(__file__).resolve().parent.parent / "references" / "ai-smell-ja.md"

# 見出しが短いのが正常な型（ラベル見出しの検査をしない）
LABEL_OK_TYPES = {"cover", "section", "agenda", "faq", "notes", "closing", "quote"}
# 見出しのそろいを数えない型
NOT_BODY_TYPES = LABEL_OK_TYPES
# 書き手が数を選べる列挙（費用の内訳・進捗・体制などデータで数が決まるものは数えない）
ENUM_KEYS = {
    "summary": ["points"], "kpi": ["items"], "bullets": ["items"], "cards": ["items"],
    "tree": ["branches"], "process": ["steps"], "risks": ["items"], "next-actions": ["items"],
    "chart": ["points"], "compare": ["left.points", "right.points"],
}
HEAD_KEYS = {"title", "lead", "subtitle"}
TABLE_KEYS = {"rows", "header"}
SKIP_KEYS = {"type", "confidential", "en", "number", "variant", "chart", "unit", "start", "end",
             "milestone", "status", "level", "score", "scores", "icon", "color"}
CONNECTIVES = ("また", "さらに", "加えて", "したがって", "そのため", "このように")
PARTICLES = set("はがをにでとへも")
EMOJI = re.compile("[\U0001F300-\U0001FAFF☀-➿⭐⭕✅❌❗‼⁉]")
DASH = re.compile("[—―]")
POLITE_END = re.compile(r"(です|ます|ました|でした|ません|ください)[。．]?$")


def load_lexicon(path=LEXICON):
    """ai-smell-ja.md の表から、確度「高」の語を (正規表現, 語, 範囲, 言い換え) で返す。"""
    words = []
    text = Path(path).read_text(encoding="utf-8")
    in_table = False
    for line in text.splitlines():
        if line.startswith("## "):
            in_table = line.strip() == "## 語の一覧"
            continue
        if not in_table or not line.startswith("|"):
            continue
        cols = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cols) < 5 or cols[1] != "高":
            continue
        word, _, scope, _, hint = cols[:5]
        if word == "<絵文字>":
            rx = EMOJI
        elif word == "<全角ダッシュ>":
            rx = DASH
        else:
            rx = re.compile(".{0,20}?".join(re.escape(p) for p in word.split("〜")))
        words.append((rx, word.strip("<>"), scope, hint))
    return words


def plain(s):
    return s.replace("**", "").strip()


def strings(slide):
    """ページの文字を (種類, 文字) で返す。種類は head / body / table。"""
    out = []

    def walk(v, kind):
        if isinstance(v, str):
            if v.strip():
                out.append((kind, plain(v)))
        elif isinstance(v, list):
            for x in v:
                walk(x, kind)
        elif isinstance(v, dict):
            for k, x in v.items():
                if k in SKIP_KEYS:
                    continue
                walk(x, "table" if k in TABLE_KEYS or kind == "table" else kind)

    for k, v in slide.items():
        if k in SKIP_KEYS:
            continue
        if k in TABLE_KEYS:
            walk(v, "table")
        elif k in HEAD_KEYS or (slide.get("type") == "cover" and k == "badge"):
            walk(v, "head")
        else:
            walk(v, "body")
    return out


def sentences(text):
    return [s for s in re.split(r"[。\n]", text) if s.strip()]


def get_path(d, dotted):
    for part in dotted.split("."):
        if not isinstance(d, dict):
            return None
        d = d.get(part)
    return d


def is_label(title):
    t = plain(title)
    return (len(t) <= 12 and not (set(t) & PARTICLES) and not re.search(r"[0-9０-９]", t)
            and not t.endswith(("。", "る", "た", "い", "す", "だ")))


def check(deck, lexicon=None):
    """文面の指摘の一覧（文字列）を返す。"""
    words = load_lexicon() if lexicon is None else lexicon
    notes = []
    titles = []      # (枚目, 見出し) 本文の型だけ
    enum_counts = []
    conn = 0
    for i, s in enumerate(deck.get("slides", []), start=1):
        if not isinstance(s, dict):
            continue
        t = s.get("type", "")
        where = f"{i}枚目（{t}）: "
        found = set()
        for kind, text in strings(s):
            for rx, word, scope, hint in words:
                if word in found or (scope == "head" and kind != "head"):
                    continue
                if word in ("全角ダッシュ", "絵文字") and kind == "table":
                    continue
                m = rx.search(text)
                if m:
                    found.add(word)
                    notes.append(f"{where}「{m.group(0)}」は機械的に見えます。{hint}")
            if kind != "table":
                conn += sum(1 for sen in sentences(text) if sen.lstrip().startswith(CONNECTIVES))
        title = s.get("title")
        if isinstance(title, str) and title.strip() and t not in NOT_BODY_TYPES:
            titles.append((i, plain(title)))
            if is_label(title):
                notes.append(f"{where}見出し「{plain(title)}」がラベルのままです。結論を言い切る文にしてください")
        for key in ENUM_KEYS.get(t, []):
            v = get_path(s, key)
            if isinstance(v, list) and len(v) >= 2:
                enum_counts.append(len(v))

    # 資料全体のそろい
    if len(titles) >= 4:
        polite = [n for n, tt in titles if POLITE_END.search(tt)]
        if len(polite) * 2 >= len(titles):
            notes.append(f"資料全体: 本文の見出し {len(titles)} 枚のうち {len(polite)} 枚が「です・ます」で終わっています。"
                         "見出しは言い切り・体言止めを中心にしてください")
    run, prev = [], None
    for n, tt in titles:
        end = re.sub(r"[。．]$", "", tt)[-2:]
        if end == prev:
            run.append(n)
        else:
            run = [n]
        prev = end
        if len(run) == 3:
            notes.append(f"資料全体: {run[0]}〜{run[-1]}枚目の見出しが同じ語尾（「…{end}」）で続いています。文の形を変えてください")
    if len(enum_counts) >= 4 and sum(1 for c in enum_counts if c == 3) * 10 >= len(enum_counts) * 7:
        threes = sum(1 for c in enum_counts if c == 3)
        notes.append(f"資料全体: 列挙 {len(enum_counts)} か所のうち {threes} か所が 3 個です。"
                     "数は内容の重さで 2〜4 に決め、毎回 3 点にそろえないでください")
    if conn >= 4:
        notes.append(f"資料全体: 文頭の接続詞（また・さらに・そのため など）が {conn} 回あります。大半は消しても意味が通ります")
    return notes


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        return 0
    with open(sys.argv[1], encoding="utf-8-sig") as f:
        deck = json.load(f)
    notes = check(deck)
    if notes:
        print(f"文面の確認（{len(notes)} 件）:")
        for n in notes:
            print("  - " + n)
    else:
        print("文面の確認: 指摘はありません")
    return 0


if __name__ == "__main__":
    sys.exit(main())
