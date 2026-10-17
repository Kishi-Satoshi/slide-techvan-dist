#!/usr/bin/env python3
"""テンプレートの元のマスター・レイアウトが残っているかを検査する。

使い方:
    python check_template.py <検査する.pptx> [--base assets/template-base.pptx]

検査内容:
  - 元テンプレート（template-base.pptx）のスライドマスターとレイアウトが、
    検査対象にすべて存在し、中身（XML）が変わっていないこと
  - 元のマスター・レイアウトが参照する画像が同じであること
  - 追加されたマスター・レイアウトの一覧を表示する（削除はエラー、追加は OK）

終了コード: 0 = 問題なし、1 = 欠落・変更あり
"""
import argparse
import copy
import hashlib
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation

HERE = Path(__file__).resolve().parent
DEFAULT_BASE = HERE.parent / "assets" / "template-base.pptx"



def _console_utf8():
    """日本語版 Windows でも表示が文字化けしないよう、出力を UTF-8 にする。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

def canon(element):
    """XML を正規化してハッシュ化する（保存し直しによる書式差を無視するため）。"""
    data = etree.tostring(element, method="c14n")
    return hashlib.sha256(data).hexdigest()


def image_hashes(part):
    """パーツが参照している画像のハッシュ一覧。"""
    hashes = []
    for rel in part.rels.values():
        if rel.reltype.endswith("/image") and not rel.is_external:
            hashes.append(hashlib.sha256(rel.target_part.blob).hexdigest())
    return sorted(hashes)


def snapshot(prs):
    """マスターごとに {マスターの指紋: {レイアウト名: 指紋}} を作る。"""
    result = []
    for master in prs.slide_masters:
        layouts = {}
        for layout in master.slide_layouts:
            layouts[layout.name] = (canon(layout._element), image_hashes(layout.part))
        # マスター全体（背景・図形・文字の設定・配色の対応）を比べる。
        # レイアウトの一覧（sldLayoutIdLst）だけは、レイアウトを追加すると変わるので除く
        m_el = copy.deepcopy(master._element)
        for lst in m_el.findall(
                "{http://schemas.openxmlformats.org/presentationml/2006/main}sldLayoutIdLst"):
            m_el.remove(lst)
        result.append(
            {
                "master": (canon(m_el), image_hashes(master.part)),
                "theme": canon(etree.fromstring(
                    master.part.part_related_by(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                    ).blob)),
                "layouts": layouts,
                "layout_names": [l.name for l in master.slide_layouts],
            }
        )
    return result


def main():
    _console_utf8()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("target", help="検査する .pptx")
    ap.add_argument("--base", default=str(DEFAULT_BASE), help="元テンプレート（既定: assets/template-base.pptx）")
    args = ap.parse_args()

    base = snapshot(Presentation(args.base))
    target = snapshot(Presentation(args.target))

    errors = []
    matched = set()
    for bi, bm in enumerate(base):
        # 元マスターと同じ中身のマスターを探す
        found = None
        for ti, tm in enumerate(target):
            if ti in matched:
                continue
            if tm["master"] == bm["master"] and tm["theme"] == bm["theme"]:
                found = ti
                break
        if found is None:
            errors.append(f"元のマスター {bi + 1} が見つからないか、内容が変更されています")
            continue
        matched.add(found)
        tm = target[found]
        for name, fp in bm["layouts"].items():
            if name not in tm["layouts"]:
                errors.append(f"元のレイアウト「{name}」が削除されています")
            elif tm["layouts"][name] != fp:
                errors.append(f"元のレイアウト「{name}」の内容が変更されています")
        added = [n for n in tm["layout_names"] if n not in bm["layouts"]]
        if added:
            print(f"  元のマスター {bi + 1} に追加されたレイアウト: {', '.join(added)}")

    for ti, tm in enumerate(target):
        if ti not in matched:
            print(f"  追加されたマスター {ti + 1}: レイアウト {', '.join(tm['layout_names'])}")

    if errors:
        print("NG: 元テンプレートの部分に問題があります")
        for e in errors:
            print("  - " + e)
        return 1
    n_layouts = sum(len(b["layouts"]) for b in base)
    print(f"OK: 元のマスター {len(base)} 個・レイアウト {n_layouts} 個はすべて元のまま残っています")
    return 0


if __name__ == "__main__":
    sys.exit(main())
