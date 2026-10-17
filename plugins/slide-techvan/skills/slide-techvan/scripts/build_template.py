#!/usr/bin/env python3
"""元テンプレートから、生成に使う template.pptx を作る（保守用）。

使い方:
    python build_template.py            # assets/template-base.pptx → assets/template.pptx

追加するもの（元のマスター・レイアウトには一切手を加えない。追加だけ）:
  - 元のマスターに「TECHVAN 本文（ラベル付き）」レイアウトを 1 つ追加
    （「ユーザー設定レイアウト」の写しから左上の旗の飾りを外し、
     英字ラベル → 見出し → リード文 の並びに合わせて見出しの位置を変えたもの）
作り直したら check_template.py で元の部分が残っていることを確かめる。
"""
import copy
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from lxml import etree  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.opc.constants import CONTENT_TYPE as CT  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402
from pptx.parts.slide import SlideLayoutPart  # noqa: E402
from pptx.util import Emu, Inches  # noqa: E402

import slidekit as K  # noqa: E402

ASSETS = HERE.parent / "assets"
BASE = ASSETS / "template-base.pptx"
OUT = ASSETS / "template.pptx"


# ---------------------------------------------------------------- ID の払い出し

class IdPool:
    """スライドマスターとレイアウトの ID（2147483648 以上で重複なし）。"""

    def __init__(self, prs):
        used = []
        for el in prs.part._element.iter(qn("p:sldMasterId")):
            used.append(int(el.get("id")))
        for master in prs.slide_masters:
            for el in master._element.iter(qn("p:sldLayoutId")):
                used.append(int(el.get("id")))
        self.next = max(used + [2147483647]) + 1

    def take(self):
        v = self.next
        self.next += 1
        return v


# ---------------------------------------------------------------- 写すときの後始末

def strip_creation_ids(el):
    """複製したときに重複しないよう、図形の作成IDなどの拡張情報を外す。"""
    for ext in list(el.iter(qn("a:extLst"))):
        ext.getparent().remove(ext)
    for ext in list(el.iter(qn("p:extLst"))):
        ext.getparent().remove(ext)


def emu(v):
    return int(Emu(Inches(v)))


# ---------------------------------------------------------------- techvan の本文レイアウト

TECHVAN_SRC = "ユーザー設定レイアウト"
TECHVAN_NEW = "TECHVAN 本文（ラベル付き）"
TECHVAN_FLAG = "直角三角形 30"


def add_techvan_body_layout(prs, ids):
    """元の本文レイアウトの写しを作り、旗の飾りを外して見出しの位置を変える。"""
    master = prs.slide_masters[0]
    src = next(l for l in master.slide_layouts if l.name == TECHVAN_SRC)
    el = copy.deepcopy(src._element)
    strip_creation_ids(el)
    el.find(qn("p:cSld")).set("name", TECHVAN_NEW)
    tree = el.find(qn("p:cSld")).find(qn("p:spTree"))
    for sp in list(tree):
        cnv = sp.find(".//" + qn("p:cNvPr"))
        ph = sp.find(".//" + qn("p:ph"))
        if cnv is not None and cnv.get("name") == TECHVAN_FLAG:
            tree.remove(sp)               # 左上の旗の飾り
        elif ph is not None and ph.get("idx") == "2":
            tree.remove(sp)               # 使わない本文のプレースホルダー
        elif ph is not None and ph.get("type") == "title":
            xfrm = sp.find(".//" + qn("a:xfrm"))
            xfrm.find(qn("a:off")).set("x", str(emu(0.55)))
            xfrm.find(qn("a:off")).set("y", str(emu(0.56)))
            xfrm.find(qn("a:ext")).set("cx", str(emu(12.23)))
            xfrm.find(qn("a:ext")).set("cy", str(emu(0.55)))
            body = sp.find(".//" + qn("a:bodyPr"))
            for k in ("lIns", "rIns"):
                body.set(k, "0")
            body.set("anchor", "ctr")
    partname = prs.part.package.next_partname("/ppt/slideLayouts/slideLayout%d.xml")
    lpart = SlideLayoutPart(partname, CT.PML_SLIDE_LAYOUT, prs.part.package, el)
    # 画像（ロゴ）などの関係を写す。rId は新しく振り直して XML を書き換える
    remap = {}
    for rId, rel in src.part.rels.items():
        if rel.is_external:
            continue
        remap[rId] = lpart.relate_to(rel.target_part, rel.reltype)
    for node in el.iter():
        for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
            v = node.get(attr)
            if v in remap:
                node.set(attr, remap[v])
    rId = master.part.relate_to(lpart, RT.SLIDE_LAYOUT)
    lst = master._element.find(qn("p:sldLayoutIdLst"))
    entry = etree.SubElement(lst, qn("p:sldLayoutId"))
    entry.set("id", str(ids.take()))
    entry.set(qn("r:id"), rId)
    print(f"追加しました: {TECHVAN_NEW}")


def main():
    K.console_utf8()
    prs = Presentation(str(BASE))
    ids = IdPool(prs)
    add_techvan_body_layout(prs, ids)
    prs.save(str(OUT))
    print(f"保存しました: {OUT}")


if __name__ == "__main__":
    main()
