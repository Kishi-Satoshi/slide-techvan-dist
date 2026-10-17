#!/usr/bin/env python3
"""作った PowerPoint を画像にして、一覧（コンタクトシート）を作る。

使い方:
    python preview.py 資料.pptx            # 資料_preview/ に1枚ずつの PNG と一覧を作る
    python preview.py 資料.pptx --dpi 90

LibreOffice（soffice）が必要。PDF を画像にするのに pdftoppm（poppler）か
PyMuPDF（python -m pip install pymupdf）のどちらかを使う。
どれもない場合はその旨を表示して終了する（そのときは PowerPoint で開いて目で確認する）。
注意: 游ゴシックなどのフォントが入っていない環境では、代わりのフォントで描かれるため、
文字幅や改行位置が実際の PowerPoint と少し違うことがある。
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path



def _console_utf8():
    """日本語版 Windows でも表示が文字化けしないよう、出力を UTF-8 にする。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(p).exists():
            return p
    return None


def pdf_rasterizer():
    """PDF を 1 ページずつ PNG にする方法（pdftoppm → PyMuPDF の順に探す）。"""
    if shutil.which("pdftoppm"):
        def run(pdf, out, dpi):
            subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(out / "slide")],
                           check=True, timeout=300)
        return run
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return None

    def run(pdf, out, dpi):
        doc = fitz.open(str(pdf))
        width = len(str(len(doc)))
        for i, page in enumerate(doc, start=1):
            page.get_pixmap(dpi=dpi).save(str(out / f"slide-{i:0{width}d}.png"))
    return run


def main():
    _console_utf8()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--dpi", type=int, default=80)
    ap.add_argument("--out", help="出力フォルダ（既定: <資料名>_preview）")
    args = ap.parse_args()

    src = Path(args.pptx).resolve()
    out = Path(args.out) if args.out else src.with_name(src.stem + "_preview")
    soffice = find_soffice()
    rasterizer = pdf_rasterizer()
    if not soffice or not rasterizer:
        missing = "LibreOffice" if not soffice else "pdftoppm または PyMuPDF"
        print(f"{missing} が見つからないため、画像にできませんでした。"
              "PowerPoint で開いて確認してください。")
        return 3
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        # LibreOffice を開いていても動くように、毎回べつの設定フォルダを使う
        profile = Path(tmp, "profile").as_uri()
        try:
            subprocess.run([soffice, f"-env:UserInstallation={profile}", "--headless",
                            "--convert-to", "pdf", "--outdir", tmp, str(src)],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=300)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            print("LibreOffice で PDF にできませんでした。PowerPoint で開いて確認してください。")
            return 3
        pdf = Path(tmp) / (src.stem + ".pdf")
        for old in out.glob("slide-*.png"):
            old.unlink()
        rasterizer(pdf, out, args.dpi)
    pages = sorted(out.glob("slide-*.png"))
    try:
        from PIL import Image
    except ImportError:
        print(f"{len(pages)} 枚の画像を作りました: {out}")
        return 0
    ims = [Image.open(p) for p in pages]
    if ims:
        w, h = ims[0].size
        cols = 3 if len(ims) > 4 else 2
        rows = (len(ims) + cols - 1) // cols
        gap = 12
        sheet = Image.new("RGB", (cols * w + (cols - 1) * gap, rows * h + (rows - 1) * gap), "#9AA0A6")
        for i, im in enumerate(ims):
            sheet.paste(im, ((i % cols) * (w + gap), (i // cols) * (h + gap)))
        sheet.save(out / "overview.png")
    print(f"{len(pages)} 枚の画像と一覧（overview.png）を作りました: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
