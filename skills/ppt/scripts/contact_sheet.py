#!/usr/bin/env python
"""PDF → 페이지 PNG(1600px) + 덱 전체 컨택트시트(4열).
사용: contact_sheet.py <pdf> <png_dir> <contact.png> [--width 1600] [--cols 4]
컨택트시트는 덱 리듬(같은 레이아웃 반복, 밀도 편차)을 한 화면에서 보려는 용도.
"""
import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import argparse
from pathlib import Path

try:
    import pymupdf as fitz  # noqa: F401
except ImportError:  # PyMuPDF 1.24.3 미만
    import fitz  # noqa: F401

fitz.TOOLS.mupdf_display_errors(False)  # 구조트리 경고 등 무해한 로그 억제
from PIL import Image, ImageDraw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("png_dir")
    ap.add_argument("contact")
    ap.add_argument("--width", type=int, default=1600)
    ap.add_argument("--cols", type=int, default=4)
    a = ap.parse_args()

    doc = fitz.open(a.pdf)
    out = Path(a.png_dir)
    out.mkdir(parents=True, exist_ok=True)
    thumbs = []
    for i, page in enumerate(doc):
        z = a.width / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
        p = out / f"slide-{i + 1:02d}.png"
        pix.save(str(p))
        im = Image.open(p).convert("RGB")
        im.thumbnail((480, 270))
        thumbs.append(im)

    if not thumbs:
        return
    tw, th = 480, 270
    pad, lab = 12, 22
    cols = min(a.cols, len(thumbs))
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (tw + pad) + pad, rows * (th + pad + lab) + pad), "#DADDE1")
    dr = ImageDraw.Draw(sheet)
    for i, im in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (tw + pad)
        y = pad + r * (th + pad + lab)
        dr.text((x, y), f"{i + 1}", fill="#333333")
        sheet.paste(im, (x, y + lab))
    sheet.save(a.contact)


if __name__ == "__main__":
    main()
