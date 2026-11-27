#!/usr/bin/env python
"""시각 점검용 확대: 렌더 PDF 의 한 장에서 영역을 잘라 2배로 저장한다(경계·seam 확인).

사용: zoom.py <render.pdf> <쪽 번호> [x y w h] [--out zoom.png] [--scale 2]
  좌표는 1920×1080 px. 영역을 생략하면 본문 프레임(120,162,1680,846) 전체.
  예: zoom.py render/deck.pdf 5 120 600 900 300   → 5쪽 글↔그림 경계 확대
"""
import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import argparse
from pathlib import Path

try:
    import pymupdf as fitz  # noqa: F401
except ImportError:  # PyMuPDF 1.24.3 미만
    import fitz  # noqa: F401

fitz.TOOLS.mupdf_display_errors(False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("page", type=int)
    ap.add_argument("rect", nargs="*", type=float)
    ap.add_argument("--out")
    ap.add_argument("--scale", type=float, default=2.0)
    a = ap.parse_args()
    x, y, w, h = a.rect if len(a.rect) == 4 else (120, 162, 1680, 846)
    doc = fitz.open(a.pdf)
    page = doc[a.page - 1]
    k = page.rect.width / 1920
    clip = fitz.Rect(x * k, y * k, (x + w) * k, (y + h) * k)
    z = a.scale * 1920 / page.rect.width
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z), clip=clip, alpha=False)
    out = Path(a.out) if a.out else Path(a.pdf).with_name(f"zoom_p{a.page}_{int(x)}_{int(y)}.png")
    pix.save(str(out))
    print(out)


if __name__ == "__main__":
    main()
