# -*- coding: utf-8 -*-
"""로고 고르기 — 사용자 설정(~/.ppt_skills/config.json)에 기록한다. 덱마다 Deck(logo=...)로 덮어쓸 수도 있다.

사용:
  python logo.py status                 지금 쓰이는 로고와 첫 사용 안내 여부
  python logo.py set <그림 파일>        PNG·JPG·SVG·PDF 로고 등록(여백 잘라 ~/.ppt_skills/logo.png 로 저장)
  python logo.py text "기관 이름"       로고 파일이 없을 때 글자 로고를 만들어 등록 [--color 404040]
  python logo.py placeholder            'LOGO' 자리표시로 되돌리기(기본값)
  python logo.py none                   로고를 아예 넣지 않기
  python logo.py preview [--out 파일]   로고 4곳(머리 띠·표지·간지·감사) 크기 견본 PNG
  python logo.py welcomed               첫 사용 안내를 마쳤다고 기록
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

if os.name == "nt" and not sys.flags.utf8_mode:            # Windows 콘솔 cp949 대비
    os.environ["PYTHONUTF8"] = "1"
    sys.exit(subprocess.call([sys.executable, *sys.argv]))

SKILL = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL / "lib"))
import logo_slots as L  # noqa: E402

USER_LOGO = L.USER_DIR / "logo.png"


def status():
    cfg = L.read_config()
    png, why = L.resolve()
    print(f"설정 파일: {L.CONFIG} ({'있음' if L.CONFIG.exists() else '없음'})")
    print(f"첫 사용 안내: {'마침' if cfg.get('welcomed') else '아직 — 두 종류(다목적 ppt / 실험결과 ppt_experiment)를 소개하고 고르게 할 것'}")
    print(f"로고: {png if png else '없음'}  ← {why}")
    if cfg.get("logo_source"):
        print(f"로고 원본: {cfg['logo_source']}")


def cmd_set(src: str):
    p = Path(src).expanduser().resolve()
    if not p.exists():
        sys.exit(f"파일이 없음: {p}")
    L.to_png(p, USER_LOGO)
    L.write_config(logo=str(USER_LOGO), logo_source=str(p), logo_set_at=datetime.now().isoformat(timespec="seconds"))
    from PIL import Image
    with Image.open(USER_LOGO) as im:
        w, h = im.size
    print(f"로고 등록: {USER_LOGO} ({w}×{h}px, 가로:세로 {w / h:.2f})")
    if max(w, h) < 300:
        print("[주의] 로고 해상도가 낮아 표지에서 흐릴 수 있음 — 더 큰 PNG나 SVG가 있으면 그것으로")


def cmd_text(name: str, color: str = "404040"):
    from PIL import Image, ImageDraw, ImageFont
    font = SKILL / "fonts" / "Paperlogy-7Bold.ttf"
    f = ImageFont.truetype(str(font), 160)
    tmp = Image.new("RGBA", (10, 10))
    l, t, r, b = ImageDraw.Draw(tmp).textbbox((0, 0), name, font=f)
    im = Image.new("RGBA", (r - l + 40, b - t + 40), (0, 0, 0, 0))
    rgb = tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
    ImageDraw.Draw(im).text((20 - l, 20 - t), name, font=f, fill=rgb + (255,))
    L.USER_DIR.mkdir(parents=True, exist_ok=True)
    im.save(USER_LOGO, "PNG")
    L.write_config(logo=str(USER_LOGO), logo_source=f"글자 로고: {name}", logo_set_at=datetime.now().isoformat(timespec="seconds"))
    print(f"글자 로고 등록: {USER_LOGO} ('{name}', #{color})")


def cmd_preview(out: str | None):
    """바탕 파일 표지·간지·본문·감사 4장을 지금 로고로 만들어 PNG 한 장으로 본다(LibreOffice 필요)."""
    import tempfile
    from pptx import Presentation
    tmp = Path(tempfile.mkdtemp())
    prs = Presentation(str(SKILL / "template" / "base.pptx"))
    print("로고:", L.apply(prs, workdir=tmp))
    for sl in list(prs.slides)[1:]:                       # 간지·감사 원형의 자리 글자는 견본에서 뺀다(실제 덱은 엔진이 다시 씀)
        for shp in list(sl.shapes):
            if shp.shape_type == 6 or (shp.has_text_frame and shp.text_frame.text.strip()):
                shp._element.getparent().remove(shp._element)
    layout = next(l for l in prs.slide_layouts if l.name == "슬라이드 1")
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    pptx = tmp / "logo_preview.pptx"
    prs.save(str(pptx))
    from preview import render
    pdf = render(pptx, tmp)
    try:
        import pymupdf as fitz  # noqa: F401
    except ImportError:  # PyMuPDF 1.24.3 미만
        import fitz  # noqa: F401
    from PIL import Image
    doc = fitz.open(str(pdf))
    ims = []
    for pg in doc:
        pix = pg.get_pixmap(dpi=40)
        ims.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    w, h = ims[0].size
    sheet = Image.new("RGB", (w * 2 + 10, h * 2 + 10), (200, 200, 200))
    order = [0, 3, 1, 2]                                  # 표지·본문 / 간지·감사
    for k, i in enumerate(order[:len(ims)]):
        sheet.paste(ims[i], ((k % 2) * (w + 10), (k // 2) * (h + 10)))
    dst = Path(out) if out else Path.cwd() / "logo_preview.png"
    sheet.save(dst)
    print("견본:", dst)


def main():
    a = sys.argv[1:]
    if not a or a[0] in ("-h", "--help"):
        print(__doc__)
        return
    c = a[0]
    if c == "status":
        status()
    elif c == "set" and len(a) >= 2:
        cmd_set(a[1])
    elif c == "text" and len(a) >= 2:
        color = a[a.index("--color") + 1].lstrip("#") if "--color" in a else "404040"
        cmd_text(a[1], color)
    elif c in ("placeholder", "none"):
        L.write_config(logo=c, logo_source="", logo_set_at=datetime.now().isoformat(timespec="seconds"))
        print("로고:", "'LOGO' 자리표시" if c == "placeholder" else "넣지 않음")
    elif c == "preview":
        cmd_preview(a[a.index("--out") + 1] if "--out" in a else None)
    elif c == "welcomed":
        L.write_config(welcomed=True, welcomed_at=datetime.now().isoformat(timespec="seconds"))
        print("첫 사용 안내 완료로 기록")
    else:
        print(__doc__)
        sys.exit(2)


if __name__ == "__main__":
    main()
