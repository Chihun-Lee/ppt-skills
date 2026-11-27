# -*- coding: utf-8 -*-
"""로고 자리 채우기 — 바탕 파일의 `LOGO:<자리>` 그림을 사용자 로고로 바꾼다.

로고를 고르는 순서 (먼저 찾은 것):
  1) Deck(logo=...) 인자   : 그림 파일 경로 | "placeholder"(LOGO 자리표시) | "none"(로고 없음)
  2) 환경변수 PPT_LOGO      : 위와 같은 값
  3) ~/.ppt_skills/config.json 의 "logo" (scripts/logo.py set/text/none 으로 기록)
  4) 아무것도 없으면 동봉 template/logo_placeholder.png ('LOGO'라고 적힌 그림)

자리(px, 1920×1080 격자)와 정렬은 SLOTS. 로고는 비율을 지킨 채 자리 상자 안에 최대 크기로 넣는다.
"""
from __future__ import annotations

import io
import json
import os
from pathlib import Path

from PIL import Image

SKILL = Path(__file__).resolve().parents[1]
PLACEHOLDER = SKILL / "template" / "logo_placeholder.png"
USER_DIR = Path(os.environ.get("PPT_SKILLS_HOME", Path.home() / ".ppt_skills"))
CONFIG = USER_DIR / "config.json"
EPX = 13439775 / 1920                       # EMU per px (바탕 파일 가로 = 1920px)

# 자리 이름: (x, y, w, h, 가로 정렬 l|c|r) — 세로는 가운데
SLOTS = {
    "header": (1560, 36, 304, 46, "r"),       # 본문 장 오른쪽 위 머리 띠
    "cover": (1280, 320, 520, 220, "c"),      # 표지 오른쪽(제목 아래, 배경 원호 위)
    "divider": (57, 38, 330, 54, "l"),        # 간지 왼쪽 위
    "closing": (760, 76, 400, 64, "c"),       # 감사 장 가운데 위
}
RASTER = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
VECTOR = {".svg", ".pdf", ".ai", ".eps"}


def read_config() -> dict:
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_config(**kw) -> dict:
    cfg = read_config()
    cfg.update(kw)
    USER_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    return cfg


def to_png(src: Path, dst: Path, trim: bool = True, max_px: int = 1600) -> Path:
    """그림·벡터 로고를 투명 PNG 로 바꾼다(여백 잘라냄). SVG·PDF·AI 는 PyMuPDF 로 그린다."""
    src = Path(src)
    ext = src.suffix.lower()
    if ext in RASTER:
        im = Image.open(src)
        im.load()
    elif ext in VECTOR:
        try:
            import pymupdf as fitz  # noqa: F401
        except ImportError:  # PyMuPDF 1.24.3 미만
            import fitz  # noqa: F401
        doc = fitz.open(str(src))
        page = doc[0]
        zoom = max(1.0, max_px / max(page.rect.width, page.rect.height))
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=True)
        im = Image.open(io.BytesIO(pix.tobytes("png")))
        doc.close()
    else:
        raise ValueError(f"로고 형식 {ext} 은(는) 못 읽음 — PNG·JPG·SVG·PDF 로 주세요(EMF·WMF·HWP 안 됨)")
    im = im.convert("RGBA")
    if trim:
        im = _trim(im)
    if max(im.size) > max_px:
        im.thumbnail((max_px, max_px), Image.LANCZOS)
    dst.parent.mkdir(parents=True, exist_ok=True)
    im.save(dst, "PNG", optimize=True)
    return dst


def _trim(im: Image.Image) -> Image.Image:
    """투명 여백 + 흰 여백을 잘라낸다(로고가 자리 상자에 꽉 차게)."""
    alpha = im.getchannel("A")
    box = alpha.point(lambda v: 255 if v > 8 else 0).getbbox()
    if box:
        im = im.crop(box)
    rgb = im.convert("RGB")
    from PIL import ImageChops
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255))).convert("L")
    box = diff.point(lambda v: 255 if v > 12 else 0).getbbox()
    if box and (box[2] - box[0]) * (box[3] - box[1]) < rgb.size[0] * rgb.size[1]:
        pad = 2
        box = (max(0, box[0] - pad), max(0, box[1] - pad), min(im.width, box[2] + pad), min(im.height, box[3] + pad))
        im = im.crop(box)
    return im


def resolve(logo=None, workdir: Path | None = None):
    """쓸 로고 PNG 경로(없음이면 None)와 출처 설명을 돌려준다."""
    src, why = logo, "Deck(logo=)"
    if src is None and os.environ.get("PPT_LOGO"):
        src, why = os.environ["PPT_LOGO"], "환경변수 PPT_LOGO"
    if src is None:
        cfg = read_config()
        if cfg.get("logo"):
            src, why = cfg["logo"], str(CONFIG)
    if src is None or str(src).lower() == "placeholder":
        return PLACEHOLDER, "LOGO 자리표시(로고 미지정)"
    if str(src).lower() == "none":
        return None, f"로고 없음({why})"
    p = Path(src).expanduser()
    if not p.exists():
        print(f"[logo] {why} 의 로고 파일이 없음: {p} — LOGO 자리표시로 대신함")
        return PLACEHOLDER, "LOGO 자리표시(지정 파일 없음)"
    if p.suffix.lower() in (".png", ".jpg", ".jpeg"):
        return p, why
    cache = Path(workdir or Path.cwd()) / "assets" / "_prep" / f"logo_{p.stem}.png"
    if not cache.exists() or cache.stat().st_mtime < p.stat().st_mtime:
        to_png(p, cache)
    return cache, why


def _fit(img: Path, slot: str):
    x, y, w, h, align = SLOTS[slot]
    with Image.open(img) as im:
        iw, ih = im.size
    s = min(w / iw, h / ih)
    nw, nh = iw * s, ih * s
    nx = x if align == "l" else (x + w - nw if align == "r" else x + (w - nw) / 2)
    ny = y + (h - nh) / 2
    return [int(round(v * EPX)) for v in (nx, ny, nw, nh)]


def apply(prs, logo=None, workdir: Path | None = None) -> str:
    """바탕 파일(prs)의 레이아웃·원형 장에 있는 LOGO:<자리> 그림을 바꾼다. 출처 설명을 돌려준다."""
    from pptx.oxml.ns import qn
    png, why = resolve(logo, workdir)
    parts = list(prs.slide_layouts) + list(prs.slides)
    n = 0
    for owner in parts:
        for shp in list(owner.shapes):
            name = shp.name or ""
            if not name.startswith("LOGO:"):
                continue
            slot = name.split(":", 1)[1]
            el = shp._element
            blip = el.find(".//" + qn("a:blip"))
            old = blip.get(qn("r:embed")) if blip is not None else None
            rid = None
            if png is None or slot not in SLOTS:
                el.getparent().remove(el)
            else:
                _, rid = owner.part.get_or_add_image_part(str(png))
                blip.set(qn("r:embed"), rid)
                for ch in list(blip):                  # 원 틀의 색 효과(밝기 100% = 흰 실루엣)·SVG 확장 제거 → 로고 원래 색
                    blip.remove(ch)
                src = el.find(".//" + qn("a:srcRect"))
                if src is not None:
                    src.getparent().remove(src)
                xfrm = el.find(".//" + qn("a:xfrm"))
                off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
                ex, ey, ew, eh = _fit(png, slot)
                off.set("x", str(ex)); off.set("y", str(ey))
                ext.set("cx", str(ew)); ext.set("cy", str(eh))
            if old and old != rid:
                try:
                    owner.part.drop_rel(old)
                except Exception:
                    pass
            n += 1
    return f"{why} · {n}곳" + (f" · {png.name}" if png else "")
