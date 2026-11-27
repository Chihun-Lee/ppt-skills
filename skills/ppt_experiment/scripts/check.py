# -*- coding: utf-8 -*-
"""ppt_experiment 자동 검사 — SKILL.md 「통과 기준」 중 기계로 볼 수 있는 것.

사용: python check.py <pptx>
  E1 fail  13pt 미만 글자
  E2 fail  문장형 종결(~다, ~습니다, ~였음 등) / warn ~함·~임·~됨(명사일 수도 있어 사람이 판단)
  E3 warn  불릿 한 줄 길이(한글 24자, 영문 40자 상당 초과)
  E4 warn  한 상자 불릿 5개 이상, 장당 불릿 11개 이상
  E5 fail  그림 비율 왜곡(1% 초과)
  E6 fail  슬라이드 밖으로 나간 요소
  E7 warn  그림자 원인(p:style) 남은 도형
종료 코드: fail 이 있으면 1
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import io
import os
import re
import subprocess
import sys
from pathlib import Path

if os.name == "nt" and not sys.flags.utf8_mode:
    os.environ["PYTHONUTF8"] = "1"
    sys.exit(subprocess.call([sys.executable, *sys.argv]))

from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt

MIN_PT = 13
END_FAIL = re.compile(r"(다|습니다|였음|하였음|되었음)[.。]?$")
END_WARN = re.compile(r"(함|임|됨)[.。]?$")
BULLETS = ("•", "➢")


def width_units(s: str) -> float:
    """한글 1, 그 밖 0.6 (영문 40자 ≈ 한글 24자)."""
    return sum(1.0 if "가" <= c <= "힣" or "㄰" <= c <= "㆏" else 0.6 for c in s)


def runs_text(p) -> str:
    return "".join(r.text for r in p.runs)


def iter_text_frames(shapes):
    for sh in shapes:
        if sh.shape_type == 6:                                 # 그룹
            yield from iter_text_frames(sh.shapes)
        elif sh.has_text_frame:
            yield sh, sh.text_frame
        elif getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield sh, cell.text_frame


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    path = Path(sys.argv[1])
    prs = Presentation(str(path))
    SW, SH = prs.slide_width, prs.slide_height
    fails, warns = [], []
    for si, s in enumerate(prs.slides, start=1):
        n_bul = 0
        for sh, tf in iter_text_frames(s.shapes):
            box_bul = 0
            for p in tf.paragraphs:
                txt = runs_text(p).strip()
                for r in p.runs:
                    if r.font.size is not None and r.font.size < Pt(MIN_PT) and r.text.strip():
                        fails.append(f"{si}쪽 E1 {r.font.size.pt:.0f}pt < {MIN_PT}pt: '{r.text.strip()[:20]}'")
                if not txt:
                    continue
                body = txt.lstrip("".join(BULLETS)).strip()
                if END_FAIL.search(body):
                    fails.append(f"{si}쪽 E2 문장형 종결: '{body[-24:]}'")
                elif END_WARN.search(body) and txt.startswith(BULLETS):
                    warns.append(f"{si}쪽 E2 종결 의심(명사면 무시): '{body[-24:]}'")
                if txt.startswith(BULLETS):
                    box_bul += 1
                    u = width_units(body)
                    if u > 24.5:
                        warns.append(f"{si}쪽 E3 불릿 길이 {u:.0f}(한글 24 기준): '{body[:30]}…'")
            if box_bul > 4:
                warns.append(f"{si}쪽 E4 한 상자 불릿 {box_bul}개 > 4 — 표로 옮길 것")
            n_bul += box_bul
        if n_bul > 10:
            warns.append(f"{si}쪽 E4 장당 불릿 {n_bul}개 > 10")
        for sh in s.shapes:
            if sh.left is not None and (sh.left < -1 or sh.top < -1 or sh.left + sh.width > SW + 9525
                                        or sh.top + sh.height > SH + 9525):
                fails.append(f"{si}쪽 E6 슬라이드 밖: {sh.name}")
            if sh.shape_type == 13:                           # 그림
                try:
                    im = Image.open(io.BytesIO(sh.image.blob))
                    iw, ih = im.size
                    cl, ct, cr, cb = sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom
                    ar = (iw * (1 - cl - cr)) / (ih * (1 - ct - cb))
                    d = (sh.width / sh.height) / ar
                    if abs(d - 1) > 0.01:
                        fails.append(f"{si}쪽 E5 그림 비율 왜곡 {d:.3f}: {sh.name}")
                except Exception:
                    pass
            if sh._element.find(qn("p:style")) is not None and sh.shape_type in (1, 9):
                warns.append(f"{si}쪽 E7 p:style 남음(그림자 원인): {sh.name}")
    for w in warns:
        print("[warn]", w)
    for f in fails:
        print("[FAIL]", f)
    print(f"검사: {len(prs.slides)}장 · fail {len(fails)} · warn {len(warns)}")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
