# -*- coding: utf-8 -*-
"""kims_ppt — KIMS PPT v4.1 엔진. STANDARD.md(v4.1)의 수치를 코드로 옮긴 단일 SSOT.

덱마다 이 파일을 복사하지 않는다. build.py 에서:

    import sys; sys.path.insert(0, r"<스킬 폴더>/lib")
    from kims_ppt import Deck

핵심 보장
  * 좌표는 1920×1080 px 격자, 글자 크기는 PowerPoint pt (이 캔버스에서 1pt = 1.81432px)
  * 줄바꿈은 엔진이 어절 경계에서 계산해 a:br 로 확정한다. 줄 폭은 Paperlogy(커닝 유·무)와
    맑은 고딕 중 가장 넓은 값 ≤ 상자 × 0.96 — Paperlogy 가 없는 PC에서도 다시 줄바꿈되지 않는다
  * 텍스트박스 높이 = 계산한 블록 높이. 모듈이 영역에 안 들어가면 LayoutError(조용한 넘침·자동 축소 없음)
  * 폰트는 Paperlogy(영문 전체명, b=0, 굵기는 이름). Paperlogy 에 없는 그리스 문자·수학 기호는 Arial run
  * 그림은 비율 보존, EXIF 회전 반영, 흰 테두리 트림. 배치는 그림이 가장 커지는 안을 고른다
  * 도형의 p:style 제거(그림자 방지), 모든 a:br·문단 끝에 서식(PowerPoint 줄 높이 보호)
"""
from __future__ import annotations

import copy
import functools
import inspect
import json
import math
import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from PIL import Image, ImageChops, ImageFont, ImageOps
from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

V4 = Path(__file__).resolve().parents[1]
_ENGINE_MTIME = Path(__file__).stat().st_mtime
BASE = V4 / "template" / "base.pptx"
FONT_DIRS_ENV = "KIMS_PPT_FONT_DIRS"


def _font_dirs() -> list:
    """폰트 탐색 순서: 환경변수 KIMS_PPT_FONT_DIRS → 스킬 동봉 fonts/ → OS 폰트 폴더(Mac·Windows·Linux).
    동봉 Paperlogy 는 글 폭 측정용이다. PowerPoint 에 보이려면 OS 에도 설치해야 한다(README)."""
    import os
    import sys as _sys
    ds = [Path(p) for p in os.environ.get(FONT_DIRS_ENV, "").split(os.pathsep) if p]
    ds.append(V4 / "fonts")
    home = Path.home()
    if _sys.platform == "darwin":
        ds += [home / "Library/Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
               Path("/System/Library/Fonts")]
        ds += sorted(Path("/Applications").glob("Microsoft */Contents/Resources/DFonts"))   # Office 동봉 맑은 고딕
    elif os.name == "nt":
        ds += [Path(os.environ.get("LOCALAPPDATA", str(home / "AppData/Local"))) / "Microsoft/Windows/Fonts",
               Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
    else:
        ds += [home / ".local/share/fonts", home / ".fonts", Path("/usr/local/share/fonts"), Path("/usr/share/fonts")]
    return [d for d in ds if d.is_dir()]


FONT_DIRS = _font_dirs()


@functools.lru_cache(maxsize=None)
def find_font(*names: str):
    """이름 목록 중 처음 찾은 폰트 파일 경로(없으면 None). 1차 정확한 이름, 2차 하위 폴더·대소문자 무시."""
    for d in FONT_DIRS:
        for n in names:
            p = d / n
            if p.exists():
                return p
    low = {n.lower() for n in names}
    for d in FONT_DIRS:
        try:
            for p in d.rglob("*"):
                if p.name.lower() in low:
                    return p
        except OSError:
            continue
    return None

# ─────────────────────────────────────────────────────────────────────────────
# 1. 캔버스·토큰 (STANDARD §2~§10)
# ─────────────────────────────────────────────────────────────────────────────
SW, SH = 13439775, 7559675
EPX = SW / 1920.0                      # EMU per px = 6999.883
PX_PT = 12700 / EPX                    # px per pt = 1.81432
PX_PER_IN = 1920 / (SW / 914400)       # 130.63 px/in


def E(v: float) -> Emu:
    return Emu(int(round(v * EPX)))


# 덱 종류별 토큰 (§5.2, §4.2)
TOK = {
    "ppt": dict(title=28, key=24, body=20, caption=16, label=16, th=16, td=16, td_dense=14, pill=16,
                toc_num=24, toc_ch=24, toc_sec=20, refs=14, source=10,
                key_body_pt=8, item_pt=4, group_pt=16, block=48, cap_gap=14,
                card_px=32, card_py=28, pill_h=52, cell_px=12, cell_py=10,
                key_lines=1, key_lines_sum=4, body_lines=4, body_wrap=1,
                narr_warn=100, narr_fail=140, vis_warn=180, vis_fail=250, narr_sum=140),
    "report": dict(title=28, key=20, body=16, caption=14, label=14, th=14, td=14, td_dense=14, pill=14,
                   toc_num=24, toc_ch=24, toc_sec=20, refs=14, source=10,
                   key_body_pt=6, item_pt=4, group_pt=12, block=40, cap_gap=14,
                   card_px=32, card_py=28, pill_h=46, cell_px=10, cell_py=8,
                   key_lines=3, key_lines_sum=5, body_lines=6, body_wrap=2,
                   narr_warn=220, narr_fail=280, vis_warn=260, vis_fail=300, narr_sum=240),
}
FIXED = dict(cover=40, cover_sub=28, cover_info=20, div_label=17, div_num=80, div_title=37,
             chapter=18, crumb=13, closing=48, div_prog=14)
C = dict(
    GREEN="71BF44", PILL="546882", PILL_ASIS="808080", ORANGE="F15922",
    STEP1="E2F0D9", STEP2="E0EFF3", STEP3="DDE8F4", TABLE_HEAD="E2F0D9", SUMMARY="F2F8EC",
    INK="262626", TEXT2="404040", GRAY="808080", MUTED="A6A6A6", LINE="BFBFBF",
    LINE_LIGHT="D9D9D9", WHITE="FFFFFF", NUM_LIGHT="C6E5B3", CLOSING="505666",
)
STEPS = [C["STEP1"], C["STEP2"], C["STEP3"]]
GUTTER = {2: 48, 3: 40, 4: 32}
FRAME_X0, FRAME_X1, FRAME_Y0, FRAME_Y1 = 120, 1800, 162, 1008
WIDE_X0, WIDE_X1, WIDE_Y1 = 80, 1840, 1016
TITLE_X, TITLE_MAXW = 120, 1307
SLACK = 0.96
MIN_GAP = 8


class LayoutError(Exception):
    """모듈이 영역에 안 들어감 — 문안을 줄이거나 장을 나눌 것(글자 크기는 줄이지 않는다)."""


class WrapError(LayoutError):
    """한 줄에 들어가야 하는 글이 넘침 — 문안을 줄일 것."""


# ─────────────────────────────────────────────────────────────────────────────
# 2. 폰트·측정
# ─────────────────────────────────────────────────────────────────────────────
WEIGHTS = {3: "Light", 4: "Regular", 6: "SemiBold", 7: "Bold", 8: "ExtraBold"}


def face(w: int) -> str:
    """최종 pptx 표기 = PowerPoint 가 저장하는 영문 전체 패밀리명(미리보기는 preview.py 가 한글명으로)."""
    return f"Paperlogy {w} {WEIGHTS[w]}"


@functools.lru_cache(maxsize=None)
def ttf(w: int) -> Path:
    name = f"Paperlogy-{w}{WEIGHTS[w]}.ttf"
    p = find_font(name)
    if p is None:
        raise FileNotFoundError(f"{name} 를 찾지 못함 — 스킬 폴더 fonts/ 가 있는지 확인(README 설치 2단계)")
    return p


MALGUN = find_font("malgun.ttf", "Malgun Gothic.ttf")
MALGUN_B = find_font("malgunbd.ttf", "Malgun Gothic Bold.ttf")
ARIAL = find_font("Arial.ttf", "arial.ttf", "LiberationSans-Regular.ttf", "Arimo-Regular.ttf")
ARIAL_B = find_font("Arial Bold.ttf", "arialbd.ttf", "LiberationSans-Bold.ttf", "Arimo-Bold.ttf")
if ARIAL is None or ARIAL_B is None:
    raise FileNotFoundError("Arial(또는 Liberation Sans) 폰트를 찾지 못함 — 그리스 문자·기호 폭 측정에 필요")
# 맑은 고딕이 없으면(Office 없는 Mac 등) Windows 대체 폰트 폭을 Paperlogy × 1.22 로 어림한다(실측 최대 1.218, 보수적)
MALGUN_RATIO = 1.22
if MALGUN is None or MALGUN_B is None:
    import warnings as _w
    _w.warn("맑은 고딕을 찾지 못해 Windows 대체 폰트 폭을 Paperlogy×1.22 로 어림합니다(줄이 조금 짧아질 수 있음)")

_PIL: dict = {}
_CMAP: dict = {}


def _font(path: Path, px: float):
    key = (str(path), round(px * 100))
    f = _PIL.get(key)
    if f is None:
        f = _PIL[key] = ImageFont.truetype(str(path), px)
    return f


def _cmap(path: Path):
    f = _CMAP.get(str(path))
    if f is None:
        try:
            import pymupdf as fitz  # noqa: F401
        except ImportError:  # PyMuPDF 1.24.3 미만
            import fitz  # noqa: F401
        f = _CMAP[str(path)] = fitz.Font(fontfile=str(path))
    return f


def in_paperlogy(ch: str, w: int = 4) -> bool:
    return bool(_cmap(ttf(w)).has_glyph(ord(ch)))


def in_arial(ch: str) -> bool:
    return bool(_cmap(ARIAL).has_glyph(ord(ch)))


def text_w(s: str, w: int, pt: float, family: str = "pl") -> float:
    """렌더 폭(px): 커닝 켠 PIL RAQM, 크기 = pt × 1.81432 (LibreOffice 대비 ±0.4%, 실험 E4)."""
    if not s:
        return 0.0
    path = ttf(w) if family == "pl" else (ARIAL_B if w >= 6 else ARIAL)
    return _font(path, pt * PX_PT).getlength(s)


def fit_w(s: str, w: int, pt: float, family: str = "pl") -> float:
    """줄 맞춤 판정 폭 = max(Paperlogy 커닝, 커닝 없음, 맑은 고딕) — §5.1.1 폴백 보장, §7.3."""
    if not s:
        return 0.0
    px = pt * PX_PT
    if family == "pl":
        f = _font(ttf(w), px)
        a = f.getlength(s)
        try:
            b = f.getlength(s, features=["-kern"])      # libraqm 이 없는 Pillow 에서는 커닝 끔 측정을 건너뛴다
        except Exception:
            b = a
    else:
        f = _font(ARIAL_B if w >= 6 else ARIAL, px)
        a = b = f.getlength(s)
    mg = MALGUN_B if w >= 6 else MALGUN
    m = _font(mg, px).getlength(s) if mg is not None else max(a, b) * MALGUN_RATIO
    return max(a, b, m)


def ink_left(text: str, w: int, pt: float) -> float:
    """첫 글자 잉크의 왼쪽 여백(px). 크기가 다른 글을 같은 왼쪽선에 맞출 때 쓴다(간지).
    getbbox()는 이 Pillow(12.x)·Paperlogy 조합에서 x 값으로 전진폭 상자를 돌려줘서 쓰지 않는다 → getmask2 잉크."""
    if not text:
        return 0.0
    m, off = _font(ttf(w), pt * PX_PT).getmask2(text[0])
    b = m.getbbox()
    return off[0] + b[0] if b else 0.0


def pitch(pt: float, spacing: float = 1.0) -> float:
    """줄 피치(px) = 1.2 × em × 배수 (LibreOffice 실측, 폰트 무관)."""
    return 1.2 * pt * PX_PT * spacing


NORMALIZE = [("μ", "µ"), ("℃", "°C"), ("⇒", "→"), (" ", " "), ("​", ""), ("\t", " "),
             ("▪", "■"), ("✓", "V"), ("★", "*")]
SUPERSCRIPT = {"²": "2", "³": "3", "¹": "1", "⁻": "-", "⁺": "+"}
SUBSCRIPT = {"₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4", "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9"}


def normalize(s: str) -> str:
    for a, b in NORMALIZE:
        s = s.replace(a, b)
    return s


# ─────────────────────────────────────────────────────────────────────────────
# 3. 스팬(부분 서식)과 줄바꿈
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class Span:
    text: str
    w: int = 4
    color: str = C["INK"]
    baseline: int = 0          # 30000 위첨자 / -25000 아래첨자
    family: str = "pl"         # "pl" Paperlogy | "arial" (그리스 문자·수학 기호)


def to_spans(text, w: int, color: str) -> list[Span]:
    """문자열 또는 [(text, {w,color}), ...] → Span. 위·아래첨자와 Paperlogy 에 없는 글자(Arial)를 분리."""
    raw = []
    if isinstance(text, str):
        raw = [Span(normalize(text), w, color)]
    else:
        for item in text:
            if isinstance(item, Span):
                raw.append(Span(normalize(item.text), item.w, item.color, item.baseline))
            elif isinstance(item, tuple):
                t, st = item
                raw.append(Span(normalize(t), st.get("w", w), st.get("color", color), st.get("baseline", 0)))
            else:
                raw.append(Span(normalize(item), w, color))
    out = []
    for sp in raw:
        buf, key = "", None
        for ch in sp.text:
            if ch in SUPERSCRIPT:
                k, c2 = (30000, "pl"), SUPERSCRIPT[ch]
            elif ch in SUBSCRIPT:
                k, c2 = (-25000, "pl"), SUBSCRIPT[ch]
            elif ch in " \u00a0\n" or in_paperlogy(ch, sp.w):
                k, c2 = (sp.baseline, "pl"), ch
            elif in_arial(ch):
                k, c2 = (sp.baseline, "arial"), ch
            else:
                raise WrapError(f"'{ch}'(U+{ord(ch):04X}) 는 Paperlogy·Arial 어디에도 없음 — 다른 글자로 바꿀 것")
            if key is not None and k != key and buf:
                out.append(Span(buf, sp.w, sp.color, key[0], key[1]))
                buf = ""
            key = k
            buf += c2
        if buf:
            out.append(Span(buf, sp.w, sp.color, key[0], key[1]))
    return out


def spans_fit(spans: list[Span], pt: float) -> float:
    return sum(fit_w(s.text, s.w, pt * (0.65 if s.baseline else 1.0), s.family) for s in spans)


def spans_w(spans: list[Span], pt: float) -> float:
    return sum(text_w(s.text, s.w, pt * (0.65 if s.baseline else 1.0), s.family) for s in spans)


def slice_spans(spans: list[Span], a: int, b: int) -> list[Span]:
    out, pos = [], 0
    for s in spans:
        s0, s1 = pos, pos + len(s.text)
        lo, hi = max(a, s0), min(b, s1)
        if lo < hi:
            out.append(Span(s.text[lo - s0:hi - s0], s.w, s.color, s.baseline, s.family))
        pos = s1
    return out


_NUM = re.compile(r"^[~±+\-(]?\d[\d.,]*(\s*[~\-–]\s*\d[\d.,]*)?[)%]?$")
_UNIT = re.compile(r"^(mm/s|mm/min|m/s|g/cm[3³]|J/mm[23²³]|W/mK|°C/s|K/s|"
                   r"%|‰|°C?|µ?m|nm|mm|cm|km|kg|g|mg|µg|t|MPa|GPa|kPa|Pa|W|kW|mW|V|kV|mV|A|mA|Hz|kHz|MHz|"
                   r"min|h|s|ms|µs|K|J|kJ|J/mm\d?|L|mL|rpm|dpi|px|pt|wt%|at%|개|장|건|회|년|개월|월|일|명|종|배|"
                   r"초|분|시간|점|쌍|차|단계|층|편|권|곳|대|원|억|만|천|조|배속|조건|시편|셀|"
                   r"epoch|epochs|ep|iter|GB|GiB|MB|MiB|TB|KB|instance|instances|fps|FPS|img|images|장면|클래스|HV|HRC|HB|%p|ea)"
                   r"(?:[)\],.]|[가-힣]{1,2})?$")          # 단위 뒤 조사(MPa로, 개를)도 한 덩어리
_PURE_NUM = re.compile(r"^[~±+\-−(]?\d[\d.,]*[%)]?[,]?$")
_LABEL = re.compile(r"^[A-Za-z_가-힣][\w\-가-힣]{0,7}$")


def _words(plain: str):
    """공백 기준 어절 (시작, 끝). NBSP 는 어절 안쪽(묶기). 숫자+단위는 한 덩어리."""
    out, i, n = [], 0, len(plain)
    while i < n:
        while i < n and plain[i] == " ":
            i += 1
        if i >= n:
            break
        j = i
        while j < n and plain[j] != " ":
            j += 1
        out.append([i, j])
        i = j
    merged = []
    for w in out:                  # 1) 숫자 + 단위
        if merged:
            prev = plain[merged[-1][0]:merged[-1][1]]
            cur = plain[w[0]:w[1]]
            if _NUM.match(prev) and _UNIT.match(cur):
                merged[-1][1] = w[1]
                continue
        merged.append(w)
    out2 = []
    for w in merged:               # 2) 짧은 라벨 + 값('경도 45', 'scale 0.2', '속도 1,320 mm/s')
        if out2:
            prev = plain[out2[-1][0]:out2[-1][1]]
            cur = plain[w[0]:w[1]]
            head = cur.split(" ", 1)
            is_val = _PURE_NUM.match(cur) or (len(head) == 2 and _NUM.match(head[0]) and _UNIT.match(head[1]))
            if _LABEL.match(prev) and is_val:
                out2[-1][1] = w[1]
                continue
        out2.append(w)
    return out2


def wrap(spans: list[Span], pt: float, avail: float, max_lines: int | None = None, what: str = ""):
    """어절 경계에서만 줄을 나눈다. greedy 로 최소 줄 수를 구한 뒤 같은 줄 수 안에서 균형 분할(고아 줄 방지)."""
    plain = "".join(s.text for s in spans)
    lim = avail * SLACK
    words = _words(plain)
    if not words:
        return []
    cache = {}

    def width(a, b):
        k = (a, b)
        if k not in cache:
            cache[k] = spans_fit(slice_spans(spans, a, b), pt)
        return cache[k]

    soft = []                      # 묶은 덩어리(숫자+단위, 라벨+값)가 줄 폭보다 넓으면 원래 어절로 되돌린다
    for a, b in words:
        if width(a, b) > lim and " " in plain[a:b]:
            i = a
            for part in plain[a:b].split(" "):
                if part:
                    soft.append([i, i + len(part)])
                i += len(part) + 1
        else:
            soft.append([a, b])
    words = soft
    for a, b in words:
        if width(a, b) > lim:
            raise WrapError(f"{what}: '{plain[a:b]}' 한 어절이 폭({avail:.0f}px)보다 김 — 문안을 줄이거나 칸을 넓힐 것")
    lines, cur = [], 0
    while cur < len(words):
        j = cur
        while j + 1 < len(words) and width(words[cur][0], words[j + 1][1]) <= lim:
            j += 1
        lines.append((cur, j))
        cur = j + 1
    n = len(lines)
    if max_lines is not None and n > max_lines:
        raise WrapError(f"{what}: {n}줄 (허용 {max_lines}줄) — '{plain[:40]}…' 문안을 줄일 것")
    if n >= 2:
        m = len(words)
        INF = float("inf")
        best = [[INF] * (m + 1) for _ in range(n + 1)]
        back = [[-1] * (m + 1) for _ in range(n + 1)]
        best[0][0] = 0.0
        for k in range(1, n + 1):
            for j in range(1, m + 1):
                for i in range(k - 1, j):
                    if best[k - 1][i] == INF:
                        continue
                    wd = width(words[i][0], words[j - 1][1])
                    if wd > lim:
                        continue
                    cost = best[k - 1][i] + (lim - wd) ** 2
                    if k < n and plain[words[j - 1][1] - 1] in ",·":   # 쉼표 뒤에서 끊는 줄 선호
                        cost -= 0.08 * lim * lim
                    if cost < best[k][j]:
                        best[k][j], back[k][j] = cost, i
        if best[n][m] < INF:
            cuts, j = [], m
            for k in range(n, 0, -1):
                i = back[k][j]
                cuts.append((i, j - 1))
                j = i
            lines = list(reversed(cuts))
    return [(words[a][0], words[b][1]) for a, b in lines]


def is_orphan(spans, pt, lines, avail) -> bool:
    if len(lines) < 2:
        return False
    plain = "".join(s.text for s in spans)
    a, b = lines[-1]
    last = plain[a:b]
    return (" " not in last.strip()) or spans_fit(slice_spans(spans, a, b), pt) < 0.3 * avail * SLACK


# ─────────────────────────────────────────────────────────────────────────────
# 4. 기하
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class Rect:
    x: float
    y: float
    w: float
    h: float

    @property
    def r(self):
        return self.x + self.w

    @property
    def b(self):
        return self.y + self.h

    @property
    def cx(self):
        return self.x + self.w / 2

    @property
    def cy(self):
        return self.y + self.h / 2

    def inset(self, dx, dy=None):
        dy = dx if dy is None else dy
        return Rect(self.x + dx, self.y + dy, self.w - 2 * dx, self.h - 2 * dy)

    def hsplit(self, weights, gutter=None):
        n = len(weights)
        g = gutter if gutter is not None else GUTTER.get(n, 32)
        tot = sum(weights)
        usable = self.w - g * (n - 1)
        out, x = [], self.x
        for wt in weights:
            ww = usable * wt / tot
            out.append(Rect(x, self.y, ww, self.h))
            x += ww + g
        return out

    def vsplit(self, weights, gutter=40):
        n = len(weights)
        tot = sum(weights)
        usable = self.h - gutter * (n - 1)
        out, y = [], self.y
        for wt in weights:
            hh = usable * wt / tot
            out.append(Rect(self.x, y, self.w, hh))
            y += hh + gutter
        return out

    def top(self, h):
        return Rect(self.x, self.y, self.w, h)

    def as_list(self):
        return [round(self.x, 1), round(self.y, 1), round(self.w, 1), round(self.h, 1)]


FRAME = Rect(FRAME_X0, FRAME_Y0, FRAME_X1 - FRAME_X0, FRAME_Y1 - FRAME_Y0)


# ─────────────────────────────────────────────────────────────────────────────
# 5. OOXML 저수준
# ─────────────────────────────────────────────────────────────────────────────
_RPR_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "effectLst",
              "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
              "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
_PPR_ORDER = ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx",
              "buFont", "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst"]


def _reorder(el, order):
    kids = list(el)
    kids.sort(key=lambda e: order.index(etree.QName(e).localname) if etree.QName(e).localname in order else 99)
    for k in kids:
        el.remove(k)
    for k in kids:
        el.append(k)


def _fill_rpr(rPr, pt, color, w, family="pl", baseline=0):
    rPr.set("lang", "ko-KR")
    rPr.set("altLang", "en-US")
    rPr.set("sz", str(int(round(pt * 100))))
    rPr.set("b", "1" if (family == "arial" and w >= 6) else "0")
    rPr.set("i", "0")
    rPr.set("kern", "100")
    rPr.set("dirty", "0")
    if baseline:
        rPr.set("baseline", str(baseline))
    for tag in ("solidFill", "latin", "ea", "cs"):
        for el in rPr.findall(qn("a:" + tag)):
            rPr.remove(el)
    sf = etree.SubElement(rPr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    fc = face(w) if family == "pl" else "Arial"
    for tag in ("latin", "ea", "cs"):
        etree.SubElement(rPr, qn("a:" + tag)).set("typeface", fc)
    _reorder(rPr, _RPR_ORDER)


def _style_run(r, sp: Span, pt: float):
    _fill_rpr(r._r.get_or_add_rPr(), pt, sp.color, sp.w, sp.family, sp.baseline)


def _add_br(p, pt, color, w):
    """a:br 에도 서식을 넣는다(PowerPoint 가 기본 18pt 를 상속해 줄이 벌어지는 것 방지)."""
    p.add_line_break()
    br = p._p.findall(qn("a:br"))[-1]
    rPr = br.find(qn("a:rPr"))
    if rPr is None:
        rPr = etree.SubElement(br, qn("a:rPr"))
    _fill_rpr(rPr, pt, color, w)


def _end_para(p, pt, color, w):
    epr = p._p.find(qn("a:endParaRPr"))
    if epr is None:
        epr = etree.SubElement(p._p, qn("a:endParaRPr"))
    _fill_rpr(epr, pt, color, w)


def _style_para(p, align="l", before_pt=0.0, bullet=None, mar_px=0.0, ind_px=0.0,
                bu_color=None, bu_pct=100, bu_w=4):
    p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
    p.line_spacing = 1.0
    p.space_before = Pt(before_pt)
    p.space_after = Pt(0)
    pPr = p._p.get_or_add_pPr()
    pPr.set("eaLnBrk", "0")
    pPr.set("latinLnBrk", "0")
    pPr.set("marL", str(int(round(mar_px * EPX))))
    pPr.set("indent", str(int(round(-ind_px * EPX))) if (bullet or ind_px) else "0")
    for tag in ("buNone", "buChar", "buFont", "buClr", "buSzPct", "buAutoNum"):
        for el in pPr.findall(qn("a:" + tag)):
            pPr.remove(el)
    if bullet:
        if bu_color:
            bc = etree.SubElement(pPr, qn("a:buClr"))
            etree.SubElement(bc, qn("a:srgbClr")).set("val", bu_color)
        etree.SubElement(pPr, qn("a:buSzPct")).set("val", str(int(bu_pct * 1000)))
        etree.SubElement(pPr, qn("a:buFont")).set("typeface", face(bu_w))
        etree.SubElement(pPr, qn("a:buChar")).set("char", bullet)
    else:
        etree.SubElement(pPr, qn("a:buNone"))
    _reorder(pPr, _PPR_ORDER)


def _strip_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _textbox(slide, rect: Rect, name: str, anchor="t"):
    tb = slide.shapes.add_textbox(E(rect.x), E(rect.y), E(rect.w), E(max(rect.h, 1)))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    bp = tf._txBody.find(qn("a:bodyPr"))
    for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
        for el in bp.findall(qn(tag)):
            bp.remove(el)
    etree.SubElement(bp, qn("a:noAutofit"))
    bp.set("anchor", {"t": "t", "m": "ctr", "b": "b"}[anchor])
    bp.set("rtlCol", "0")
    return tb


def _rgb(hexs):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(hexs)


def _fill_line(shape, fill=None, line=None, line_pt=1.0):
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = _rgb(line)
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()


def _shape(slide, kind, rect: Rect, name, fill=None, line=None, line_pt=1.0, radius=None):
    sh = slide.shapes.add_shape(kind, E(rect.x), E(rect.y), E(rect.w), E(rect.h))
    sh.name = name
    _strip_style(sh)
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = min(0.5, radius / max(1, min(rect.w, rect.h)))
    _fill_line(sh, fill, line, line_pt)
    return sh


# ─────────────────────────────────────────────────────────────────────────────
# 6. 그림
# ─────────────────────────────────────────────────────────────────────────────
def _load_rgb(src: Path, trim=True) -> Image.Image:
    """EXIF 회전 반영 → 알파는 흰 배경 합성 → 균일 테두리 트림(>1.5%)."""
    im = ImageOps.exif_transpose(Image.open(src))
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        bg.alpha_composite(im)
        im = bg.convert("RGB")
    else:
        im = im.convert("RGB")
    if trim:
        corner = im.getpixel((0, 0))
        diff = ImageChops.difference(im, Image.new("RGB", im.size, corner)).convert("L").point(lambda v: 255 if v > 12 else 0)
        box = diff.getbbox()
        if box:
            l, t, r, b = box
            if (l + (im.width - r)) / im.width > 0.015 or (t + (im.height - b)) / im.height > 0.015:
                pad = 6
                im = im.crop((max(0, l - pad), max(0, t - pad), min(im.width, r + pad), min(im.height, b + pad)))
    return im


_DIMS: dict = {}


def img_dims(src, trim=True) -> tuple[int, int]:
    """배치 계산용 크기(EXIF 회전·트림 반영). 파일을 쓰지 않는다."""
    src = Path(src)
    key = (str(src), src.stat().st_mtime, trim)
    if key not in _DIMS:
        im = _load_rgb(src, trim)
        _DIMS[key] = (im.width, im.height)
    return _DIMS[key]


def prep_image(src: Path, cache_dir: Path, trim=True, transparent=False, jpeg=False) -> tuple[Path, int, int]:
    """삽입용 사본을 만든다 → (경로, 폭, 높이). 사진은 jpeg=True(품질 90), 도표는 PNG.
    transparent=True 면 가장자리와 이어진 근백색(모든 채널 ≥ 244) 영역만 투명하게 한다(4-연결 라벨링).
    차트·도식이 템플릿 배경 무늬 위에서 흰 사각형으로 드러나지 않게 하려는 것이고, 글자·선 안쪽의 흰 틈과
    검은 픽셀은 건드리지 않는다. 사진에는 쓰지 않는다."""
    src = Path(src)
    cache_dir.mkdir(parents=True, exist_ok=True)
    dst = cache_dir / (src.stem + ("_v4t.png" if transparent else ("_v4.jpg" if jpeg else "_v4.png")))
    if dst.exists() and dst.stat().st_mtime >= src.stat().st_mtime and dst.stat().st_mtime >= _ENGINE_MTIME:
        with Image.open(dst) as im:
            return dst, im.width, im.height
    im = _load_rgb(src, trim)
    if transparent:
        import numpy as np
        from scipy import ndimage
        a = np.asarray(im)
        white = a.min(axis=2) >= 253          # 순백만(옅은 데이터 색·컬러맵 끝 띠는 남긴다)
        lab, _ = ndimage.label(white)
        edge = np.unique(np.concatenate([lab[0], lab[-1], lab[:, 0], lab[:, -1]]))
        edge = edge[edge > 0]
        alpha = np.where(np.isin(lab, edge), 0, 255).astype(np.uint8)
        rgba = im.convert("RGBA")
        rgba.putalpha(Image.fromarray(alpha))
        rgba.save(dst, dpi=(144, 144))
        return dst, im.width, im.height
    if jpeg:
        im.save(dst, quality=90, dpi=(144, 144))
    else:
        im.save(dst, dpi=(144, 144))
    return dst, im.width, im.height


def contain(img_w, img_h, box: Rect, align="c") -> Rect:
    s = min(box.w / img_w, box.h / img_h)
    w, h = img_w * s, img_h * s
    x = {"c": box.x + (box.w - w) / 2, "l": box.x, "r": box.r - w}[align]
    return Rect(x, box.y + (box.h - h) / 2, w, h)


def fit_photo(img_w, img_h, box: Rect, max_crop=0.20):
    """사진용: 슬롯 비율에 맞춰 가운데 cover 크롭(한 변 ≤ 20%) 후 contain. 반환 (그림 rect, (l,t,r,b) 크롭)."""
    ir, br = img_w / img_h, box.w / box.h
    cl = ct = 0.0
    if ir > br:     # 가로로 더 김 → 좌우 크롭
        cl = min(max_crop, (1 - br / ir) / 2)
    else:
        ct = min(max_crop, (1 - ir / br) / 2)
    vw, vh = img_w * (1 - 2 * cl), img_h * (1 - 2 * ct)
    r = contain(vw, vh, box)
    return r, (cl, ct, cl, ct)


# ─────────────────────────────────────────────────────────────────────────────
# 7. 텍스트 블록 (여러 문단 = 한 텍스트박스, ST-06)
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class Para:
    text: object
    pt: float
    w: int = 4
    color: str = C["TEXT2"]
    bullet: str | None = None      # "■"(요지) / "•"(세부) / None
    level: int = 0                 # 1 = 요지 글자 시작선까지 들여씀
    before_pt: float = 0.0
    align: str = "l"
    max_lines: int | None = None
    role: str = "body"
    hang_px: float = 0.0           # 글머리 없는 매달린 들여쓰기(References '[n] ' 폭)
    pad_px: float = 0.0            # 모든 줄 왼쪽 여백(목차 번호 없는 항목을 챕터명 시작선에 맞출 때)


KEY_BULLET = dict(char="■", pct=60, color=C["INK"])
BODY_BULLET = dict(char="•", pct=100, color=C["TEXT2"])


def hang(pt):
    """매달린 들여쓰기 = 0.85em."""
    return 0.85 * pt * PX_PT


def layout_paras(paras: list[Para], width: float, what="text", key_pt=None):
    """문단 → (문단별 줄, 총 높이). 첫 문단의 before 는 0(LibreOffice 무시)."""
    out, total = [], 0.0
    for i, p in enumerate(paras):
        spans = to_spans(p.text, p.w, p.color)
        kp = key_pt or p.pt
        if p.bullet:
            ind = hang(p.pt)
            mar = p.pad_px + ind + (hang(kp) if p.level == 1 else 0.0)
        else:
            ind = p.hang_px
            mar = p.pad_px + p.hang_px + (hang(kp) if p.level == 1 else 0.0)
        avail = width - mar
        plain = "".join(sp.text for sp in spans)
        if "\n" in plain:          # 작성자가 넣은 강제 줄바꿈: 구간마다 따로 나누고 합친다
            lines, pos = [], 0
            for seg in plain.split("\n"):
                sub = slice_spans(spans, pos, pos + len(seg))
                for a, b in wrap(sub, p.pt, avail, None, what=what):
                    lines.append((pos + a, pos + b))
                pos += len(seg) + 1
            if p.max_lines is not None and len(lines) > p.max_lines:
                raise WrapError(f"{what}: {len(lines)}줄 (허용 {p.max_lines}줄) — 문안을 줄일 것")
        else:
            lines = wrap(spans, p.pt, avail, p.max_lines, what=what)
        before = 0.0 if i == 0 else p.before_pt
        total += before * PX_PT + len(lines) * pitch(p.pt)
        out.append(dict(para=p, spans=spans, lines=lines, mar=mar, ind=ind, before=before,
                        orphan=is_orphan(spans, p.pt, lines, avail)))
    return out, total


def natural_width(paras: list[Para], key_pt=None) -> float:
    best = 0.0
    for p in paras:
        spans = to_spans(p.text, p.w, p.color)
        kp = key_pt or p.pt
        mar = p.pad_px + ((hang(p.pt) + (hang(kp) if p.level == 1 else 0)) if p.bullet else (p.hang_px + (hang(kp) if p.level == 1 else 0)))
        best = max(best, spans_fit(spans, p.pt) / SLACK + mar + 2)
    return best


def _emit_line(p, seg, pt):
    """한 줄의 스팬을 런으로 만든다(스팬 하나 = 런 하나). 마지막 스팬을 돌려준다."""
    last = None
    for sp in seg:
        r = p.add_run()
        r.text = sp.text.replace("\u00a0", " ")
        _style_run(r, sp, pt)
        last = sp
    return last


def emit_text(slide, rect: Rect, laid, name: str, anchor="t"):
    tb = _textbox(slide, rect, name, anchor)
    tf = tb.text_frame
    for n_i, item in enumerate(laid):
        p = tf.paragraphs[0] if n_i == 0 else tf.add_paragraph()
        pa = item["para"]
        bu = None
        if pa.bullet:
            bu = KEY_BULLET if pa.bullet == KEY_BULLET["char"] else BODY_BULLET
        _style_para(p, pa.align, item["before"], bullet=pa.bullet, mar_px=item["mar"], ind_px=item["ind"],
                    bu_color=bu["color"] if bu else None, bu_pct=bu["pct"] if bu else 100, bu_w=pa.w)
        spans = item["spans"]
        last = spans[-1] if spans else Span("", pa.w, pa.color)
        for li, (a, b) in enumerate(item["lines"]):
            last = _emit_line(p, slice_spans(spans, a, b), pa.pt) or last
            if li < len(item["lines"]) - 1:
                _add_br(p, pa.pt, last.color, last.w if last.family == "pl" else pa.w)
        _end_para(p, pa.pt, pa.color, pa.w)
    return tb


def _laid_meta(laid):
    out = []
    for it in laid:
        plain = "".join(s.text for s in it["spans"])
        out.append(dict(pt=it["para"].pt, w=it["para"].w, role=it["para"].role,
                        lines=[plain[a:b] for a, b in it["lines"]], orphan=it["orphan"]))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 8. 덱
# ─────────────────────────────────────────────────────────────────────────────
class Deck:
    """v4.1 덱. kind = 'ppt'(발표용) | 'report'(보고용). 로고는 logo_slots.py."""

    def __init__(self, kind="ppt", base: Path = BASE, workdir: Path | None = None, logo=None):
        """logo = 로고 그림 경로 | "placeholder" | "none" | None(사용자 설정 ~/.ppt_skills/config.json → 없으면 'LOGO' 자리표시)."""
        assert kind in TOK, kind
        self.kind = kind
        self.T = TOK[kind]
        self.prs = Presentation(str(base))
        assert self.prs.slide_width == SW
        self.protos = list(self.prs.slides)                      # [표지, 간지, 감사]
        self.layout = next(l for l in self.prs.slide_layouts if l.name == "슬라이드 1")
        self.workdir = Path(workdir) if workdir else Path.cwd()
        # 로고 자리(LOGO:<자리>) 채우기. 사용자가 고친 pptx 를 바탕으로 열 때(base ≠ 기본)는 logo 를 줬을 때만 바꾼다
        self.logo_info = ""
        if Path(base).resolve() == BASE.resolve() or logo is not None:
            import importlib.util as _u
            _sp = _u.spec_from_file_location("logo_slots", Path(__file__).with_name("logo_slots.py"))
            _lg = _u.module_from_spec(_sp)
            _sp.loader.exec_module(_lg)
            self.logo_info = _lg.apply(self.prs, logo, self.workdir)
            print("[logo]", self.logo_info)
        self.cache = self.workdir / "assets" / "_prep"
        self.meta = {"kind": kind, "engine": "v4.1", "slides": []}
        self.warnings: list[str] = []
        self._open: list = []
        self.toc_items: list[str] = []

    # ── 내부 ──
    def _new(self, layout=None):
        s = self.prs.slides.add_slide(layout or self.layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        return s

    def _clone(self, proto):
        dst = self._new(proto.slide_layout)
        rid = {}
        for key, rel in list(proto.part.rels.items()):
            if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
                continue
            if rel.is_external:
                rid[key] = dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                rid[key] = dst.part.relate_to(rel.target_part, rel.reltype)
        for el in proto.shapes._spTree:
            if etree.QName(el).localname in ("nvGrpSpPr", "grpSpPr"):
                continue
            new = copy.deepcopy(el)
            for node in new.iter():
                for attr in list(node.attrib):
                    if attr.startswith("{%s}" % "http://schemas.openxmlformats.org/officeDocument/2006/relationships"):
                        v = node.get(attr)
                        if v in rid:
                            node.set(attr, rid[v])
            dst.shapes._spTree.append(new)
        bg = proto._element.find(qn("p:cSld")).find(qn("p:bg"))
        if bg is not None:
            dst._element.find(qn("p:cSld")).insert(0, copy.deepcopy(bg))
        return dst

    def _record(self, slide, kind, **kw):
        m = dict(index=len(self.prs.slides) - len(self.protos), kind=kind, boxes=[], pics=[], modules=[], **kw)
        self.meta["slides"].append(m)
        slide._element.find(qn("p:cSld")).set("name", f"v4:{kind}")
        return m

    def _text(self, slide, m, rect, paras, name, anchor="t"):
        laid, h = layout_paras(paras, rect.w, f"{m['index']}쪽 {name}")
        r = Rect(rect.x, rect.y, rect.w, h)
        emit_text(slide, r, laid, name, anchor)
        m["boxes"].append(dict(name=name, rect=r.as_list(), paras=_laid_meta(laid)))
        return r

    # ── 슬라이드 유형 ──
    def cover(self, title, subtitle="", info=(), title_lines=2):
        """표지(원형 1 복제): 제목 40pt ≤2줄, 행사·날짜 28pt(가운데), 소속·이름 20pt Bold."""
        s = self._clone(self.protos[0])
        for shp in list(s.shapes):
            if shp.has_text_frame:
                shp._element.getparent().remove(shp._element)
        m = self._record(s, "cover")
        paras = [Para(title, FIXED["cover"], 7, C["INK"], align="c", max_lines=title_lines, role="cover")]
        if subtitle:
            paras.append(Para(subtitle, FIXED["cover_sub"], 4, C["TEXT2"], align="c", before_pt=10, max_lines=1, role="cover_sub"))
        n_lines = len(layout_paras(paras[:1], 1200, "표지 제목")[0][0]["lines"])
        self._text(s, m, Rect(360, 92 if n_lines == 1 else 60, 1200, 0), paras, "v4:cover_title")
        if info:
            ip = [Para(t, FIXED["cover_info"], 7, C["INK"], before_pt=0 if i == 0 else 8, max_lines=1, role="cover_info")
                  for i, t in enumerate(info)]
            self._text(s, m, Rect(120, 430, 1100, 0), ip, "v4:cover_info")
        return s

    def toc(self, chapters, title="목차", front=(), back=()):
        """목차. chapters = [(챕터명, [섹션명, ...]), ...]. 섹션명 = 본문 제목 바 문자열.
        front·back = 챕터 밖의 번호 없는 항목(예: front=["결과 요약"], back=["References"]).
        챕터 4개 이하이고 합계 5열 이하면 챕터와 나란한 번호 없는 열(이름이 챕터명과 같은 높이), 넘으면 2열 목록의 앞·뒤 줄."""
        T = self.T
        sl = self.slide(title, chapter=None, kind="toc")
        n = len(chapters)
        front, back = list(front), list(back)
        self.toc_items = front + [sec for _, secs in chapters for sec in secs] + back

        def extra(items):
            return [Para(t, T["toc_ch"], 6, C["INK"], before_pt=0 if i == 0 else 12, max_lines=1, role="toc_extra")
                    for i, t in enumerate(items)]

        spec = [("x", t, []) for t in front] + [("c", nm, secs) for nm, secs in chapters] + [("x", t, []) for t in back]
        if n <= 4 and len(spec) <= 5:
            k_all = len(spec)
            g = 96 if k_all <= 2 else GUTTER[3]
            col_paras, offs = [], []
            num = 0
            for kind_, name, secs in spec:
                if kind_ == "c":
                    num += 1
                    paras = [Para(f"{num:02d}", T["toc_num"], 7, C["GREEN"], max_lines=1, role="toc_num"),
                             Para(name, T["toc_ch"], 6, C["INK"], before_pt=4, max_lines=2, role="toc_ch")]
                    paras += [Para(sec, T["toc_sec"], 4, C["TEXT2"], BODY_BULLET["char"], 0,
                                   12 if i == 0 else T["item_pt"], max_lines=2, role="toc_sec") for i, sec in enumerate(secs)]
                    offs.append(0.0)
                else:           # 번호 없는 앞·뒤 항목: 챕터명과 같은 높이에서 시작
                    paras = [Para(name, T["toc_ch"], 6, C["INK"], max_lines=2, role="toc_extra")]
                    offs.append(pitch(T["toc_num"]) + 4 * PX_PT)
                col_paras.append(paras)
            # 열 폭 = 내용 자연 폭, 묶음은 가로 가운데(구분선이 거터 가운데에 오게). 넘치면 비례 분할
            nat = [max(natural_width(pp), 160) for pp in col_paras]
            g = 96
            tot = sum(nat) + g * (k_all - 1)
            if tot <= FRAME.w:
                cols, x = [], FRAME.x + (FRAME.w - tot) / 2
                for wv in nat:
                    cols.append(Rect(x, FRAME.y, wv, FRAME.h))
                    x += wv + g
            else:
                g = GUTTER[3]
                cols = FRAME.hsplit(nat, g)
            laid_all, hs = [], []
            for paras, col, off in zip(col_paras, cols, offs):
                laid, h = layout_paras(paras, col.w, "목차")
                laid_all.append(laid)
                hs.append(h + off)
            H = max(hs)
            if H > FRAME.h:
                raise LayoutError(f"목차 높이 {H:.0f}px > {FRAME.h:.0f}px — 섹션을 줄일 것")
            y = FRAME.y + max(0, (FRAME.h - H) / 2 - 24)
            for k, (laid, col, off, h) in enumerate(zip(laid_all, cols, offs, hs)):
                r = Rect(col.x, y + off, col.w, h - off)
                nm = f"v4:toc{k}"
                emit_text(sl.s, r, laid, nm)
                sl.m["boxes"].append(dict(name=nm, rect=r.as_list(), paras=_laid_meta(laid)))
                if k:
                    x = col.x - g / 2
                    ln = sl.s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(y), E(x), E(y + H))
                    ln.name = "v4:toc_div"
                    _strip_style(ln)
                    ln.line.color.rgb = _rgb(C["LINE_LIGHT"])
                    ln.line.width = Pt(0.75)
        else:
            per = math.ceil(n / 2)
            areas = FRAME.hsplit([1, 1], 96)
            pad = text_w("01  ", 7, T["toc_ch"])          # 번호 없는 항목을 챕터명 시작선에 맞춘다

            def unnumbered(t, before):
                return Para(t, T["toc_ch"], 6, C["INK"], before_pt=before, max_lines=1, role="toc_extra", pad_px=pad)
            laid_cols = []
            for ci in range(2):
                paras = [unnumbered(t, 0 if i == 0 else 18) for i, t in enumerate(front)] if ci == 0 else []
                for k, (name, secs) in enumerate(chapters[ci * per:(ci + 1) * per]):
                    num = ci * per + k + 1
                    paras.append(Para([(f"{num:02d}  ", {"w": 7, "color": C["GREEN"]}), (name, {"w": 6, "color": C["INK"]})],
                                      T["toc_ch"], 6, C["INK"], before_pt=0 if not paras else 18, max_lines=1, role="toc_ch"))
                    for sec in secs:
                        paras.append(Para(sec, T["toc_sec"], 4, C["TEXT2"], BODY_BULLET["char"], 0, T["item_pt"],
                                          max_lines=1, role="toc_sec", pad_px=pad))
                if ci == 1:
                    paras += [unnumbered(t, 18) for t in back]
                laid, h = layout_paras(paras, areas[ci].w, "목차", key_pt=T["toc_ch"])
                if h > areas[ci].h:
                    raise LayoutError(f"목차 {h:.0f}px > 본문 높이 {areas[ci].h:.0f}px")
                laid_cols.append((laid, h))
            H = max(h for _, h in laid_cols)
            y = FRAME.y + max(0, (FRAME.h - H) / 2 - 24)
            for ci, (laid, h) in enumerate(laid_cols):
                r = Rect(areas[ci].x, y, areas[ci].w, h)
                emit_text(sl.s, r, laid, f"v4:toc{ci}")
                sl.m["boxes"].append(dict(name=f"v4:toc{ci}", rect=r.as_list(), paras=_laid_meta(laid)))
        sl.balance = False
        sl.m["toc"] = [[c, list(x)] for c, x in chapters]
        sl.m["toc_front"], sl.m["toc_back"] = front, back
        self.meta["toc_items"] = self.toc_items
        return sl

    def divider(self, num: int, title: str, chapters=None):
        """간지(원형 2 복제): CHAPTER 17pt, 번호 80pt 두 톤, 제목 37pt 1줄. chapters 를 주면 진행 표시."""
        s = self._clone(self.protos[1])
        for shp in list(s.shapes):
            if shp.shape_type == 6 or shp.has_text_frame:
                shp._element.getparent().remove(shp._element)
        m = self._record(s, "divider", num=num, title=title)
        X = 124.4        # 원형 간지의 가로선 시작 x — 세 줄의 잉크 왼쪽선
        r1 = self._text(s, m, Rect(X - ink_left("C", 6, FIXED["div_label"]), 338, 400, 0),
                        [Para("CHAPTER", FIXED["div_label"], 6, C["MUTED"], max_lines=1, role="div_label")], "v4:div_label")
        digits = f"{num:02d}"
        r2 = self._text(s, m, Rect(X - ink_left(digits, 7, FIXED["div_num"]), r1.b + 6, 420, 0),
                        [Para([(digits[0], {"color": C["NUM_LIGHT"]}), (digits[1:], {"color": C["GREEN"]})],
                              FIXED["div_num"], 7, C["GREEN"], max_lines=1, role="div_num")], "v4:div_num")
        self._text(s, m, Rect(X - ink_left(title, 7, FIXED["div_title"]), max(556, r2.b + 8), 900, 0),
                   [Para(title, FIXED["div_title"], 7, C["INK"], max_lines=1, role="div_title")], "v4:div_title")
        if chapters:
            ps = [Para(f"{i:02d}  {nm}", FIXED["div_prog"], 7 if i == num else 4, C["GREEN"] if i == num else C["MUTED"],
                       before_pt=0 if i == 1 else 4, max_lines=1, role="div_prog") for i, nm in enumerate(chapters, start=1)]
            self._text(s, m, Rect(X - ink_left("0", 7, FIXED["div_prog"]), 700, 800, 0), ps, "v4:div_progress")
        return s

    def slide(self, title: str, chapter: int | None = None, crumb: str | None = None, kind="content"):
        """본문 슬라이드. title = 목차 항목명(섹션 내 반복). kind = content|summary|conclusion|toc|refs."""
        s = self._new()
        m = self._record(s, kind, title=title, chapter=chapter)
        return Slide(self, s, m, title, chapter, crumb)

    def references(self, refs, title="References"):
        sl = self.slide(title, kind="refs")
        pt = self.T["refs"]
        lab = max(spans_fit(to_spans(f"[{i}] ", 4, C["TEXT2"]), pt) for i in range(1, len(refs) + 1))
        paras = [Para(f"[{i}] {t}", pt, 4, C["TEXT2"], before_pt=0 if i == 1 else 8, max_lines=2, role="refs",
                      hang_px=lab) for i, t in enumerate(refs, start=1)]
        sl.text(FRAME, paras, name="v4:refs")
        sl.balance = False
        return sl

    def closing(self, text="감사합니다"):
        s = self._clone(self.protos[2])
        for shp in list(s.shapes):
            if shp.has_text_frame and shp.text_frame.text.strip():
                shp._element.getparent().remove(shp._element)
        m = self._record(s, "closing")
        self._text(s, m, Rect(260, 300, 1400, 0),
                   [Para(text, FIXED["closing"], 6, C["CLOSING"], align="c", max_lines=1, role="closing")], "v4:closing")
        return s

    # ── 저장 ──
    def save(self, path) -> Path:
        for sl in self._open:
            sl.finalize()
        self._open = []
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        lst = self.prs.slides._sldIdLst
        for sld in list(lst)[: len(self.protos)]:
            self.prs.part.drop_rel(sld.rId)
            lst.remove(sld)
        self.protos = []
        from native import positive_ax_ids              # 네이티브 차트 음수 축 ID → 작은 양수(OOXML 스키마, lib/native.py)
        positive_ax_ids(self.prs)
        self.prs.save(str(path))
        Presentation(str(path))                         # G18: 다시 열리는지
        self.meta["warnings"] = self.warnings
        path.with_suffix(".v4.json").write_text(json.dumps(self.meta, ensure_ascii=False, indent=1))
        for w in self.warnings:
            print("[warn]", w)
        print(f"saved {path} ({len(self.prs.slides)} slides)")
        return path


# ─────────────────────────────────────────────────────────────────────────────
# 9. 본문 슬라이드와 모듈
# ─────────────────────────────────────────────────────────────────────────────
def _module(always=False):
    """공개 모듈 표식. row() 밖의 바깥 호출에서 rect(좌표)를 직접 주면 그 장은 세로 균형 이동을 하지 않는다
    (직접 놓은 요소는 흐름 밖이라 같이 움직이지 않으므로). 모듈 안에서 부르는 하위 모듈은 세지 않는다."""
    def deco(fn):
        sig = inspect.signature(fn)

        @functools.wraps(fn)
        def call(self, *a, **k):
            if self._depth == 0 and not self._in_row:
                if always or sig.bind(self, *a, **k).arguments.get("rect") is not None:
                    self._manual = True
            self._depth += 1
            try:
                return fn(self, *a, **k)
            finally:
                self._depth -= 1
        return call
    return deco


class Slide:
    def __init__(self, deck: Deck, s, m, title, chapter, crumb):
        self.d, self.s, self.m = deck, s, m
        self.T = deck.T
        self.cursor = FRAME.y
        self.n = 0
        self._flow = []               # [(모듈 종류, [요소])] — 자동 흐름으로 놓인 모듈
        self.pending_key = None       # 요지 줄: 다음 전폭 불릿과 한 상자로 합치려고 보류(ST-06)
        self._manual = False          # row() 밖에서 rect= 로 직접 놓은 모듈이 있으면 세로 균형 이동을 하지 않는다
        self._in_row = False
        self._depth = 0
        self.balance = True
        self._header(title, chapter, crumb)
        deck._open.append(self)

    # ── 헤더 ──
    def _header(self, title, chapter, crumb):
        if chapter is not None:
            laid, _ = layout_paras([Para(str(chapter), FIXED["chapter"], 6, C["WHITE"], align="c", max_lines=1, role="chapter")], 92.6)
            emit_text(self.s, Rect(8.7, 12.3, 92.6, 87.1), laid, "v4:chapter", anchor="m")
        if crumb:
            laid, _ = layout_paras([Para(crumb, FIXED["crumb"], 6, C["GRAY"], max_lines=1, role="crumb")], 900)
            emit_text(self.s, Rect(TITLE_X, 20, 900, 28), laid, "v4:crumb", anchor="m")
        spans = to_spans(title, 7, C["INK"])
        tw = spans_fit(spans, self.T["title"])
        if tw > TITLE_MAXW * SLACK:
            raise WrapError(f"제목 '{title}' 폭 {tw:.0f}px > {TITLE_MAXW * SLACK:.0f}px — 줄일 것(국문 20자 권장)")
        laid, _ = layout_paras([Para(title, self.T["title"], 7, C["INK"], max_lines=1, role="title")], TITLE_MAXW)
        r = Rect(TITLE_X, 56.8 - 1.5, TITLE_MAXW, 70.3)
        emit_text(self.s, r, laid, "v4:title", anchor="m")
        self.m["boxes"].append(dict(name="v4:title", rect=r.as_list(), paras=_laid_meta(laid)))

    # ── 영역·흐름 ──
    def _area(self) -> Rect:
        return Rect(FRAME.x, self.cursor, FRAME.w, FRAME.b - self.cursor)

    def body(self) -> Rect:
        """다음 모듈이 놓일 영역. 보류 중인 요지가 있으면 그 아래(요지 높이 + 블록 간격)부터.
        그래프 슬롯 크기를 미리 잴 때 쓴다: fig = plt.figure(figsize=chart_figsize(s.body()))."""
        if self.pending_key is not None:
            _, h = layout_paras([self._key_para(self.pending_key[0], align=self.pending_key[1])], FRAME.w,
                                key_pt=self.T["key"])
            y = self.cursor + h + self.T["block"]
            return Rect(FRAME.x, y, FRAME.w, FRAME.b - y)
        return self._area()

    def _begin(self, rect, name, register=True, flush=True):
        if register:
            if rect is None and flush and name not in ("bullets",):
                self._flush_key()
            self.m["modules"].append(name)
        return dict(mark=len(self.s.shapes._spTree), cursor=self.cursor, nb=len(self.m["boxes"]),
                    np=len(self.m["pics"]), auto=rect is None, name=name)

    def _end(self, tok, used: Rect):
        if not tok["auto"]:
            return used
        self.cursor = max(self.cursor, used.b + self.T["block"])
        els = list(self.s.shapes._spTree)[tok["mark"]:]
        self._flow.append((tok["name"], els, tok["nb"], tok["np"]))
        return used

    def _fit(self, need_h, area: Rect, what):
        if need_h > area.h + 0.5:
            raise LayoutError(f"[{self.m['index']}쪽 {what}] 필요 {need_h:.0f}px > 영역 {area.h:.0f}px"
                              " — 문안·그림을 줄이거나 장을 나눌 것")

    def _flush_key(self):
        """보류한 요지를 단독 상자로 놓는다(다음 모듈이 전폭 불릿이 아닐 때)."""
        if self.pending_key is None:
            return
        text, align = self.pending_key
        self.pending_key = None
        self._depth += 1
        try:
            r = self.text(Rect(FRAME.x, self.cursor, FRAME.w, FRAME.b - self.cursor),
                          [self._key_para(text, align=align)], name="v4:key")
        finally:
            self._depth -= 1
        self.cursor = r.b + self.T["block"]

    def finalize(self):
        """위쪽 글 묶음(요지·불릿, 요지 대신 쓰는 결론 박스)은 고정하고, 그 아래 그림·표·카드 묶음이 짧으면
        조금 내린다(§4.3): 요지→묶음 간격이 아래 여백의 절반을 넘지 않고, 묶음이 작을수록 요지에 붙는다."""
        self._flush_key()
        if not self.balance or not self._flow or self._manual:
            return
        first = next((i for i, f in enumerate(self._flow) if f[0] not in ("bullets", "note")), None)
        if first is None:
            return
        bottom = self.cursor - self.T["block"]
        leftover = FRAME.b - bottom
        nb0, np0 = self._flow[first][2], self._flow[first][3]
        tops = [b["rect"][1] for b in self.m["boxes"][nb0:]] + [b["rect"][1] for b in self.m["pics"][np0:]]
        group_h = bottom - (min(tops) if tops else bottom)
        if leftover < 0.04 * FRAME.h:
            return
        r = group_h / leftover
        if r < 0.8:
            # 묶음이 남는 공간보다 작으면 옮겨도 섬처럼 뜬다 → 요지에 붙여 두고 알린다
            self.d.warnings.append(f"{self.m['index']}쪽: 본문 아래 빈 공간 {leftover:.0f}px ({leftover / FRAME.h:.0%})"
                                   " — 그림·표를 키우거나 추가할 것")
        # 요지→묶음 간격(블록 + dy)이 아래 여백(L − dy)의 절반을 넘지 않게. 묶음이 작을수록 덜 내린다(문턱에서 0부터 연속)
        B = self.T["block"]
        full = max(0.0, min((0.5 * leftover - B) / 1.5, 0.4 * leftover))
        dy = full * min(1.0, max(0.0, (r - 0.8) / 0.8))
        if dy < 1:
            return
        for _, els, _, _ in self._flow[first:]:
            for el in els:
                off = el.find(".//" + qn("a:off"))
                if off is not None:
                    off.set("y", str(int(off.get("y")) + int(round(dy * EPX))))
        for b in self.m["boxes"][nb0:]:
            b["rect"][1] = round(b["rect"][1] + dy, 1)
        for b in self.m["pics"][np0:]:
            b["rect"][1] = round(b["rect"][1] + dy, 1)
        self.m["balanced_dy"] = round(dy, 1)

    # ── 좌우 분할 블록 ──
    @contextmanager
    def row(self, weights, gutter=None, valign="auto"):
        """좌우 분할을 한 흐름 블록으로 쓴다.

            with s.row([3, 2]) as (L, R):
                s.figure(img, "캡션", rect=L)
                s.bullets(["...", "..."], rect=R)

        칸 안 모듈은 rect 로 칸을 받는다. 블록 높이 = 가장 높은 칸.
        valign='auto'(기본, §9.5): 짧은 칸이 가장 높은 칸의 50% 이상이면 윗선 맞춤, 미만이면 세로 가운데.
        't' = 모두 윗선, 'm' = 모두 세로 가운데."""
        self._flush_key()
        tok = self._begin(None, "row", flush=False)
        self.m.setdefault("layouts", []).append("row" + ":".join(str(w) for w in weights))
        area = self._area()
        cols = area.hsplit(weights, gutter)
        self._in_row = True
        try:
            yield cols
        finally:
            self._in_row = False
        els = list(self.s.shapes._spTree)[tok["mark"]:]
        col_of, bottoms = {}, [area.y] * len(cols)
        for el in els:
            off, ext = el.find(".//" + qn("a:off")), el.find(".//" + qn("a:ext"))
            if off is None or ext is None:
                continue
            x, y = int(off.get("x")) / EPX, int(off.get("y")) / EPX
            w, h = int(ext.get("cx")) / EPX, int(ext.get("cy")) / EPX
            cx = x + w / 2
            ci = min(range(len(cols)), key=lambda i: 0 if cols[i].x <= cx <= cols[i].r else min(abs(cx - cols[i].x), abs(cx - cols[i].r)))
            col_of[id(el)] = ci
            bottoms[ci] = max(bottoms[ci], y + h)
        heights = [b - area.y for b in bottoms]
        block_h = max(heights)
        if valign in ("m", "auto"):
            for ci, hc in enumerate(heights):
                dy = (block_h - hc) / 2
                if dy < 1 or hc <= 0 or (valign == "auto" and hc >= 0.5 * block_h):
                    continue
                for el in els:
                    if col_of.get(id(el)) == ci:
                        off = el.find(".//" + qn("a:off"))
                        off.set("y", str(int(off.get("y")) + int(round(dy * EPX))))
                for b in self.m["boxes"][tok["nb"]:] + self.m["pics"][tok["np"]:]:
                    bx = b["rect"][0] + b["rect"][2] / 2
                    if cols[ci].x - 1 <= bx <= cols[ci].r + 1:
                        b["rect"][1] = round(b["rect"][1] + dy, 1)
        self._end(tok, Rect(area.x, area.y, area.w, block_h))

    def below(self, used: Rect, area: Rect, gap=None) -> Rect:
        """칸(area) 안에서 이미 놓은 모듈(used) 아래 남은 영역. 칸 안에 모듈을 세로로 쌓을 때 쓴다."""
        g = self.T["block"] if gap is None else gap
        return Rect(area.x, used.b + g, area.w, area.b - used.b - g)

    def notes(self, text: str):
        """발표자 노트(슬라이드에 보이지 않음). 내부 데이터 경로·계산 근거를 남기는 곳(§9.2: 슬라이드 출처는 논문만)."""
        ns = self.s.notes_slide
        ns.notes_text_frame.text = text
        self.m["notes"] = text

    # ── 공개 도형 도우미 (p:style 제거·토큰 색) ──
    @_module()
    def shape(self, kind, rect: Rect, fill=None, line=None, line_pt=1.0, radius=None, name="v4:deco"):
        return _shape(self.s, kind, rect, name, fill, line, line_pt, radius)

    @_module(always=True)
    def line(self, x1, y1, x2, y2, color=None, pt=1.0, name="v4:line"):
        ln = self.s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
        ln.name = name
        _strip_style(ln)
        ln.line.color.rgb = _rgb(color or C["LINE"])
        ln.line.width = Pt(pt)
        return ln

    # ── 텍스트 ──
    @_module()
    def text(self, rect: Rect, paras: list[Para], name="v4:text", anchor="t", flow=False) -> Rect:
        tok = self._begin(None if flow else rect, "bullets" if flow else name, register=flow)
        laid, h = layout_paras(paras, rect.w, f"{self.m['index']}쪽 {name}", key_pt=self.T["key"])
        self._fit(h, rect, name)
        y = rect.y + ((rect.h - h) / 2 if anchor == "m" else (rect.h - h if anchor == "b" else 0))
        r = Rect(rect.x, y, rect.w, h)
        emit_text(self.s, r, laid, name)
        self.m["boxes"].append(dict(name=name, rect=r.as_list(), paras=_laid_meta(laid)))
        self._orphans(laid, name)
        if flow:
            self._end(tok, r)
        return r

    def _orphans(self, laid, name):
        for it in laid:
            if it["orphan"]:
                self.d.warnings.append(f"{self.m['index']}쪽 {name}: 고아 줄 '{''.join(s.text for s in it['spans'])[:30]}…'"
                                       " — 문안을 줄이거나 칸을 넓힐 것(게이트 G6 fail)")

    def _key_para(self, text, before_pt=0.0, align="l", max_lines=1):
        b = KEY_BULLET["char"] if align == "l" else None
        return Para(text, self.T["key"], 6, C["INK"], b, 0, before_pt, align, max_lines, "key")

    def _body_para(self, text, level=1, before_pt=None, max_lines=None):
        return Para(text, self.T["body"], 4, C["TEXT2"], BODY_BULLET["char"], level,
                    self.T["item_pt"] if before_pt is None else before_pt,
                    max_lines=self.T["body_wrap"] if max_lines is None else max_lines, role="body")

    def _bullet_paras(self, items, indent=True):
        paras = []
        lvl = 1 if indent else 0
        for it in items:
            if isinstance(it, tuple):
                k, subs = it
                paras.append(self._key_para(k, before_pt=self.T["group_pt"] if paras else 0))
                for j, sub in enumerate(subs):
                    paras.append(self._body_para(sub, 1, self.T["key_body_pt"] if j == 0 else self.T["item_pt"]))
            else:
                paras.append(self._body_para(it, lvl, self.T["item_pt"] if paras else 0))
        return paras

    def key(self, text):
        """요지 줄 ■ (1줄, y=162 고정). 바로 다음이 전폭 불릿이면 한 텍스트박스로 합친다(ST-06)."""
        self._flush_key()
        self.m.setdefault("keys", []).append(text)
        self.pending_key = (text, "l")

    def message(self, text):
        """메시지 장: 요지를 가운데 정렬(같은 크기·굵기, 강조 없음). 덱당 2장 이하."""
        self._flush_key()
        self.m.setdefault("keys", []).append(text)
        self.m["message"] = True
        self.pending_key = (text, "c")
        self._flush_key()

    @_module()
    def bullets(self, items, rect: Rect | None = None, anchor="t", indent=None) -> Rect:
        """items: 문자열(세부 •) 또는 (요지, [세부...]) 튜플. 한 텍스트박스에 문단으로 쌓는다.
        indent(기본): 흐름 배치면 세부 • 를 요지 글자 시작선에, rect(칸) 안이면 칸 왼쪽선에 둔다."""
        if indent is None:
            indent = rect is None
        paras = self._bullet_paras(items, indent)
        if rect is None and self.pending_key is not None and self.pending_key[1] == "l":
            k = self._key_para(self.pending_key[0])
            self.pending_key = None
            if paras and paras[0].role == "body":
                paras[0].before_pt = self.T["key_body_pt"]
            elif paras:
                paras[0].before_pt = self.T["group_pt"]
            paras = [k] + paras
        elif rect is not None:
            pass
        else:
            self._flush_key()
        self.n += 1
        area = rect or self._area()
        return self.text(area, paras, name=f"v4:bul{self.n}", anchor=anchor, flow=rect is None)

    # ── 그림 ──
    def _pic(self, path, rect: Rect, name, border=False, crop=None, transparent=False, jpeg=False):
        p, iw, ih = prep_image(Path(path), self.d.cache, transparent=transparent, jpeg=jpeg)
        pic = self.s.shapes.add_picture(str(p), E(rect.x), E(rect.y), E(rect.w), E(rect.h))
        pic.name = name
        if crop:
            pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = crop
        if border:
            pic.line.color.rgb = _rgb(C["LINE"])
            pic.line.width = Pt(0.75)
        vis_w = iw * (1 - (crop[0] + crop[2] if crop else 0))
        self.m["pics"].append(dict(name=name, rect=rect.as_list(), src=str(path), px=[iw, ih],
                                   scale=round(rect.w / vis_w, 3), crop=list(crop) if crop else None))
        return pic

    def _cap_h(self):
        return self.T["cap_gap"] + pitch(self.T["caption"])

    def _caption(self, text, img: Rect, lim: Rect, name):
        t = text if text.startswith(("<", "[")) else f"<{text}>"
        tw = spans_fit(to_spans(t, 4, C["TEXT2"]), self.T["caption"]) / SLACK + 2
        w = min(max(img.w, tw), lim.w)
        x = min(max(img.cx - w / 2, lim.x), lim.r - w)
        r = Rect(x, img.b + self.T["cap_gap"], w, pitch(self.T["caption"]))
        return self.text(r, [Para(t, self.T["caption"], 4, C["TEXT2"], align="c", max_lines=1, role="caption")], name=name)

    def _place(self, path, box: Rect, caption, lim: Rect, photo=False, border=False, valign="t", clear_bg=False):
        """그림 1장을 box(캡션 제외 영역)에 contain(사진은 크롭 허용)으로 놓고 캡션을 단다.
        box 폭의 90% 이상을 쓰는 도표는 가운데가 아니라 box 왼쪽선에 붙인다(§4.1 좌측 시작선)."""
        iw, ih = img_dims(path)
        if photo:
            img, crop = fit_photo(iw, ih, box)
        else:
            img, crop = contain(iw, ih, box), None
            if img.w >= 0.9 * box.w:
                img = Rect(box.x, img.y, img.w, img.h)
        if valign == "t":
            img = Rect(img.x, box.y, img.w, img.h)
        k = len(self.m["pics"]) + 1
        is_jpg = Path(path).suffix.lower() in (".jpg", ".jpeg")
        self._pic(path, img, f"v4:fig{k}", border, crop, transparent=clear_bg and not (photo or border or is_jpg),
                  jpeg=photo or is_jpg)
        bottom = img.b
        if caption:
            c = self._caption(caption, img, lim, f"v4:cap{k}")
            bottom = c.b
        return img, bottom

    def _img_area(self, path, box, photo=False):
        iw, ih = img_dims(path)
        r = fit_photo(iw, ih, box)[0] if photo else contain(iw, ih, box)
        return r.w * r.h

    @_module()
    def figure(self, path, caption=None, rect: Rect | None = None, photo=False, border=False, clear_bg=False) -> Rect:
        """그림 1장 크게 + 아래 캡션. clear_bg=True 면 가장자리와 이어진 순백 배경을 투명하게(기본은 흰 배경 그대로)."""
        tok = self._begin(rect, "figure")
        area = rect or self._area()
        cap = self._cap_h() if caption else 0
        img, bottom = self._place(path, Rect(area.x, area.y, area.w, area.h - cap), caption, area, photo, border,
                                  clear_bg=clear_bg)
        return self._end(tok, Rect(area.x, area.y, area.w, bottom - area.y))

    @_module()
    def figures(self, items, rect: Rect | None = None, photo=False) -> Rect:
        """그림 2~4장. items = [(경로, 캡션), ...]. 한 줄 / 1+2 / 2×2 중 그림 면적이 큰 배치를 고른다(§9.5)."""
        tok = self._begin(rect, "figures")
        area = rect or self._area()
        n = len(items)
        has_cap = any(c for _, c in items)
        cap = self._cap_h() if has_cap else 0
        g = GUTTER.get(n, 32)
        dims = [img_dims(p) for p, _ in items]
        asp = [w / h for w, h in dims]
        cands = []
        gv = 32

        def score(area_px, used_h):
            # 그림 면적 × 세로 채움 가중(아래가 크게 비는 배치를 벌점) — §4.3·§9.5
            fill = min(1.0, used_h / max(1.0, area.h))
            return area_px * (0.5 + 0.5 * fill)

        # 한 줄
        h = min(area.h - cap, (area.w - g * (n - 1)) / sum(asp))
        cands.append(("row", score(sum(a * h * h for a in asp), h + cap)))
        if n == 3:   # 1+2
            lw = (area.w - g) * 0.6
            rw = area.w - g - lw
            cell_h = (area.h - 2 * cap - gv) / 2
            iw0, ih0 = img_dims(items[0][0])
            big = contain(iw0, ih0, Rect(0, 0, lw, area.h - cap))
            a0 = big.w * big.h
            a1 = self._img_area(items[1][0], Rect(0, 0, rw, cell_h), photo)
            a2 = self._img_area(items[2][0], Rect(0, 0, rw, cell_h), photo)
            cands.append(("1+2", score(a0 + a1 + a2, max(big.h + cap, 2 * (cell_h + cap) + gv))))
        if n == 4:   # 2×2
            cw = (area.w - g) / 2
            cell_h = (area.h - 2 * cap - gv) / 2
            hs_ = []
            tot = 0.0
            for p_, _ in items:
                iw_, ih_ = img_dims(p_)
                rr = contain(iw_, ih_, Rect(0, 0, cw, cell_h))
                tot += rr.w * rr.h
                hs_.append(rr.h)
            used = max(hs_[0], hs_[1]) + max(hs_[2], hs_[3]) + 2 * cap + gv
            cands.append(("2x2", score(tot, used)))
        mode = max(cands, key=lambda c: c[1])[0]
        self.m.setdefault("layouts", []).append(mode)
        bottom = area.y
        if mode == "row":
            if h < 120:
                raise LayoutError(f"[{self.m['index']}쪽 figures] 그림 높이 {h:.0f}px — 그림 수를 줄일 것")
            widths = [a * h for a in asp]
            total = sum(widths) + g * (n - 1)
            if total >= 0.9 * area.w:
                if photo:          # 사진은 가로로 조금 크롭해 양 끝(120~1800)에 정확히 맞춘다
                    k_ = (area.w - g * (n - 1)) / sum(widths)
                    widths = [wd * k_ for wd in widths]
                    total = area.w
                x0 = area.x
            else:
                x0 = area.x + (area.w - total) / 2
            x = x0
            for (path, c), wd in zip(items, widths):
                box = Rect(x, area.y, wd, h)
                _, b = self._place(path, box, c, Rect(x - g / 2, area.y, wd + g, area.h), photo)
                bottom = max(bottom, b)
                x += wd + g
            used = Rect(x0, area.y, total, bottom - area.y)
        elif mode == "1+2":
            lw = (area.w - g) * 0.6
            rw = area.w - g - lw
            cell_h = (area.h - 2 * cap - gv) / 2
            _, b0 = self._place(items[0][0], Rect(area.x, area.y, lw, area.h - cap), items[0][1],
                                Rect(area.x, area.y, lw, area.h), photo)
            _, b1 = self._place(items[1][0], Rect(area.x + lw + g, area.y, rw, cell_h), items[1][1],
                                Rect(area.x + lw + g, area.y, rw, area.h), photo)
            _, b2 = self._place(items[2][0], Rect(area.x + lw + g, b1 + gv, rw, cell_h), items[2][1],
                                Rect(area.x + lw + g, area.y, rw, area.h), photo)
            used = Rect(area.x, area.y, area.w, max(b0, b2) - area.y)
        else:
            cw = (area.w - g) / 2
            cell_h = (area.h - 2 * cap - gv) / 2
            y = area.y
            bottom = y
            for r_ in range(2):
                row_b = y
                for c_ in range(2):
                    path, cp = items[r_ * 2 + c_]
                    x = area.x + c_ * (cw + g)
                    _, b = self._place(path, Rect(x, y, cw, cell_h), cp, Rect(x, area.y, cw, area.h), photo)
                    row_b = max(row_b, b)
                y = row_b + gv
                bottom = row_b
            used = Rect(area.x, area.y, area.w, bottom - area.y)
        return self._end(tok, used)

    @_module()
    def text_figure(self, items, path, caption=None, rect: Rect | None = None, photo=False, side="left") -> Rect:
        """글 + 그림 1장. ① 그림 왼쪽 60% + 설명 오른쪽(06-24) ② 설명 위 + 그림 전폭 아래 — 그림이 큰 쪽을 고른다.
        보고용은 ②(위 글·아래 그림)를 ①의 85% 이상이면 우선한다(06-18 통일성)."""
        tok = self._begin(rect, "text_figure", flush=False)
        area = rect or self._area()
        paras = self._bullet_paras(items, indent=False)
        iw, ih = img_dims(path)
        cap = self._cap_h() if caption else 0
        g = GUTTER[2]
        tw = min(max(natural_width(paras, self.T["key"]), 360), 0.42 * area.w)
        try:
            laid_h, th = layout_paras(paras, tw, "text_figure", key_pt=self.T["key"])
        except WrapError:
            th = float("inf")            # 옆 칸(최대 42%)에 줄 수 안에 안 들어감 → 좌우 배치 불가, 위 글·아래 그림으로
        fw = area.w - tw - g
        img_h = (fit_photo(iw, ih, Rect(0, 0, fw, area.h - cap))[0] if photo else contain(iw, ih, Rect(0, 0, fw, area.h - cap)))
        laid_v, tv = layout_paras(self._bullet_paras(items, indent=True), area.w, "text_figure", key_pt=self.T["key"])
        vbox = Rect(0, 0, area.w, max(1.0, area.h - tv - self.T["block"] - cap))
        img_v = fit_photo(iw, ih, vbox)[0] if photo else contain(iw, ih, vbox)
        a_h = img_h.w * img_h.h if th <= area.h else 0
        a_v = img_v.w * img_v.h
        prefer_v = (self.d.kind == "report" and a_v >= 0.85 * a_h) or iw / ih >= 1.6
        horiz = not prefer_v and a_h >= a_v and th <= area.h
        self.m.setdefault("layouts", []).append("A" if horiz else "R")
        if rect is None and self.pending_key is not None:
            if horiz:
                self._flush_key()
                area = self._area()
                tok["mark"], tok["nb"], tok["np"] = len(self.s.shapes._spTree), len(self.m["boxes"]), len(self.m["pics"])
                fw = area.w - tw - g
            else:
                kp = self._key_para(self.pending_key[0])
                self.pending_key = None
                top = self._bullet_paras(items, indent=True)
                if top:
                    top[0].before_pt = self.T["key_body_pt"] if top[0].role == "body" else self.T["group_pt"]
                self.n += 1
                t = self.text(Rect(area.x, area.y, area.w, area.h), [kp] + top, name=f"v4:bul{self.n}")
                fig = self.figure(path, caption, rect=Rect(area.x, t.b + self.T["block"], area.w,
                                                             area.b - t.b - self.T["block"]), photo=photo)
                return self._end(tok, Rect(area.x, area.y, area.w, fig.b - area.y))
        if horiz:
            if side == "left":
                fbox, tx = Rect(area.x, area.y, fw, area.h), area.x + fw + g
            else:
                fbox, tx = Rect(area.x + tw + g, area.y, fw, area.h), area.x
            fig = self.figure(path, caption, rect=fbox, photo=photo)
            img_r = Rect(*self.m["pics"][-1]["rect"])
            ty = img_r.cy - th / 2 if th < 0.5 * img_r.h else img_r.y
            self.n += 1
            self.text(Rect(tx, ty, tw, th), paras, name=f"v4:bul{self.n}")
            used = Rect(area.x, area.y, area.w, max(fig.h, ty + th - area.y))
        else:
            self.n += 1
            t = self.text(Rect(area.x, area.y, area.w, tv), self._bullet_paras(items, indent=True), name=f"v4:bul{self.n}")
            fig = self.figure(path, caption, rect=Rect(area.x, t.b + self.T["block"], area.w,
                                                         area.b - t.b - self.T["block"]), photo=photo)
            used = Rect(area.x, area.y, area.w, fig.b - area.y)
        return self._end(tok, used)

    def text_then(self, items):
        """세부 불릿만 먼저 놓고(요지와 한 상자) 이어서 다른 모듈을 쓸 때의 편의 함수 = bullets(items)."""
        return self.bullets(items)

    @_module()
    def photo(self, path, caption, notes, rect: Rect | None = None) -> Rect:
        """사진 1장 크게(왼쪽) + 설명 2~4줄(오른쪽, 06-20). 사진은 크롭 허용."""
        tok = self._begin(rect, "photo")
        area = rect or self._area()
        paras = [self._body_para(t, 0, 0 if i == 0 else self.T["item_pt"], max_lines=2) for i, t in enumerate(notes)]
        tw = min(max(natural_width(paras), 360), 0.40 * area.w)
        fw = area.w - tw - GUTTER[2]
        fig = self.figure(path, caption, rect=Rect(area.x, area.y, fw, area.h), photo=True)
        img_r = Rect(*self.m["pics"][-1]["rect"])
        laid, th = layout_paras(paras, tw, "photo notes")
        ty = img_r.cy - th / 2 if th < 0.5 * img_r.h else img_r.y
        self.text(Rect(area.x + fw + GUTTER[2], ty, tw, th), paras, name="v4:notes")
        return self._end(tok, Rect(area.x, area.y, area.w, max(fig.h, ty + th - area.y)))

    # ── 표 ──
    @_module()
    def table(self, header, rows, rect: Rect | None = None, col_w=None, align=None, summary_last=False,
              width=None, dense=None, fill=False) -> Rect:
        """진짜 표(add_table). 셀 줄도 엔진이 확정(a:br), 행 높이 = 줄 수 × 피치 + 여백. 셀 줄바꿈은 1회까지.

        정렬(§9.4): 첫 열 왼쪽, 수치·단위만 있는 열과 8자 이하 짧은 값 열은 가운데, 그 밖 글 열은 왼쪽.
                    머리는 열 정렬을 따른다(머리와 내용이 떨어져 보이지 않게).
        폭: 자연 폭이 영역의 50%를 넘으면 전폭, 아니면 자연 폭 × 1.3(영역의 50% 이상)을 가운데에.
            남는 폭은 가운데 정렬 열에 고르게, 왼쪽 정렬 글 열은 자연 폭의 +35%까지만.
        dense: None 이면 기본 크기로 안 들어갈 때만 밀집(본문 14pt, 머리는 그대로).
        fill=True: 남는 높이를 모든 행(머리 포함)에 고르게, 행마다 +0.5줄까지."""
        tok = self._begin(rect, "table")
        T = self.T
        area = rect or self._area()
        ncol = len(header)
        allrows = [list(map(str, header))] + [[str(c) for c in r] for r in rows]
        px_, py_ = T["cell_px"], T["cell_py"]
        if align is None:
            align = []
            for j in range(ncol):
                vals = [r[j] for r in allrows[1:] if r[j]]
                if j == 0:
                    align.append("l")
                elif vals and all(is_numcell(v) and len(v) <= 28 for v in vals):
                    align.append("c")
                elif vals and max(len(v) for v in vals) <= 8:
                    align.append("c")
                else:
                    align.append("l")

        def plan(dense_):
            pt_h, pt = T["th"], (T["td_dense"] if dense_ else T["td"])
            if col_w is not None:
                cw = list(col_w)
            else:
                nat = []
                for j in range(ncol):
                    wmax = max(spans_fit(to_spans(r[j], 7 if i == 0 else 4, C["INK"]), pt_h if i == 0 else pt)
                               for i, r in enumerate(allrows))
                    nat.append(wmax / SLACK + 2 * px_ + 2)
                full = sum(nat) > 0.5 * area.w
                tot = min(width or (area.w if full else max(sum(nat) * 1.3, 0.5 * area.w)), area.w)
                if sum(nat) > tot:
                    mins = []
                    for j in range(ncol):
                        mw = 0.0
                        for i, r in enumerate(allrows):
                            if not r[j]:
                                continue
                            wgt, size = (7, pt_h) if i == 0 else (4, pt)
                            sp = to_spans(r[j], wgt, C["INK"])
                            words = _words("".join(x.text for x in sp))
                            longest = max(spans_fit(slice_spans(sp, a, b), size) for a, b in words) if words else 0.0
                            mw = max(mw, longest, 0.55 * spans_fit(sp, size))
                        mins.append(mw / SLACK + 2 * px_ + 2)
                    cw = _shrink_cols(nat, tot, mins)
                else:
                    cw, extra = list(nat), tot - sum(nat)
                    lefts = [j for j in range(ncol) if align[j] == "l"]
                    centers = [j for j in range(ncol) if align[j] != "l"]
                    if centers:
                        for j in lefts:
                            add = min(0.35 * nat[j], extra * nat[j] / sum(nat))
                            cw[j] += add
                            extra -= add
                        for j in centers:
                            cw[j] += extra / len(centers)
                    else:
                        for j in range(ncol):
                            cw[j] += min(0.35 * nat[j], extra * nat[j] / sum(nat))
                        if full:            # 왼쪽선·오른쪽 끝을 다른 전폭 모듈과 맞춘다. 빈 폭은 마지막 열 오른쪽에(시선 이동 없음)
                            cw[-1] += tot - sum(cw)
            if sum(cw) > area.w + 0.5:
                raise LayoutError(f"[{self.m['index']}쪽 table] 폭 {sum(cw):.0f} > {area.w:.0f}")
            cells_, rh = [], []
            for i, r in enumerate(allrows):
                row_, hmax = [], 0
                for j, val in enumerate(r):
                    wgt, size = (7, pt_h) if i == 0 else (4, pt)
                    spans = to_spans(val, wgt, C["INK"]) if val else []
                    ln = wrap(spans, size, cw[j] - 2 * px_, 2, what=f"{self.m['index']}쪽 표 셀 '{val[:12]}'") if val else []
                    row_.append((spans, ln, wgt, size))
                    hmax = max(hmax, max(1, len(ln)) * pitch(size) + 2 * py_)
                cells_.append(row_)
                rh.append(hmax)
            return pt_h, pt, cw, cells_, rh

        if dense is None:
            try:
                pt_h, pt, cw, cells, row_h = plan(False)
                if sum(row_h) > area.h + 0.5:
                    raise LayoutError("tall")
                dense = False
            except (LayoutError, WrapError):
                pt_h, pt, cw, cells, row_h = plan(True)
                dense = True
        else:
            pt_h, pt, cw, cells, row_h = plan(dense)
        tot_w, tot_h = sum(cw), sum(row_h)
        self._fit(tot_h, area, "table")
        if fill and tot_h < area.h:
            each = (area.h - tot_h) / len(row_h)
            for i in range(len(row_h)):
                row_h[i] += min(each, 0.5 * pitch(pt_h if i == 0 else pt))
            tot_h = sum(row_h)
        x = area.x + ((area.w - tot_w) / 2 if tot_w < 0.9 * area.w else 0)
        gf = self.s.shapes.add_table(len(allrows), ncol, E(x), E(area.y), E(tot_w), E(tot_h))
        gf.name = f"v4:tbl{len(self.m['modules'])}"
        tbl = gf.table
        tbl.first_row = False
        tbl.horz_banding = False
        sid = tbl._tbl.tblPr.find(qn("a:tableStyleId"))
        if sid is not None:
            tbl._tbl.tblPr.remove(sid)
        for j, c_w in enumerate(cw):
            tbl.columns[j].width = E(c_w)
        for i, rh_ in enumerate(row_h):
            tbl.rows[i].height = E(rh_)
        last_i = len(allrows) - 1
        for i, r in enumerate(cells):
            for j, (spans, ln, wgt, size) in enumerate(r):
                cell = tbl.cell(i, j)
                cell.margin_left = cell.margin_right = E(px_)
                cell.margin_top = cell.margin_bottom = E(py_)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                bg = C["TABLE_HEAD"] if i == 0 else (C["SUMMARY"] if summary_last and i == last_i else C["WHITE"])
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(bg)
                p = cell.text_frame.paragraphs[0]
                _style_para(p, align[j])
                for li, (a, b) in enumerate(ln):
                    _emit_line(p, slice_spans(spans, a, b), size)
                    if li < len(ln) - 1:
                        _add_br(p, size, C["INK"], wgt)
                _end_para(p, size, C["INK"], wgt)
                _cell_borders(cell, C["LINE"], 0.75)
        used = Rect(x, area.y, tot_w, tot_h)
        self.m["boxes"].append(dict(name=gf.name, rect=used.as_list(), dense=dense, table=[[
            ["".join(s_.text for s_ in sp)[a:b] for a, b in ln] for sp, ln, _, _ in r] for r in cells]))
        return self._end(tok, used)

    # ── 카드·비교·흐름 ──
    @_module()
    def cards(self, items, rect: Rect | None = None, pills=None, valign="t", fill=False) -> Rect:
        """필 머리 카드 N개(등폭·등높이, 회색 테두리). items = [(제목, [줄...]), ...]. pills = 필 색 목록.
        카드 글은 항상 윗변 + 필 절반 + 16px 에서 시작한다(§9.3, 카드끼리 첫 줄이 맞음).
        fill=True: 카드 높이를 영역(rect) 높이까지 늘린다. 이웃 칸과 높이를 맞출 때 쓴다.
        자연 높이의 1.8배를 넘으면 빈 카드로 보이므로 경고한다."""
        tok = self._begin(rect, "cards")
        T = self.T
        area = rect or self._area()
        n = len(items)
        cols = area.hsplit([1] * n)
        pills = pills or [C["PILL"]] * n
        laid_all, hs = [], []
        for (title, lines), col in zip(items, cols):
            inner = col.w - 2 * T["card_px"]
            paras = [self._body_para(t, 0, T["item_pt"] if k else 0) for k, t in enumerate(lines)]
            laid, h = layout_paras(paras, inner, f"{self.m['index']}쪽 카드 '{title}'")
            self._orphans(laid, f"카드 '{title}'")
            laid_all.append(laid)
            hs.append(h)
        ph = T["pill_h"]
        top_pad = ph / 2 + 16
        card_h = top_pad + max(hs) + T["card_py"]
        total = card_h + ph / 2
        self._fit(total, area, "cards")
        if fill and total < area.h:        # 영역을 채우는 같은 높이 카드(글은 윗변 기준)
            nat_h = card_h
            card_h = area.h - ph / 2
            total = card_h + ph / 2
            if card_h > 1.8 * nat_h:
                self.d.warnings.append(f"{self.m['index']}쪽 카드: fill 로 자연 높이의 {card_h / nat_h:.1f}배 — 빈 카드로 보임."
                                       " 내용을 더하거나 fill 을 끄고 그림·표를 더할 것")
        y0 = area.y + ((area.h - total) / 2 if valign == "m" else 0)
        for (title, _), col, laid, pc in zip(items, cols, laid_all, pills):
            card = Rect(col.x, y0 + ph / 2, col.w, card_h)
            _shape(self.s, MSO_SHAPE.ROUNDED_RECTANGLE, card, "v4:card", C["WHITE"], C["LINE_LIGHT"], 1.0, 24)
            tw = spans_fit(to_spans(title, 7, C["WHITE"]), T["pill"]) / SLACK
            pw = min(card.w - 2 * 16, max(tw + 48, 0.4 * card.w))
            pill = Rect(card.cx - pw / 2, y0, pw, ph)
            _shape(self.s, MSO_SHAPE.ROUNDED_RECTANGLE, pill, "v4:pill", pc, None, radius=ph / 2)
            self.text(pill, [Para(title, T["pill"], 7, C["WHITE"], align="c", max_lines=1, role="pill")],
                      name="v4:pilltext", anchor="m")
            h = sum(it["before"] * PX_PT + len(it["lines"]) * pitch(it["para"].pt) for it in laid)
            r = Rect(card.x + T["card_px"], card.y + top_pad, card.w - 2 * T["card_px"], h)
            emit_text(self.s, r, laid, "v4:cardtext")
            self.m["boxes"].append(dict(name="v4:cardtext", rect=r.as_list(), paras=_laid_meta(laid)))
        return self._end(tok, Rect(area.x, y0, area.w, total))

    @_module()
    def compare(self, left, right, rect: Rect | None = None) -> Rect:
        """비교 2단: (제목, [줄...]) × 2. 현행 필 회색, 개선 필 #546882."""
        return self.cards([left, right], rect, pills=[C["PILL_ASIS"], C["PILL"]])

    @_module()
    def flow(self, steps, rect: Rect | None = None, numbered=True) -> Rect:
        """단계 흐름 3~5개: steps = [(제목, 설명 or None), ...]. 등폭 옅은 면 상자 + 회색 셰브런."""
        tok = self._begin(rect, "flow")
        T = self.T
        area = rect or self._area()
        n = len(steps)
        g = GUTTER[2]              # 48 (셰브런 20px + 좌우 14px)
        bw = (area.w - g * (n - 1)) / n
        inner = bw - 2 * 20
        laid_all, hs = [], []
        for k, (title, desc) in enumerate(steps, start=1):
            head = f"{k}  {title}" if numbered else title
            paras = [Para(head, T["pill"], 7, C["INK"], align="c", max_lines=2, role="step")]
            if desc:
                paras.append(Para(desc, T["label"], 4, C["INK"], align="c", before_pt=4, max_lines=2, role="step_desc"))
            laid, h = layout_paras(paras, inner, f"{self.m['index']}쪽 흐름 '{title}'")
            self._orphans(laid, f"흐름 '{title}'")
            laid_all.append(laid)
            hs.append(h)
        bh = max(hs) + 2 * 22
        self._fit(bh, area, "flow")
        x = area.x
        for i, laid in enumerate(laid_all):
            box = Rect(x, area.y, bw, bh)
            _shape(self.s, MSO_SHAPE.ROUNDED_RECTANGLE, box, "v4:step", STEPS[i % 3], None, radius=16)
            h = sum(it["before"] * PX_PT + len(it["lines"]) * pitch(it["para"].pt) for it in laid)
            r = Rect(box.x + 20, box.y + 22, inner, h)          # 모든 상자 같은 윗선 → 단계 제목이 한 줄에
            emit_text(self.s, r, laid, "v4:steptext")
            self.m["boxes"].append(dict(name="v4:steptext", rect=r.as_list(), paras=_laid_meta(laid)))
            if i < n - 1:
                aw, ah = 20, 28
                _shape(self.s, MSO_SHAPE.CHEVRON, Rect(box.r + (g - aw) / 2, box.cy - ah / 2, aw, ah), "v4:arrow", C["LINE"], None)
            x += bw + g
        return self._end(tok, Rect(area.x, area.y, area.w, bh))

    @_module()
    def timeline(self, items, rect: Rect | None = None, orient="v") -> Rect:
        """일정: items = [(날짜, 내용), ...]. orient='v' 세로(왼쪽 선 + 점), 'h' 가로(위 날짜, 아래 내용)."""
        tok = self._begin(rect, "timeline")
        T = self.T
        area = rect or self._area()
        dot = 14
        if orient == "v":
            x_line = area.x + dot / 2
            tx = area.x + dot + 20
            paras = []
            for i, (dt, tx_) in enumerate(items):
                paras.append(Para([(dt + "  ", {"w": 7, "color": C["INK"]}), (tx_, {"w": 4, "color": C["TEXT2"]})],
                                  T["body"], 4, C["TEXT2"], before_pt=0 if i == 0 else 10, max_lines=2, role="timeline"))
            laid, h = layout_paras(paras, area.w - (tx - area.x), "timeline")
            self._orphans(laid, "타임라인")
            self._fit(h, area, "timeline")
            r = Rect(tx, area.y, area.w - (tx - area.x), h)
            emit_text(self.s, r, laid, "v4:timeline")
            self.m["boxes"].append(dict(name="v4:timeline", rect=r.as_list(), paras=_laid_meta(laid)))
            ys, y = [], area.y
            for it in laid:
                y += it["before"] * PX_PT
                ys.append(y + pitch(it["para"].pt) / 2)
                y += len(it["lines"]) * pitch(it["para"].pt)
            if len(ys) > 1:
                self.line(x_line, ys[0], x_line, ys[-1], C["LINE_LIGHT"], 1.5, "v4:tl_line")
            for yy in ys:
                self.shape(MSO_SHAPE.OVAL, Rect(x_line - dot / 2, yy - dot / 2, dot, dot), C["GREEN"], None, name="v4:tl_dot")
            return self._end(tok, Rect(area.x, area.y, area.w, h))
        n = len(items)
        colw = area.w / n
        top_h = pitch(T["body"])
        y_line = area.y + top_h + 12 + dot / 2
        self.line(area.x + colw / 2, y_line, area.r - colw / 2, y_line, C["LINE_LIGHT"], 1.5, "v4:tl_line")
        bottom = y_line
        for i, (dt, tx_) in enumerate(items):
            cx = area.x + colw * (i + 0.5)
            self.shape(MSO_SHAPE.OVAL, Rect(cx - dot / 2, y_line - dot / 2, dot, dot), C["GREEN"], None, name="v4:tl_dot")
            self.text(Rect(area.x + colw * i + 8, area.y, colw - 16, top_h),
                      [Para(dt, T["body"], 7, C["INK"], align="c", max_lines=1, role="tl_date")], name="v4:tl_date")
            r = self.text(Rect(area.x + colw * i + 8, y_line + dot / 2 + 12, colw - 16, area.b - y_line - 20),
                          [Para(tx_, T["label"], 4, C["TEXT2"], align="c", max_lines=3, role="tl_text")], name="v4:tl_text")
            bottom = max(bottom, r.b)
        return self._end(tok, Rect(area.x, area.y, area.w, bottom - area.y))

    @_module()
    def stat(self, main, sub, rect: Rect) -> Rect:
        """수치 2줄(평문, 숫자를 키우지 않음): 주문장 = 요지 토큰, 보조 = 세부 토큰."""
        return self.text(rect, [Para(main, self.T["key"], 6, C["INK"], max_lines=1, role="stat"),
                                Para(sub, self.T["body"], 4, C["TEXT2"], before_pt=4, max_lines=2, role="stat_sub")],
                         name="v4:stat")

    @_module()
    def note(self, text, rect: Rect | None = None) -> Rect:
        """결론 박스(요약·결론 장 전용, 장당 1개): 흰 면 + 주황 1.5pt 테두리, 요지 토큰 1~2줄."""
        if self.m["kind"] not in ("summary", "conclusion"):
            self.d.warnings.append(f"{self.m['index']}쪽: 결론 박스는 요약·결론 장 전용(kind='summary'|'conclusion')")
        tok = self._begin(rect, "note")
        T = self.T
        area = rect or self._area()
        inner_w = area.w - 2 * T["card_px"]
        laid, h = layout_paras([Para(text, T["key"], 6, C["INK"], align="c", max_lines=2, role="note")], inner_w)
        self._orphans(laid, "결론 박스")
        bh = h + 2 * T["card_py"]
        self._fit(bh, area, "note")
        box = Rect(area.x, area.y, area.w, bh)
        _shape(self.s, MSO_SHAPE.ROUNDED_RECTANGLE, box, "v4:note", C["WHITE"], C["ORANGE"], 1.5, 20)
        r = Rect(box.x + T["card_px"], box.y + T["card_py"], inner_w, h)
        emit_text(self.s, r, laid, "v4:notetext")
        self.m["boxes"].append(dict(name="v4:notetext", rect=r.as_list(), paras=_laid_meta(laid)))
        return self._end(tok, box)

    def source(self, text):
        """출처(논문·외부 figure 인용 장에만): 우하단 10pt Light, 1줄, 오른쪽 끝 x=1780."""
        if self.d.kind == "report":
            self.d.warnings.append(f"{self.m['index']}쪽: 보고용은 출처 줄 대신 캡션·본문 끝 [n] + References(§9.2)")
        spans = to_spans(text, 3, C["GRAY"])
        w = spans_fit(spans, self.T["source"]) / SLACK + 4
        if w > 1100:
            raise WrapError(f"{self.m['index']}쪽 출처가 김({w:.0f}px) — '저자 et al., 저널 권, 쪽 (연도)' 형식으로 줄일 것")
        r = Rect(1780 - w, 1044.7, w, 23.1)
        laid, _ = layout_paras([Para(text, self.T["source"], 3, C["GRAY"], align="r", max_lines=1, role="source")], w)
        emit_text(self.s, r, laid, "v4:source", anchor="m")
        self.m["source"] = text
        return r


_UNIT_RE = re.compile(r"(MPa|GPa|kPa|Pa|µm|um|nm|mm|cm|km|kg|mg|kW|mW|kV|mV|mA|kHz|MHz|Hz|ms|µs|min|kJ|HV|HRC|HB|wt|at|"
                      r"rpm|dpi|px|pt|GB|GiB|MB|MiB|TB|KB|fps|epochs|epoch|ep|W|V|J|g|h|s|m|"
                      r"개월|시간|개|장|초|분|건|회|명|종|배|점|차|쌍|시편|조건|층|편|권|곳|대|원|억|만|천|조|월|일|년|주)")


def is_numcell(v: str) -> bool:
    """숫자·단위만 있는 표 칸(데이터)인가. 숫자·기호를 지운 뒤 단위만 남으면 참. 영문 코드(B1, r08)·글은 거짓.
    표 수치 열 정렬(§9.4)과 검사기 G9 글자 수 면제(§8)가 같이 쓴다."""
    if not re.search(r"\d", v):
        return False
    rest = re.sub(r"[\d\s.,~±+\-−%/()×:°→²³]", "", v)
    return not _UNIT_RE.sub("", rest)


def _shrink_cols(nat, total, mins=None):
    """자연 폭 합이 total 을 넘으면 여유(현재 폭 − 최소 폭)가 큰 열부터 줄인다. 최소 폭 = 가장 긴 어절·두 줄 기준."""
    mins = mins or [n * 0.55 for n in nat]
    cols = list(nat)
    over = sum(cols) - total
    while over > 0.5:
        j = max(range(len(cols)), key=lambda k: cols[k] - mins[k])
        room = cols[j] - mins[j]
        if room <= 0:
            break
        d = min(room, over)
        cols[j] -= d
        over -= d
    return cols


def _cell_borders(cell, color, w_pt):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        ln = etree.SubElement(tcPr, qn("a:" + tag), w=str(int(w_pt * 12700)), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
    order = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr", "cell3D", "noFill", "solidFill", "gradFill",
             "blipFill", "pattFill", "grpFill", "headers", "extLst"]
    _reorder(tcPr, order)


# ─────────────────────────────────────────────────────────────────────────────
# 10. 차트 도우미 (FT-05: 그래프 글씨는 영어·Arial·크게)
# ─────────────────────────────────────────────────────────────────────────────
def chart_figsize(rect: Rect, caption=True, kind="ppt"):
    """슬롯에 1:1 로 들어갈 matplotlib figsize(인치). 이 크기로 그리면 그림 속 pt = 슬라이드 pt."""
    T = TOK[kind]
    h = rect.h - ((T["cap_gap"] + pitch(T["caption"])) if caption else 0)
    return (rect.w / PX_PER_IN, h / PX_PER_IN)


def mpl_style():
    """matplotlib rcParams — Arial, 눈금 12pt, 축 제목·범례 14pt, 흰 배경, 위·오른쪽 축 없음.
    fig = plt.figure(figsize=chart_figsize(slot)); ... fig.savefig(path, dpi=200)  (bbox_inches='tight' 금지)"""
    import cycler
    return {
        "font.family": "Arial", "font.size": 12, "axes.titlesize": 14, "axes.labelsize": 14,
        "xtick.labelsize": 12, "ytick.labelsize": 12, "legend.fontsize": 14, "legend.frameon": False,
        "axes.spines.top": False, "axes.spines.right": False, "axes.linewidth": 1.0,
        "figure.facecolor": "white", "axes.facecolor": "white", "savefig.facecolor": "white",
        "axes.prop_cycle": cycler.cycler(color=["#71BF44", "#546882", "#71B2C1", "#808080"]),
        "text.color": "#262626", "axes.labelcolor": "#262626", "axes.edgecolor": "#262626",
        "xtick.color": "#262626", "ytick.color": "#262626",
        "figure.constrained_layout.use": True,
    }
