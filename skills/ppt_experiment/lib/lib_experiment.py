# -*- coding: utf-8 -*-
"""ppt_experiment 헬퍼 라이브러리 — 실험보고(자료 공유용) PPT 프리미티브.

스타일: 원작자의 실험보고 샘플 pptx 에서 뽑은 디자인 토큰.
덱 폴더로 복사하지 말고 build.py 에서 sys.path 로 불러 쓴다(SKILL.md 4절).
캔버스: 1280×720 px 그리드 (1px = 9525 EMU, 16:9)
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image
import copy

# ── 디자인 토큰 (샘플 실측) ────────────────────────────────
NAVY = RGBColor(0x00, 0x20, 0x60)        # 제목띠·섹션헤더·표 헤더
BLUE_LIGHT = RGBColor(0xEA, 0xF2, 0xFA)  # 데이터표 헤더행
BLUE_MID = RGBColor(0xC4, 0xDC, 0xF2)    # 강조행 (Average 등)
GRAY_BAND = RGBColor(0xE7, 0xE6, 0xE6)   # 개요 섹션 하이라이트
BAND_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)  # 결과 슬라이드 제목띠
INK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x00, 0x00)         # 주의·이상치 강조 전용
GRAY_LINE = RGBColor(0xBF, 0xBF, 0xBF)

FONT = "맑은 고딕"   # pptx 에 적는 글꼴(Windows 기본). Mac·Linux 미리보기는 LibreOffice 가 대체 글꼴로 그린다

# ── 폰트 하한 (원작자 지정: 최소 13pt) ──────
# 어떤 호출이 더 작은 값을 넘겨도 _set_font가 이 값으로 올려 잡는다.
MIN_FONT_PT = 13     # 본문·표·이미지라벨 공통 하한 — 어떤 호출도 이보다 작아지지 않는다

PX = lambda v: Emu(int(v * 9525))
MARGIN = 40          # 좌우 기본 마진(px)
CONTENT_W = 1200


def new_prs():
    prs = Presentation()
    prs.slide_width = PX(1280)
    prs.slide_height = PX(720)
    return prs


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _flat(shape):
    """테마 style 참조(p:style) 제거 — LibreOffice가 그리는 기본 그림자 차단."""
    sp = shape._element
    style = sp.find(qn('p:style'))
    if style is not None:
        sp.remove(style)
    shape.shadow.inherit = False


def _text_w(text, size=14):
    """섹션 헤더 폭 근사 (한글 = 전각)."""
    units = sum(1.9 if ord(c) > 0x2E80 else 1.0 for c in text)
    return int(units * size * 0.72) + 34


def _set_font(run, size, bold=False, color=INK, italic=False, floor=None):
    f = run.font
    f.name = FONT
    f.size = Pt(max(size, floor if floor is not None else MIN_FONT_PT))
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    # 한글(East Asian) 폰트 명시
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', FONT)


def textbox(slide, x, y, w, h, lines, size=14, bold=False, color=INK,
            align=PP_ALIGN.LEFT, line_spacing=1.15, anchor=MSO_ANCHOR.TOP):
    """lines: str 또는 [(text, {size,bold,color}) …] 리스트."""
    tb = slide.shapes.add_textbox(PX(x), PX(y), PX(w), PX(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if isinstance(line, tuple):
            text, opt = line
        else:
            text, opt = line, {}
        r = p.add_run()
        r.text = text
        _set_font(r, opt.get('size', size), opt.get('bold', bold),
                  opt.get('color', color))
    return tb


# ── 제목띠 ────────────────────────────────────────────────
def title_band(slide, text, style="light"):
    """style='navy': 개요 슬라이드용(네이비 바탕+흰 글씨),
    'light': 결과 슬라이드용(연회색 바탕+네이비 글씨)."""
    navy = style == "navy"
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PX(0), PX(16), PX(1280), PX(52))
    band.fill.solid()
    band.fill.fore_color.rgb = NAVY if navy else BAND_LIGHT
    band.line.fill.background()
    _flat(band)
    tf = band.text_frame
    tf.margin_left = PX(24)
    tf.margin_right = PX(24)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    _set_font(r, 22, True, WHITE if navy else NAVY)
    return band


# ── 섹션 헤더 ─────────────────────────────────────────────
def section_numbered(slide, x, y, text, w=None):
    """개요 슬라이드용: 회색 하이라이트 박스 + '1. Xxx' 검정 볼드."""
    w = w or _text_w(text, 16)
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PX(x), PX(y), PX(w), PX(32))
    box.fill.solid()
    box.fill.fore_color.rgb = GRAY_BAND
    box.line.fill.background()
    _flat(box)
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = PX(10)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    _set_font(r, 16, True, INK)
    return box


def section_diamond(slide, x, y, text):
    """결과 슬라이드용: '❖ Xxx' 네이비 볼드."""
    return textbox(slide, x, y, min(640, 1280 - x - 8), 26, f"❖  {text}", size=15, bold=True, color=NAVY)


# ── 불릿 ─────────────────────────────────────────────────
def kv_bullets(slide, x, y, w, items, size=14, gap=1.45):
    """'• key : value' 개조식 목록. items: str 리스트 (문장 금지, 명사구·수치만)."""
    lines = [f"•  {it}" for it in items]
    return textbox(slide, x, y, w, len(items) * size * 2, lines,
                   size=size, line_spacing=gap)


def arrow_notes(slide, x, y, w, items, size=14, color=INK):
    """'➢' caveat·조치필요 사항 전용."""
    lines = [f"➢  {it}" for it in items]
    return textbox(slide, x, y, w, len(items) * size * 2.2, lines,
                   size=size, line_spacing=1.4, color=color)


# ── 표 ───────────────────────────────────────────────────
def _cell_border(cell, weight=6350, color="000000"):
    """모든 변에 0.5pt 검정 테두리. 스키마 순서상 a:lnX는 fill 앞이어야 함."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
        ln = tcPr.makeelement(qn(tag), {'w': str(weight), 'cap': 'flat'})
        fill = ln.makeelement(qn('a:solidFill'), {})
        clr = fill.makeelement(qn('a:srgbClr'), {'val': color})
        fill.append(clr)
        ln.append(fill)
        tcPr.append(ln)
    # lnL~lnB가 solidFill보다 앞에 오도록 fill을 맨 뒤로 이동
    for f in tcPr.findall(qn('a:solidFill')):
        tcPr.remove(f)
        tcPr.append(f)


def _style_cell(cell, text, size, bold, font_color, fill_color, align=PP_ALIGN.CENTER):
    if fill_color is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill_color
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = PX(4)
    cell.margin_top = cell.margin_bottom = PX(1)
    tf = cell.text_frame
    tf.word_wrap = True
    # 줄바꿈은 문단으로 나눈다 — run 안의 줄바꿈은 LibreOffice에서 첫 줄 정렬이 어긋난다
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        _set_font(r, size, bold, font_color)
    _cell_border(cell)


def exp_table(slide, x, y, w, headers, rows, style="navy",
              col_ratios=None, size=13, row_h=30, header_h=32,
              highlight_last=False, align=PP_ALIGN.CENTER):
    """실험보고 표.
    style='navy'  : 헤더 네이비 채움 + 흰 글씨 (시편 정보표)
    style='blue'  : 헤더 연블루(EAF2FA) + 검정 글씨 (데이터표)
    highlight_last: 마지막 행(Average 등) BLUE_MID + bold
    """
    n_rows, n_cols = len(rows) + 1, len(headers)
    gt = slide.shapes.add_table(n_rows, n_cols, PX(x), PX(y),
                                PX(w), PX(header_h + row_h * len(rows)))
    tbl = gt.table
    # 기본 스타일(밴딩) 제거
    tbl.first_row = False
    tbl.horz_banding = False
    if col_ratios:
        total = sum(col_ratios)
        for i, ratio in enumerate(col_ratios):
            tbl.columns[i].width = PX(int(w * ratio / total))
    tbl.rows[0].height = PX(header_h)
    for ri in range(1, n_rows):
        tbl.rows[ri].height = PX(row_h)

    hdr_fill = NAVY if style == "navy" else BLUE_LIGHT
    hdr_color = WHITE if style == "navy" else INK
    for ci, h in enumerate(headers):
        _style_cell(tbl.cell(0, ci), h, size, True, hdr_color, hdr_fill)
    for ri, row in enumerate(rows, start=1):
        last = highlight_last and ri == n_rows - 1
        for ci, val in enumerate(row):
            _style_cell(tbl.cell(ri, ci), val, size, last, INK,
                        BLUE_MID if last else WHITE, align=align)
    return gt


# ── 이미지 ────────────────────────────────────────────────
def add_image_fit(slide, path, x, y, w, h, border=False):
    """비율 보존 등비축소 + 박스 중앙 정렬. 반환: (draw_x, draw_y, draw_w, draw_h)."""
    iw, ih = Image.open(path).size
    scale = min(w / iw, h / ih)
    dw, dh = iw * scale, ih * scale
    dx, dy = x + (w - dw) / 2, y + (h - dh) / 2
    slide.shapes.add_picture(str(path), PX(dx), PX(dy), PX(dw), PX(dh))
    if border:
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PX(dx), PX(dy), PX(dw), PX(dh))
        box.fill.background()
        box.line.color.rgb = GRAY_LINE
        box.line.width = Pt(0.5)
        _flat(box)
    return dx, dy, dw, dh


def image_row(slide, x, y, w, h, images, gap=8, label_size=13, label_pos="top"):
    """이미지 스트립: images=[(path, label), …] 균등 배치, 라벨은 이미지 위/아래."""
    n = len(images)
    cell_w = (w - gap * (n - 1)) / n
    label_h = label_size * 1.9
    img_h = h - label_h
    for i, (path, label) in enumerate(images):
        cx = x + i * (cell_w + gap)
        if label_pos == "top":
            textbox(slide, cx, y, cell_w, label_h, label, size=max(label_size, MIN_FONT_PT),
                    bold=True, align=PP_ALIGN.CENTER)
            add_image_fit(slide, path, cx, y + label_h, cell_w, img_h)
        else:
            add_image_fit(slide, path, cx, y, cell_w, img_h)
            textbox(slide, cx, y + img_h + 2, cell_w, label_h, label,
                    size=label_size, bold=True, align=PP_ALIGN.CENTER)

# ── 연구노트 전용 표지 (원작자 지정) ────────────
def note_cover(slide, *, month, project, account=None, author=None,
               pi=None, period=None, hero=None, topics=None):
    """연구노트 표지 — '09월 연구노트' + 과제명 + 메타.

    ppt_experiment 는 원칙적으로 표지를 두지 않지만, 전자연구노트로 등록되는
    월간 노트는 첫 장에 대상 과제·월이 보여야 하므로 예외로 허용한다.
    """
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PX(0), PX(0), PX(1280), PX(250))
    band.fill.solid(); band.fill.fore_color.rgb = NAVY
    band.line.fill.background(); _flat(band)
    band.text_frame.text = ""

    textbox(slide, MARGIN + 8, 74, 700, 56, month, size=44, bold=True, color=WHITE)
    textbox(slide, MARGIN + 10, 146, 760, 32, "연구노트", size=26, bold=True,
            color=RGBColor(0xC4, 0xDC, 0xF2))

    y = 292
    textbox(slide, MARGIN, y, 1200, 40, project, size=24, bold=True, color=NAVY)
    y += 52
    add_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      PX(MARGIN), PX(y), PX(90), PX(4))
    add_rect.fill.solid(); add_rect.fill.fore_color.rgb = BLUE_MID
    add_rect.line.fill.background(); _flat(add_rect)
    add_rect.text_frame.text = ""
    y += 28

    meta = []
    if account: meta.append(["계정번호", account])
    if period:  meta.append(["연구기간", period])
    if author:  meta.append(["작성자", author])
    if pi:      meta.append(["연구책임자", pi])
    if meta:
        exp_table(slide, MARGIN, y, 560, ["항목", "내용"], meta,
                  style="navy", col_ratios=[0.30, 0.70], align=PP_ALIGN.LEFT)

    if topics:
        section_diamond(slide, 660, y - 34, "주요 내용")
        kv_bullets(slide, 664, y + 4, 580, topics)

    if hero:
        add_image_fit(slide, hero, 660, 40, 580, 200, border=False)
    return slide
