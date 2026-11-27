#!/usr/bin/env python
"""v4.1 자동 게이트 — STANDARD.md §12.2 G1~G18.

사용: check.py <deck.pptx> [--pdf render/deck.pdf] [--json out.json] [--full] [--state review_state.json]
  * pptx 옆의 <deck>.v4.json (엔진 메타: 예측 줄·박스·그림·모듈)을 함께 읽는다
  * --pdf 가 있으면 렌더 기반 검사(폰트 임베드, 예측 줄 == 렌더 줄, 상자 밖 글자, 빈 띠)를 더한다
  * 증분 출력이 기본: review_state.json(scripts/review_state.py)에 장별 내용 해시와 결과를 남기고, 다음 회차에는
    처음 보는 장·고친 장·지난 fail 장만 출력한다(fail 은 항상 출력). 전 장을 보려면 --full
종료 코드: fail 이 하나라도 있으면 1. 게이트는 시각 점검을 대신하지 않는다.
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import argparse
import io
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree

LIB = Path(__file__).resolve().parents[1] / "lib"
sys.path.insert(0, str(LIB))
import kims_ppt as K  # noqa: E402

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
R = "{%s}" % NS["r"]

FACE = re.compile(r"^Paperlogy ([3467]) (Light|Regular|SemiBold|Bold)$")
HEADER = ("v4:title", "v4:chapter", "v4:crumb", "v4:source")
CHROME_SMALL = {K.FIXED["crumb"], K.TOK["ppt"]["source"], K.TOK["ppt"]["refs"], K.FIXED["div_prog"]}
ALLOWED_COLORS = {K.C[k] for k in ("INK", "TEXT2", "GRAY", "MUTED", "WHITE", "GREEN", "NUM_LIGHT", "CLOSING")}
FORBIDDEN = ["혁신적", "획기적", "핵심적", "게임체인저", "시너지", "전례 없는", "비약적", "청사진", "대전환",
             "눈부신", "총체적", "새로운 지평", "한 단계 도약", "미래를 선도", "중요한 이정표", "초격차", "선도적",
             "티어1", "티어2", "티어 1", "티어 2", "Tier 1", "Tier 2",
             "delve", "holistic", "pivotal", "cutting-edge", "groundbreaking", "transformative", "harness",
             "unleash", "seamless", "leverage", "tapestry", "landscape", "realm", "elevate", "streamline"]
PLACEHOLDER = re.compile(r"(TODO|TBD|XXX|lorem|초안|보완 영역|작성 중|\(미정\)|placeholder)", re.I)
VERSION = re.compile(r"(\bv\d+(\.\d+)?\b|\b[Ww]ave\s?\d\b)")
GEN_LABEL = re.compile(r"(?<![\d.])[A-Za-z가-힣]+\s?[1-4]\.0(?![\d%])")
DASH = re.compile(r"[—–]|\s-\s")
KEY_DATE = re.compile(r"(?<![\d.])\d{1,2}/\d{1,2}(?![\d/])|\d{1,2}월\s?\d{1,2}일")
ENDS_DA = re.compile(r"(다|해요|세요|에요|예요|어요|아요|네요|나요|군요|죠)\s*[.!]?$")
ENDS_PERIOD = re.compile(r"[가-힣A-Za-z)\]]\.$")
WHY_HOW = re.compile(r"^(왜|어떻게)\s*[:?]")


class Report:
    def __init__(self):
        self.items = []

    def fail(self, g, s, m):
        self.items.append(dict(gate=g, level="fail", slide=s, msg=m))

    def warn(self, g, s, m):
        self.items.append(dict(gate=g, level="warn", slide=s, msg=m))


def px(v):
    return int(v) / K.EPX


def rect_of(el):
    # xfrm 안의 off/ext만 본다. PowerPoint로 저장한 그림은 blipFill 의 a:extLst/a:ext(uri, cx 없음)가
    # xfrm 보다 앞에 와서 .//a:ext 가 그것을 집는다(2026-11-02 원작자 최종본에서 확인)
    for xf in el.iter(A + "xfrm", P + "xfrm"):
        off, ext = xf.find(A + "off"), xf.find(A + "ext")
        if off is not None and ext is not None and ext.get("cx") is not None:
            return K.Rect(px(off.get("x")), px(off.get("y")), px(ext.get("cx")), px(ext.get("cy")))
    return None


def name_of(el):
    nv = el.find(".//" + P + "cNvPr")
    return nv.get("name") if nv is not None else ""


def load(pptx: Path):
    z = zipfile.ZipFile(pptx)
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    out = []
    for s in pres.find(P + "sldIdLst"):
        t = rid2t[s.get(R + "id")]
        name = "ppt/" + t if not t.startswith("ppt/") else t
        x = etree.fromstring(z.read(name))
        srel = etree.fromstring(z.read(name.replace("slides/", "slides/_rels/") + ".rels"))
        out.append((name, x, {r.get("Id"): r.get("Target") for r in srel}))
    return z, out


def run_face(rPr):
    faces = {rPr.find(A + t).get("typeface") for t in ("latin", "ea", "cs") if rPr.find(A + t) is not None}
    return faces


def check_pptx(pptx: Path, meta: dict, rep: Report):
    z, slides = load(pptx)
    kind = meta.get("kind", "ppt")
    T = K.TOK[kind]
    allowed_body = {v for k, v in T.items() if isinstance(v, (int, float)) and k in (
        "title", "key", "body", "caption", "label", "th", "td", "td_dense", "pill", "toc_num", "toc_ch", "toc_sec",
        "refs", "source")} | {K.FIXED["chapter"], K.FIXED["crumb"]}
    allowed_sizes = allowed_body | set(K.FIXED.values())      # 표지·간지·감사 장(원형)만: 40·28·20·17·80·37·48·14
    min_body = T["caption"] if kind == "ppt" else 14
    toc_items = set(meta.get("toc_items", []))
    titles, sigs = [], []
    for si, (name, x, emb) in enumerate(slides, start=1):
        m = meta["slides"][si - 1] if si - 1 < len(meta.get("slides", [])) else {}
        sk = m.get("kind", "content")
        tree = x.find(".//" + P + "spTree")
        shapes = []
        for el in tree.iter(P + "sp", P + "pic", P + "graphicFrame", P + "cxnSp"):
            nm = name_of(el)
            if nm.startswith("v4:"):
                shapes.append((nm, el, rect_of(el)))
        # ── 도형 단위: G11 장식·G18 무결성
        for nm, el, r in shapes:
            if el.find(P + "style") is not None:
                rep.fail("G11", si, f"{nm}: p:style 잔존(그림자 원인)")
            if el.find(".//" + A + "outerShdw") is not None:
                rep.fail("G11", si, f"{nm}: 그림자")
            GREENS = ("71BF44", "8DC63F", "6FA32E", "47B521")
            if etree.QName(el).localname == "cxnSp":
                clr = el.find(".//" + A + "ln//" + A + "srgbClr")
                if r and r.h < 1 and clr is not None and clr.get("val", "").upper() in GREENS:
                    rep.fail("G11", si, f"{nm}: 녹색 가로줄")
            if etree.QName(el).localname == "sp" and r and r.h <= 12 and r.w >= 12 * max(r.h, 1):
                fillc = el.find(P + "spPr/" + A + "solidFill/" + A + "srgbClr")
                if fillc is not None and fillc.get("val", "").upper() in GREENS:
                    rep.fail("G11", si, f"{nm}: 녹색 가로 막대(도형) — 가로 구분선은 회색 0.75pt만(§9.3)")
            if r and (r.w < 0 or r.h < 0):
                rep.fail("G18", si, f"{nm}: 음수 크기")
        if sum(1 for nm, _, _ in shapes if nm == "v4:note") > 1:
            rep.fail("G11", si, "결론 박스 2개 이상(장당 강조 장치 1개)")
        if len(shapes) > 25:
            rep.warn("G17", si, f"엔진 도형 {len(shapes)}개 > 25")
        # ── 글자 단위: G1·G2·G8·G11·G13·G18
        narr = vis = colon = middot = 0
        for nm, el, r in shapes:
            txb = el.find(".//" + P + "txBody")
            if txb is None:
                txb = el.find(".//" + A + "tbl")
            if txb is None:
                continue
            is_tbl = etree.QName(el).localname == "graphicFrame"
            rows = el.findall(".//" + A + "tr") if is_tbl else [None]
            paras = []
            if is_tbl:
                for ri, tr in enumerate(rows):
                    for p in tr.iter(A + "p"):
                        paras.append((p, f"row{ri}"))
            else:
                for p in txb.iter(A + "p"):
                    paras.append((p, None))
            for p, tag in paras:
                runs = p.findall(A + "r")
                txt = "".join(r.findtext(A + "t") or "" for r in runs).strip()
                bu = p.find(A + "pPr/" + A + "buChar")
                bu_c = bu.get("char") if bu is not None else None
                for br in p.findall(A + "br"):
                    brp = br.find(A + "rPr")
                    if brp is None or not brp.get("sz"):
                        rep.fail("G18", si, f"{nm}: a:br 서식(sz) 누락")
                if runs and p.find(A + "endParaRPr") is None:
                    rep.fail("G18", si, f"{nm}: endParaRPr 누락")
                for rn in runs:
                    rPr = rn.find(A + "rPr")
                    t = rn.findtext(A + "t") or ""
                    if rPr is None or not t:
                        continue
                    faces = run_face(rPr)
                    fc = next(iter(faces)) if len(faces) == 1 else None
                    mm = FACE.match(fc or "")
                    if fc == "Arial":
                        bad = [ch for ch in t if ch.strip() and K.in_paperlogy(ch)]
                        if bad:
                            rep.fail("G1", si, f"{nm}: Arial run 에 Paperlogy 로 쓸 수 있는 글자 {bad[:5]}")
                        w = 7 if rPr.get("b") == "1" else 4
                    elif mm:
                        w = int(mm.group(1))
                        if rPr.get("b") == "1":
                            rep.fail("G1", si, f"{nm}: b=1 (Paperlogy 는 굵기를 이름으로)")
                        miss = [ch for ch in t if ch.strip() and not K.in_paperlogy(ch, w)]
                        if miss:
                            rep.fail("G13", si, f"{nm}: Paperlogy 에 없는 글자 {sorted(set(miss))}")
                    else:
                        rep.fail("G1", si, f"{nm}: 폰트 표기 {sorted(faces)}")
                        w = 4
                    if any(f.startswith("+") for f in faces):
                        rep.fail("G1", si, f"{nm}: 테마 글꼴 참조 {sorted(faces)}")
                    sz = int(rPr.get("sz", "0")) / 100
                    if sz not in (allowed_sizes if sk in ("cover", "divider", "closing") else allowed_body):
                        rep.fail("G2", si, f"{nm}: 크기 {sz}pt 는 {'이 장의 ' if sz in allowed_sizes else ''}토큰 밖"
                                 + (" (표지·간지 전용 크기, 본문 장에서 숫자 키우기 금지)" if sz in allowed_sizes else ""))
                    if sz < min_body and sz not in CHROME_SMALL and not nm.startswith(("v4:refs", "v4:source", "v4:crumb", "v4:div_progress")):
                        rep.fail("G2", si, f"{nm}: {sz}pt < 본문 최소 {min_body}pt")
                    clr = rPr.find(A + "solidFill/" + A + "srgbClr")
                    cv = clr.get("val", "").upper() if clr is not None else ""
                    if cv == "000000":
                        rep.fail("G11", si, f"{nm}: 순검정 글자")
                    elif cv and cv not in ALLOWED_COLORS:
                        rep.fail("G11", si, f"{nm}: 허용 밖 글자색 #{cv}")
                    elif cv in (K.C["GREEN"], K.C["NUM_LIGHT"]) and not nm.startswith(("v4:toc", "v4:div")):
                        rep.fail("G11", si, f"{nm}: 녹색 글자 '{t[:12]}' — 녹색 글자는 목차·간지 번호에만(§10)")
                    if w >= 6:
                        bold_bad = (bu_c == "•") or nm.startswith(("v4:cap", "v4:cardtext", "v4:notes")) \
                            or (tag and tag != "row0") or (nm.startswith("v4:steptext") and False)
                        if bold_bad:
                            rep.fail("G11", si, f"{nm}: 본문·캡션·표 본문에 굵은 글자 '{t[:15]}'")
                if not txt:
                    continue
                # 글자 수 (숫자·단위만 있는 표 칸은 데이터라 세지 않는다)
                chars = len(re.sub(r"\s", "", txt))
                if is_tbl and tag != "row0" and K.is_numcell(txt):
                    chars = 0
                if not nm.startswith(HEADER) and not nm.startswith(("v4:toc", "v4:refs", "v4:cover", "v4:div", "v4:closing")):
                    vis += chars
                    if nm.startswith(("v4:bul", "v4:key", "v4:notes", "v4:stat", "v4:notetext")):
                        narr += chars
                    colon += len(re.findall(r":(?=\s|[가-힣A-Za-z])", txt))
                    middot += txt.count("·")
                if nm.startswith(("v4:refs", "v4:source")):
                    continue
                if DASH.search(txt):
                    rep.fail("G8", si, f"{nm}: 대시·띄운 하이픈 '{txt[:30]}'")
                if ";" in txt:
                    rep.fail("G8", si, f"{nm}: 세미콜론")
                if ENDS_PERIOD.search(txt) and not re.search(r"(et al|Fig|vs|No|Eq)\.$", txt):
                    rep.fail("G8", si, f"{nm}: 마침표 종결 '{txt[-20:]}'")
                if ENDS_DA.search(txt) and txt not in ("감사합니다",):
                    rep.fail("G8", si, f"{nm}: 문장형 종결 '{txt[-20:]}'")
                if "?" in txt:
                    rep.fail("G8", si, f"{nm}: 물음표 '{txt[:30]}'")
                if WHY_HOW.search(txt):
                    rep.fail("G8", si, f"{nm}: '왜/어떻게' 라벨")
                for wd in FORBIDDEN:
                    if wd.lower() in txt.lower():
                        rep.fail("G8", si, f"{nm}: 금칙어 '{wd}'")
                if PLACEHOLDER.search(txt):
                    rep.fail("G8", si, f"{nm}: 자리표시 문자열 '{txt[:30]}'")
                if VERSION.search(txt):
                    rep.warn("G8", si, f"{nm}: 버전 표기 '{VERSION.search(txt).group(0)}' (외부 보고면 삭제)")
                if GEN_LABEL.search(txt):
                    rep.warn("G8", si, f"{nm}: 세대 라벨 의심 '{GEN_LABEL.search(txt).group(0)}'")
                if (nm == "v4:title" or bu_c == "■") and ":" in txt:
                    rep.fail("G7", si, f"{nm}: 제목·요지에 콜론")
                # 발표용 중간제목(■) = 짧은 주제 이름표 (2026-10-27 원작자 지시: 결론·날짜·→ 연쇄 없이 간결하게)
                if kind == "ppt" and bu_c == "■" and sk == "content":
                    n_key = len(re.sub(r"\s", "", txt))
                    if "→" in txt:
                        rep.fail("G7", si, f"중간제목에 → 연쇄 '{txt[:24]}' — 이름표로 쓰고 흐름은 도형(flow)으로")
                    if KEY_DATE.search(txt):
                        rep.fail("G7", si, f"중간제목에 날짜 '{KEY_DATE.search(txt).group(0)}' — 날짜는 빼거나 일정 장으로")
                    if n_key > 22:
                        rep.warn("G7", si, f"중간제목 {n_key}자 '{txt[:24]}…' — 8~20자 이름표로, 결론·수치는 세부(•)로")
                    elif txt.count(",") >= 1 and n_key > 14:
                        rep.warn("G7", si, f"중간제목 쉼표 나열 '{txt[:24]}' — 결론·수치 나열이면 세부로")
        if colon > 2:
            rep.fail("G8", si, f"본문 콜론 {colon}개 (> 2)")
        if middot > 2:
            rep.warn("G8", si, f"가운뎃점 {middot}개 — 나열은 쉼표")
        # ── G9 글자량
        if sk in ("content", "summary", "conclusion"):
            nf = T["narr_sum"] if sk != "content" else T["narr_fail"]
            if narr > nf:
                rep.fail("G9", si, f"서술 글자 {narr}자 > {nf}")
            elif sk == "content" and narr > T["narr_warn"]:
                rep.warn("G9", si, f"서술 글자 {narr}자 > {T['narr_warn']}")
            if vis > T["vis_fail"]:
                rep.fail("G9", si, f"보이는 글자 {vis}자 > {T['vis_fail']}")
            elif vis > T["vis_warn"]:
                rep.warn("G9", si, f"보이는 글자 {vis}자 > {T['vis_warn']}")
        # ── G6 줄 (메타)
        key_lines = body_lines = 0
        for b in m.get("boxes", []):
            for row_ in b.get("table", []):             # 표 셀 고아 줄(두 줄 셀의 마지막 줄이 한 어절)
                for cell_lines in row_:
                    if len(cell_lines) >= 2 and " " not in cell_lines[-1].strip():
                        rep.warn("G6", si, f"{b['name']}: 표 셀 고아 줄 '{' / '.join(cell_lines)[:30]}'")
            for p in b.get("paras", []):
                n = len(p["lines"])
                if p["role"] == "key":
                    key_lines += n
                    if n > 1:
                        rep.fail("G6", si, f"요지 {n}줄 '{p['lines'][0][:20]}…'")
                if p["role"] == "body" and b["name"].startswith("v4:bul"):
                    body_lines += n
                if p["role"] == "body" and n > T["body_wrap"]:
                    rep.fail("G6", si, f"세부 {n}줄 > {T['body_wrap']} '{p['lines'][0][:20]}…'")
                if p.get("orphan"):
                    rep.fail("G6", si, f"고아 줄 '{p['lines'][-1]}'")
                # G15 폴백 폭 재검증
                wbox = b["rect"][2]
                for ln_ in p["lines"]:
                    try:
                        sp = K.to_spans(ln_, p["w"], K.C["INK"])
                        if K.spans_fit(sp, p["pt"]) > wbox + 1:
                            rep.fail("G15", si, f"{b['name']}: 맑은 고딕 폭 초과 '{ln_[:20]}'")
                    except Exception:
                        pass
        if sk == "content":
            if key_lines > T["key_lines"]:
                rep.fail("G9", si, f"요지 {key_lines}줄 > {T['key_lines']}")
            if body_lines > T["body_lines"]:
                rep.fail("G9", si, f"세부 {body_lines}줄 > {T['body_lines']}")
        elif sk in ("summary", "conclusion") and key_lines > T["key_lines_sum"]:
            rep.fail("G9", si, f"요지 {key_lines}줄 > {T['key_lines_sum']}")
        # ── G7 제목
        if sk == "content" and m.get("title"):
            t = m["title"]
            titles.append(t)
            if toc_items and t not in toc_items:
                rep.fail("G7", si, f"제목 '{t}' 이 목차 항목에 없음")
            if re.search(r"[.?:!]$", t) or ENDS_DA.search(t):
                rep.fail("G7", si, f"제목 문장형/부호 '{t}'")
            hang = len(re.findall(r"[가-힣]", t))
            if hang > 20:
                rep.warn("G7", si, f"국문 제목 {hang}자 > 20 '{t}'")
        # ── G4 겹침·G5 프레임·G16 분할
        def is_text(el):
            return etree.QName(el).localname == "sp" and el.find(".//" + A + "t") is not None
        texts = [(n, r) for n, el, r in shapes if r and is_text(el) and not n.startswith(("v4:pilltext",))
                 and n not in HEADER]
        conts = [(n, r) for n, el, r in shapes if r and n.split(":")[1] in ("card", "note", "step", "pill")]
        pics = [(n, r) for n, el, r in shapes if r and etree.QName(el).localname == "pic"]
        tbls = [(n, r) for n, el, r in shapes if r and etree.QName(el).localname == "graphicFrame"]

        def inter(a, b, tol=0.5):
            return min(a.r, b.r) - max(a.x, b.x) > tol and min(a.b, b.b) - max(a.y, b.y) > tol

        def inside(a, b, tol=1.0):
            return a.x >= b.x - tol and a.y >= b.y - tol and a.r <= b.r + tol and a.b <= b.b + tol

        free_texts = [(n, r) for n, r in texts if not any(inside(r, c) for _, c in conts)]
        objs = free_texts + pics + tbls
        for i in range(len(objs)):
            for j in range(i + 1, len(objs)):
                (n1, r1), (n2, r2) = objs[i], objs[j]
                if inter(r1, r2):
                    rep.fail("G4", si, f"겹침 {n1} ↔ {n2}")
        for n1, r1 in texts:
            for n2, r2 in conts:
                if inter(r1, r2) and not inside(r1, r2):
                    rep.fail("G4", si, f"{n1} 이 {n2} 테두리를 넘음")
        boxes_c = [(n, r) for n, r in conts if not n.startswith("v4:pill")]
        for n1, r1 in boxes_c:
            for n2, r2 in pics + tbls:
                if inter(r1, r2):
                    rep.fail("G4", si, f"겹침 {n1} ↔ {n2}")
        for i in range(len(boxes_c)):
            for j in range(i + 1, len(boxes_c)):
                (n1, r1), (n2, r2) = boxes_c[i], boxes_c[j]
                if inter(r1, r2):
                    rep.fail("G4", si, f"겹침 {n1} ↔ {n2}")
        for n1, r1 in free_texts:
            for n2, r2 in pics + tbls:
                gap = max(r2.x - r1.r, r1.x - r2.r, r2.y - r1.b, r1.y - r2.b)
                if 0 <= gap < K.MIN_GAP and not n1.startswith("v4:cap"):
                    rep.warn("G4", si, f"간격 {gap:.1f}px < {K.MIN_GAP}px: {n1} ↔ {n2}")
        if sk in ("content", "summary", "conclusion", "toc", "refs"):
            for n, el, r in shapes:
                if not r or n in HEADER or n.startswith("v4:toc_div"):
                    continue
                wide = etree.QName(el).localname in ("pic", "graphicFrame")
                x0, x1, y1 = (K.WIDE_X0, K.WIDE_X1, K.WIDE_Y1) if wide else (K.FRAME_X0, K.FRAME_X1, K.FRAME_Y1)
                if r.x < x0 - 1 or r.r > x1 + 1 or r.y < K.FRAME_Y0 - 1 or r.b > y1 + 1:
                    rep.fail("G5", si, f"{n} 이 본문 프레임 밖 {r.as_list()}")
        stack = sorted([(n, r) for n, r in free_texts if n.startswith(("v4:bul", "v4:key", "v4:notes", "v4:text", "v4:stat"))],
                       key=lambda t: t[1].y)
        for (n1, r1), (n2, r2) in zip(stack, stack[1:]):
            if abs(r1.x - r2.x) <= 4 and abs(r1.w - r2.w) <= 4 and 0 <= r2.y - r1.b < T["block"] - 1:
                rep.fail("G16", si, f"같은 폭 텍스트박스가 나뉘어 쌓임: {n1} → {n2} (한 상자로)")
        # ── G10 그림
        for n, el, r in shapes:
            if etree.QName(el).localname != "pic" or not r:
                continue
            blip = el.find(".//" + A + "blip")
            rid = blip.get(R + "embed") if blip is not None else None
            tgt = emb.get(rid)
            if not tgt:
                continue
            try:
                from PIL import Image
                im = Image.open(io.BytesIO(z.read("ppt/" + tgt.replace("../", ""))))
                iw, ih = im.size
            except Exception:
                continue
            src = el.find(".//" + A + "srcRect")
            cl = ct = cr = cb = 0.0
            if src is not None:
                cl, ct, cr, cb = (int(src.get(k, "0")) / 100000 for k in ("l", "t", "r", "b"))
            vis_ar = (iw * (1 - cl - cr)) / (ih * (1 - ct - cb))
            d = (r.w / r.h) / vis_ar
            if abs(d - 1) > 0.01:
                rep.fail("G10", si, f"{n}: 비율 왜곡 {d:.3f}")
            up = r.w / (iw * (1 - cl - cr))
            if up > 1.6:
                rep.fail("G10", si, f"{n}: {up:.2f}배 확대(흐림)")
            elif up > 1.25:
                rep.warn("G10", si, f"{n}: {up:.2f}배 확대")
        capd = {n: r for n, _, r in shapes if n.startswith("v4:cap")}
        picd = {n: r for n, _, r in shapes if n.startswith("v4:fig")}
        for cn, cr_ in capd.items():
            pr = picd.get("v4:fig" + cn.replace("v4:cap", ""))
            if pr is None:
                continue
            if abs(cr_.y - (pr.b + T["cap_gap"])) > 2:
                rep.fail("G10", si, f"{cn}: 그림 아래 {T['cap_gap']}px 에 있지 않음")
            if abs(cr_.cx - pr.cx) > 2 and cr_.w < K.FRAME.w - 2:
                rep.warn("G10", si, f"{cn}: 그림 가운데와 {cr_.cx - pr.cx:+.0f}px 어긋남")
        sigs.append((si, sk, tuple(m.get("modules", [])) + tuple(m.get("layouts", [])) + (f"그림{len(m.get('pics', []))}",),
                     m.get("title")))
    # ── G7 반복률
    if len(titles) >= 8:
        dup = 1 - len(set(titles)) / len(titles)
        if dup < 0.25 or dup > 0.7:
            rep.warn("G7", 0, f"대제목 반복률 {dup:.0%} (참고 25~70%)")
    # ── G14 반복 (/ppt 만)
    if kind == "ppt":
        run = 1
        for (a, ka, sa, ta), (b, kb, sb, tb) in zip(sigs, sigs[1:]):
            if ka == kb == "content" and sa == sb and sa and ta != tb:
                run += 1
                if run >= 3:
                    rep.warn("G14", b, f"같은 구성 {run}장 연속 {sa}")
            else:
                run = 1


def check_pdf(pdf: Path, meta: dict, rep: Report):
    try:
        import pymupdf as fitz  # noqa: F401
    except ImportError:  # PyMuPDF 1.24.3 미만
        import fitz  # noqa: F401
    fitz.TOOLS.mupdf_display_errors(False)
    doc = fitz.open(str(pdf))
    for si, page in enumerate(doc, start=1):
        m = meta["slides"][si - 1] if si - 1 < len(meta.get("slides", [])) else {}
        s = 1920 / page.rect.width
        d = page.get_text("rawdict")
        fonts = Counter()
        lines = []
        for b in d["blocks"]:
            if b.get("type") != 0:
                continue
            for ln in b["lines"]:
                txt = "".join(ch["c"] for sp in ln["spans"] for ch in sp["chars"])
                if not txt.strip():
                    continue
                for sp in ln["spans"]:
                    fn = re.sub(r"^[A-Z]{6}\+", "", sp["font"])
                    fonts[fn] += sum(1 for ch in sp["chars"] if not ch["c"].isspace())
                bb = [v * s for v in ln["bbox"]]
                lines.append((txt, K.Rect(bb[0], bb[1], bb[2] - bb[0], bb[3] - bb[1])))
        for fn, n in fonts.items():
            if "Paperlogy" not in fn and "Arial" not in fn and n > 0:
                rep.fail("G1", si, f"렌더 폰트 폴백 {fn} ({n}자)")
        for b in m.get("boxes", []):
            if "paras" not in b:
                continue
            r = K.Rect(*b["rect"])
            pred = [ln_ for p in b["paras"] for ln_ in p["lines"]]
            got = []
            for txt, lr in lines:
                if r.x - 4 <= lr.cx <= r.r + 4 and r.y - 4 <= lr.cy <= r.b + 4:
                    got.append((lr.y, txt))
                    if lr.r > r.r + 6 or lr.x < r.x - 6:
                        rep.fail("G3", si, f"{b['name']}: 글자가 상자 밖(가로) '{txt[:20]}'")
            norm = lambda t_: re.sub(r"[\s\u00a0■•]+", "", t_)      # 렌더 PDF 글에는 불릿 기호가 섞인다
            g_ = [x for x in (norm(t) for _, t in sorted(got)) if x]
            p_ = [x for x in (norm(t) for t in pred) if x]
            if g_ and p_ and "".join(g_) == "".join(p_):
                if g_ != p_:
                    rep.fail("G6", si, f"{b['name']}: 렌더 줄 {len(g_)} ≠ 예측 {len(p_)} (재줄바꿈)")
            elif g_ and p_:
                rep.warn("G6", si, f"{b['name']}: 렌더 글자열이 예측과 달라 줄 비교를 못 함 — 확대해서 볼 것")
        if m.get("kind") in ("content", "summary", "conclusion"):
            pix = page.get_pixmap(matrix=fitz.Matrix(480 / page.rect.width, 480 / page.rect.width),
                                  colorspace=fitz.csGRAY, alpha=False)
            k = 480 / 1920
            x0, x1 = int(K.FRAME.x * k), int(K.FRAME.r * k)
            y0, y1 = int(K.FRAME.y * k), int(K.FRAME.b * k)
            W = pix.width
            rows = [any(v < 240 for v in pix.samples[y * W + x0: y * W + x1]) for y in range(y0, y1)]   # 240: 카드 테두리(217)도 잉크, 배경 무늬(≥250)는 아님
            best = cur = 0
            for ink in rows:
                cur = 0 if ink else cur + 1
                best = max(best, cur)
            frac = best / max(1, len(rows))
            pics = m.get("pics") or []
            has_pic = bool(pics)
            # §4.3 면제: 비율을 지킨 그림이 이미 폭을 95% 이상 쓰거나 확대 한도(1.25배)에 닿았으면 더 키울 수 없다 → warn
            span = (max(pp["rect"][0] + pp["rect"][2] for pp in pics) - min(pp["rect"][0] for pp in pics)) if pics else 0
            capped = pics and (span >= 0.95 * K.FRAME.w or any((pp.get("scale") or 0) >= 1.25 for pp in pics))
            if frac > 0.30 and has_pic and capped:
                rep.warn("G12", si, f"그림 장 빈 띠 {frac:.0%} — 그림이 폭·확대 한도에 닿아 면제(§4.3), 그림 수·배치 검토")
            elif frac > 0.30 and has_pic:
                rep.fail("G12", si, f"그림 장 빈 띠 {frac:.0%} (> 30%) — 배치를 바꾸거나 그림을 키울 것")
            elif frac > 0.30:
                rep.warn("G12", si, f"본문 빈 띠 {frac:.0%} — 표·글 장은 그림·도식 추가 검토")
            elif best / k > 100:
                rep.warn("G12", si, f"본문 빈 띠 {best / k:.0f}px (> 100px)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--pdf")
    ap.add_argument("--json")
    ap.add_argument("--full", action="store_true", help="지난 회차 기록과 상관없이 전 장 결과를 출력")
    ap.add_argument("--state", help="검토 기록 파일(기본 pptx 폴더의 review_state.json)")
    a = ap.parse_args()
    pptx = Path(a.pptx)
    mp = pptx.with_suffix(".v4.json")
    meta = json.loads(mp.read_text()) if mp.exists() else {"slides": []}
    rep = Report()
    check_pptx(pptx, meta, rep)
    if a.pdf:
        check_pdf(Path(a.pdf), meta, rep)
    # 증분 출력(STANDARD §12.1): 계산은 전 장(덱 단위 규칙 때문), 출력은 처음 보는 장·고친 장·지난 fail 장만.
    # fail 은 어느 장이든 숨기지 않는다. 바뀌지 않고 지난 회차에 통과한 장의 warn 은 이미 판단했으므로 생략
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import review_state as RS
    show, todo, skip = None, [], []
    if not a.full:
        todo, skip, _ = RS.queue(pptx, a.state, kind="gate")
        if skip:
            show = set(todo) | {0}
    acc = set() if a.full else RS.accepted_pages(pptx, a.state)
    # 수용한 fail(내용이 그대로인 장)은 출력·종료 코드에서 뺀다. 장이 바뀌면 해시가 달라져 다시 나온다
    shown = [i for i in rep.items if (show is None or i["slide"] in show or i["level"] == "fail")
             and not (i["slide"] in acc and i["slide"] not in todo)]
    fails = [i for i in rep.items if i["level"] == "fail" and not (i["slide"] in acc and i["slide"] not in todo)]
    warns = [i for i in shown if i["level"] == "warn"]
    for i in sorted(shown, key=lambda i: (i["slide"], i["gate"])):
        print(f"[{i['level']:4}] {i['gate']:3} p{i['slide']:<3} {i['msg']}")
    gates = sorted({i["gate"] for i in shown})
    print(f"\n== {pptx.name}: fail {len(fails)} · warn {len(warns)}" + (f"  ({', '.join(gates)})" if gates else "  (전 게이트 통과)"))
    if show is not None:
        print(f"   증분 검사: 출력 {len(todo)}쪽({', '.join(map(str, todo)) or '없음'}), 건너뜀 {len(skip)}쪽"
              f"(바뀌지 않았고 지난 회차 통과 또는 fail 수용, {len(rep.items) - len(shown)}건 생략). 전 장은 --full")
    (Path(a.json) if a.json else pptx.with_suffix(".check.json")).write_text(json.dumps(rep.items, ensure_ascii=False, indent=1))
    RS.record_gate(pptx, rep.items, a.state)
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
