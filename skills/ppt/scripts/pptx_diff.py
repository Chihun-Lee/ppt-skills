#!/usr/bin/env python
"""사용자가 손본 pptx와 이전 생성본을 장·도형 단위로 비교한다 (STANDARD §13 OUT-03 ②, 읽기 전용).

사용: pptx_diff.py <이전 생성본.pptx> <사용자 수정본.pptx> [--json out.json]
  * 장 수가 다르면(사용자가 장을 지우거나 넣음) 글 유사도로 장을 짝지어 비교한다(align, 2026-10-29)
  * 장 수가 같으면 같은 번호의 장끼리 비교한다
  * 도형은 이름으로 짝짓되, 같은 이름이 여러 개면(엔진 v4:step 등) 나온 순서로 짝짓는다
  * 보는 것: 추가·삭제된 도형, 위치·크기(0.5 pt 넘게), 글(문단 단위), 글자 크기·굵기·글꼴·색, 채움·선, 그림 바이트, 표 칸, 노트
출력: 장별 '동일' 또는 바뀐 항목 목록. 끝에 고친 장 번호 요약
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import argparse
import hashlib
import json
import sys
from collections import defaultdict

from pptx import Presentation
from pptx.util import Emu


def _rgb(color):
    try:
        if color is not None and color.type is not None:
            return str(color.rgb) if color.type == 1 else f"theme:{color.theme_color}"
    except Exception:
        return None
    return None


def _fill(sh):
    try:
        f = sh.fill
        return _rgb(f.fore_color) if f.type == 1 else (str(f.type) if f.type is not None else None)
    except Exception:
        return None


def _line(sh):
    try:
        ln = sh.line
        return (_rgb(ln.color) if ln.fill.type == 1 else None, round(Emu(ln.width).pt, 2) if ln.width else 0)
    except Exception:
        return None


def describe(sh):
    d = {"name": sh.name, "type": str(sh.shape_type).split(" ")[0],
         "box": tuple(round(Emu(v).pt, 1) if v is not None else None for v in (sh.left, sh.top, sh.width, sh.height)),
         "fill": _fill(sh), "line": _line(sh), "paras": [], "fmt": [], "img": None}
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            d["paras"].append("".join(r.text for r in p.runs))
            for r in p.runs:
                f = r.font
                d["fmt"].append((f.size.pt if f.size else None, f.bold, f.name, _rgb(f.color)))
    if getattr(sh, "has_table", False) and sh.has_table:
        d["paras"] = [" | ".join(c.text for c in row.cells) for row in sh.table.rows]
    if sh.shape_type == 13:
        try:
            d["img"] = hashlib.md5(sh.image.blob).hexdigest()[:10]
        except Exception:
            d["img"] = "linked/none"
    if sh.shape_type == 6:
        d["paras"] = [s.text_frame.text for s in sh.shapes if s.has_text_frame]
    return d


def pair(a_shapes, b_shapes):
    """이름별로 나온 순서대로 짝짓는다."""
    qa, qb = defaultdict(list), defaultdict(list)
    for s in a_shapes:
        qa[s["name"]].append(s)
    for s in b_shapes:
        qb[s["name"]].append(s)
    pairs, gone, new = [], [], []
    for nm in list(qa) + [n for n in qb if n not in qa]:
        la, lb = qa.get(nm, []), qb.get(nm, [])
        for i in range(max(len(la), len(lb))):
            x = la[i] if i < len(la) else None
            y = lb[i] if i < len(lb) else None
            if x and y:
                pairs.append((x, y))
            elif x:
                gone.append(x)
            else:
                new.append(y)
    return pairs, gone, new


def diff_shape(x, y):
    out = []
    if any(abs((p or 0) - (q or 0)) > 0.5 for p, q in zip(x["box"], y["box"])):
        out.append(f"위치·크기 {x['box']} → {y['box']}")
    if x["paras"] != y["paras"]:
        for i in range(max(len(x["paras"]), len(y["paras"]))):
            p = x["paras"][i] if i < len(x["paras"]) else None
            q = y["paras"][i] if i < len(y["paras"]) else None
            if p != q:
                out.append(f"글[{i}] {p!r} → {q!r}")
    elif x["fmt"] != y["fmt"]:
        out.append(f"글자 서식 {x['fmt'][:4]} → {y['fmt'][:4]}")
    if x["fill"] != y["fill"]:
        out.append(f"채움 {x['fill']} → {y['fill']}")
    if x["line"] != y["line"]:
        out.append(f"선 {x['line']} → {y['line']}")
    if x["img"] != y["img"]:
        out.append(f"그림 {x['img']} → {y['img']}")
    return out


def notes(sl):
    return sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else ""


def slide_text(sl) -> str:
    """짝짓기용 장 글(도형 글 + 노트 앞부분)."""
    return " | ".join(" ".join(describe(sh)["paras"]) for sh in sl.shapes) + " || " + notes(sl)[:200]


def align(sa, sb, gap: float = 0.35):
    """장 추가·삭제가 있어도 내용이 비슷한 장끼리 순서대로 짝짓는다(글 유사도 합 최대, 빈 짝은 gap 감점).
    돌려주는 값 = [(이전 i | None, 수정본 j | None), ...]"""
    from difflib import SequenceMatcher
    ta, tb = [slide_text(s) for s in sa], [slide_text(s) for s in sb]
    n, m = len(ta), len(tb)
    sim = [[SequenceMatcher(None, x, y, autojunk=False).ratio() for y in tb] for x in ta]
    dp = [[float("-inf")] * (m + 1) for _ in range(n + 1)]
    bt = [[None] * (m + 1) for _ in range(n + 1)]
    dp[0][0] = 0.0
    for i in range(n + 1):
        for j in range(m + 1):
            for di, dj, w, tag in ((1, 1, None, "m"), (1, 0, -gap, "d"), (0, 1, -gap, "a")):
                if i >= di and j >= dj and dp[i - di][j - dj] > float("-inf"):
                    v = dp[i - di][j - dj] + (sim[i - 1][j - 1] if tag == "m" else w)
                    if v > dp[i][j]:
                        dp[i][j], bt[i][j] = v, tag
    out, i, j = [], n, m
    while i or j:
        tag = bt[i][j]
        if tag == "m":
            out.append((i - 1, j - 1)); i, j = i - 1, j - 1
        elif tag == "d":
            out.append((i - 1, None)); i -= 1
        else:
            out.append((None, j - 1)); j -= 1
    return out[::-1]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("before")
    ap.add_argument("after")
    ap.add_argument("--json")
    a = ap.parse_args()
    A, B = Presentation(a.before), Presentation(a.after)
    sa, sb = list(A.slides), list(B.slides)
    print(f"이전 {len(sa)}장 / 수정본 {len(sb)}장" + ("" if len(sa) == len(sb) else "  ※ 장 수가 다름 — 내용으로 장을 짝지어 비교"))
    report, changed = {}, []
    pairs_ = align(sa, sb) if len(sa) != len(sb) else [(i, i) for i in range(len(sa))]
    for i, j in pairs_:
        if i is None or j is None:
            what = f"수정본 {j + 1}장: 수정본에만 있음" if i is None else f"이전 {i + 1}장: 수정본에서 삭제"
            sl = sb[j] if i is None else sa[i]
            print(f"\n=== {what} — {slide_text(sl)[:70]!r}")
            changed.append(f"+{j + 1}" if i is None else f"-{i + 1}")
            continue
        label = f"{i + 1}장" if i == j else f"이전 {i + 1}장 = 수정본 {j + 1}장"
        da = [describe(s) for s in sa[i].shapes]
        db = [describe(s) for s in sb[j].shapes]
        items = []
        pairs, gone, new = pair(da, db)
        for s in gone:
            items.append(f"[삭제] {s['name']} ({s['type']}) {' / '.join(s['paras'])[:80]!r}")
        for s in new:
            items.append(f"[추가] {s['name']} ({s['type']}) {' / '.join(s['paras'])[:80]!r} 채움 {s['fill']}")
        for x, y in pairs:
            for c in diff_shape(x, y):
                items.append(f"[변경] {x['name']}: {c}")
        if notes(sa[i]) != notes(sb[j]):
            items.append(f"[노트] {notes(sa[i])[:60]!r} → {notes(sb[j])[:60]!r}")
        if items:
            changed.append(i + 1)
            print(f"\n=== {label}: {len(items)}건")
            for it in items:
                print("  ", it)
        report[label] = items
    print(f"\n고친 장: {changed or '없음'}")
    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            json.dump({"changed": changed, "slides": report}, f, ensure_ascii=False, indent=1)
    sys.exit(0)


if __name__ == "__main__":
    main()
