#!/usr/bin/env python
"""검토 회차 기록 — 걸린 장·바뀐 장만 다음 회차에 다시 본다 (STANDARD §12.1, 2026-10-29 "검토에 걸리는 부분만 다음 이터레이션에").

장마다 내용 해시(장 XML + 그 장이 쓰는 그림·차트·레이아웃 부품 바이트, 노트 제외)를 키로 결과를 남긴다.
  gate     = pass | warn | fail   (check.py 가 기록)
  visual   = todo | ok | flag     (시각 점검 뒤 mark 로 기록)
  accepted = True                 (판단을 마친 fail: PowerPoint 저장 바탕 장의 G18 등 거짓 fail, 원문 유지 fail — mark --accept)
다음 회차 대상(queue) = 처음 보는 해시(새 장·고친 장) ∪ 수용 안 한 gate fail ∪ visual todo·flag.
내용이 같으면 위치가 바뀌어도(장 삽입·삭제·재배치) 다시 보지 않는다.

사용:
  review_state.py queue <deck.pptx> [--png render/png] [--state 파일]   다음 회차에 볼 장과 PNG 경로
  review_state.py mark  <deck.pptx> --ok 3,5-7 [--flag "9=글 넘침"] [--accept 1-6] [--state 파일]
  review_state.py show  <deck.pptx> [--state 파일]                       장별 기록
  review_state.py reset <deck.pptx> [--state 파일]                       기록 지움(다음 회차 전 장)
기록 파일 기본값 = pptx 와 같은 폴더의 review_state.json (여러 덱이 같이 써도 해시 키라 섞이지 않음).
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

SKIP_REL = ("notesSlide", "package")  # 노트는 화면에 안 보임. 차트의 내장 엑셀(package)은 빌드 시각이 들어가 매번 바뀌지만
                                      # 화면은 차트 XML 캐시 값으로 그려지므로 뺀다


def _canon(name: str, data: bytes) -> bytes:
    """XML 부품은 정규화(C14N)해서 PowerPoint 저장과 python-pptx 저장의 속성·이름공간 순서 차이를 없앤다."""
    if name.endswith((".xml", ".rels")):
        try:
            from lxml import etree
            return etree.tostring(etree.fromstring(data), method="c14n")
        except Exception:
            return data
    return data


def _rels(z: zipfile.ZipFile, part: str) -> list[tuple[str, str]]:
    d, f = posixpath.split(part)
    rp = posixpath.join(d, "_rels", f + ".rels")
    if rp not in z.namelist():
        return []
    x = z.read(rp).decode("utf-8", "replace")
    out = []
    for m in re.finditer(r"<Relationship\b[^>]*>", x):
        t = m.group(0)
        typ = re.search(r'Type="([^"]+)"', t)
        tgt = re.search(r'Target="([^"]+)"', t)
        if not typ or not tgt or 'TargetMode="External"' in t:
            continue
        out.append((typ.group(1).rsplit("/", 1)[-1], posixpath.normpath(posixpath.join(d, tgt.group(1)))))
    return out


def page_hashes(pptx: Path) -> list[str]:
    """쪽 순서대로 장 내용 해시(16자). 장 XML + 관계 부품(그림·차트·차트 데이터·레이아웃·마스터) 바이트."""
    z = zipfile.ZipFile(pptx)
    pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    prels = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "replace")):
        rid = re.search(r'Id="([^"]+)"', m.group(0))
        tgt = re.search(r'Target="([^"]+)"', m.group(0))
        if rid and tgt:
            prels[rid.group(1)] = posixpath.normpath(posixpath.join("ppt", tgt.group(1)))
    out = []
    for rid in re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"', pres):
        part = prels[rid]
        comps = [hashlib.sha1(_canon(part, z.read(part))).hexdigest()]
        seen, stack = {part}, [part]
        while stack:                                   # 차트 → 차트 데이터·스타일처럼 한 단계 더 따라간다
            cur = stack.pop()
            for typ, tgt in _rels(z, cur):
                if typ in SKIP_REL or tgt in seen or tgt not in z.namelist():
                    continue
                seen.add(tgt)
                comps.append(typ + ":" + hashlib.sha1(_canon(tgt, z.read(tgt))).hexdigest())
                if typ in ("chart", "chartUserShapes", "diagramData", "diagramDrawing"):
                    stack.append(tgt)
        # 관계 목록 순서는 저장 프로그램마다 달라서(PowerPoint·python-pptx) 장 XML 뒤 부품은 정렬해서 합친다
        out.append(hashlib.sha1("|".join([comps[0]] + sorted(comps[1:])).encode()).hexdigest()[:16])
    return out


def state_path(pptx: Path, state: str | None) -> Path:
    return Path(state) if state else Path(pptx).with_name("review_state.json")


def load_state(p: Path) -> dict:
    if p.exists():
        try:
            return json.loads(p.read_text())
        except Exception:
            pass
    return {"version": 1, "pages": {}}


def save_state(p: Path, st: dict) -> None:
    p.write_text(json.dumps(st, ensure_ascii=False, indent=1))


def queue(pptx: Path, state: str | None = None, kind: str = "visual") -> tuple[list[int], list[int], list[str]]:
    """(다음 회차에 볼 쪽, 건너뛸 쪽, 해시). 쪽 번호는 1부터.
    kind="gate"   : 처음 보는 해시 ∪ 지난 게이트 fail  (check.py --changed 가 이 쪽만 출력)
    kind="visual" : 위 ∪ 시각 점검 todo·flag           (PNG 를 다시 볼 쪽)"""
    hs = page_hashes(pptx)
    st = load_state(state_path(pptx, state))["pages"]
    todo, skip = [], []
    for n, h in enumerate(hs, 1):
        r = st.get(h)
        open_fail = r is not None and r.get("gate") == "fail" and not r.get("accepted")
        need = r is None or open_fail or (kind == "visual" and r.get("visual", "todo") != "ok")
        (todo if need else skip).append(n)
    return todo, skip, hs


def record_gate(pptx: Path, items: list[dict], state: str | None = None) -> None:
    """check.py 결과(전 장)를 해시별로 남긴다. 처음 보는 해시는 visual=todo 로 시작."""
    p = state_path(pptx, state)
    st = load_state(p)
    hs = page_hashes(pptx)
    worst = {}
    for i in items:
        s = i.get("slide", 0)
        if s and 1 <= s <= len(hs):
            lv = i["level"]
            worst[s] = "fail" if lv == "fail" or worst.get(s) == "fail" else "warn"
    for n, h in enumerate(hs, 1):
        r = st["pages"].setdefault(h, {"visual": "todo"})
        r["gate"] = worst.get(n, "pass")
        r["last_page"] = n
        r["deck"] = Path(pptx).name
    save_state(p, st)


def _pages(spec: str) -> list[int]:
    out = []
    for part in filter(None, (s.strip() for s in spec.split(","))):
        if "-" in part:
            a, b = part.split("-", 1)
            out += list(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return out


def accepted_pages(pptx: Path, state: str | None = None) -> set[int]:
    """fail 을 수용한 쪽(내용이 그대로인 동안만)."""
    st = load_state(state_path(pptx, state))["pages"]
    return {n for n, h in enumerate(page_hashes(pptx), 1) if st.get(h, {}).get("accepted")}


def mark(pptx: Path, ok: list[int], flags: dict[int, str], state: str | None = None,
         accept: list[int] | None = None) -> None:
    p = state_path(pptx, state)
    st = load_state(p)
    hs = page_hashes(pptx)
    for n in accept or []:
        st["pages"].setdefault(hs[n - 1], {"gate": "fail", "visual": "todo"})["accepted"] = True
    for n in ok:
        r = st["pages"].setdefault(hs[n - 1], {"gate": "pass"})
        r.update(visual="ok", note="", last_page=n)
    for n, note in flags.items():
        r = st["pages"].setdefault(hs[n - 1], {"gate": "pass"})
        r.update(visual="flag", note=note, last_page=n)
    save_state(p, st)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["queue", "mark", "show", "reset"])
    ap.add_argument("pptx")
    ap.add_argument("--state")
    ap.add_argument("--png", help="PNG 폴더(slide-NN.png). 주면 볼 파일 경로를 함께 적는다")
    ap.add_argument("--ok", default="", help="시각 점검 통과 쪽: 3,5-7")
    ap.add_argument("--flag", action="append", default=[], help='걸린 쪽: "9=글 넘침" (여러 번)')
    ap.add_argument("--accept", default="", help="fail 을 판단 끝에 수용한 쪽(거짓 fail·원문 유지): 1-6,13")
    a = ap.parse_args()
    pptx = Path(a.pptx)
    sp = state_path(pptx, a.state)
    if a.cmd == "reset":
        st = load_state(sp)
        for h in page_hashes(pptx):
            st["pages"].pop(h, None)
        save_state(sp, st)
        print(f"기록 지움: {pptx.name} (다음 회차 전 장)")
        return
    if a.cmd == "mark":
        flags = {}
        for f in a.flag:
            n, _, note = f.partition("=")
            flags[int(n)] = note.strip() or "걸림"
        mark(pptx, _pages(a.ok), flags, a.state, _pages(a.accept))
        print(f"기록: ok {a.ok or '-'} · flag {', '.join(map(str, flags)) or '-'} · fail 수용 {a.accept or '-'}")
        return
    todo, skip, hs = queue(pptx, a.state)
    st = load_state(sp)["pages"]
    if a.cmd == "show":
        for n, h in enumerate(hs, 1):
            r = st.get(h, {})
            acc = " 수용" if r.get("accepted") else ""
            print(f"p{n:<3} {h}  gate {r.get('gate', '-'):4}{acc}  visual {r.get('visual', '-'):4}  {r.get('note', '')}")
        return
    total = len(hs)
    print(f"다음 회차 {len(todo)}/{total}쪽: {', '.join(map(str, todo)) or '없음'}")
    if skip:
        print(f"건너뜀 {len(skip)}쪽(바뀌지 않았고 지난 회차 통과): {', '.join(map(str, skip))}")
    for n in todo:
        r = st.get(hs[n - 1])
        why = "새 장·고친 장" if r is None else ("게이트 fail" if r.get("gate") == "fail" and not r.get("accepted") else
                                              f"시각 {r.get('visual', 'todo')}" + (f" ({r['note']})" if r.get("note") else ""))
        png = f"  {Path(a.png) / f'slide-{n:02d}.png'}" if a.png else ""
        print(f"  p{n:<3} {why}{png}")
    sys.exit(0)


if __name__ == "__main__":
    main()
