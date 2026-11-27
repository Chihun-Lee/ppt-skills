# -*- coding: utf-8 -*-
"""수정 가능한 형태(네이티브) 요소 도우미 — python-pptx 네이티브 차트용 (STANDARD §9.1 IM-05).

  chart_fonts(chart, 14)   차트 글꼴 일괄 지정. 기본 Arial(FT-05: 그래프 글씨는 영어 + Arial)
  positive_ax_ids(prs)     저장 직전 1회. python-pptx 차트 템플릿의 음수 축 ID를 차트마다 작은 양수로 다시 매김

"""
from __future__ import annotations

from pptx.oxml.ns import qn
from pptx.util import Pt


def chart_fonts(chart, size_pt: float = 14, family: str = "Arial"):
    """축·범례·데이터 레이블 글꼴을 한 번에. latin·ea·cs를 같은 이름으로 둬 동아시아 대체 글꼴로 새지 않게 한다."""
    chart.font.size = Pt(size_pt)
    chart.font.name = family
    txPr = chart._chartSpace.find(qn("c:txPr"))
    if txPr is not None:
        for rpr in txPr.iter(qn("a:defRPr")):
            for tag in ("a:ea", "a:cs"):
                el = rpr.find(qn(tag))
                if el is None:
                    el = rpr.makeelement(qn(tag), {})
                    rpr.append(el)
                el.set("typeface", family)


def positive_ax_ids(prs) -> int:
    """음수 c:axId·c:crossAx(스키마 xs:unsignedInt 위반, OOXML 검증기 오류)를 차트마다 100000001부터 다시 매긴다.
    짝(축 ID ↔ crossAx)은 같은 값끼리 유지된다. v + 2**32 처럼 2**31 을 넘기면 LibreOffice가 32비트로 읽어
    축 짝이 깨지고 축 제목이 뒤바뀐다(2026-10-27 확인). 바뀐 속성 수를 돌려준다."""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    n, seen = 0, set()
    for sl in prs.slides:
        for rel in sl.part.rels.values():
            if rel.reltype != RT.CHART or id(rel.target_part) in seen:
                continue
            seen.add(id(rel.target_part))
            els = list(rel.target_part._element.iter(qn("c:axId"), qn("c:crossAx")))
            if not any(int(el.get("val")) < 0 for el in els):
                continue
            new = {}
            for el in els:
                v = el.get("val")
                new.setdefault(v, str(100000001 + len(new)))
                el.set("val", new[v])
                n += 1
    return n
