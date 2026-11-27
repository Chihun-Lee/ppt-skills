# -*- coding: utf-8 -*-
"""ppt 스킬 예시 덱(발표용) — 그림·수치는 모두 합성 예시이며 실제 실험 결과가 아니다.

빌드: python <스킬>/scripts/make.py <이 폴더>
모듈 예시: 표지·목차·간지·요지+세부+그림·단계+그림·글+그림·좌우 분할(네이티브 차트 + 표)·사진·결론 박스+비교+일정·감사
"""
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = Path(os.environ.get("KIMS_PPT_SKILL", HERE.parents[1]))   # 실제 덱에서는 스킬 폴더 절대경로를 직접 쓴다
sys.path.insert(0, str(SKILL / "lib"))
from pptx.chart.data import XyChartData  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION  # noqa: E402
from pptx.util import Pt  # noqa: E402

from kims_ppt import E, Deck  # noqa: E402
from native import chart_fonts  # noqa: E402

A = HERE / "assets"


def native_chart(s, rect, series):
    """네이티브 XY 꺾은선 차트(PowerPoint 에서 데이터 편집 가능). rect 칸을 통째로 쓴다."""
    cd = XyChartData()
    for name, pts in series:
        ser = cd.add_series(name)
        for x, y in pts:
            ser.add_data_point(x, y)
    gf = s.s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES, E(rect.x), E(rect.y), E(rect.w), E(rect.h), cd)
    ch = gf.chart
    chart_fonts(ch, 14)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    for ax, title in ((ch.category_axis, "Aging time (h)"), (ch.value_axis, "YS (MPa)")):
        ax.has_title = True
        ax.axis_title.text_frame.text = title
        ax.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
        ax.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
        ax.has_major_gridlines = False
    ch.value_axis.minimum_scale, ch.value_axis.maximum_scale = 200, 360
    ch.category_axis.minimum_scale, ch.category_axis.maximum_scale = 0, 50
    ch.category_axis.major_unit = 10
    return gf


d = Deck("ppt", workdir=HERE)
d.cover("알루미늄 합금 시효 조건 탐색", "ppt 스킬 예시 덱 2026.11",
        ["○○기관 소속 연구본부", "홍길동 선임연구원"])
d.toc([("배경", ["시효 조건 현황"]),
       ("방법", ["실험 설계", "예측 모델 검증"]),
       ("결과", ["조건별 결과", "결론 및 향후 일정"])])

d.divider(1, "배경")
s = d.slide("시효 조건 현황", chapter=1)
s.key("시효 온도·시간별 경도 분포")
s.bullets(["8온도 × 8시간 64조건 중 최고 경도 139 HV", "고온일수록 최고 경도 시점이 앞당겨짐"])
s.figure(A / "hardness_map.png", "시효 온도·시간별 경도 지도(합성 예시)")

d.divider(2, "방법")
s = d.slide("실험 설계", chapter=2)
s.key("조건 탐색 절차")
s.flow([("조건 후보", "온도 8, 시간 8"), ("1차 시험", "경도 64조건"), ("모델 학습", "GP 회귀"),
        ("조건 추천", "불확실성 기준")])
s.figure(A / "doe_points.png", "후보 64조건과 추천 12조건(합성 예시)")

s = d.slide("예측 모델 검증", chapter=2)
s.key("예측과 실측 항복강도 비교")
s.text_figure(["시편 36개, 평균 오차 약 10 MPa", "예측보다 실측이 평균 3% 높음", "순서는 대체로 일치"],
              A / "parity.png", "예측과 실측 항복강도(합성 예시)")

d.divider(3, "결과")
s = d.slide("조건별 결과", chapter=3)
s.key("온도별 항복강도 변화")
with s.row([3, 2]) as (L, R):
    native_chart(s, L, [("160 °C", [(1, 228), (4, 262), (12, 301), (24, 334), (48, 341)]),
                        ("180 °C", [(1, 241), (4, 289), (12, 327), (24, 318), (48, 296)]),
                        ("200 °C", [(1, 266), (4, 318), (12, 305), (24, 281), (48, 259)])])
    s.table(["온도", "최고 YS", "시간"],
            [["160 °C", "341 MPa", "48 h"], ["180 °C", "327 MPa", "12 h"], ["200 °C", "318 MPa", "4 h"]],
            rect=R, fill=True)

s = d.slide("조건별 결과", chapter=3)
s.key("시편 단면 관찰")
s.photo(A / "micrograph.jpg", "결정립 조직(합성 예시 이미지)",
        ["평균 결정립 크기 약 60 µm", "결정립계 석출물 관찰", "조건별 차이는 추가 분석"])

s = d.slide("결론 및 향후 일정", chapter=3, kind="conclusion")
s.note("180 °C, 12 h 조건에서 항복강도 3% 향상")
s.compare(("기존 조건", ["200 °C, 4 h", "항복강도 318 MPa"]),
          ("추천 조건", ["180 °C, 12 h", "항복강도 327 MPa"]))
with s.row([1, 1]) as (L, R):
    s.timeline([("12월", "추천 조건 재현 시험"), ("1월", "피로 시험"), ("2월", "부품 시제품 적용")], rect=L)
    s.table(["산출물", "시점"], [["재현 시험 결과", "12월"], ["피로 수명 자료", "1월"], ["시제품 평가표", "2월"]],
            rect=R, fill=True)

d.closing()
d.save(HERE / "output" / "ppt_예시덱.pptx")
