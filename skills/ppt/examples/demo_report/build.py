# -*- coding: utf-8 -*-
"""ppt 스킬 예시 덱(보고용, REPORT.md) — 그림·수치는 모두 합성 예시이며 실제 실험 결과가 아니다.

빌드: python <스킬>/scripts/make.py <이 폴더>      (그림은 ../demo/assets 를 같이 쓴다)
"""
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = Path(os.environ.get("KIMS_PPT_SKILL", HERE.parents[1]))   # 실제 덱에서는 스킬 폴더 절대경로를 직접 쓴다
sys.path.insert(0, str(SKILL / "lib"))
from kims_ppt import Deck  # noqa: E402

A = HERE.parent / "demo" / "assets"
if not A.exists():                      # doctor --demo 처럼 이 폴더만 복사한 경우
    A = SKILL / "examples" / "demo" / "assets"
CH = ["배경", "결과", "계획"]

d = Deck("report", workdir=HERE)
d.cover("알루미늄 합금 시효 조건 탐색 결과", "ppt 스킬 보고용 예시 2026.11", ["○○기관 소속 연구본부", "홍길동 선임연구원"])
d.toc([("배경", ["탐색 범위"]), ("결과", ["예측 모델 검증", "추천 조건"]), ("계획", ["향후 일정"])],
      front=["결과 요약"], back=["References"])

s = d.slide("결과 요약", kind="summary")
s.bullets([("추천 조건 180 °C, 12 h에서 항복강도 327 MPa", ["기존 200 °C, 4 h 대비 3% 향상", "64조건 중 12조건만 인장 시험"]),
           ("예측 모델 평균 오차 약 10 MPa", ["시편 36개 기준, 순서는 대체로 일치 [1]"])])
s.table(["구분", "기존 조건", "추천 조건"],
        [["시효 조건", "200 °C, 4 h", "180 °C, 12 h"], ["항복강도", "318 MPa", "327 MPa"]])

d.divider(1, "배경", chapters=CH)
s = d.slide("탐색 범위", chapter=1)
s.key("후보 64조건 중 추천 12조건 시험")
s.bullets(["시효 온도 150~220 °C, 시효 시간 1~48 h", "불확실성이 큰 조건부터 추천"])
s.figure(A / "doe_points.png", "후보 64조건과 추천 12조건(합성 예시)")

d.divider(2, "결과", chapters=CH)
s = d.slide("예측 모델 검증", chapter=2)
s.key("시편 36개, 평균 오차 약 10 MPa")
s.bullets(["예측보다 실측이 평균 3% 높음, 순서는 대체로 일치 [1]"])
s.figure(A / "parity.png", "예측과 실측 항복강도(합성 예시)")

s = d.slide("추천 조건", chapter=2)
s.key("180 °C, 12 h에서 항복강도 327 MPa")
s.text_figure([("결정립계 석출물 관찰", ["평균 결정립 크기 약 60 µm", "조건별 차이는 추가 분석"])],
              A / "micrograph.jpg", "결정립 조직(합성 예시 이미지)")

d.divider(3, "계획", chapters=CH)
s = d.slide("향후 일정", chapter=3, kind="conclusion")
s.note("추천 조건 재현 시험 후 피로 시험 착수")
s.timeline([("12월", "추천 조건 재현 시험"), ("1월", "피로 시험"), ("2월", "부품 시제품 적용")], orient="h")
s.table(["산출물", "시점"], [["재현 시험 결과", "12월"], ["피로 수명 자료", "1월"], ["시제품 평가표", "2월"]])

d.references(["Example A. et al., Example Journal 1, 1 (2025)"])
d.save(HERE / "output" / "ppt_보고용예시.pptx")
