---
name: ppt_experiment
description: 실험결과 정리용 PPT(발표용 아님)를 python-pptx로 만든다. 실험 조건·실험 환경·실험 결과·(필요하면) 최종 결론만 표·그래프·사진·개조식 key:value로 나열하고 문장형 서술은 쓰지 않는다. 사용자가 "실험결과 PPT로 정리해", "실험보고 PPT", "분석결과 PPT 만들어줘", "측정 데이터 PPT로 정리", "/ppt_experiment"라고 하면 사용. 네이비 #002060 + 연블루 표, 맑은 고딕, 표지·로고·장식 없음. 발표·학회·보고용 다목적 PPT는 형제 스킬 ppt.
---

# ppt_experiment — 실험결과 정리 PPT (자료 공유용)

실험·분석 결과를 **표 + 그래프 + 사진 + 개조식 최소 텍스트**로만 정리한다. 발표용 스토리텔링(형제 스킬 `ppt`)과 다르다 — **꼭 필요한 전달 내용만, 데이터가 주인공**.

- **`<스킬>`** = 이 SKILL.md가 있는 폴더. 명령과 build.py에는 실제 절대경로를 넣는다.
- 라이브러리 `<스킬>/lib/lib_experiment.py`(덱마다 복사하지 않는다), 한 번에 빌드·렌더·검사 `<스킬>/scripts/make.py`, 검사기 `<스킬>/scripts/check.py`, 예시 `<스킬>/examples/demo/build.py`(합성 데이터).

## 0. 처음 한 번

```bash
python "<스킬>/scripts/welcome.py" status        # Windows 는 python 대신 py 일 수 있다
```
`첫 사용 안내: 아직`이면 만들기 전에 사용자에게 **한 번** 묻는다(요청이 실험 정리로 보여도 처음 한 번은 묻는다).

- 이 묶음에는 두 가지가 있다.
  - **① 다목적 PPT** (형제 스킬 `ppt`, 보통 `<스킬>/../ppt/SKILL.md`): 발표·학회·세미나·보고자료. 표지·목차·간지·그림 중심 본문, 로고 자리
  - **② 실험결과 정리 PPT** (이 스킬): 조건·환경·결과를 표·그래프·사진으로 나열. 표지·로고·장식 없음
- ①을 고르면 형제 스킬 SKILL.md를 읽고 그 0.1절(로고 질문 포함)로 넘어간다.
- ②를 고르면 `python "<스킬>/scripts/welcome.py" welcomed`로 기록하고 아래 절차로 간다. 로고는 묻지 않는다(이 양식은 원래 로고를 넣지 않는다. 사용자가 원하면 `add_image_fit`으로 제목띠 오른쪽 끝에 작게 넣는다).

환경 점검: 형제 스킬의 `python "<스킬>/../ppt/scripts/doctor.py"`가 파이썬 패키지·LibreOffice를 함께 본다. LibreOffice가 없으면 빌드·검사만 되고 미리보기 PNG 점검을 못 하므로, 사용자에게 PowerPoint로 열어 봐 달라고 하고 그 사실을 보고에 적는다.

## 1. 절대 규칙 — 글은 최소로

- **문장형 서술 금지**: 종결어미(~다/~임/~함/~됨) 금지. 명사구·`key : value`·수치만.
  - ❌ "SEM 분석은 20kV에서 수행하였다" → ⭕ `SEM Conditions : 20 kV, 20 nA`
  - ❌ "인덱싱률이 94% 이상으로 양호하였다" → ⭕ `Indexing : 94.3–99.4 % (pass)`
- **불릿 1줄 ≤ 24자**(한글 기준, 영문 ≤ 40자). 넘으면 쪼개지 말고 **지운다**. 두 줄이 되는 불릿은 실패.
- **한 블록 불릿 ≤ 4개**, 슬라이드당 불릿 합계 ≤ 10개. 남는 건 표로 옮긴다.
- **조사·접속어 생략**: "~에 대한", "~를 통한", "그리고", "따라서" 삭제.
  ❌ `전해연마를 통한 EBSD 전처리 품질 확보` → ⭕ `EP 전처리 → EBSD 합격`
- **요약 말투**: 약어·기호 적극 사용(`→` 인과·순서, `/` 병렬, `·` 나열, `≤ ≥ ±`). 숫자·단위·고유명사가 대부분이어야 정상.
- `➢` 화살표 불릿은 **주의·조치 필요·결론**에만, 장당 ≤ 3줄.
- 표지·목차·장식·로고·푸터·페이지 번호 **없음**. 첫 장부터 바로 내용.
  - 예외 — 월간 연구노트처럼 첫 장에 과제·월이 보여야 하는 문서는 `note_cover()` 표지 1장을 둔다(과제정보 표는 넣지 않는다: `account/author/pi/period` 생략).
- 수치는 측정 원본 그대로(반올림·연출 금지). 이상치·주의값만 RED 강조.

## 2. 절대 규칙 — 글자 크기 하한

**어떤 요소도 13pt 미만으로 찍지 않는다.** `MIN_FONT_PT = 13`을 `_set_font`가 강제한다.

| 요소 | 크기 |
|---|---|
| 제목띠 `title_band` | **22** |
| 번호 섹션 `section_numbered` | **16** |
| ❖ 섹션 `section_diamond` | **15** |
| 불릿 `kv_bullets` · `arrow_notes` | **14** |
| 표 `exp_table` (행 높이 30 / 머리 32) | **13** |
| 그림 라벨 `image_row` | **13** |

- 넘치면 글자를 줄이지 말고 **문구를 지우거나 표로 옮긴다**.
- 표 열이 많아 13pt가 안 들어가면 **열을 합치거나 행을 나눠 장을 하나 더 쓴다**.

## 3. 슬라이드 구조 (3종만)

1. **개요 1장** (`title_band(..., style="navy")`): 실험 조건·환경
   - 번호 섹션 `section_numbered`: `1. Specimens` / `2. Conditions` / `3. Output Files` 식
   - `kv_bullets`로 장비·조건, 오른쪽에 시편 모식도·장비 사진(`add_image_fit`), 아래 `arrow_notes`로 주의점
2. **결과 N장 — 같은 레이아웃 반복** (`style="light"`, 예: `Sample #1. A1 (160 °C, 8 h)`)
   - `section_diamond`(❖): 정보표(navy) → 데이터표(blue, Average 행 강조) → 그림 줄(`image_row`)
   - 시편·조건마다 **좌표까지 같은 배치**(덱의 핵심 패턴)
3. **최종 결론 1장 — 필요할 때만**: 전 시편 비교표 1개 + `arrow_notes` 결론 2~4줄

분량 = 개요 1 + 시편·조건 수 N + (결론 0~1). 그 이상 만들지 않는다.

## 4. 작업 위치와 절차

```
<덱 이름>/                사용자 작업 폴더 아래(스킬 폴더 안에 만들지 않는다)
├── build.py              덱 정의
├── images/               쓸 그림 사본(원본은 건드리지 않는다)
├── output/               <이름>.pptx + <이름>_미리보기.pdf   ← 전달물
└── render/               미리보기 PNG·contact.png(자동)
```

1. **자료 읽기**: 사용자 사진(현미경·맵·그래프 캡처)·CSV·엑셀·수치를 먼저 읽고 `images/`에 복사한다. 자료에 없는 수치를 지어내지 않는다.
   - ⚠️ 폰카 사진은 MPO 형식일 수 있다 — `Image.open(f).format` 확인 후 JPEG/PNG로 다시 저장. 스크린샷 위아래 검은 띠는 잘라 낸다.
   - **그림은 있으면 쓴다** — 표만 있는 장보다 실측 그림이 한 장이라도 붙은 쪽이 낫다. 그림이 없는 수치는 matplotlib으로 그린다(흰 배경, 장식 최소, 축·범례는 영어).
2. **build.py 작성** — 예시 `<스킬>/examples/demo/build.py`를 본보기로:
   ```python
   import sys; sys.path.insert(0, r"<스킬>/lib")        # ← 실제 절대경로
   from pathlib import Path
   from lib_experiment import *
   HERE = Path(__file__).resolve().parent
   prs = new_prs()
   s = blank_slide(prs); title_band(s, "OOO 시편 EBSD 분석 결과", style="navy")
   ...
   (HERE / "output").mkdir(exist_ok=True); prs.save(HERE / "output" / "실험결과.pptx")
   ```
3. **빌드·렌더·검사**: `python "<스킬>/scripts/make.py" <덱 폴더>`
   - `render/contact.png`로 흐름을 보고 `render/png/slide-NN.png`를 **전부 직접 연다**. 고치고 다시(최대 3회).
   - 검사 fail 0이 될 때까지 고친다. warn(E2 종결 의심·E3 길이·E4 개수)은 읽고 판단한다.
4. **전달**: `output/<이름>_미리보기.pdf`를 사용자에게 열어 보여 준다(Mac `open`, Windows `start ""`). 보고에는 pptx·PDF 경로, 장 수, **표 수치를 원본과 대조했다는 사실**, 남은 주의점을 쓴다.
- **이미 건넨 pptx는 같은 이름으로 다시 저장하지 않는다** — `_v2`, `_v3`. 같은 폴더에 `~$<파일명>`(PowerPoint 잠금 파일)이 있으면 열려 있는 것이니 멈추고 알린다.

## 5. 도우미 API (`lib_experiment.py`)

캔버스 1280×720 px 격자(16:9), 흰 바탕, `MARGIN = 40`.

| 함수 | 용도 |
|---|---|
| `new_prs()` / `blank_slide(prs)` | 1280×720 프레젠테이션 / 빈 슬라이드 |
| `title_band(s, text, style)` | 위 제목띠. `"navy"`=개요·결론, `"light"`=결과 반복 장 |
| `section_numbered(s, x, y, "1. Xxx")` | 회색 강조 섹션(개요용, 폭 자동) |
| `section_diamond(s, x, y, "Xxx")` | ❖ 네이비 섹션(결과용) |
| `kv_bullets(s, x, y, w, items)` | `• key : value` 목록 |
| `arrow_notes(s, x, y, w, items)` | `➢` 주의·결론 전용 |
| `exp_table(s, x, y, w, headers, rows, style, col_ratios, highlight_last)` | 표. `"navy"`=정보표, `"blue"`=데이터표, `highlight_last`=Average 행 강조. 전 셀 0.5pt 검정 테두리·줄무늬 제거 내장 |
| `add_image_fit(s, path, x, y, w, h, border)` | 비율 유지 그림(상자 가운데) |
| `image_row(s, x, y, w, h, [(path, label), …])` | 라벨 달린 그림 줄 |
| `textbox(s, …)` | 낮은 수준 글상자(`(text, {size, bold, color})` 튜플로 혼합) |
| `note_cover(s, month=, project=, hero=, topics=)` | 월간 연구노트 표지(1절 예외) |

디자인 값: NAVY `#002060`(제목띠·❖·표 머리) · BLUE_LIGHT `#EAF2FA`(데이터표 머리) · BLUE_MID `#C4DCF2`(강조 행) · GRAY_BAND `#E7E6E6` · BAND_LIGHT `#F2F2F2` · RED `#C00000`(주의 전용). 도형은 모두 평면(`_flat()`이 그림자 원인 제거).

## 6. 통과 기준

- 문장형 글 0건(검사 E2), 13pt 미만 0건(E1), 두 줄 넘는 불릿 0건(눈 점검), 불릿 1줄 ≤ 24자(E3)
- 표 수치 = 사용자 원본, 지어낸 데이터 없음(예시 덱 제외)
- 결과 장 레이아웃이 좌표까지 같음
- 그림 비율 왜곡 없음(E5), 글 넘침 없음, 표 줄무늬 없음, 그림자 없음(E7)

## 7. 함정

- **표 줄무늬**: python-pptx 기본 표 스타일의 파란 줄무늬 — `exp_table`이 지운다. `add_table`을 직접 쓰지 않는다.
- **셀 테두리**: python-pptx가 지원하지 않아 `_cell_border`가 XML로 넣는다(a:lnL~lnB가 solidFill **앞**).
- **LibreOffice 그림자**: `shadow.inherit=False`만으로는 부족 — `_flat()`으로 p:style 제거.
- **글꼴**: pptx에는 맑은 고딕으로 적는다(Windows 기본). Mac·Linux 미리보기는 LibreOffice가 대체 글꼴로 그려 폭이 조금 다를 수 있으니, 글자 수에 여유를 두고 PowerPoint에서 한 번 확인한다.
- **그림 늘이기 금지**: 반드시 `add_image_fit`(비율 유지).
