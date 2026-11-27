# ppt — 다목적 발표·학회·보고 PPT 스킬

이 폴더는 [ppt-skills](https://github.com/Chihun-Lee/ppt-skills) 레포의 스킬 하나다. 설치·첫 사용·로고 설정은 레포의 `README.md`와 `INSTALL.md`를 따른다.

- Claude가 읽는 절차: `SKILL.md`(발표용) · `REPORT.md`(보고용) · `STANDARD.md`(모든 수치의 기준서)
- 점검: `python scripts/doctor.py` (`--demo`면 예시 덱까지 빌드)
- 로고: `python scripts/logo.py status | set <파일> | text "이름" | placeholder | none | preview`
- 예시: `examples/demo`(발표용 12장) · `examples/demo_report`(보고용 11장). 그림·수치는 합성 데이터
- 글꼴: `fonts/` Paperlogy 9종(SIL OFL 1.1, `fonts/OFL.txt`)
