# 설치 절차 — Claude Code가 실행하는 순서

사용자가 이 레포 링크(`https://github.com/Chihun-Lee/ppt-skills`)를 주며 「이 스킬 쓰게 해줘」, 「설치해줘」, 「이거 참고해서 PPT 만들어줘」라고 하면 Claude Code가 아래 명령을 순서대로 실행해 설치한다. 사용자가 명령을 직접 칠 필요는 없다.

- 명령은 Windows의 Git Bash·PowerShell과 macOS·Linux 셸에서 그대로 돌아가도록 `$HOME`으로 썼다.
- 단계마다 출력을 확인한다. 실패하면 맨 아래 「문제 해결」을 본다.
- 사용자가 PPT 내용까지 같이 요청했으면(예: 「이걸로 학회 발표 PPT 만들어줘」) 설치를 마친 뒤 Claude Code를 다시 시작하지 않고 **같은 대화에서** 설치된 `SKILL.md`를 직접 읽어 이어서 만든다(5단계).

## 설치가 바꾸는 것

| 위치 | 내용 |
|---|---|
| `~/.claude/skills/ppt`, `~/.claude/skills/ppt_experiment` | 스킬 두 개. Windows는 복사, macOS·Linux는 연결(심링크) |
| `~/.ppt_skills/` | 받은 원본(`src/`), 파이썬 가상환경(`venv/`), 사용자 설정·로고(`config.json`, `logo.png`), 설치 기록(`install.json`) |
| 사용자 글꼴 폴더 | Paperlogy 9종(SIL OFL 1.1). macOS `~/Library/Fonts`, Windows `%LOCALAPPDATA%\Microsoft\Windows\Fonts`(+ HKCU 글꼴 등록), Linux `~/.local/share/fonts` |

- 전역 파이썬, 시스템 설정, 관리자 권한은 쓰지 않는다.
- 같은 이름의 스킬(예: 사용자가 직접 만든 `ppt`)이 이미 있으면 지우지 않고 `~/.ppt_skills/backup/`으로 옮긴다. **이미 `ppt`라는 스킬이 있으면 설치 전에 사용자에게 알리고 동의를 받는다.**

## 1. Python 3.9 이상 확인

| OS | 명령 |
|---|---|
| Windows | `py -3 --version` (안 되면 `python --version`) |
| macOS·Linux | `python3 --version` |

- 「Python 3.x」가 나오면 2단계로 간다. 아래 명령의 `py -3`(Windows)·`python3`(macOS·Linux)은 여기서 된 명령으로 바꿔 쓴다.
- Windows에서 「Python was not found … Microsoft Store」가 나오면 설치되지 않은 것이다.
- Python이 없으면 **사용자에게 설치해도 되는지 묻고** 설치한다.
  - Windows(관리자 권한 불필요): `winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements`
    지금 세션에는 새 PATH가 반영되지 않으므로 이후 `py -3` 자리에 `"$HOME/AppData/Local/Programs/Python/Python312/python.exe"`를 쓴다.
  - winget이 없으면 https://www.python.org/downloads/ 설치를 안내한다(첫 화면 「Add python.exe to PATH」 체크, 설치 뒤 Claude Code 재시작).
  - macOS: `brew install python` 또는 python.org.

## 2. 받기

```bash
git clone --depth 1 https://github.com/Chihun-Lee/ppt-skills.git "$HOME/.ppt_skills/src"
```

- 이미 `~/.ppt_skills/src`가 있으면 갱신한다: `git -C "$HOME/.ppt_skills/src" pull --ff-only`
- git이 없으면 zip으로 받는다.
  - Windows: `curl.exe -fL -o "$HOME/ppt-skills.zip" https://github.com/Chihun-Lee/ppt-skills/archive/refs/heads/main.zip` 뒤 `py -3 -m zipfile -e "$HOME/ppt-skills.zip" "$HOME/.ppt_skills"`
  - macOS·Linux: `curl -fL -o "$HOME/ppt-skills.zip" https://github.com/Chihun-Lee/ppt-skills/archive/refs/heads/main.zip` 뒤 `python3 -m zipfile -e "$HOME/ppt-skills.zip" "$HOME/.ppt_skills"`
  - 원본 폴더는 `$HOME/.ppt_skills/ppt-skills-main`이 된다. 3단계의 `src`를 이 경로로 바꾼다.

## 3. 설치

```bash
py -3 "$HOME/.ppt_skills/src/install.py"          # Windows
python3 "$HOME/.ppt_skills/src/install.py"        # macOS·Linux
```

- 파이썬 패키지를 내려받으므로 2~5분 걸린다. **Bash 도구 timeout을 600000(10분)으로 준다.**
- 하는 일: ① 스킬 두 개를 넣는다 ② `~/.ppt_skills/venv`에 python-pptx·lxml·Pillow·PyMuPDF·numpy·scipy·matplotlib 설치 ③ Paperlogy 글꼴 설치 ④ 점검(doctor)과 예시 덱 2종 빌드·검사
- 성공: 출력 끝에 `✓ 설치 완료`가 나온다. 점검의 ✗가 LibreOffice뿐이면 설치는 된 것이다(4단계).
- pip가 SSL·프록시 오류로 실패하면(회사망):
  ```bash
  py -3 "$HOME/.ppt_skills/src/install.py" --pip-args "--trusted-host pypi.org --trusted-host files.pythonhosted.org"
  ```

## 4. LibreOffice (미리보기 PDF·렌더 검사용, 권장)

3단계 점검에 「LibreOffice(soffice)를 찾지 못함」이 있으면 **사용자에게 설치해도 되는지 묻는다.**
- Windows: `winget install -e --id TheDocumentFoundation.LibreOffice --accept-package-agreements --accept-source-agreements`(설치 창에서 관리자 승인을 물을 수 있다)
- macOS: `brew install --cask libreoffice`
- 원하지 않으면 건너뛴다. pptx는 만들어지고, 미리보기 PNG 점검 대신 사용자가 PowerPoint로 열어 확인한다.
- 설치했으면 점검을 다시 본다: `py -3 "$HOME/.claude/skills/ppt/scripts/doctor.py"`(macOS·Linux는 `python3`)

## 5. 첫 사용 안내 — 반드시

설치가 끝나면 사용자에게 아래를 **한 번에** 알리고 묻는다(사용자가 이미 내용을 요청했어도 처음 한 번은 묻는다).

> 설치가 끝났습니다. PPT 스킬이 두 가지 있습니다. 어느 쪽으로 만들까요?
> **① 다목적 PPT** — 발표·학회·세미나·내부 회의 발표자료, 발표 없이 읽히는 보고자료. 표지·목차·간지·그림 중심 본문·감사 장이 있고 로고가 들어갑니다.
> **② 실험결과 정리 PPT** — 실험 조건·환경·결과를 표·그래프·사진으로만 나열하는 데이터 공유용. 문장·표지·로고가 없습니다.
> ①이면 로고 파일(PNG·JPG·SVG·PDF)이 있으면 알려 주세요. 파일이 없으면 기관·팀 이름으로 글자 로고를 만들 수 있고, 둘 다 없으면 'LOGO'라고 적힌 자리표시가 들어갑니다(나중에 바꿀 수 있습니다).

답에 따라 기록한다(Windows는 `python3` 대신 `py -3`).
```bash
python3 "$HOME/.claude/skills/ppt/scripts/logo.py" set "<로고 파일>"     # 파일을 받았을 때
python3 "$HOME/.claude/skills/ppt/scripts/logo.py" text "기관·팀 이름"   # 이름만 받았을 때
python3 "$HOME/.claude/skills/ppt/scripts/logo.py" welcomed              # 안내를 마쳤다고 기록(①·② 모두)
```
- ②를 고르면 로고는 묻지 않는다(그 양식은 로고를 넣지 않는다). `welcomed`만 기록하면 된다 — `logo.py welcomed`와 `ppt_experiment/scripts/welcome.py welcomed`는 같은 설정(`~/.ppt_skills/config.json`)에 같은 표시를 남기므로 어느 쪽을 써도 된다.
- 사용자에게 물을 수 없는 실행(자동화·예약 작업)이면 요청 내용으로 고르고, 로고는 자리표시로 두고, 그렇게 했다고 결과 보고에 적는다.
- 그다음 고른 스킬의 `SKILL.md`(`~/.claude/skills/ppt/SKILL.md` 또는 `~/.claude/skills/ppt_experiment/SKILL.md`)를 읽고 그 절차로 PPT를 만든다. `<스킬>` 자리는 그 SKILL.md가 있는 폴더의 절대경로다.
- 이후 새 대화에서는 Claude Code를 다시 시작하면 `/ppt`, `/ppt_experiment`로 바로 부를 수 있다고 알린다.

## 설치 없이 쓰기 (claude.ai 웹·앱처럼 설치할 수 없을 때)

- 레포를 받을 수 있으면 작업 폴더 아래에 받아(`git clone` 또는 zip) `skills/ppt/SKILL.md`나 `skills/ppt_experiment/SKILL.md`를 읽고 따른다. 파이썬 패키지는 `requirements.txt`로 그 환경에 설치하고, 글꼴은 `skills/ppt/fonts/`가 측정용으로 자동으로 쓰인다.
- LibreOffice가 없는 환경이면 미리보기 렌더·PNG 점검은 빠진다고 사용자에게 알린다.
- 이 경우에도 5단계의 두 종류 안내와 로고 질문은 똑같이 한다(`logo.py`는 `~/.ppt_skills/`에 기록하므로 그 환경 홈 폴더에 쓸 수 있어야 한다. 못 쓰면 build.py에서 `Deck(..., logo=파일)`로 직접 준다).

## 업데이트·삭제

- **업데이트**: 사용자가 「ppt 스킬 업데이트해줘」라고 하면
  ```bash
  git -C "$HOME/.ppt_skills/src" pull --ff-only
  python3 "$HOME/.ppt_skills/src/install.py"        # Windows: py -3
  ```
  로고·설정은 `~/.ppt_skills`에 그대로 남는다. zip으로 받았으면 2단계 zip 명령으로 다시 받아 덮어쓴 뒤 install.py.
- **삭제**: `install.py --uninstall`은 이 설치가 넣은 스킬 두 개만 지운다. 설정·로고·가상환경까지 지우려면 `~/.ppt_skills` 폴더도 지운다(글꼴은 남는다).

## 문제 해결

| 증상 | 조치 |
|---|---|
| `py`·`python` 둘 다 없음, Microsoft Store 메시지 | 1단계 Python 설치 |
| `git` 없음 | 2단계 zip 방법 |
| pip `SSL: CERTIFICATE_VERIFY_FAILED`, 프록시 오류 | 3단계 `--pip-args` |
| `venv를 만들지 못함` (Linux) | `sudo apt install python3-venv`를 사용자에게 안내 |
| 점검 「Paperlogy 가 OS 에 설치되지 않음」 | `install.py`를 `--no-fonts` 없이 다시 실행. 그래도 안 되면 `~/.claude/skills/ppt/fonts/`의 .ttf 9개를 사용자가 직접 설치(더블클릭 → 설치) |
| 게이트 G1 fail 「Paperlogy 외 폰트」 | 글꼴 설치 직후면 LibreOffice를 한 번 껐다 켜고 다시 빌드 |
| `LayoutError`·`WrapError` | 정상 동작. 글이 넘친다는 뜻이라 문안을 줄이거나 장을 나눈다 |
| 자가검사 실패 | 출력 끝 부분을 사용자에게 보여 준다. `doctor.py`로 빠진 것을 확인 |
| 다시 시작해도 `/ppt`가 안 보임 | `~/.claude/skills/ppt/SKILL.md`가 있는지 확인. 없으면 3단계 다시 |
| 원래 쓰던 `ppt` 스킬이 사라짐 | `~/.ppt_skills/backup/`에 있다. 되돌리려면 `install.py --uninstall` 뒤 그 폴더를 `~/.claude/skills/`로 옮긴다 |
