# -*- coding: utf-8 -*-
"""ppt 스킬 환경 점검: 파이썬 패키지, 글꼴, LibreOffice. 문제가 있으면 고치는 방법을 함께 출력한다.

사용: python doctor.py           점검만
      python doctor.py --demo    점검 후 예시 덱(examples/demo)을 임시 폴더에 복사해 끝까지 빌드
종료 코드: 필수 항목 실패 1, 통과 0
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import importlib
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
if os.name == "nt" and not sys.flags.utf8_mode:       # Windows: ✓·✗ 출력이 cp949 에서 깨지지 않게
    os.environ["PYTHONUTF8"] = "1"
    sys.exit(subprocess.call([sys.executable, *sys.argv]))
OK, NO, OPT = "✓", "✗", "△"
fails = 0


def say(mark, what, fix=""):
    global fails
    if mark == NO:
        fails += 1
    print(f" {mark} {what}" + (f"\n     → {fix}" if fix else ""))


def main():
    print(f"ppt 스킬 점검 — 스킬 폴더: {SKILL}")
    print(f" · Python {sys.version.split()[0]} ({sys.executable})")
    if sys.version_info < (3, 9):
        say(NO, "Python 3.9 이상 필요", "python.org 에서 3.10~3.12 설치")

    print("\n[1] 파이썬 패키지")
    pip = f'"{sys.executable}" -m pip install -r "{SKILL / "requirements.txt"}"'
    for mod, name, need in [("pptx", "python-pptx", True), ("lxml", "lxml", True), ("PIL", "Pillow", True),
                            ("pymupdf", "PyMuPDF", True), ("numpy", "numpy", True), ("scipy", "scipy", True),
                            ("matplotlib", "matplotlib", False)]:
        try:
            m = importlib.import_module(mod)
            say(OK, f"{name} {getattr(m, '__version__', getattr(m, 'VersionBind', ''))}")
        except Exception:
            say(NO if need else OPT, f"{name} 없음" + ("" if need else " (선택: matplotlib 그래프)"), pip)
    try:
        from PIL import features
        if features.check("raqm"):
            say(OK, "Pillow 텍스트 엔진 raqm (커닝 측정)")
        else:
            say(OPT, "Pillow raqm 없음 — 커닝 끔 폭 측정을 건너뛴다(결과에는 거의 영향 없음)")
    except Exception:
        pass

    print("\n[2] 글꼴")
    sys.path.insert(0, str(SKILL / "lib"))
    try:
        import kims_ppt as K
        for w in (3, 4, 6, 7):
            say(OK, f"Paperlogy {w} 측정용: {K.ttf(w)}")
        say(OK if K.MALGUN else OPT, f"맑은 고딕(Windows 대체 폭 보장): {K.MALGUN or '없음 — Paperlogy×1.22 로 어림'}")
        say(OK, f"Arial(그리스 문자·기호): {K.ARIAL}")
    except Exception as e:  # noqa: BLE001
        say(NO, f"엔진 불러오기 실패: {e}", pip)
    installed = _paperlogy_installed()
    if installed:
        say(OK, f"Paperlogy 가 OS 에 설치됨: {installed}")
    else:
        say(NO, "Paperlogy 가 OS 에 설치되지 않음 (PowerPoint·LibreOffice 표시용)",
            f"{SKILL / 'fonts'} 의 .ttf 9개를 모두 선택 → 더블클릭(Mac 서체 관리자 '설치' / Windows '설치')")

    print("\n[3] LibreOffice (미리보기 PDF·PNG, 렌더 검사)")
    try:
        from preview import _find_soffice
        so = _find_soffice()
        say(OK, f"soffice: {so}")
    except Exception as e:  # noqa: BLE001
        say(NO, f"{e}", "https://www.libreoffice.org 에서 설치(Mac: brew install --cask libreoffice). "
                        "없으면 빌드·XML 검사만 되고 미리보기·렌더 검사는 못 한다")

    print("\n[4] 로고·첫 사용")
    try:
        import logo_slots as LG
        png, why = LG.resolve()
        say(OK, f"로고: {png or '넣지 않음'} ({why})")
        if not LG.read_config().get("welcomed"):
            print("     · 첫 사용 안내 아직 — SKILL.md 0절대로 두 종류를 소개하고 고르게 한다")
    except Exception as e:  # noqa: BLE001
        say(OPT, f"로고 설정 읽기 실패: {e}")

    print("\n결과:", "필수 항목 통과" if fails == 0 else f"{fails}개 항목을 고쳐야 함")
    if "--demo" in sys.argv and fails == 0:
        demo()
    sys.exit(1 if fails else 0)


def _paperlogy_installed():
    home = Path.home()
    if sys.platform == "darwin":
        dirs = [home / "Library/Fonts", Path("/Library/Fonts")]
    elif os.name == "nt":
        dirs = [Path(os.environ.get("LOCALAPPDATA", str(home / "AppData/Local"))) / "Microsoft/Windows/Fonts",
                Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
    else:
        dirs = [home / ".local/share/fonts", home / ".fonts", Path("/usr/share/fonts"), Path("/usr/local/share/fonts")]
    for d in dirs:
        if d.is_dir():
            hits = list(d.rglob("Paperlogy-4Regular.ttf")) or [p for p in d.glob("*") if p.name.lower().startswith("paperlogy")]
            if hits:
                return d
    return None


def demo():
    tmp = Path(tempfile.mkdtemp()) / "demo"
    shutil.copytree(SKILL / "examples" / "demo", tmp, ignore=shutil.ignore_patterns("output", "render"))
    print(f"\n[예시 덱 빌드] {tmp}")
    env = dict(os.environ, KIMS_PPT_SKILL=str(SKILL))
    r = subprocess.run([sys.executable, str(SKILL / "scripts" / "make.py"), str(tmp), "--full"], env=env)
    print("\n예시 덱:", tmp / "output", "(종료 코드", r.returncode, ")")


if __name__ == "__main__":
    main()
