#!/usr/bin/env python
"""덱 한 번에 만들기: build.py 실행 → 미리보기 렌더(PDF·PNG·컨택트시트) → 게이트 검사.

사용: make.py <덱 폴더> [--no-render] [--no-check] [--full]
  * 덱 폴더의 build.py 는 Deck.save() 로 pptx 경로를 출력해야 한다(마지막 'saved <경로>' 줄)
  * 게이트와 시각 점검 목록은 증분이 기본: 처음 보는 장·고친 장·지난 회차에 걸린 장만(scripts/review_state.py)
  * 산출: <덱 폴더>/render/<이름>.pdf, render/png/slide-NN.png, render/contact.png, <pptx>.check.json
          + pptx 옆 <이름>_미리보기.pdf (전달용 사본)
종료 코드: 빌드 실패 2, 게이트 fail 1, 통과 0
"""
import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
# Windows: 한글·기호 출력이 cp949 에서 깨지지 않게 UTF-8 모드로 다시 실행(자식 build.py·check.py 도 물려받음)
if os.name == "nt" and not sys.flags.utf8_mode:
    os.environ["PYTHONUTF8"] = "1"
    sys.exit(subprocess.call([sys.executable, *sys.argv]))
PY = sys.executable
sys.path.insert(0, str(HERE.parent / "lib"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck_dir")
    ap.add_argument("--no-render", action="store_true")
    ap.add_argument("--no-check", action="store_true")
    ap.add_argument("--full", action="store_true", help="게이트 결과를 전 장 출력(기본은 증분: 바뀐 장·걸린 장만)")
    a = ap.parse_args()
    d = Path(a.deck_dir).resolve()
    r = subprocess.run([PY, "build.py"], cwd=d, capture_output=True, text=True, encoding="utf-8", errors="replace")
    sys.stdout.write(r.stdout)
    if r.returncode != 0:
        sys.stderr.write(r.stderr)
        print("\n[빌드 실패] 위 오류(LayoutError/WrapError)는 문안을 줄이거나 장을 나눠서 해결한다.")
        sys.exit(2)
    m = re.findall(r"^saved (.+\.pptx) \(", r.stdout, re.M)
    if not m:
        print("build.py 출력에서 'saved <pptx>' 를 찾지 못함")
        sys.exit(2)
    pptx = Path(m[-1])
    pdf = None
    if not a.no_render:
        from preview import render
        pdf = render(pptx, d / "render")
        print("render:", pdf)
        print("contact:", d / "render" / "contact.png")
        if pptx.parent.resolve() != (d / "render").resolve():
            dst = pptx.with_name(pptx.stem + "_미리보기.pdf")
            shutil.copyfile(pdf, dst)
            print("pdf:", dst)
    if a.no_check:
        return
    cmd = [PY, str(HERE / "check.py"), str(pptx)] + (["--pdf", str(pdf)] if pdf else []) + (["--full"] if a.full else [])
    c = subprocess.run(cmd)
    if pdf:
        # 이번 회차에 볼 PNG(처음 보는 장·고친 장·지난 회차에 걸린 장). 통과한 장은 review_state.py mark --ok 로 기록
        subprocess.run([PY, str(HERE / "review_state.py"), "queue", str(pptx), "--png", str(d / "render" / "png")])
    sys.exit(c.returncode)


if __name__ == "__main__":
    main()
