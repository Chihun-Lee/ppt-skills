# -*- coding: utf-8 -*-
"""ppt_experiment 한 번에: build.py 실행 → pptx → 미리보기 PDF·PNG → 검사(check.py).

사용: python make.py <덱 폴더>              build.py 를 돌리고 output/ 의 가장 최근 pptx 를 렌더·검사
      python make.py <덱 폴더> --no-build   빌드 없이 렌더·검사만
      python make.py <덱 폴더> --no-render  빌드·검사만(LibreOffice 없을 때)

결과: <덱>/output/<이름>.pptx, <이름>_미리보기.pdf / <덱>/render/png/slide-NN.png, render/contact.png
종료 코드: 검사 fail 이 있으면 1
"""
from __future__ import annotations

import _venv  # noqa: F401,E402  (패키지가 없으면 ~/.ppt_skills/venv 파이썬으로 다시 실행)
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

if os.name == "nt" and not sys.flags.utf8_mode:             # Windows 콘솔 cp949 대비
    os.environ["PYTHONUTF8"] = "1"
    sys.exit(subprocess.call([sys.executable, *sys.argv]))

SKILL = Path(__file__).resolve().parents[1]


def find_soffice() -> str | None:
    cands = [os.environ.get("SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
             "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
             "/Applications/LibreOffice.app/Contents/MacOS/soffice",
             r"C:\Program Files\LibreOffice\program\soffice.exe",
             r"C:\Program Files (x86)\LibreOffice\program\soffice.exe", "/usr/bin/soffice"]
    return next((c for c in cands if c and Path(c).exists()), None)


def render(pptx: Path, deck: Path) -> Path | None:
    so = find_soffice()
    if not so:
        print("[render] LibreOffice(soffice)를 찾지 못함 — 미리보기 생략. PowerPoint 로 직접 열어 확인할 것")
        return None
    out = deck / "render"
    shutil.rmtree(out, ignore_errors=True)
    (out / "png").mkdir(parents=True)
    prof = Path(tempfile.mkdtemp())
    subprocess.run([so, "--headless", f"-env:UserInstallation={prof.as_uri()}", "--convert-to", "pdf",
                    "--outdir", str(out), str(pptx)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   timeout=600, check=False)
    shutil.rmtree(prof, ignore_errors=True)
    pdf = out / f"{pptx.stem}.pdf"
    if not pdf.exists():
        print("[render] PDF 변환 실패")
        return None
    try:
        import pymupdf as fitz  # noqa: F401
    except ImportError:  # PyMuPDF 1.24.3 미만
        import fitz  # noqa: F401
    from PIL import Image
    doc = fitz.open(str(pdf))
    thumbs = []
    for i, pg in enumerate(doc, start=1):
        pix = pg.get_pixmap(dpi=110)
        p = out / "png" / f"slide-{i:02d}.png"
        pix.save(str(p))
        im = Image.open(p).convert("RGB")
        im.thumbnail((640, 360))
        thumbs.append(im)
    doc.close()
    if thumbs:
        cols = min(3, len(thumbs))
        rows = (len(thumbs) + cols - 1) // cols
        w, h = thumbs[0].size
        sheet = Image.new("RGB", (cols * (w + 12) + 12, rows * (h + 12) + 12), (210, 210, 210))
        for k, im in enumerate(thumbs):
            sheet.paste(im, (12 + (k % cols) * (w + 12), 12 + (k // cols) * (h + 12)))
        sheet.save(out / "contact.png")
    dst = pptx.with_name(f"{pptx.stem}_미리보기.pdf")
    shutil.copy2(pdf, dst)
    print(f"[render] {len(thumbs)}장 → {out / 'png'}  ·  미리보기 PDF {dst}")
    return dst


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    deck = Path(args[0]).resolve()
    if "--no-build" not in sys.argv:
        b = deck / "build.py"
        if not b.exists():
            sys.exit(f"build.py 없음: {b}")
        env = dict(os.environ, PPT_EXPERIMENT_SKILL=str(SKILL), PYTHONUTF8="1")
        r = subprocess.run([sys.executable, "build.py"], cwd=deck, env=env)
        if r.returncode:
            sys.exit(f"build.py 실패(종료 코드 {r.returncode})")
    outs = sorted((deck / "output").glob("*.pptx"), key=lambda p: p.stat().st_mtime)
    outs = [p for p in outs if not p.name.startswith("~$")]
    if not outs:
        sys.exit(f"{deck / 'output'} 에 pptx 가 없음 — build.py 의 저장 경로를 output/ 으로")
    pptx = outs[-1]
    print(f"[build] {pptx}")
    if "--no-render" not in sys.argv:
        render(pptx, deck)
    r = subprocess.run([sys.executable, str(SKILL / "scripts" / "check.py"), str(pptx)])
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()
