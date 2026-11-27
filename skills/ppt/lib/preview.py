# -*- coding: utf-8 -*-
"""미리보기 렌더: 최종 pptx 는 그대로 두고, LibreOffice 용 사본만 보정해 PDF·PNG 를 만든다.

  1) 영문 폰트명 → 한글 패밀리명 (원작자 Mac LibreOffice 는 한글명만 등록: 실험 E1).
     PC마다 다르므로 렌더 PDF에 Paperlogy 가 안 박히면 영문명으로 다시 렌더하고 ~/.kims_ppt_preview_mode 에 기억
       'Paperlogy N Weight' → '페이퍼로지 N Weight'  (슬라이드·레이아웃·마스터·테마·네이티브 차트 전부)
  2) 같은 크기·같은 충돌 클래스 굵기 혼용 → 두 번째 굵기부터 +0.2pt (LibreOffice 폰트 캐시 충돌, 실험 E1)
       클래스(b=0): N={1,4,5,6} L={2,3} B={7,8} K={9}
  3) soffice → PDF, 페이지 PNG(1600px), 컨택트시트

사용: python preview.py deck.pptx [출력폴더]   → 출력폴더/<이름>.pdf, png/slide-NN.png, contact.png
"""
from __future__ import annotations

import collections
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
EN2KO = [(re.compile(r"^Paperlogy (\d) (\w+)$"), r"페이퍼로지 \1 \2")]
PL = re.compile(r"^(?:페이퍼로지|Paperlogy) (\d) \w+$")
CLASS = {"1": "N", "4": "N", "5": "N", "6": "N", "2": "L", "3": "L", "7": "B", "8": "B", "9": "K"}
STEP = 20


def _find_soffice() -> str:
    """LibreOffice 실행 파일: 환경변수 SOFFICE → PATH → OS별 기본 설치 위치."""
    import os
    cands = [os.environ.get("SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
             "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
             "/Applications/LibreOffice.app/Contents/MacOS/soffice",
             r"C:\Program Files\LibreOffice\program\soffice.exe",
             r"C:\Program Files (x86)\LibreOffice\program\soffice.exe", "/usr/bin/soffice"]
    for c in cands:
        if c and Path(c).exists():
            return c
    raise RuntimeError("LibreOffice(soffice)를 찾지 못함 — 설치하거나 환경변수 SOFFICE 에 경로 지정(README 설치 3단계)")


# 미리보기 폰트명: 'ko' = 한글 패밀리명(페이퍼로지 N …), 'en' = 영문명 그대로.
# LibreOffice 가 어느 이름으로 등록하는지는 PC(OS 언어)마다 다르다 → 렌더 PDF에 Paperlogy 가 안 박히면 반대 모드로 한 번 더 렌더하고 기억한다.
MODE_FILE = Path.home() / ".kims_ppt_preview_mode"
PY_TOOLS = Path(__file__).resolve().parents[1] / "scripts"


def _face(rpr):
    for tag in ("a:ea", "a:latin"):
        el = rpr.find(tag, NS)
        if el is not None and el.get("typeface"):
            return el.get("typeface")
    return None


def fix(src: Path, dst: Path, mode: str = "ko") -> dict:
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    xml_parts = [n for n in parts if re.match(r"ppt/(slides|slideLayouts|slideMasters|theme|charts)/[^/]+\.xml$", n)]
    trees = {n: etree.fromstring(parts[n]) for n in xml_parts}
    renamed = collections.Counter()
    for n, root in trees.items():
        for el in root.iter("{%s}latin" % NS["a"], "{%s}ea" % NS["a"], "{%s}cs" % NS["a"],
                            "{%s}font" % NS["a"], "{%s}buFont" % NS["a"]):
            tf = el.get("typeface") or ""
            for pat, rep in (EN2KO if mode == "ko" else []):
                if pat.match(tf):
                    el.set("typeface", pat.sub(rep, tf))
                    renamed[tf] += 1
                    break
    usage = collections.defaultdict(collections.Counter)
    runs = []
    for n in sorted(k for k in trees if "/slides/" in k):
        for rpr in trees[n].iter("{%s}rPr" % NS["a"], "{%s}endParaRPr" % NS["a"]):
            f = _face(rpr)
            sz = rpr.get("sz")
            m = PL.match(f or "")
            if not m or not sz:
                continue
            w = m.group(1)
            cls = "B" if rpr.get("b") in ("1", "true") else CLASS[w]
            key = (cls, int(sz))
            usage[key][w] += 1
            runs.append((rpr, key, w))
    used = collections.defaultdict(set)
    for (cls, sz) in usage:
        used[cls].add(sz)
    remap = {}
    for (cls, sz), cnt in sorted(usage.items()):
        for w, _ in cnt.most_common()[1:]:
            k = 1
            while any(abs(sz + k * STEP - u) < 15 for u in used[cls]):
                k += 1
            remap[(cls, sz, w)] = sz + k * STEP
            used[cls].add(sz + k * STEP)
    nudged = collections.Counter()
    for rpr, (cls, sz), w in runs:
        new = remap.get((cls, sz, w))
        if new:
            rpr.set("sz", str(new))
            nudged[f"w{w}@{sz / 100:g}→{new / 100:g}"] += 1
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for n, data in parts.items():
        if n in trees:
            data = etree.tostring(trees[n], xml_declaration=True, encoding="UTF-8", standalone=True)
        zout.writestr(n, data)
    zout.close()
    return {"renamed": dict(renamed), "nudged": dict(nudged)}


def _has_paperlogy(pdf: Path) -> bool:
    """렌더 PDF에 Paperlogy 글꼴이 실제로 박혔는지(대체 폰트로 새지 않았는지)."""
    try:
        try:
            import pymupdf as fitz  # noqa: F401
        except ImportError:  # PyMuPDF 1.24.3 미만
            import fitz  # noqa: F401
        doc = fitz.open(str(pdf))
        names = {f[3] for pno in range(min(3, len(doc))) for f in doc.get_page_fonts(pno)}
        doc.close()
        return any("paperlogy" in n.lower() or "페이퍼로지" in n for n in names)
    except Exception:
        return True


def render(pptx, out_dir=None, width=1600) -> Path:
    pptx = Path(pptx).resolve()
    out = Path(out_dir) if out_dir else pptx.parent / "render"
    out.mkdir(parents=True, exist_ok=True)
    name = pptx.stem
    soffice = _find_soffice()
    first = MODE_FILE.read_text().strip() if MODE_FILE.exists() else "ko"
    modes = [first, "en" if first == "ko" else "ko"]
    for i, mode in enumerate(modes):
        tmp = Path(tempfile.mkdtemp())
        prev = tmp / f"{name}.pptx"
        log = fix(pptx, prev, mode)
        profile = Path(tempfile.mkdtemp())
        subprocess.run([soffice, "--headless", f"-env:UserInstallation={profile.as_uri()}",
                        "--convert-to", "pdf", "--outdir", str(tmp), str(prev)],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False, timeout=600)
        shutil.rmtree(profile, ignore_errors=True)
        pdf_tmp = tmp / f"{name}.pdf"
        if not pdf_tmp.exists():
            raise RuntimeError(f"LibreOffice PDF 변환 실패: {pptx}")
        if _has_paperlogy(pdf_tmp) or i == 1:
            if _has_paperlogy(pdf_tmp):
                MODE_FILE.write_text(mode)
            else:
                print("[preview] 렌더 PDF에 Paperlogy 가 없음 — OS에 Paperlogy 를 설치했는지 확인(README 설치 2단계)")
            break
        shutil.rmtree(tmp, ignore_errors=True)
    log["font_mode"] = mode
    pdf = out / f"{name}.pdf"
    shutil.move(str(pdf_tmp), pdf)
    shutil.rmtree(tmp, ignore_errors=True)
    png_dir = out / "png"
    shutil.rmtree(png_dir, ignore_errors=True)
    subprocess.run([sys.executable, str(PY_TOOLS / "contact_sheet.py"), str(pdf), str(png_dir), str(out / "contact.png"),
                    "--width", str(width)], check=True)
    (out / "preview_fix.json").write_text(str(log))
    return pdf


if __name__ == "__main__":
    p = render(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(p)
