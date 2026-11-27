#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ppt-skills 설치 — Windows·macOS·Linux 공통.

    python install.py                  스킬 2개(ppt, ppt_experiment)를 ~/.claude/skills/ 에 넣고,
                                       ~/.ppt_skills/venv 에 파이썬 패키지, Paperlogy 글꼴을 사용자 글꼴 폴더에 설치, 자가검사
    python install.py --link           복사 대신 연결(git pull 하면 바로 최신). Windows 는 디렉터리 정션
    python install.py --copy           연결 대신 복사 (Windows 기본)
    python install.py --dest <폴더>    다른 스킬 폴더(예: 프로젝트의 .claude/skills)
    python install.py --no-fonts       글꼴 설치 건너뜀
    python install.py --no-selftest    예시 덱 빌드 건너뜀
    python install.py --pip-args "--trusted-host pypi.org --trusted-host files.pythonhosted.org"   회사망 SSL 오류
    python install.py --uninstall      스킬 2개만 지움(~/.ppt_skills 의 설정·로고·가상환경은 남김)

Claude Code 에게 맡길 때의 절차는 INSTALL.md.
- 이 설치가 복사한 스킬 폴더에는 표식 파일(.ppt-skills.json)을 둔다. 다시 설치하면 표식이 있는 폴더는 그냥 바꾼다.
- 표식이 없는 같은 이름의 스킬(직접 만든 ppt 스킬 등)은 지우지 않고 ~/.ppt_skills/backup/ 으로 옮겨 둔다.
- 관리자 권한·전역 파이썬은 쓰지 않는다.
"""
import argparse
import datetime
import json
import os
import shutil
import stat
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parent
REPO_URL = "https://github.com/Chihun-Lee/ppt-skills"
SKILLS = [d for d in sorted((REPO / "skills").iterdir()) if (d / "SKILL.md").exists()]
WIN = os.name == "nt"
HOME = Path(os.environ.get("PPT_SKILLS_HOME") or (Path.home() / ".ppt_skills")).expanduser()
VENV = HOME / "venv"
VPY = VENV / ("Scripts/python.exe" if WIN else "bin/python")
MARK = ".ppt-skills.json"
PKGS = ["python-pptx>=0.6.21", "lxml", "Pillow>=9.2", "PyMuPDF", "numpy", "scipy", "matplotlib"]

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass


def version():
    try:
        return (REPO / "VERSION").read_text(encoding="utf-8").strip()
    except Exception:  # noqa: BLE001
        return "?"


# ── 스킬 넣기 ────────────────────────────────────────────────────────────────
def is_link(p):
    if p.is_symlink():
        return True
    if WIN and p.exists():
        try:
            return bool(os.readlink(p))
        except (OSError, ValueError):
            return False
    return False


def target_of(p):
    try:
        return Path(os.readlink(p)).resolve()
    except OSError:
        return None


def link(src, t):
    if WIN:
        r = subprocess.run(["cmd", "/c", "mklink", "/J", str(t), str(src)], capture_output=True)
        if r.returncode:
            raise OSError((r.stdout + r.stderr).decode("mbcs", errors="replace"))
    else:
        os.symlink(src, t, target_is_directory=True)


def _force_writable(func, path, _exc):
    os.chmod(path, stat.S_IWRITE)
    func(path)


def remove(t):
    if is_link(t):
        os.rmdir(t) if WIN else t.unlink()
    elif t.is_dir():
        if sys.version_info >= (3, 12):
            shutil.rmtree(t, onexc=_force_writable)
        else:
            shutil.rmtree(t, onerror=_force_writable)
    elif t.exists():
        t.unlink()


def backup(t):
    stamp = f"{datetime.datetime.now():%Y%m%d%H%M%S}"
    bk = HOME / "backup" / f"{t.name}.{stamp}"
    bk.parent.mkdir(parents=True, exist_ok=True)
    if is_link(t):
        old = target_of(t)
        (bk.parent / f"{t.name}.{stamp}.link.txt").write_text(f"{t} → {old}\n", encoding="utf-8")
        remove(t)
        print(f"  기존 {t.name} 연결 해제 (원래 대상 {old} — 기록: {bk.parent})")
    else:
        shutil.move(str(t), str(bk))
        print(f"  기존 {t.name} 백업 → {bk}")


def ours(t):
    return not is_link(t) and (t / MARK).is_file()


def put_skills(dest: Path, use_link: bool, ver: str):
    dest.mkdir(parents=True, exist_ok=True)
    print(f"[1] 스킬 → {dest}  ({'연결' if use_link else '복사'})")
    for d in SKILLS:
        t = dest / d.name
        if use_link and is_link(t) and target_of(t) == d.resolve():
            print(f"  = {d.name} (이미 연결됨)")
            continue
        if ours(t):
            remove(t)
        elif t.exists() or is_link(t):
            backup(t)
        if use_link:
            link(d, t)
            print(f"  + {d.name} → {d}")
        else:
            shutil.copytree(d, t, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store"))
            (t / MARK).write_text(json.dumps({"version": ver, "source": str(REPO)}, ensure_ascii=False) + "\n",
                                  encoding="utf-8")
            print(f"  + {d.name} (복사)")


# ── 파이썬 패키지 ─────────────────────────────────────────────────────────────
def setup_venv(pip_args: str):
    print(f"\n[2] 파이썬 패키지 → {VENV}")
    if not VPY.exists():
        r = subprocess.run([sys.executable, "-m", "venv", str(VENV)])
        if r.returncode or not VPY.exists():
            sys.exit("venv를 만들지 못함 — Linux 면 python3-venv 패키지가 필요하다(INSTALL.md 문제 해결)")
    extra = pip_args.split() if pip_args else []
    subprocess.run([str(VPY), "-m", "pip", "install", "-q", "--upgrade", "pip", *extra])
    r = subprocess.run([str(VPY), "-m", "pip", "install", "-q", *extra, *PKGS])
    if r.returncode:
        sys.exit("pip 설치 실패 — 회사망이면 --pip-args 로 다시(INSTALL.md 문제 해결)")
    print("  ✓", ", ".join(p.split(">")[0] for p in PKGS))


# ── 글꼴 ─────────────────────────────────────────────────────────────────────
def install_fonts():
    src = REPO / "skills" / "ppt" / "fonts"
    ttfs = sorted(src.glob("Paperlogy-*.ttf"))
    print(f"\n[3] Paperlogy 글꼴 {len(ttfs)}종 (SIL OFL 1.1) → 사용자 글꼴 폴더")
    if sys.platform == "darwin":
        dst = Path.home() / "Library" / "Fonts"
    elif WIN:
        dst = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local"))) / "Microsoft" / "Windows" / "Fonts"
    else:
        dst = Path.home() / ".local" / "share" / "fonts"
    dst.mkdir(parents=True, exist_ok=True)
    new = 0
    for f in ttfs:
        t = dst / f.name
        if not t.exists():
            shutil.copy2(f, t)
            new += 1
        if WIN:
            _win_register_font(f.stem, t)
    if not WIN and sys.platform != "darwin" and shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f", str(dst)], capture_output=True)
    print(f"  ✓ {dst} (새로 {new}개, 이미 있음 {len(ttfs) - new}개)")
    if new:
        print("  · 열려 있던 PowerPoint·LibreOffice 는 다시 켜야 새 글꼴이 보인다")


def _win_register_font(stem: str, path: Path):
    """Windows 사용자 글꼴 등록(HKCU, 관리자 권한 불필요, Windows 10 1809 이상)."""
    import winreg
    w = stem.split("-", 1)[1] if "-" in stem else stem          # 4Regular
    name = f"Paperlogy {w[0]} {w[1:]} (TrueType)"
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows NT\CurrentVersion\Fonts")
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(path))
    winreg.CloseKey(key)
    try:
        import ctypes
        ctypes.windll.gdi32.AddFontResourceW(str(path))
    except Exception:  # noqa: BLE001
        pass


# ── 자가검사 ─────────────────────────────────────────────────────────────────
def selftest(dest: Path) -> bool:
    print("\n[4] 자가검사 — 예시 덱 2종 빌드·검사")
    env = dict(os.environ, PYTHONUTF8="1", PPT_SKILLS_NO_VENV="1")
    ok = True
    r = subprocess.run([str(VPY), str(dest / "ppt" / "scripts" / "doctor.py")], env=env)
    soffice_ok = r.returncode == 0
    tmp = Path(tempfile.mkdtemp(prefix="ppt_skills_test_"))
    for skill, sub, demo_env in (("ppt", "demo", "KIMS_PPT_SKILL"), ("ppt_experiment", "demo", "PPT_EXPERIMENT_SKILL")):
        d = tmp / skill
        shutil.copytree(dest / skill / "examples" / sub, d, ignore=shutil.ignore_patterns("output", "render", "images"))
        args = [str(VPY), str(dest / skill / "scripts" / "make.py"), str(d)]
        if skill == "ppt":
            args.append("--full")
        if not soffice_ok:
            args.append("--no-render")
        e = dict(env, **{demo_env: str(dest / skill)}, PPT_LOGO="placeholder")
        r = subprocess.run(args, env=e, capture_output=True, text=True, encoding="utf-8", errors="replace")
        tail = "\n".join((r.stdout + r.stderr).strip().splitlines()[-6:])
        outs = sorted((d / "output").glob("*.pptx"))
        good = r.returncode == 0 and outs
        ok &= bool(good)
        print(f"  {'✓' if good else '✗'} {skill}: {outs[0] if outs else '(pptx 없음)'}")
        if not good:
            print("    " + tail.replace("\n", "\n    "))
    if not soffice_ok:
        print("  · LibreOffice 가 없거나 점검 항목이 남아 미리보기 렌더는 건너뜀(INSTALL.md 4단계)")
    print(f"  예시 출력 폴더: {tmp}")
    return ok


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dest", default=str(Path.home() / ".claude" / "skills"))
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--link", action="store_true")
    g.add_argument("--copy", action="store_true")
    ap.add_argument("--no-fonts", action="store_true")
    ap.add_argument("--no-selftest", action="store_true")
    ap.add_argument("--no-setup", action="store_true", help="파이썬 패키지·글꼴·자가검사 모두 건너뜀")
    ap.add_argument("--pip-args", default=os.environ.get("PPT_PIP_ARGS", ""))
    ap.add_argument("--uninstall", action="store_true")
    a = ap.parse_args()
    if sys.version_info < (3, 9):
        sys.exit("Python 3.9 이상이 필요하다 — https://www.python.org/downloads/")
    dest = Path(a.dest).expanduser().resolve()
    ver = version()

    if a.uninstall:
        for d in SKILLS:
            t = dest / d.name
            if ours(t) or (is_link(t) and target_of(t) == d.resolve()):
                remove(t)
                print(f"  - {t}")
            elif t.exists():
                print(f"  = {t} (이 설치가 넣은 것이 아니라 그대로 둠)")
        print(f"\n스킬을 지웠다. 설정·로고·가상환경까지 지우려면 {HOME} 폴더를 지운다(글꼴은 그대로 둔다).")
        return

    use_link = a.link or (not a.copy and not WIN)
    print(f"ppt-skills {ver} 설치\n")
    put_skills(dest, use_link, ver)
    HOME.mkdir(parents=True, exist_ok=True)
    ok = True
    if not a.no_setup:
        setup_venv(a.pip_args)
        if not a.no_fonts:
            install_fonts()
        if not a.no_selftest:
            ok = selftest(dest)
    (HOME / "install.json").write_text(json.dumps({
        "version": ver, "source": str(REPO), "git": (REPO / ".git").exists(), "repo_url": REPO_URL,
        "dest": str(dest), "mode": "link" if use_link else "copy", "venv": str(VENV),
        "skills": [d.name for d in SKILLS], "installed": f"{datetime.datetime.now():%Y-%m-%d %H:%M}",
    }, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    if not ok:
        sys.exit("\n자가검사 실패 — 위 출력을 확인한다(INSTALL.md 문제 해결)")
    me = "py -3" if WIN else "python3"
    print(f"\n✓ 설치 완료 (ppt-skills {ver})")
    print("  - Claude Code 를 다시 시작하면(새 대화) /ppt 와 /ppt_experiment 스킬이 보인다.")
    print("  - 처음 쓸 때 Claude 가 두 종류(① 다목적 발표·학회·보고 PPT / ② 실험결과 정리 PPT) 중 고르게 하고, 로고 파일을 묻는다.")
    print("    로고가 없으면 'LOGO' 자리표시가 들어간다.")
    print(f"  - 설정·로고·가상환경: {HOME}")
    print(f"  - 점검: {me} \"{dest / 'ppt' / 'scripts' / 'doctor.py'}\"")


if __name__ == "__main__":
    main()
