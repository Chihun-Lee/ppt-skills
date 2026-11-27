# -*- coding: utf-8 -*-
"""필요한 패키지가 없으면 install.py 가 만든 ~/.ppt_skills/venv 의 파이썬으로 같은 명령을 다시 실행한다.

스크립트 맨 앞에서 `import _venv` 한 줄로 쓴다. 끄려면 환경변수 PPT_SKILLS_NO_VENV=1.
"""
import os
import subprocess
import sys
from pathlib import Path


def _venv_dir() -> Path:
    return Path(os.environ.get("PPT_SKILLS_HOME") or (Path.home() / ".ppt_skills")) / "venv"


def _venv_python():
    v = _venv_dir()
    p = v / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    return p if p.exists() else None


if not os.environ.get("PPT_SKILLS_NO_VENV"):
    try:
        import pptx  # noqa: F401
        try:
            import pymupdf as fitz  # noqa: F401
        except ImportError:  # PyMuPDF 1.24.3 미만
            import fitz  # noqa: F401
        import PIL  # noqa: F401
    except ImportError:
        _vp = _venv_python()
        _in_venv = Path(sys.prefix).resolve() == _venv_dir().resolve()
        if _vp and not _in_venv:
            _env = dict(os.environ, PPT_SKILLS_NO_VENV="1")
            if os.name == "nt":
                _env["PYTHONUTF8"] = "1"
            sys.exit(subprocess.call([str(_vp), *sys.argv], env=_env))
