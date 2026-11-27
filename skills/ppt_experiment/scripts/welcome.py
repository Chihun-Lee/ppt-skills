# -*- coding: utf-8 -*-
"""첫 사용 안내 기록 — 형제 스킬 ppt 와 같은 설정 파일(~/.ppt_skills/config.json)을 쓴다.

사용: python welcome.py status     안내를 마쳤는지
      python welcome.py welcomed   안내를 마쳤다고 기록
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

if os.name == "nt" and not sys.flags.utf8_mode:
    os.environ["PYTHONUTF8"] = "1"
    sys.exit(subprocess.call([sys.executable, *sys.argv]))

USER_DIR = Path(os.environ.get("PPT_SKILLS_HOME", Path.home() / ".ppt_skills"))
CONFIG = USER_DIR / "config.json"


def read() -> dict:
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    cfg = read()
    if cmd == "status":
        print(f"설정 파일: {CONFIG} ({'있음' if CONFIG.exists() else '없음'})")
        print(f"첫 사용 안내: {'마침' if cfg.get('welcomed') else '아직 — 두 종류(다목적 ppt / 실험결과 ppt_experiment)를 소개하고 고르게 할 것'}")
    elif cmd == "welcomed":
        cfg.update(welcomed=True, welcomed_at=datetime.now().isoformat(timespec="seconds"))
        USER_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
        print("첫 사용 안내 완료로 기록")
    else:
        print(__doc__)
        sys.exit(2)


if __name__ == "__main__":
    main()
