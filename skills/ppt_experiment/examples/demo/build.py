# -*- coding: utf-8 -*-
"""ppt_experiment 예시 덱 — 그림·수치는 모두 합성 예시이며 실제 실험 결과가 아니다.

구성: ① 개요(시편·조건·산출물) ② 시편별 결과 2장(같은 레이아웃 반복) ③ 요약·결론
실행: python <스킬>/scripts/make.py <이 폴더>
"""
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = Path(os.environ.get("PPT_EXPERIMENT_SKILL", HERE.parents[1]))   # 실제 덱에서는 스킬 폴더 절대경로를 직접 쓴다
sys.path.insert(0, str(SKILL / "lib"))

import numpy as np  # noqa: E402
from lib_experiment import *  # noqa: E402,F401,F403

IMG = HERE / "images"
OUT = HERE / "output" / "실험결과_예시.pptx"

SAMPLES = {   # 합성 예시값: 시편 → (시효 조건, [(YS, UTS, El) × 3])
    "A1": ("160 °C, 8 h", [(281, 318, 14.2), (285, 321, 13.6), (279, 316, 14.9)]),
    "A2": ("180 °C, 8 h", [(302, 334, 12.1), (298, 331, 12.8), (305, 337, 11.7)]),
}


# ── 합성 그림 ────────────────────────────────────────────────────────────────
def gen_images():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    IMG.mkdir(exist_ok=True)
    rng = np.random.default_rng(7)
    plt.rcParams.update({"font.size": 13, "axes.spines.top": False, "axes.spines.right": False})

    for sid, (_, vals) in SAMPLES.items():
        # 응력-변형률 곡선 3개
        fig, ax = plt.subplots(figsize=(4.2, 3.2), dpi=160)
        for k, (ys, uts, el) in enumerate(vals, start=1):
            e = np.linspace(0, el, 300)
            s = np.where(e < ys / 700, e * 700, ys + (uts - ys) * (1 - np.exp(-(e - ys / 700) / (el * 0.25))))
            s = np.where(e > el * 0.92, s - (e - el * 0.92) * 40, s)
            ax.plot(e, s, lw=1.6, label=f"#{k}")
        ax.set_xlabel("Strain (%)"); ax.set_ylabel("Stress (MPa)"); ax.set_ylim(0, 380)
        ax.legend(frameon=False, loc="lower right")
        fig.tight_layout(); fig.savefig(IMG / f"ss_{sid}.png", facecolor="white"); plt.close(fig)

        # 결정립 조직(보로노이 합성)
        n = 70 if sid == "A1" else 110
        pts = rng.uniform(0, 1, (n, 2))
        yy, xx = np.mgrid[0:1:300j, 0:1:400j]
        d = (xx[..., None] - pts[:, 0] * 1.33 / 1.33) ** 2 + (yy[..., None] - pts[:, 1]) ** 2
        lab = d.argmin(-1)
        shade = rng.uniform(90, 200, n)[lab]
        edge = (np.abs(np.diff(lab, axis=0, prepend=lab[:1])) + np.abs(np.diff(lab, axis=1, prepend=lab[:, :1]))) > 0
        img = np.where(edge, 40, shade).astype(np.uint8)
        Image.fromarray(img, "L").convert("RGB").save(IMG / f"om_{sid}.png")

        # 경도 지도 3×3
        fig, ax = plt.subplots(figsize=(4.0, 3.2), dpi=160)
        base = 98 if sid == "A1" else 106
        h = base + rng.normal(0, 1.6, (3, 3))
        im = ax.imshow(h, cmap="viridis", vmin=92, vmax=112)
        for (i, j), v in np.ndenumerate(h):
            ax.text(j, i, f"{v:.0f}", ha="center", va="center", color="white", fontsize=13)
        ax.set_xticks([]); ax.set_yticks([])
        cb = fig.colorbar(im, ax=ax, fraction=0.046, pad=0.05)
        cb.set_label("HV0.5", labelpad=10)
        fig.tight_layout(); fig.savefig(IMG / f"hv_{sid}.png", facecolor="white"); plt.close(fig)

    # 시편 도면(인장 시편 모식도)
    fig, ax = plt.subplots(figsize=(5, 2.2), dpi=160)
    x = [0, 30, 38, 82, 90, 120, 120, 90, 82, 38, 30, 0, 0]
    y = [0, 0, 3, 3, 0, 0, 12.5, 12.5, 9.5, 9.5, 12.5, 12.5, 0]
    ax.fill(x, y, color="#C4DCF2", ec="#002060", lw=1.4)
    ax.annotate("", (38, -3), (82, -3), arrowprops=dict(arrowstyle="<->", lw=1))
    ax.text(60, -6.5, "GL 25 mm", ha="center", fontsize=12)
    ax.set_xlim(-5, 125); ax.set_ylim(-9, 15); ax.axis("off")
    fig.tight_layout(); fig.savefig(IMG / "specimen.png", facecolor="white"); plt.close(fig)


def avg(rows):
    a = np.mean(np.array(rows), axis=0)
    return [f"{a[0]:.0f}", f"{a[1]:.0f}", f"{a[2]:.1f}"]


# ── ① 개요 ───────────────────────────────────────────────────────────────────
def slide_overview(prs):
    s = blank_slide(prs)
    title_band(s, "Al-Mg-Si 시효 시편 인장·경도 시험 (예시)", style="navy")
    section_numbered(s, MARGIN, 84, "1. Specimens")
    kv_bullets(s, MARGIN + 8, 126, 500, ["Alloy : Al-Mg-Si 압출재 (합성)",
                                         "A1 : 160 °C, 8 h", "A2 : 180 °C, 8 h"])
    section_numbered(s, MARGIN, 262, "2. Test Conditions")
    kv_bullets(s, MARGIN + 8, 304, 520, ["Tensile : ASTM E8, 1 mm/min",
                                         "Hardness : HV0.5, 3×3 grid",
                                         "Temp. : 23 °C"])
    section_numbered(s, MARGIN, 440, "3. Output Files")
    kv_bullets(s, MARGIN + 8, 482, 400, ["*.csv : 하중-변위 원자료", "PPT report"])
    textbox(s, 640, 84, 600, 24, "* Tensile specimen (sub-size)", size=14, bold=True, color=NAVY)
    add_image_fit(s, IMG / "specimen.png", 640, 116, 600, 260, border=True)
    arrow_notes(s, 640, 430, 600, ["시편당 n = 3, 평균 = 산술평균",
                                   "A2 #3 그립부 미끄럼 → 재시험 예정"])


# ── ② 시편별 결과 (같은 좌표 반복) ─────────────────────────────────────────────
def slide_sample(prs, num, sid):
    cond, rows = SAMPLES[sid]
    s = blank_slide(prs)
    title_band(s, f"Sample #{num}. {sid} ({cond})", style="light")
    section_diamond(s, MARGIN, 84, "Sample Information")
    exp_table(s, MARGIN, 114, 700, ["ID", "합금", "시효", "시편 수", "규격"],
              [[sid, "Al-Mg-Si", cond, "3", "ASTM E8"]], style="navy",
              col_ratios=[10, 20, 24, 14, 18])
    section_diamond(s, MARGIN, 196, "Tensile Results")
    data = [[f"#{k}", f"{ys}", f"{uts}", f"{el:.1f}"] for k, (ys, uts, el) in enumerate(rows, start=1)]
    exp_table(s, MARGIN, 226, 700, ["No.", "YS (MPa)", "UTS (MPa)", "El. (%)"],
              data + [["Average"] + avg(rows)], style="blue", highlight_last=True)
    section_diamond(s, 800, 84, "Hardness Map")
    add_image_fit(s, IMG / f"hv_{sid}.png", 800, 114, 300, 240)
    kv_bullets(s, 1110, 150, 160, ["9점 평균", "HV0.5"])
    section_diamond(s, MARGIN, 392, "Curves & Microstructure")
    image_row(s, MARGIN, 424, 1200, 270, [
        (IMG / f"ss_{sid}.png", "Stress-Strain"),
        (IMG / f"om_{sid}.png", "OM (x200)"),
    ], label_pos="top")


# ── ③ 요약·결론 ──────────────────────────────────────────────────────────────
def slide_summary(prs):
    s = blank_slide(prs)
    title_band(s, "Summary", style="navy")
    section_diamond(s, MARGIN, 84, "결과 요약 (평균, n = 3)")
    body = [[sid, cond] + avg(rows) for sid, (cond, rows) in SAMPLES.items()]
    exp_table(s, MARGIN, 116, 1200, ["ID", "시효", "YS (MPa)", "UTS (MPa)", "El. (%)"],
              body, style="blue")
    section_diamond(s, MARGIN, 260, "결론")
    arrow_notes(s, MARGIN + 8, 296, 1180, ["A2 : YS +20 MPa, 연신율 -2 %p",
                                           "후속 : 200 °C 조건 추가 시험"])


def build():
    gen_images()
    prs = new_prs()
    slide_overview(prs)
    slide_sample(prs, 1, "A1")
    slide_sample(prs, 2, "A2")
    slide_summary(prs)
    OUT.parent.mkdir(exist_ok=True)
    prs.save(OUT)
    print(f"saved: {OUT}")


if __name__ == "__main__":
    build()
