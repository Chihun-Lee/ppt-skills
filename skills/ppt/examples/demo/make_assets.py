# -*- coding: utf-8 -*-
"""예시 덱용 그림 만들기 (모두 합성 데이터 — 실제 실험 결과가 아니다).

assets/ 에 이미 만들어 둔 그림이 들어 있으므로 다시 돌릴 필요는 없다. matplotlib 그래프를 슬롯 크기로 그리는 방법의 예시다.
사용: python make_assets.py
"""
import os
import sys
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
SKILL = Path(os.environ.get("KIMS_PPT_SKILL", HERE.parents[1]))
sys.path.insert(0, str(SKILL / "lib"))
import matplotlib  # noqa: E402

matplotlib.use("Agg")
import matplotlib as mpl  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
from kims_ppt import Rect, chart_figsize, mpl_style  # noqa: E402

A = HERE / "assets"
A.mkdir(exist_ok=True)
mpl.rcParams.update(mpl_style())
rng = np.random.default_rng(7)


def hardness_map():
    """시효 온도 × 시간 경도 지도(합성). 전폭 슬롯(요지 + 세부 2줄 아래)에 맞춘 크기."""
    T = np.linspace(150, 220, 8)
    t = np.array([1, 2, 4, 8, 12, 24, 36, 48])
    TT, tt = np.meshgrid(T, t, indexing="ij")
    peak_t = 40 * np.exp(-(TT - 150) / 30)
    hv = 95 + 45 * np.exp(-((np.log(tt) - np.log(peak_t)) ** 2) / 1.2) * (1 - (TT - 150) / 260)
    fig = plt.figure(figsize=chart_figsize(Rect(120, 0, 1680, 640)), layout="none")
    ax = fig.add_axes([0.07, 0.17, 0.8, 0.77])
    im = ax.imshow(hv, origin="lower", aspect="auto", cmap="viridis",
                   extent=[-0.5, len(t) - 0.5, T[0] - 5, T[-1] + 5])
    ax.set_xticks(range(len(t)), [str(v) for v in t])
    ax.set_xlabel("Aging time (h)")
    ax.set_ylabel("Aging temp. (°C)")
    cax = fig.add_axes([0.89, 0.17, 0.015, 0.77])
    fig.colorbar(im, cax=cax).set_label("Hardness (HV)")
    fig.savefig(A / "hardness_map.png", dpi=200)
    plt.close(fig)


def parity():
    """예측-실측 비교(합성): 가로 = 예측, 세로 = 실측."""
    pred = rng.uniform(260, 380, 36)
    meas = pred * 1.03 + rng.normal(0, 9, pred.size)
    fig = plt.figure(figsize=chart_figsize(Rect(0, 0, 900, 700)), layout="none")
    ax = fig.add_axes([0.16, 0.15, 0.78, 0.8])
    lo, hi = 250, 400
    ax.plot([lo, hi], [lo, hi], color="#808080", lw=1.2, label="1:1")
    ax.plot([lo, hi], [lo * 1.03, hi * 1.03], color="#808080", lw=1.2, ls="--", label="+3 %")
    ax.scatter(pred, meas, s=36, color="#546882", label="Specimens", zorder=3)
    mae = np.mean(np.abs(meas - pred))
    ax.text(258, 385, f"MAE {mae:.1f} MPa", fontsize=14)
    ax.set_xlim(lo, hi)
    ax.set_ylim(lo, hi)
    ax.set_xlabel("Predicted YS (MPa)")
    ax.set_ylabel("Measured YS (MPa)")
    ax.legend(loc="lower right")
    fig.savefig(A / "parity.png", dpi=200)
    plt.close(fig)


def micrograph():
    """결정립 조직처럼 보이는 합성 이미지(보로노이 + 잡음). 사진 모듈 예시용."""
    from PIL import Image, ImageFilter
    W, H = 1200, 900
    n = 140
    pts = rng.uniform(0, 1, (n, 2)) * [W, H]
    gray = rng.uniform(120, 215, n)
    yy, xx = np.mgrid[0:H, 0:W]
    best = np.full((H, W), np.inf)
    second = np.full((H, W), np.inf)
    lab = np.zeros((H, W), int)
    for i, (px, py) in enumerate(pts):
        d = (xx - px) ** 2 + (yy - py) ** 2
        m = d < best
        second = np.where(m, best, np.minimum(second, d))
        lab = np.where(m, i, lab)
        best = np.where(m, d, best)
    img = gray[lab]
    edge = (np.sqrt(second) - np.sqrt(best)) < 2.2
    img = np.where(edge, 45, img) + rng.normal(0, 6, (H, W))
    im = Image.fromarray(np.clip(img, 0, 255).astype("uint8")).filter(ImageFilter.GaussianBlur(0.8))
    im.convert("RGB").save(A / "micrograph.jpg", quality=90)


def doe_points():
    """후보 64조건(회색)과 추천 12조건(파랑) — 단계 흐름 아래 전폭 슬롯."""
    T = np.linspace(150, 220, 8)
    t = np.array([1, 2, 4, 8, 12, 24, 36, 48])
    fig = plt.figure(figsize=chart_figsize(Rect(120, 0, 1680, 590)), layout="none")
    ax = fig.add_axes([0.07, 0.17, 0.9, 0.7])
    gx, gy = np.meshgrid(range(len(t)), T)
    ax.scatter(gx, gy, s=60, color="#BFBFBF", label="Candidates (64)")
    pick = [(2, 210), (3, 200), (3, 190), (4, 180), (4, 190), (5, 170), (5, 180), (6, 160), (6, 170), (7, 150), (7, 160), (4, 200)]
    ax.scatter([a for a, _ in pick], [b for _, b in pick], s=110, color="#546882", label="Recommended (12)", zorder=3)
    ax.set_xticks(range(len(t)), [str(v) for v in t])
    ax.set_xlabel("Aging time (h)")
    ax.set_ylabel("Aging temp. (°C)")
    ax.set_ylim(143, 227)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.0), ncol=2)
    fig.savefig(A / "doe_points.png", dpi=200)
    plt.close(fig)


if __name__ == "__main__":
    doe_points()
    hardness_map()
    parity()
    micrograph()
    print("assets:", sorted(p.name for p in A.iterdir()))
