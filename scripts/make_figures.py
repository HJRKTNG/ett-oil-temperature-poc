"""結果の図を reports/figures/ に書き出す（run_baselines.py と run_lgbm.py の後に実行）。

- res_mae_by_horizon.png       設備×ホライズン別 MAE（テスト）: ベースライン vs LightGBM
- res_ablation_h6.png          主ホライズンの特徴量アブレーション
- res_alarm_timeline_*.png     高温の事前検知: 成功例・失敗例の時系列（ETTh2・6h・条件 B）
- res_error_by_month_hour.png  誤差の季節・時刻依存（ETTh2・6h）
- res_importance_*.png         特徴量重要度（上位 15）
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
REP = ROOT / "reports"
FIG = REP / "figures"
plt.rcParams.update({"font.family": "Arial", "figure.dpi": 130, "axes.grid": True, "grid.alpha": 0.3})
C = {"B0_persistence": "#9aa0ad", "B1_seasonal24h": "#c79a1e", "B2_ridge(B)": "#5c7a3a", "M1_lgbm(B)": "#3f55a8", "M1_lgbm(A)": "#8ea0e6"}

base = pd.read_csv(REP / "results_baselines.csv")
lgb = pd.read_csv(REP / "results_lgbm.csv")
res = pd.concat([base, lgb], ignore_index=True)
test = res[res.split == "test"]

# 1) MAE by horizon
fig, axes = plt.subplots(1, 2, figsize=(11, 3.6), sharey=False)
for ax, name in zip(axes, ["ETTh1", "ETTh2"]):
    sub = test[test.dataset == name]
    models = [m for m in C if m in set(sub.model)]
    hs = [1, 6, 24]
    w = 0.8 / len(models)
    for i, m in enumerate(models):
        vals = [sub[(sub.model == m) & (sub.horizon_h == h)].MAE.values[0] for h in hs]
        ax.bar(np.arange(len(hs)) + i * w - 0.4 + w / 2, vals, w, label=m, color=C[m])
        for x, v in zip(np.arange(len(hs)) + i * w - 0.4 + w / 2, vals):
            ax.text(x, v + 0.03, f"{v:.2f}", ha="center", va="bottom", fontsize=7)
    ax.set_xticks(range(len(hs)))
    ax.set_xticklabels([f"{h}h ahead" for h in hs])
    ax.set_ylabel("MAE [°C] (test period)")
    ax.set_title(name)
axes[0].legend(fontsize=8, loc="upper left")
fig.suptitle("Forecast error by horizon: baselines vs LightGBM (test = Nov 2017–Jun 2018)")
fig.tight_layout()
fig.savefig(FIG / "res_mae_by_horizon.png")
plt.close(fig)

# 2) Ablation
abl = pd.read_csv(REP / "results_ablation.csv")
fig, ax = plt.subplots(figsize=(8, 3.4))
order = ["1_ot_only", "2_+calendar", "3_+load_origin (B)", "4_+load_target (A)"]
for i, (name, col) in enumerate([("ETTh1", "#3f55a8"), ("ETTh2", "#b9642a")]):
    for j, part in enumerate(["valid", "test"]):
        sub = abl[(abl.dataset == name) & (abl.split == part)].set_index("features").loc[order]
        ax.plot(range(len(order)), sub.MAE, marker="o", color=col, ls="-" if part == "test" else "--", label=f"{name} {part}")
ax.set_xticks(range(len(order)))
ax.set_xticklabels(["OT history only", "+ calendar", "+ loads up to s (cond. B)", "+ loads at T (cond. A)"], fontsize=8)
ax.set_ylabel("MAE [°C], 6h ahead")
ax.set_title("What each feature group adds (LightGBM, residual learning, 6h ahead)")
ax.legend(fontsize=8)
fig.tight_layout()
fig.savefig(FIG / "res_ablation_h6.png")
plt.close(fig)

# 3) Alarm timelines (ETTh2, 6h, cond B)
pred = pd.read_csv(REP / "predictions" / "ETTh2_h6_B.csv", index_col=0, parse_dates=True)
h = 6
thr = float(pd.read_csv(ROOT / "data" / "ETTh2.csv", parse_dates=["date"]).set_index("date").loc[:"2017-06-30", "OT"].quantile(0.95))
te = pred[pred.split == "test"]
actual = te.y > thr
alarm = (te.pred > thr) & (te.ot_now <= thr)
f = actual.astype(bool)
starts = f & ~f.shift(1, fill_value=False)
ends = f & ~f.shift(-1, fill_value=False)
events = list(zip(te.index[starts], te.index[ends]))


def plot_event(start, end, tag):
    lo, hi = start - pd.Timedelta(hours=36), end + pd.Timedelta(hours=18)
    w = te.loc[lo:hi]
    fig, ax = plt.subplots(figsize=(9, 3.2))
    # 実測は対象時刻で描く（行の起点 + h）
    ax.plot(w.index + pd.Timedelta(hours=h), w.y, color="k", lw=1.2, label="actual OT at T")
    ax.plot(w.index + pd.Timedelta(hours=h), w.pred, color="#3f55a8", lw=1.2, label="predicted OT at T (issued 6h earlier)")
    ax.axhline(thr, color="#a33f3f", ls="--", lw=0.9, label=f"threshold {thr:.1f}°C (train 95th pct)")
    al = w[alarm.reindex(w.index).fillna(False)]
    ax.scatter(al.index + pd.Timedelta(hours=h), al.pred, color="#b9642a", s=22, zorder=5, label="alarm issued (OT[s] ≤ thr, pred > thr)")
    ax.axvspan(start + pd.Timedelta(hours=h), end + pd.Timedelta(hours=h), color="#a33f3f", alpha=0.08)
    ax.set_ylabel("OT [°C]")
    ax.set_title(f"ETTh2, 6h-ahead alarms — {tag} (event {start + pd.Timedelta(hours=h):%Y-%m-%d %H:%M} → {end + pd.Timedelta(hours=h):%m-%d %H:%M})")
    ax.legend(fontsize=7, loc="upper left")
    fig.tight_layout()
    fig.savefig(FIG / f"res_alarm_timeline_{tag}.png")
    plt.close(fig)


# 成功例 = 開始前に警報あり、失敗例 = 警報なし（それぞれ最初の 1 件）
succ = fail = None
for s0, e0 in events:
    win = te.loc[s0:e0]
    onset_rows = (win.ot_now <= thr)
    if not onset_rows.any():
        continue
    hit = (alarm.loc[s0:e0] & onset_rows).any()
    if hit and succ is None:
        succ = (s0, e0)
    if not hit and fail is None:
        fail = (s0, e0)
if succ:
    plot_event(*succ, "detected")
if fail:
    plot_event(*fail, "missed")

# 4) Error by month / hour (ETTh2, 6h, B vs persistence)
err = (te.pred - te.y).abs()
err_p = (te.ot_now - te.y).abs()
fig, axes = plt.subplots(1, 2, figsize=(11, 3.3))
bym = pd.DataFrame({"LightGBM": err.groupby(te.index.month).mean(), "persistence": err_p.groupby(te.index.month).mean()})
bym.plot(ax=axes[0], marker="o", color=["#3f55a8", "#9aa0ad"])
axes[0].set_title("ETTh2, 6h: MAE by month (test)")
axes[0].set_xlabel("month")
axes[0].set_ylabel("MAE [°C]")
byh = pd.DataFrame({"LightGBM": err.groupby((te.index + pd.Timedelta(hours=h)).hour).mean(), "persistence": err_p.groupby((te.index + pd.Timedelta(hours=h)).hour).mean()})
byh.plot(ax=axes[1], marker=".", color=["#3f55a8", "#9aa0ad"])
axes[1].set_title("ETTh2, 6h: MAE by target hour of day (test)")
axes[1].set_xlabel("hour of target time T")
fig.tight_layout()
fig.savefig(FIG / "res_error_by_month_hour.png")
plt.close(fig)

# 5) Feature importance
for name in ["ETTh1", "ETTh2"]:
    imp = pd.read_csv(REP / "predictions" / f"{name}_h6_B_importance.csv", index_col=0).iloc[:, 0].head(15)[::-1]
    fig, ax = plt.subplots(figsize=(6.5, 4))
    ax.barh(imp.index, imp.values / imp.values.sum(), color="#3f55a8")
    ax.set_xlabel("share of total gain")
    ax.set_title(f"{name}, 6h ahead (cond. B): top-15 features by gain")
    fig.tight_layout()
    fig.savefig(FIG / f"res_importance_{name}_h6.png")
    plt.close(fig)

print("figures written:", sorted(p.name for p in FIG.glob("res_*.png")))
