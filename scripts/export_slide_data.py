"""報告スライドに載せる数値と系列を `reports/slide_data.json` に書き出す。

スライド生成（`scripts/build_slides.js`）は数値をこの JSON からだけ読む。
数値を手で写さないため。入力は `reports/results_*.csv`・`reports/predictions/`・`data/`。

使い方: uv run python scripts/export_slide_data.py
"""

from __future__ import annotations

import json
import re
import sys
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from ett_poc.data import TARGET, Split, load  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
H = 6  # 主ホライズン

MODELS = {
    "B0_persistence": "persistence",
    "B1_seasonal24h": "seasonal",
    "B2_ridge(B)": "ridge",
    "M1_lgbm(B)": "lgbm",
    "M1_lgbm(A)": "lgbm_A",
}


def r(x: float, n: int = 2) -> float:
    """四捨五入（README・既存の図と同じ half-up。CSV の 4.435 → 4.44）。"""
    q = Decimal(1).scaleb(-n)
    return float(Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP))


def mae_table(ndigits: int | None = 2) -> dict:
    """テスト期間の MAE。ndigits=None なら CSV の値そのまま（改善率の計算用）。"""
    base = pd.read_csv(REPORTS / "results_baselines.csv")
    lgbm = pd.read_csv(REPORTS / "results_lgbm.csv")
    both = pd.concat([base, lgbm], ignore_index=True)
    test = both[both["split"] == "test"]
    out: dict = {}
    for (ds, h, model), g in test.groupby(["dataset", "horizon_h", "model"]):
        if model not in MODELS:
            continue
        v = float(g["MAE"].iloc[0])
        out.setdefault(ds, {}).setdefault(str(h), {})[MODELS[model]] = v if ndigits is None else r(v, ndigits)
    return out


def detection() -> dict:
    """設備 2・6 時間先・q95 の高温区間の事前検知（手法ごと）。"""
    base = pd.read_csv(REPORTS / "results_baselines.csv")
    lgbm = pd.read_csv(REPORTS / "results_lgbm.csv")
    sel = lambda d: d[(d["dataset"] == "ETTh2") & (d["horizon_h"] == H) & (d["split"] == "test")]
    out = {}
    for _, row in sel(base).iterrows():
        events = int(row["q95_onset_events"])
        recall = row["q95_onset_recall"]
        out[MODELS[row["model"]]] = {
            "events": events,
            "detected": int(round(float(recall) * events)) if pd.notna(recall) else 0,
            "notices_per_day": r(row["q95_notices/day"]),
            "false_per_day": r(row["q95_false_notices/day"]),
        }
    for _, row in sel(lgbm).iterrows():
        out[MODELS[row["model"]]] = {
            "events": int(row["q95_onset_events"]),
            "detected": int(row["q95_onset_detected"]),
            "notices": int(row["q95_notices"]),
            "false_notices": int(row["q95_false_notices"]),
            "notices_per_day": r(row["q95_notices/day"]),
            "false_per_day": r(row["q95_false_notices/day"]),
            "lead_h": r(row["q95_lead_h"], 1),
            "positives": int(row["q95_positives"]),
            "point_recall": r(row["q95_point_recall"], 3),
            "q90": {
                "events": int(row["q90_onset_events"]),
                "detected": int(row["q90_onset_detected"]),
                "notices": int(row["q90_notices"]),
                "false_notices": int(row["q90_false_notices"]),
                "notices_per_day": r(row["q90_notices/day"]),
                "lead_h": r(row["q90_lead_h"], 1),
                "positives": int(row["q90_positives"]),
            },
        }
    return out


def ablation() -> dict:
    df = pd.read_csv(REPORTS / "results_ablation.csv")
    test = df[df["split"] == "test"]
    return {ds: {row["features"]: r(row["MAE"], 3) for _, row in g.iterrows()} for ds, g in test.groupby("dataset")}


def monthly() -> dict:
    df = pd.read_csv(REPORTS / "results_stability.csv")
    out = {}
    for ds, g in df.groupby("dataset"):
        piv = g.pivot(index="target_month", columns="model", values="MAE").sort_index()
        out[ds] = {"months": list(piv.index)} | {MODELS[c]: [r(v, 3) for v in piv[c]] for c in piv.columns}
    return out


def weekly() -> dict:
    """週単位で LightGBM が勝った割合（`reports/results_stability.md` の集計行を読む）。"""
    text = (REPORTS / "results_stability.md").read_text()
    pat = re.compile(r"## (ETTh\d).*?weeks: (\d+); LightGBM better than persistence in (\d+)% of weeks, "
                     r"better than ridge in (\d+)% of weeks", re.S)
    return {m[0]: {"weeks": int(m[1]), "vs_persistence": int(m[2]) / 100, "vs_ridge": int(m[3]) / 100}
            for m in pat.findall(text)}


def thresholds_raw() -> dict:
    """学習期間の油温の 95・90 パーセンタイル（`run_lgbm.py` と同じ定義）。"""
    out = {}
    for ds in ["ETTh1", "ETTh2"]:
        df = load(ds)
        tr_time = Split.default(df.index).masks(df.index)[0]
        out[ds] = {q: float(df.loc[tr_time, TARGET].quantile(q / 100)) for q in (95, 90)}
    return out


def diurnal(ds: str = "ETTh2") -> dict:
    """季節別の時刻別平均油温（全期間・EDA と同じ定義）。"""
    df = load(ds)
    season = df.index.month.map({12: "DJF", 1: "DJF", 2: "DJF", 3: "MAM", 4: "MAM", 5: "MAM",
                                 6: "JJA", 7: "JJA", 8: "JJA", 9: "SON", 10: "SON", 11: "SON"})
    g = df[TARGET].groupby([season, df.index.hour]).mean().unstack(0)
    return {
        "hours": list(range(24)),
        "series": {s: [r(v) for v in g[s]] for s in ["DJF", "MAM", "JJA", "SON"]},
        "amplitude": {s: r(g[s].max() - g[s].min()) for s in ["DJF", "MAM", "JJA", "SON"]},
        "peak_hour": {s: int(g[s].idxmax()) for s in ["DJF", "MAM", "JJA", "SON"]},
        "trough_hour": {s: int(g[s].idxmin()) for s in ["DJF", "MAM", "JJA", "SON"]},
    }


def event_window(day: str, thr: float) -> dict:
    """設備 2・6 時間先・条件 B の、対象時刻 T で見た 1 日分の実測と予測。"""
    p = pd.read_csv(REPORTS / "predictions" / "ETTh2_h6_B.csv", parse_dates=["date"]).set_index("date")
    p.index = p.index + pd.Timedelta(hours=H)  # 起点 s → 対象時刻 T
    origin_ot = p["ot_now"]
    w = p.loc[f"{day} 00:00":f"{day} 23:00"]
    # evaluate.onset_metrics と同じ判定: 超過 = 実測 > 閾値、警報 = 起点で閾値以下かつ予測 > 閾値
    alarm = (w["pred"] > thr) & (origin_ot.loc[w.index] <= thr)
    return {
        "hours": [t.strftime("%H") for t in w.index],
        "actual": [r(v, 1) for v in w["y"]],
        "pred": [r(v, 1) for v in w["pred"]],
        "exceed_hours": [t.hour for t in w.index[w["y"] > thr]],
        "alarm_hours": [t.hour for t in w.index[alarm]],
        "pred_by_hour": {t.hour: r(v, 1) for t, v in w["pred"].items()},
    }


def median_week(ds: str = "ETTh2") -> dict:
    """テスト期間で、週の誤差が中央値に最も近い 1 週間の実測と予測（設備 2・6 時間先・条件 B）。

    都合のよい週を選ばないよう、168 時点そろった週（月〜日、予測先の時刻で区切る）の MAE の
    中央値に最も近い週を機械的に選ぶ。
    """
    p = pd.read_csv(REPORTS / "predictions" / f"{ds}_h{H}_B.csv", parse_dates=["date"])
    t = p[p["split"] == "test"].copy()
    t["T"] = t["date"] + pd.Timedelta(hours=H)  # 起点 s → 対象時刻 T
    t["err"] = (t["pred"] - t["y"]).abs()
    t["week"] = t["T"].dt.to_period("W-SUN")
    g = t.groupby("week")["err"].agg(["size", "mean"])
    full = g[g["size"] == 7 * 24]
    med = full["mean"].median()
    wk = (full["mean"] - med).abs().idxmin()
    w = t[t["week"] == wk].sort_values("T")
    return {
        "start": str(w["T"].iloc[0]),
        "end": str(w["T"].iloc[-1]),
        "mae": r(w["err"].mean()),
        "median_weekly_mae": r(med),
        "n_full_weeks": int(len(full)),
        "labels": [f"{ts.month}/{ts.day} {ts.hour:02d}" for ts in w["T"]],
        "actual": [r(v, 1) for v in w["y"]],
        "pred": [r(v, 1) for v in w["pred"]],
    }


def main() -> None:
    thr = thresholds_raw()
    data = {
        "horizon_h": H,
        "test_period": ["2017-11-01", "2018-06-26"],
        "thresholds": {ds: {f"q{q}": r(v, 1) for q, v in t.items()} for ds, t in thr.items()},
        "mae": mae_table(),
        "mae_raw": mae_table(None),
        "n_test_h6": {ds: int(g["n"].iloc[0]) for ds, g in pd.read_csv(REPORTS / "results_lgbm.csv").query(
            "horizon_h == 6 and split == 'test' and model == 'M1_lgbm(B)'").groupby("dataset")},
        "detection_ETTh2_h6": detection(),
        "ablation_h6": ablation(),
        "monthly_h6": monthly(),
        "weekly_win_rate_h6": weekly(),
        "diurnal_ETTh2": diurnal("ETTh2"),
        "event_detected": event_window("2018-05-15", thr["ETTh2"][95]),
        "event_missed": event_window("2018-06-04", thr["ETTh2"][95]),
        "week_median_ETTh2_h6": median_week("ETTh2"),
    }
    out = REPORTS / "slide_data.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=1))
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
