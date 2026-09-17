"""M1: LightGBM をホライズン別・条件 A/B で学習し、検証・テストで評価する。
主ホライズン 6h でアブレーション（特徴量群を段階的に追加）も行う。

出力: reports/results_lgbm.md, reports/results_lgbm.csv, reports/predictions/<dataset>_h<h>_<cond>.csv
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from ett_poc.data import STEPS_PER_HOUR, TARGET, Split, load  # noqa: E402
from ett_poc.evaluate import onset_metrics, regression_metrics, state_metrics  # noqa: E402
from ett_poc.features import ABLATION, CONDITION_A, CONDITION_B, columns_for, make_features, make_target  # noqa: E402
from ett_poc.models import fit_lgbm, predict_residual  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
PRED = ROOT / "reports" / "predictions"
PRED.mkdir(parents=True, exist_ok=True)
HORIZONS = [1, 6, 24]
MAIN_H = 6


def run_one(name, df, sph, split, h, cols, residual):
    X, groups = make_features(df, sph, h)
    y = make_target(df, h, sph)
    m = X[cols].notna().all(axis=1) & y.notna()
    Xa, ya = X.loc[m, cols], y[m]
    ot_now = X.loc[m, "ot_now"]
    tr, va, te = split.masks_by_target(Xa.index, pd.Timedelta(hours=h))
    target = (ya - ot_now) if residual else ya
    model = fit_lgbm(Xa[tr], target[tr], Xa[va], target[va])
    pred = predict_residual(model, Xa, ot_now, residual)
    return Xa, ya, ot_now, pred, model, (tr, va, te)


def metrics_row(name, h, label, part, ya, pred, ot_now, mask, thresholds, sph, extra=None):
    m = regression_metrics(ya[mask], pred[mask])
    row = {"dataset": name, "horizon_h": h, "model": label, "split": part,
           "MAE": round(m["MAE"], 3), "RMSE": round(m["RMSE"], 3), "n": m["n"]}
    for q, thr in thresholds.items():
        st = state_metrics(ya[mask], pred[mask], thr, 24 * sph)
        on = onset_metrics(ot_now[mask], ya[mask], pred[mask], thr, 24 * sph)
        row.update({
            f"q{q}_events": st["events"], f"q{q}_recall": round(st["recall"], 3),
            f"q{q}_onset_events": on["events_with_onset"], f"q{q}_onset_recall": round(on["event_recall"], 3),
            f"q{q}_notices/day": round(on["notices_per_day"], 2), f"q{q}_false_notices/day": round(on["false_notices_per_day"], 2),
            f"q{q}_lead_h": round(on["mean_lead_h"], 1),
        })
    if extra:
        row.update(extra)
    return row


def main() -> None:
    rows, abl_rows = [], []
    for name in ["ETTh1", "ETTh2"]:
        df = load(name)
        sph = STEPS_PER_HOUR[name]
        split = Split.default(df.index)
        tr_time, _, _ = split.masks(df.index)
        thresholds = {95: float(df.loc[tr_time, TARGET].quantile(0.95)), 90: float(df.loc[tr_time, TARGET].quantile(0.90))}

        for h in HORIZONS:
            _, groups = make_features(df, sph, h)
            for cond_name, cond in [("B", CONDITION_B), ("A", CONDITION_A)]:
                cols = columns_for(groups, cond)
                # 残差学習 vs 直接: 検証 MAE で選ぶ（テストは見ない）
                best = None
                for residual in (True, False):
                    Xa, ya, ot_now, pred, model, (tr, va, te) = run_one(name, df, sph, split, h, cols, residual)
                    va_mae = regression_metrics(ya[va], pred[va])["MAE"]
                    if best is None or va_mae < best[0]:
                        best = (va_mae, residual, Xa, ya, ot_now, pred, model, (tr, va, te))
                va_mae, residual, Xa, ya, ot_now, pred, model, (tr, va, te) = best
                label = f"M1_lgbm({cond_name})"
                extra = {"residual": residual, "best_iter": int(model.best_iteration)}
                for part, mask in [("valid", va), ("test", te)]:
                    rows.append(metrics_row(name, h, label, part, ya, pred, ot_now, mask, thresholds, sph, extra))
                out = pd.DataFrame({"ot_now": ot_now, "y": ya, "pred": pred, "split": pd.Series("train", index=Xa.index).mask(va, "valid").mask(te, "test")})
                out.to_csv(PRED / f"{name}_h{h}_{cond_name}.csv")
                imp = pd.Series(model.feature_importance("gain"), index=cols).sort_values(ascending=False)
                imp.to_csv(PRED / f"{name}_h{h}_{cond_name}_importance.csv")
                print(name, h, cond_name, "residual" if residual else "direct", "valid MAE", round(va_mae, 3), "iters", model.best_iteration)

            if h == MAIN_H:
                for abl_name, cond in ABLATION.items():
                    cols = columns_for(groups, cond)
                    Xa, ya, ot_now, pred, model, (tr, va, te) = run_one(name, df, sph, split, h, cols, residual=True)
                    for part, mask in [("valid", va), ("test", te)]:
                        m = regression_metrics(ya[mask], pred[mask])
                        abl_rows.append({"dataset": name, "features": abl_name, "n_features": len(cols), "split": part,
                                         "MAE": round(m["MAE"], 3), "RMSE": round(m["RMSE"], 3)})

    res = pd.DataFrame(rows)
    abl = pd.DataFrame(abl_rows)
    out = ["# LightGBM (auto-generated by scripts/run_lgbm.py)", "",
           "条件 B = 起点 s までの観測のみ、条件 A = B ＋ 対象時刻 T の負荷 6 種（課題ルール）。",
           "`residual` = 目的変数を OT[T]−OT[s] にしたか（検証 MAE で選択）。閾値・列の意味は results_baselines.md と同じ。", ""]
    for part in ["valid", "test"]:
        out.append(f"## {part}")
        out.append(res[res.split == part].drop(columns=["split"]).to_markdown(index=False))
        out.append("")
    out.append(f"## Ablation (h={MAIN_H}h, residual learning)")
    out.append(abl.to_markdown(index=False))
    (ROOT / "reports" / "results_lgbm.md").write_text("\n".join(out), encoding="utf-8")
    res.to_csv(ROOT / "reports" / "results_lgbm.csv", index=False)
    abl.to_csv(ROOT / "reports" / "results_ablation.csv", index=False)
    print("\n".join(out))


if __name__ == "__main__":
    main()
