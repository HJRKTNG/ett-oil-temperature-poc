"""評価指標。回帰の精度と、閾値超えの事前検知の両方を持つ。"""
from __future__ import annotations

import numpy as np
import pandas as pd


def regression_metrics(y_true: pd.Series, y_pred: pd.Series) -> dict[str, float]:
    err = (y_pred - y_true).dropna()
    return {
        "MAE": float(err.abs().mean()),
        "RMSE": float(np.sqrt((err**2).mean())),
        "n": int(err.shape[0]),
    }


def exceedance_metrics(
    y_true: pd.Series, y_pred: pd.Series, threshold: float, steps_per_day: int
) -> dict[str, float]:
    """「予測時点で、ホライズン先に閾値を超えていると警報を出したか」の精度。

    - y_true / y_pred は同じ時刻（= 予測対象時刻）で揃えたもの。
    - 誤警報/日 = 偽陽性の件数 ÷ 評価期間の日数。実運用で「無視されない警報」かを見る。
    """
    m = pd.concat([y_true.rename("y"), y_pred.rename("p")], axis=1).dropna()
    actual = m["y"] > threshold
    alarm = m["p"] > threshold
    tp = int((actual & alarm).sum())
    fp = int((~actual & alarm).sum())
    fn = int((actual & ~alarm).sum())
    days = max(len(m) / steps_per_day, 1e-9)
    return {
        "threshold": float(threshold),
        "positives": int(actual.sum()),
        "precision": tp / (tp + fp) if tp + fp else float("nan"),
        "recall": tp / (tp + fn) if tp + fn else float("nan"),
        "false_alarms_per_day": fp / days,
        "n": int(len(m)),
    }


def improvement_over(baseline_mae: float, model_mae: float) -> float:
    """持続予測などの基準に対する MAE 改善率（正なら改善）。"""
    return 1.0 - model_mae / baseline_mae if baseline_mae else float("nan")
