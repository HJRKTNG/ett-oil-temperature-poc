"""評価指標。

- 回帰: MAE / RMSE。
- 高温状態の予測: 対象時刻 T に閾値超えかを当てる（時点単位）。
- 高温になる前の検知（onset）: 起点 s では閾値以下、対象 T で超え。
  連続した超過は 1 イベントにまとめ、連続した警報は 1 通知に集約して数える。
"""
from __future__ import annotations

import numpy as np
import pandas as pd


def regression_metrics(y_true: pd.Series, y_pred: pd.Series) -> dict[str, float]:
    err = (y_pred - y_true).dropna()
    return {"MAE": float(err.abs().mean()), "RMSE": float(np.sqrt((err**2).mean())), "n": int(err.shape[0])}


def _runs(flag: pd.Series) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    """True が連続する区間の [開始, 終了] 一覧。"""
    f = flag.astype(bool)
    starts = f & ~f.shift(1, fill_value=False)
    ends = f & ~f.shift(-1, fill_value=False)
    return list(zip(f.index[starts], f.index[ends]))


def state_metrics(y_true: pd.Series, y_pred: pd.Series, threshold: float, steps_per_day: int) -> dict[str, float]:
    """時点単位: 対象時刻 T に閾値を超えるかの警報精度。"""
    m = pd.concat([y_true.rename("y"), y_pred.rename("p")], axis=1).dropna()
    actual, alarm = m["y"] > threshold, m["p"] > threshold
    tp, fp, fn = int((actual & alarm).sum()), int((~actual & alarm).sum()), int((actual & ~alarm).sum())
    days = max(len(m) / steps_per_day, 1e-9)
    return {
        "positives": int(actual.sum()),
        "events": len(_runs(actual)),
        "precision": tp / (tp + fp) if tp + fp else float("nan"),
        "recall": tp / (tp + fn) if tp + fn else float("nan"),
        "false_alarm_points_per_day": fp / days,
        "n": int(len(m)),
    }


def onset_metrics(
    ot_origin: pd.Series, y_true: pd.Series, y_pred: pd.Series, threshold: float, steps_per_day: int,
    horizon: pd.Timedelta = pd.Timedelta(0),
) -> dict[str, float]:
    """高温になる前の検知。

    - 対象: 起点 s で OT[s] <= 閾値、かつ対象 T で OT[T] > 閾値 の行（onset 行）。
    - イベント: 実際の超過区間（連続超過を 1 つ）のうち、その開始前 h 以内に onset 行を持つもの。
      イベント検知 = そのイベントの onset 行のどれかで警報（予測 > 閾値）が出た。
    - 通知: 起点 s で閾値以下のときに出た警報を、連続分を 1 つにまとめて数える。
      通知のうち、対象 T が実際に超えていないものを誤通知とする。
    """
    m = pd.concat([ot_origin.rename("s"), y_true.rename("y"), y_pred.rename("p")], axis=1).dropna()
    below_now = m["s"] <= threshold
    actual, alarm = m["y"] > threshold, m["p"] > threshold
    onset_rows = below_now & actual
    alarm_rows = below_now & alarm
    days = max(len(m) / steps_per_day, 1e-9)

    # 実際の超過イベント（連続超過を 1 つに）。onset 行を含むイベントだけを対象にする。
    events = _runs(actual)
    detected = 0
    lead_times: list[float] = []
    n_events_with_onset = 0
    for start, end in events:
        rows = m.loc[start:end]
        on = onset_rows.loc[start:end]
        if not on.any():
            continue  # 期間の先頭から超えていた等、起点が閾値以下の行が無い
        n_events_with_onset += 1
        hit = (on & alarm_rows.loc[start:end])
        if hit.any():
            detected += 1
            first = hit[hit].index[0]
            # 行は起点 s で並ぶ。イベント開始行 start は「OT[start+h] が最初に閾値を超える」起点。
            # 初回警報の起点 first から、実際の超過時刻 start+h までが先行時間。
            lead_times.append(float((start + horizon - first) / pd.Timedelta(hours=1)))
    notices = _runs(alarm_rows)
    false_notices = sum(1 for a, b in notices if not actual.loc[a:b].any())
    return {
        "onset_points": int(onset_rows.sum()),
        "events_with_onset": n_events_with_onset,
        "events_detected": detected,
        "event_recall": detected / n_events_with_onset if n_events_with_onset else float("nan"),
        "notices_per_day": len(notices) / days,
        "false_notices_per_day": false_notices / days,
        "notice_precision": 1 - false_notices / len(notices) if notices else float("nan"),
        "mean_lead_h": float(np.mean(lead_times)) if lead_times else float("nan"),
    }


def improvement_over(baseline_mae: float, model_mae: float) -> float:
    """基準に対する MAE 改善率（正なら改善）。"""
    return 1.0 - model_mae / baseline_mae if baseline_mae else float("nan")
