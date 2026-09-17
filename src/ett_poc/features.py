"""特徴量の作成。

記法: 起点 s（予測を行う時刻）、対象 T = s + h。目的変数は OT[T]。行は s で揃える。

方針:
- 条件 B（運用条件）: 起点 s までの観測だけを使う。ラグ・移動統計は shift 後に計算。
- 条件 A（課題条件）: 条件 B に **対象時刻 T の負荷 6 種** を足す（課題ルール
  「t=T の油温を予測する際には t=T での特徴量を使用してよい」の解釈）。
- 特徴量群は名前で切り分けられるようにし、アブレーションに使う。
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from .data import LOAD_COLS, TARGET

LAGS_H = [1, 2, 3, 6, 12, 24, 48, 168]
ROLL_H = [6, 24, 168]

#: 特徴量群（アブレーション用）
GROUP_OT_HISTORY = "ot_history"
GROUP_CALENDAR = "calendar"
GROUP_LOAD_ORIGIN = "load_origin"
GROUP_LOAD_TARGET = "load_target"


def make_features(df: pd.DataFrame, steps_per_hour: int, horizon_h: int) -> tuple[pd.DataFrame, dict[str, list[str]]]:
    """起点 s の行に対する特徴量と、列名の群分け。"""
    sph = steps_per_hour
    ot = df[TARGET]
    X = pd.DataFrame(index=df.index)
    groups: dict[str, list[str]] = {GROUP_OT_HISTORY: [], GROUP_CALENDAR: [], GROUP_LOAD_ORIGIN: [], GROUP_LOAD_TARGET: []}

    # --- OT の履歴（起点 s まで）
    X["ot_now"] = ot
    for lag in LAGS_H:
        X[f"ot_lag{lag}h"] = ot.shift(lag * sph)
    X["ot_d1h"] = ot - ot.shift(1 * sph)
    X["ot_d24h"] = ot - ot.shift(24 * sph)
    for w in ROLL_H:
        r = ot.rolling(w * sph, min_periods=w * sph)
        X[f"ot_mean{w}h"] = r.mean()
        X[f"ot_std{w}h"] = r.std()
        X[f"ot_max{w}h"] = r.max()
        X[f"ot_min{w}h"] = r.min()
    groups[GROUP_OT_HISTORY] = list(X.columns)

    # --- カレンダー（対象時刻 T の周期。時刻は既知なので未来の観測ではない）
    tidx = X.index + pd.Timedelta(hours=horizon_h)
    hour = tidx.hour + tidx.minute / 60.0
    doy = tidx.dayofyear
    X["t_hour_sin"] = np.sin(2 * np.pi * hour / 24)
    X["t_hour_cos"] = np.cos(2 * np.pi * hour / 24)
    X["t_doy_sin"] = np.sin(2 * np.pi * doy / 365.25)
    X["t_doy_cos"] = np.cos(2 * np.pi * doy / 365.25)
    X["t_dow"] = tidx.dayofweek
    groups[GROUP_CALENDAR] = ["t_hour_sin", "t_hour_cos", "t_doy_sin", "t_doy_cos", "t_dow"]

    # --- 負荷（起点 s まで）
    cols = []
    for c in LOAD_COLS:
        X[f"{c}_now"] = df[c]
        X[f"{c}_d1h"] = df[c] - df[c].shift(1 * sph)
        X[f"{c}_mean24h"] = df[c].rolling(24 * sph, min_periods=24 * sph).mean()
        cols += [f"{c}_now", f"{c}_d1h", f"{c}_mean24h"]
    groups[GROUP_LOAD_ORIGIN] = cols

    # --- 負荷（対象時刻 T。条件 A だけで使う）
    cols = []
    for c in LOAD_COLS:
        X[f"{c}_at_T"] = df[c].shift(-horizon_h * sph)
        cols.append(f"{c}_at_T")
    groups[GROUP_LOAD_TARGET] = cols
    return X, groups


def make_target(df: pd.DataFrame, horizon_h: int, steps_per_hour: int) -> pd.Series:
    """起点 s の行に対する目的変数 OT[T]。"""
    return df[TARGET].shift(-horizon_h * steps_per_hour).rename(f"y_h{horizon_h}")


def columns_for(groups: dict[str, list[str]], names: list[str]) -> list[str]:
    cols: list[str] = []
    for n in names:
        cols += groups[n]
    return cols


CONDITION_B = [GROUP_OT_HISTORY, GROUP_CALENDAR, GROUP_LOAD_ORIGIN]
CONDITION_A = CONDITION_B + [GROUP_LOAD_TARGET]

#: アブレーションの段階（主ホライズンで比較）
ABLATION = {
    "1_ot_only": [GROUP_OT_HISTORY],
    "2_+calendar": [GROUP_OT_HISTORY, GROUP_CALENDAR],
    "3_+load_origin (B)": CONDITION_B,
    "4_+load_target (A)": CONDITION_A,
}


def check_no_leakage(df: pd.DataFrame, steps_per_hour: int, horizon_h: int) -> None:
    """条件 B の特徴量が未来を見ていないことを、末尾を切っても値が変わらないことで確かめる。"""
    X_full, g = make_features(df, steps_per_hour, horizon_h)
    cut = df.iloc[: len(df) - 200 * steps_per_hour]
    X_cut, _ = make_features(cut, steps_per_hour, horizon_h)
    cols = columns_for(g, CONDITION_B)
    common = X_cut.index
    diff = (X_full.loc[common, cols] - X_cut[cols]).abs().max().max()
    if not (np.isnan(diff) or diff == 0):
        raise AssertionError(f"condition-B features depend on the future: max diff {diff}")
