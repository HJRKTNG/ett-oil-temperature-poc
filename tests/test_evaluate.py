"""警報指標の定義どおりに数えているかを、小さな人工系列で確かめる。

行は起点 s、y は OT[s+h]。閾値 10。
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from ett_poc.data import Split  # noqa: E402
from ett_poc.evaluate import onset_metrics, state_metrics  # noqa: E402

H = pd.Timedelta(hours=6)
THR = 10.0


def _series(values):
    idx = pd.date_range("2020-01-01", periods=len(values), freq="h")
    return pd.Series(values, index=idx, dtype=float)


def test_all_negative_gives_no_events_and_no_notices():
    y = _series([5] * 30)
    p = _series([5] * 30)
    s = _series([5] * 30)
    on = onset_metrics(s, y, p, THR, 24, H)
    assert on["events_with_onset"] == 0 and on["notices"] == 0
    assert np.isnan(on["event_recall"]) and np.isnan(on["mean_lead_h"])
    st = state_metrics(y, p, THR, 24)
    assert st["positives"] == 0 and st["events"] == 0


def test_one_notice_can_cover_two_events_and_lead_time_uses_first_hit():
    # 実際の超過: 行 10-12 と 行 15-16（2 区間）。起点 OT はずっと閾値以下。
    y = _series([5] * 30)
    y.iloc[10:13] = 12
    y.iloc[15:17] = 12
    s = _series([5] * 30)
    # 予測: 行 8-16 で連続して閾値超え → 通知は 1 通、2 区間とも検知
    p = _series([5] * 30)
    p.iloc[8:17] = 12
    on = onset_metrics(s, y, p, THR, 24, H)
    assert on["events_with_onset"] == 2 and on["events_detected"] == 2
    assert on["notices"] == 1 and on["false_notices"] == 0
    # 先行時間: 区間 1 は最初の的中行 = 行 10（起点）→ 超過時刻 = 行 10 + 6h → 6h。
    # 区間 2 も最初の的中行 = 行 15 → 6h。平均 6h。
    assert on["mean_lead_h"] == 6.0


def test_false_notice_is_a_notice_that_never_hits():
    y = _series([5] * 30)
    y.iloc[20:22] = 12
    s = _series([5] * 30)
    p = _series([5] * 30)
    p.iloc[3:5] = 12   # 誤通知（対象時刻は一度も超えない）
    p.iloc[20:22] = 12  # 真の通知
    on = onset_metrics(s, y, p, THR, 24, H)
    assert on["notices"] == 2 and on["false_notices"] == 1
    assert on["notice_precision"] == 0.5


def test_event_without_onset_rows_is_excluded():
    # 起点 OT が既に閾値超え（高温が続いている）の区間は onset の対象外
    y = _series([12] * 30)
    s = _series([12] * 30)
    p = _series([12] * 30)
    on = onset_metrics(s, y, p, THR, 24, H)
    assert on["events_with_onset"] == 0 and on["notices"] == 0


def test_split_requires_both_origin_and_target_in_period():
    idx = pd.date_range("2020-01-01", periods=24 * 400, freq="h")
    split = Split(train_end=pd.Timestamp("2020-06-01"), valid_end=pd.Timestamp("2020-09-01"))
    tr, va, te = split.masks_by_target(idx, H)
    origin = pd.Series(idx, index=idx)
    target = origin + H
    assert (target[tr] < split.train_end).all()
    assert (origin[va] >= split.train_end).all() and (target[va] < split.valid_end).all()
    assert (origin[te] >= split.valid_end).all()
    # 境界の先頭 h 時間はどこにも入らない
    gap = (~tr) & (~va) & (~te)
    assert gap.sum() == 2 * 6
