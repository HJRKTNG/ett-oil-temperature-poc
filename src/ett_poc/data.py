"""データの読込と時系列分割。

ETT（Electricity Transformer Temperature）は zhouhaoyi/ETDataset で公開されている
ベンチマーク。列は 6 種の負荷（HUFL/HULL/MUFL/MULL/LUFL/LULL）と油温 OT。
h 系列は 1 時間刻み、m 系列は 15 分刻み。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).resolve().parents[2] / "data"
LOAD_COLS = ["HUFL", "HULL", "MUFL", "MULL", "LUFL", "LULL"]
TARGET = "OT"

#: 各系列の 1 時間あたりのステップ数（h = 1 時間刻み、m = 15 分刻み）
STEPS_PER_HOUR = {"ETTh1": 1, "ETTh2": 1, "ETTm1": 4, "ETTm2": 4}


def load(name: str) -> pd.DataFrame:
    """`data/<name>.csv` を日時インデックスで読む。欠損・重複・飛びが無いことを確認する。"""
    df = pd.read_csv(DATA_DIR / f"{name}.csv", parse_dates=["date"]).set_index("date").sort_index()
    if df.index.duplicated().any():
        raise ValueError(f"{name}: duplicated timestamps")
    if df.isna().any().any():
        raise ValueError(f"{name}: missing values")
    step = df.index.to_series().diff().dropna()
    if (step != step.mode()[0]).any():
        raise ValueError(f"{name}: irregular timestamps")
    return df


@dataclass(frozen=True)
class Split:
    """時系列の 3 分割（暦で固定）。

    学習 = 最初の 12 か月、検証 = 次の 4 か月、テスト = 残り全部。
    区切りは **対象時刻 T** で行う（`masks_by_target`）。起点 s で切ると、学習末尾の
    ラベル OT[s+h] が検証期間に入り込むため。
    """

    train_end: pd.Timestamp
    valid_end: pd.Timestamp

    @classmethod
    def default(cls, index: pd.DatetimeIndex) -> "Split":
        start = index[0]
        return cls(
            train_end=start + pd.DateOffset(months=12),
            valid_end=start + pd.DateOffset(months=16),
        )

    def masks_by_target(
        self, origin_index: pd.DatetimeIndex, horizon: pd.Timedelta
    ) -> tuple[pd.Series, pd.Series, pd.Series]:
        """起点 s の行に対し、対象時刻 T = s + h がどの期間に入るかで分ける。"""
        target = pd.Series(origin_index + horizon, index=origin_index)
        train = target < self.train_end
        valid = (target >= self.train_end) & (target < self.valid_end)
        test = target >= self.valid_end
        return train, valid, test

    def masks(self, index: pd.DatetimeIndex) -> tuple[pd.Series, pd.Series, pd.Series]:
        """時刻そのもので分ける（EDA 用）。"""
        return self.masks_by_target(index, pd.Timedelta(0))
