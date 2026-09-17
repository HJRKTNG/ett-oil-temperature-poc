"""モデルの共通部品。学習期間で fit し、起点ごとに予測する。"""
from __future__ import annotations

import lightgbm as lgb
import numpy as np
import pandas as pd
from sklearn.linear_model import Ridge
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

LGB_PARAMS = dict(
    objective="regression_l1",  # MAE を主指標にするので L1
    learning_rate=0.03,
    num_leaves=31,
    min_data_in_leaf=50,
    feature_fraction=0.8,
    bagging_fraction=0.8,
    bagging_freq=1,
    lambda_l2=1.0,
    verbose=-1,
    seed=0,
)


def fit_ridge(X: pd.DataFrame, y: pd.Series, alpha: float = 1.0):
    model = make_pipeline(StandardScaler(), Ridge(alpha=alpha))
    model.fit(X, y)
    return model


def fit_lgbm(
    X_tr: pd.DataFrame, y_tr: pd.Series, X_va: pd.DataFrame, y_va: pd.Series, params: dict | None = None
) -> lgb.Booster:
    """検証期間で早期終了。テストは見ない。"""
    p = dict(LGB_PARAMS)
    if params:
        p.update(params)
    dtr = lgb.Dataset(X_tr, y_tr)
    dva = lgb.Dataset(X_va, y_va, reference=dtr)
    return lgb.train(
        p, dtr, num_boost_round=3000, valid_sets=[dva],
        callbacks=[lgb.early_stopping(100, verbose=False)],
    )


def predict_residual(model, X: pd.DataFrame, ot_now: pd.Series, residual: bool) -> pd.Series:
    """残差学習（目的変数 = OT[T] − OT[s]）なら OT[s] を足し戻す。"""
    if hasattr(model, "best_iteration"):
        pred = model.predict(X, num_iteration=model.best_iteration)
    else:
        pred = model.predict(X)
    pred = pd.Series(np.asarray(pred), index=X.index)
    return pred + ot_now if residual else pred
