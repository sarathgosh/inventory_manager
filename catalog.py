from __future__ import annotations

from typing import Any

import pandas as pd


MISSING_OPTION = "— Not set —"


def value_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def database_options(
    frame: pd.DataFrame,
    column: str,
    *,
    include_missing: bool = True,
) -> list[str]:
    if frame.empty or column not in frame.columns:
        return [MISSING_OPTION] if include_missing else []
    values = sorted(
        {value_text(value) for value in frame[column].tolist() if value_text(value)},
        key=str.casefold,
    )
    has_missing = frame[column].map(value_text).eq("").any()
    if include_missing and (has_missing or not values):
        values.append(MISSING_OPTION)
    return values


def filter_by_option(frame: pd.DataFrame, column: str, selected: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    text_series = frame[column].map(value_text)
    if selected == MISSING_OPTION:
        return frame.loc[text_series.eq("")].copy()
    return frame.loc[text_series.eq(value_text(selected))].copy()


def option_to_value(selected: str | None) -> str | None:
    if selected is None or selected == MISSING_OPTION:
        return None
    return value_text(selected) or None


def numeric_options(frame: pd.DataFrame, column: str) -> list[str]:
    if frame.empty or column not in frame.columns:
        return []
    numeric = pd.to_numeric(frame[column], errors="coerce").dropna().unique().tolist()
    return [f"{float(value):g}" for value in sorted(float(value) for value in numeric)]


def latest_record(frame: pd.DataFrame) -> pd.Series | None:
    if frame.empty:
        return None
    if "id" in frame.columns:
        return frame.sort_values("id").iloc[-1]
    return frame.iloc[-1]
