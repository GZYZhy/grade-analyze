from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd


@dataclass
class TrendResult:
    exam_order: List[str]
    total_scores: List[Optional[float]]
    subject_scores: Dict[str, List[Optional[float]]]


def build_student_trend(df: pd.DataFrame) -> TrendResult:
    if df.empty:
        return TrendResult([], [], {})
    df = df.copy()
    df["exam_name"] = df["exam_name"].astype(str)
    exams = df["exam_name"].drop_duplicates().tolist()
    total_scores = []
    for exam in exams:
        exam_df = df[df["exam_name"] == exam]
        val = exam_df["total_score"].dropna()
        if not val.empty:
            total_scores.append(val.iloc[0])
        else:
            total_row = exam_df[exam_df["subject"] == "总分"]["score"].dropna()
            total_scores.append(total_row.iloc[0] if not total_row.empty else None)

    subject_scores: Dict[str, List[Optional[float]]] = {}
    for subject in sorted(df["subject"].dropna().unique().tolist()):
        if subject == "总分":
            continue
        subject_scores[subject] = []
        for exam in exams:
            exam_df = df[(df["exam_name"] == exam) & (df["subject"] == subject)]
            val = exam_df["score"].dropna()
            subject_scores[subject].append(val.iloc[0] if not val.empty else None)

    return TrendResult(exams, total_scores, subject_scores)


def compute_improvement(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df = df.sort_values(by=["student_name", "exam_name"])
    df["score_delta"] = df.groupby(["student_name", "subject"])["score"].diff()
    df["rank_delta"] = df.groupby(["student_name", "subject"])["rank"].diff()
    df["total_delta"] = df.groupby(["student_name"])["total_score"].diff()
    df["class_rank_delta"] = df.groupby(["student_name"])["class_rank"].diff()
    df["grade_rank_delta"] = df.groupby(["student_name"])["grade_rank"].diff()
    return df
