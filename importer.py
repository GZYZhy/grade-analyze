from __future__ import annotations

from typing import Dict, List, Optional, Tuple

import pandas as pd

FORCED_RAW_SUBJECTS = {"语文", "数学", "英语", "物理", "历史"}


def guess_subject_pairs(columns: List[str]) -> List[Tuple[str, Optional[str], Optional[str], Optional[str]]]:
    pairs = []
    used = set()
    normalized = {str(col).replace(" ", ""): col for col in columns}
    if any(key in normalized for key in ["总分赋分", "总分原始", "总分"]):
        total_key = "总分赋分" if "总分赋分" in normalized else "总分原始" if "总分原始" in normalized else "总分"
        total_col = normalized[total_key]
        total_rank = normalized.get("总分名次")
        pairs.append(("总分", total_col, None, total_rank))
        used.add(total_col)
        if total_rank:
            used.add(total_rank)
    preferred_subjects = ["语文", "数学", "英语", "物理", "化学", "生物"]

    for subject in preferred_subjects:
        score_col = None
        for cand in [f"{subject}赋分", f"{subject}得分", subject, f"{subject}原始"]:
            if cand in normalized:
                score_col = normalized[cand]
                break
        if score_col:
            rank_col = normalized.get(f"{subject}名次")
            raw_col = normalized.get(f"{subject}原始")
            pairs.append((subject, score_col, raw_col, rank_col))
            used.add(score_col)
            if raw_col:
                used.add(raw_col)
            if rank_col:
                used.add(rank_col)
    for col in columns:
        if col in used:
            continue
        normalized_col = str(col).replace(" ", "")
        if normalized_col.endswith("得分"):
            subject = normalized_col.replace("得分", "")
            if subject and subject != "总分":
                rank_col = normalized.get(f"{subject}名次")
                raw_col = normalized.get(f"{subject}原始")
                pairs.append((subject or col, col, raw_col, rank_col))
                used.add(col)
                if raw_col:
                    used.add(raw_col)
                if rank_col:
                    used.add(rank_col)
            continue

        if normalized_col.endswith("赋分"):
            subject = normalized_col.replace("赋分", "")
            if subject and subject != "总分":
                rank_col = normalized.get(f"{subject}名次")
                raw_col = normalized.get(f"{subject}原始")
                pairs.append((subject or col, col, raw_col, rank_col))
                used.add(col)
                if raw_col:
                    used.add(raw_col)
                if rank_col:
                    used.add(rank_col)
            continue

        rank_candidate = f"{normalized_col}名次"
        if rank_candidate in normalized and normalized_col not in ("总分", "总分赋分", "总分原始"):
            raw_col = normalized.get(f"{normalized_col}原始")
            pairs.append((col, col, raw_col, normalized[rank_candidate]))
            used.add(col)
            if raw_col:
                used.add(raw_col)
            used.add(normalized[rank_candidate])
    return pairs


def _safe_value(value):
    if pd.isna(value):
        return None
    return value


def parse_input_sheet(
    df: pd.DataFrame,
    mapping: Dict[str, str],
    subject_map: List[Tuple[str, str, Optional[str], Optional[str]]],
) -> List[Dict[str, object]]:
    records: List[Dict[str, object]] = []
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    student_col = mapping.get("student_name") or ""
    if student_col and student_col in df.columns:
        df[student_col] = df[student_col].ffill()

    for _, row in df.iterrows():
        student_name = _safe_value(row.get(mapping.get("student_name", "")))
        exam_name = _safe_value(row.get(mapping.get("exam_name", "")))
        if not student_name or not exam_name:
            continue

        base_info = {
            "student_name": str(student_name).strip(),
            "exam_name": str(exam_name).strip(),
            "total_score": _safe_value(row.get(mapping.get("total_score", ""))),
            "total_raw": _safe_value(row.get(mapping.get("total_raw", ""))),
            "grade_rank": _safe_value(row.get(mapping.get("grade_rank", ""))),
            "class_rank": _safe_value(row.get(mapping.get("class_rank", ""))),
        }

        for subject, score_col, raw_col, rank_col in subject_map:
            score_val = _safe_value(row.get(score_col)) if score_col else None
            raw_val = _safe_value(row.get(raw_col)) if raw_col else None
            if subject in FORCED_RAW_SUBJECTS:
                raw_val = raw_val if raw_val is not None else score_val
                score_val = None
            rank_val = _safe_value(row.get(rank_col)) if rank_col else None
            if score_val is None and raw_val is None and rank_val is None:
                continue
            records.append(
                {
                    **base_info,
                    "subject": subject,
                    "score": score_val,
                    "score_raw": raw_val,
                    "rank": rank_val,
                }
            )

    return records


def parse_usual_sheet(
    df: pd.DataFrame,
    exam_name: str,
    mapping: Dict[str, str],
    subject_map: List[Tuple[str, str, Optional[str], Optional[str]]],
) -> List[Dict[str, object]]:
    records: List[Dict[str, object]] = []
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    student_col = mapping.get("student_name") or ""
    if student_col and student_col in df.columns:
        df[student_col] = df[student_col].ffill()

    for _, row in df.iterrows():
        student_name = _safe_value(row.get(mapping.get("student_name", "")))
        if not student_name:
            continue

        base_info = {
            "student_name": str(student_name).strip(),
            "exam_name": str(exam_name).strip(),
            "total_score": _safe_value(row.get(mapping.get("total_score", ""))),
            "total_raw": _safe_value(row.get(mapping.get("total_raw", ""))),
            "grade_rank": _safe_value(row.get(mapping.get("grade_rank", ""))),
            "class_rank": _safe_value(row.get(mapping.get("class_rank", ""))),
        }

        for subject, score_col, raw_col, rank_col in subject_map:
            score_val = _safe_value(row.get(score_col)) if score_col else None
            raw_val = _safe_value(row.get(raw_col)) if raw_col else None
            if subject in FORCED_RAW_SUBJECTS:
                raw_val = raw_val if raw_val is not None else score_val
                score_val = None
            rank_val = _safe_value(row.get(rank_col)) if rank_col else None
            if score_val is None and raw_val is None and rank_val is None:
                continue
            records.append(
                {
                    **base_info,
                    "subject": subject,
                    "score": score_val,
                    "score_raw": raw_val,
                    "rank": rank_val,
                }
            )

    return records
