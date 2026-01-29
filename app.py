from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import bcrypt
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from streamlit_autorefresh import st_autorefresh

from analysis import build_student_trend
from data_access import DataStore
from importer import guess_subject_pairs, parse_input_sheet, parse_usual_sheet
from reporting import export_student_reports, export_stats_excel

FORCED_RAW_SUBJECTS = {"语文", "数学", "英语", "物理", "历史"}

APP_NAME = "成绩分析工具"
APP_VERSION = "1.0"
APP_AUTHOR = "公子语"
APP_COPYRIGHT = "(c) 2026 ZhangWeb"

EXPORT_DIR = Path(__file__).resolve().parent / "exports"


def init_page() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="favicon.ico", layout="wide")
    st.markdown(
        """
        <style>
        .block-container { padding-top: 4.2rem; padding-bottom: 1.5rem; }
        div[data-testid="stVerticalBlock"] > div { gap: 0.6rem; }
        .app-header { margin-top: 0.2rem; margin-bottom: 0rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div class='app-header' style='font-size:16px; font-weight:600'>{APP_NAME}"
        f"<span style='font-weight:400; color:#6b7280; margin-left:12px;'>"
        f"版本 {APP_VERSION} · 作者 {APP_AUTHOR} · 版权 {APP_COPYRIGHT}"
        f"</span></div>",
        unsafe_allow_html=True,
    )


def to_dataframe(rows: List[object]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame([dict(row) for row in rows])


def pick_index(options: List[str], preferred: str, fallback: int = 0) -> int:
    try:
        return options.index(preferred)
    except ValueError:
        pass
    normalized = [o.replace(" ", "") for o in options]
    preferred_norm = preferred.replace(" ", "")
    try:
        return normalized.index(preferred_norm)
    except ValueError:
        return fallback


def pick_index_by_pos(options: List[str], pos: int, fallback: int = 0) -> int:
    if 0 <= pos < len(options):
        return pos
    return fallback


def load_subject_aliases(store: DataStore) -> Dict[str, str]:
    raw = store.get_setting("subject_aliases")
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        return {str(k).strip(): str(v).strip() for k, v in data.items() if str(k).strip()}
    except json.JSONDecodeError:
        return {}


def apply_subject_alias(subject: str, aliases: Dict[str, str]) -> str:
    if not subject:
        return subject
    return aliases.get(subject, subject)


def normalize_forced_subjects(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    mask = df["subject"].isin(FORCED_RAW_SUBJECTS)
    if "score_raw" in df.columns:
        df.loc[mask, "score"] = df.loc[mask, "score_raw"].fillna(df.loc[mask, "score"])
    return df


def check_password(store: DataStore) -> None:
    enabled = store.get_setting("password_enabled") == "1"
    if not enabled:
        return
    if st.session_state.get("password_ok"):
        return

    st.info("此软件已开启打开密码保护，请输入密码。")
    pwd = st.text_input("密码", type="password")
    submit = st.button("验证", type="primary")
    if submit:
        saved = store.get_setting("password_hash")
        if saved and bcrypt.checkpw(pwd.encode("utf-8"), saved.encode("utf-8")):
            st.session_state["password_ok"] = True
            st.success("验证通过。")
            st.rerun()
        else:
            st.error("密码错误。")
    st.stop()


def ensure_class(store: DataStore) -> int:
    classes = store.get_classes()
    if not classes:
        default = store.add_class("默认班级")
        return default.id
    return classes[0].id


def select_class(store: DataStore) -> int:
    classes = store.get_classes()
    if not classes:
        return ensure_class(store)
    labels = {f"{c.name}": c.id for c in classes}
    selected = st.sidebar.selectbox("切换存档（班级）", list(labels.keys()))
    return labels[selected]


def import_input_sheet(store: DataStore, class_id: int) -> None:
    st.subheader("输入成绩单导入")
    aliases = load_subject_aliases(store)
    upload = st.file_uploader("上传输入成绩单（Excel）", type=["xlsx"])
    if not upload:
        return
    xls = pd.ExcelFile(upload)
    sheet_name = st.selectbox("选择工作表", xls.sheet_names, index=0)
    df = pd.read_excel(upload, sheet_name=sheet_name)
    st.write("预览", df.head(10))

    columns = [str(c).strip() for c in df.columns]
    mapping = {
        "exam_name": st.selectbox(
            "考试名称列",
            columns,
            index=pick_index_by_pos(columns, 0, pick_index(columns, "考试名称")),
        ),
        "student_name": st.selectbox(
            "学生姓名列",
            columns,
            index=pick_index_by_pos(columns, 1, pick_index(columns, "学生姓名")),
        ),
        "total_score": st.selectbox(
            "总分赋分列（可选）",
            [""] + columns,
            index=pick_index_by_pos([""] + columns, 3, pick_index([""] + columns, "总分赋分")),
        ),
        "total_raw": st.selectbox(
            "总分原始列（可选）",
            [""] + columns,
            index=pick_index_by_pos([""] + columns, 4, pick_index([""] + columns, "总分原始")),
        ),
        "grade_rank": st.selectbox(
            "年级排名列（可选）",
            [""] + columns,
            index=pick_index_by_pos([""] + columns, 5, pick_index([""] + columns, "总分名次")),
        ),
        "class_rank": st.selectbox(
            "班级排名列（可选）",
            [""] + columns,
            index=pick_index_by_pos([""] + columns, 6, pick_index([""] + columns, "班级名次")),
        ),
    }

    default_subject_map = []
    if len(columns) >= 20:
        def col_at(idx: int) -> Optional[str]:
            return columns[idx] if 0 <= idx < len(columns) else None

        default_subject_map = [
            ("语文", col_at(6), None, col_at(7)),
            ("数学", col_at(8), None, col_at(9)),
            ("英语", col_at(10), None, col_at(11)),
            ("物理", col_at(12), None, col_at(13)),
            ("总分", col_at(2), None, col_at(4)),
            ("化学", col_at(14), col_at(15), col_at(16)),
            ("生物", col_at(17), col_at(18), col_at(19)),
        ]
    guessed = default_subject_map or guess_subject_pairs(columns)
    subject_df = pd.DataFrame(guessed, columns=["科目", "赋分列", "原始列", "名次列"])
    st.markdown("**科目映射（可编辑）**")
    edited = st.data_editor(subject_df, num_rows="dynamic")
    subject_map = []
    for _, row in edited.iterrows():
        subject = str(row.get("科目", "")).strip()
        score_col = str(row.get("赋分列", "")).strip() or None
        raw_col = str(row.get("原始列", "")).strip() or None
        rank_col = str(row.get("名次列", "")).strip() or None
        if subject and (score_col or raw_col):
            subject_map.append((apply_subject_alias(subject, aliases), score_col, raw_col, rank_col))

    if st.button("导入", type="primary"):
        records = parse_input_sheet(df, mapping, subject_map)
        persist_records(store, class_id, records)
        st.success(f"导入完成，共 {len(records)} 条科目记录。")


def import_usual_sheet(store: DataStore, class_id: int) -> None:
    st.subheader("通常成绩单导入")
    aliases = load_subject_aliases(store)
    upload = st.file_uploader("上传通常成绩单（Excel）", type=["xlsx"], key="usual")
    if not upload:
        return
    xls = pd.ExcelFile(upload)
    sheet_name = st.selectbox("选择工作表", xls.sheet_names, index=0, key="usual_sheet")
    use_second_row = st.checkbox("第一行是表头说明，映射从第二行开始", value=False)
    header_row = 1 if use_second_row else 0
    df = pd.read_excel(upload, sheet_name=sheet_name, header=header_row)
    st.write("预览", df.head(10))

    exam_name = st.text_input("考试名称", value=sheet_name)
    columns = [str(c).strip() for c in df.columns]
    mapping = {
        "student_name": st.selectbox(
            "学生姓名列",
            columns,
            index=pick_index(columns, "学生姓名"),
            key="u_student",
        ),
        "total_score": st.selectbox(
            "总分赋分列（可选）",
            [""] + columns,
            index=pick_index([""] + columns, "总分赋分"),
            key="u_total",
        ),
        "total_raw": st.selectbox(
            "总分原始列（可选）",
            [""] + columns,
            index=pick_index([""] + columns, "总分原始"),
            key="u_raw",
        ),
        "grade_rank": st.selectbox(
            "年级排名列（可选）",
            [""] + columns,
            index=pick_index([""] + columns, "总分名次"),
            key="u_grade",
        ),
        "class_rank": st.selectbox(
            "班级排名列（可选）",
            [""] + columns,
            index=pick_index([""] + columns, "班级名次"),
            key="u_class",
        ),
    }

    guessed = guess_subject_pairs(columns)
    subject_df = pd.DataFrame(guessed, columns=["科目", "赋分列", "原始列", "名次列"])
    st.markdown("**科目映射（可编辑）**")
    edited = st.data_editor(subject_df, num_rows="dynamic", key="usual_subjects")
    subject_map = []
    for _, row in edited.iterrows():
        subject = str(row.get("科目", "")).strip()
        score_col = str(row.get("赋分列", "")).strip() or None
        raw_col = str(row.get("原始列", "")).strip() or None
        rank_col = str(row.get("名次列", "")).strip() or None
        if subject and (score_col or raw_col):
            subject_map.append((apply_subject_alias(subject, aliases), score_col, raw_col, rank_col))

    clear_before = st.checkbox("导入前清空当前班级数据（用通常成绩单作为全量数据）")

    if st.button("导入", type="primary", key="usual_import"):
        if clear_before:
            store.clear_class_data(class_id)
        records = parse_usual_sheet(df, exam_name, mapping, subject_map)
        persist_records(store, class_id, records)
        st.success(f"导入完成，共 {len(records)} 条科目记录。")


def persist_records(store: DataStore, class_id: int, records: List[Dict[str, object]]) -> None:
    if not records:
        return
    student_map: Dict[str, int] = {}
    exam_map: Dict[str, int] = {}

    for r in records:
        student = store.upsert_student(r["student_name"], class_id)
        exam = store.upsert_exam(r["exam_name"], class_id)
        student_map[r["student_name"]] = student.id
        exam_map[r["exam_name"]] = exam.id

    insert_records = []
    for r in records:
        if r.get("subject") in FORCED_RAW_SUBJECTS and r.get("score") is None:
            if r.get("score_raw") is not None:
                r["score"] = r.get("score_raw")
        insert_records.append(
            {
                "student_id": student_map[r["student_name"]],
                "exam_id": exam_map[r["exam_name"]],
                "subject": r["subject"],
                "score": r.get("score"),
                "score_raw": r.get("score_raw"),
                "rank": r.get("rank"),
                "total_score": r.get("total_score"),
                "total_raw": r.get("total_raw"),
                "grade_rank": r.get("grade_rank"),
                "class_rank": r.get("class_rank"),
            }
        )
    store.insert_scores(insert_records)


def render_trend_page(store: DataStore, class_id: int) -> None:
    students = store.get_students(class_id)
    if not students:
        st.info("暂无学生数据，请先导入成绩单。")
        return

    aliases = load_subject_aliases(store)

    names = [s.name for s in students]
    if "slide_index" not in st.session_state:
        st.session_state["slide_index"] = 0

    if "trend_autoplay" not in st.session_state:
        st.session_state["trend_autoplay"] = False
    if "trend_interval" not in st.session_state:
        st.session_state["trend_interval"] = 6
    if "trend_metric" not in st.session_state:
        st.session_state["trend_metric"] = "赋分"
    if "trend_total_rank_source" not in st.session_state:
        st.session_state["trend_total_rank_source"] = "年级排名"

    if st.session_state.get("trend_autoplay"):
        st_autorefresh(interval=st.session_state.get("trend_interval", 6) * 1000, key="autoplay")
        st.session_state["slide_index"] = (st.session_state["slide_index"] + 1) % len(names)
    else:
        st_autorefresh(interval=24 * 60 * 60 * 1000, key="autoplay")

    selected = st.selectbox("选择学生", names, index=st.session_state["slide_index"])

    components.html(
        """
        <script>
        const key = "grade_analyze_scroll";
        const saved = sessionStorage.getItem(key);
        if (saved) {
            window.scrollTo(0, parseInt(saved, 10));
        }
        window.addEventListener("scroll", () => {
            sessionStorage.setItem(key, window.scrollY.toString());
        });
        </script>
        """,
        height=0,
    )

    student = next(s for s in students if s.name == selected)
    rows = store.get_scores_by_student(student.id)
    df = to_dataframe(rows)
    if df.empty:
        st.warning("该学生暂无成绩记录。")
        return

    df["subject"] = df["subject"].apply(lambda s: apply_subject_alias(s, aliases))
    df = normalize_forced_subjects(df)
    trend = build_student_trend(df)

    tab_chart, tab_table, tab_settings = st.tabs(["图表", "表格", "设置"])

    df_display = df.copy()
    df_display = df_display.sort_values(by=["exam_name", "subject"]).reset_index(drop=True)

    total_from_subject = (
        df_display[df_display["subject"] == "总分"]
        .groupby("exam_name")["score"]
        .first()
    )
    rank_from_subject = (
        df_display[df_display["subject"] == "总分"]
        .groupby("exam_name")["rank"]
        .first()
    )

    total_series = df_display.groupby("exam_name")["total_score"].first()
    grade_rank_series = df_display.groupby("exam_name")["grade_rank"].first()
    class_rank_series = df_display.groupby("exam_name")["class_rank"].first()

    total_series = total_series.fillna(total_from_subject)
    grade_rank_series = grade_rank_series.fillna(rank_from_subject)

    exam_rows = df_display.pivot_table(
        index="exam_name",
        columns="subject",
        values="score",
        aggfunc="first",
    )
    raw_rows = df_display.pivot_table(
        index="exam_name",
        columns="subject",
        values="score_raw",
        aggfunc="first",
    )
    exam_rows["总分"] = total_series
    exam_rows["班级名次"] = class_rank_series
    exam_rows["年级名次"] = grade_rank_series
    if not raw_rows.empty:
        for subject in raw_rows.columns:
            if subject == "总分":
                continue
            exam_rows[f"{subject}(原始)"] = raw_rows[subject]
    for subject in FORCED_RAW_SUBJECTS:
        raw_col_name = f"{subject}(原始)"
        if raw_col_name in exam_rows.columns:
            exam_rows[subject] = exam_rows[raw_col_name]
            exam_rows = exam_rows.drop(columns=[raw_col_name])
    exam_rows = exam_rows.reset_index()
    subject_cols = [
        col
        for col in exam_rows.columns
        if col not in ["exam_name", "总分", "班级名次", "年级名次", "总分"]
    ]
    exam_rows = exam_rows[["exam_name", "总分", "班级名次", "年级名次", *subject_cols]]
    exam_rows = exam_rows.rename(columns={"exam_name": "考试名称"})

    all_subjects = sorted([apply_subject_alias(s, aliases) for s in store.list_subjects(class_id) if s != "总分"])
    series_options_all = ["总分"] + list(dict.fromkeys(all_subjects))

    metric = st.session_state.get("trend_metric", "赋分")

    def build_metric_series(df_in: pd.DataFrame) -> Dict[str, List[Optional[float]]]:
        exams_local = df_in["exam_name"].drop_duplicates().tolist()
        series_map: Dict[str, List[Optional[float]]] = {}

        def series_for_subject(subject: str, value_col: str) -> List[Optional[float]]:
            values = []
            for exam in exams_local:
                exam_df = df_in[(df_in["exam_name"] == exam) & (df_in["subject"] == subject)]
                val = exam_df[value_col].dropna()
                values.append(val.iloc[0] if not val.empty else None)
            return values

        if metric == "名次":
            series_map["总分"] = []
            for exam in exams_local:
                exam_df = df_in[df_in["exam_name"] == exam]
                rank_col = "grade_rank" if st.session_state.get("trend_total_rank_source") == "年级排名" else "class_rank"
                val = exam_df[rank_col].dropna()
                series_map["总分"].append(val.iloc[0] if not val.empty else None)
            for subject in sorted(df_in["subject"].dropna().unique().tolist()):
                if subject == "总分":
                    continue
                series_map[subject] = series_for_subject(subject, "rank")
        elif metric == "原始":
            series_map["总分"] = []
            for exam in exams_local:
                exam_df = df_in[df_in["exam_name"] == exam]
                val = exam_df["total_raw"].dropna()
                series_map["总分"].append(val.iloc[0] if not val.empty else None)
            for subject in sorted(df_in["subject"].dropna().unique().tolist()):
                if subject == "总分":
                    continue
                series_map[subject] = series_for_subject(subject, "score_raw")
        else:
            series_map["总分"] = []
            for exam in exams_local:
                exam_df = df_in[df_in["exam_name"] == exam]
                val = exam_df["total_score"].dropna()
                if not val.empty:
                    series_map["总分"].append(val.iloc[0])
                else:
                    total_row = exam_df[exam_df["subject"] == "总分"]["score"].dropna()
                    series_map["总分"].append(total_row.iloc[0] if not total_row.empty else None)
            for subject in sorted(df_in["subject"].dropna().unique().tolist()):
                if subject == "总分":
                    continue
                series_map[subject] = series_for_subject(subject, "score")

        return series_map

    series_map = build_metric_series(df)
    series_options = list(series_map.keys())
    selected_all = st.session_state.get("trend_series", ["总分"])
    display_series = [s for s in selected_all if s in series_options]
    if not display_series:
        display_series = ["总分"] if "总分" in series_options else series_options[:1]

    with tab_table:
        with st.expander("展开查看表格", expanded=False):
            st.dataframe(exam_rows, height=360)

    with tab_chart:
        selected_series = display_series
        st.session_state["trend_series"] = selected_series

        palette = [
            "#2F6FED",
            "#F2994A",
            "#27AE60",
            "#9B51E0",
            "#EB5757",
            "#219653",
            "#56CCF2",
            "#BB6BD9",
        ]

        fig = go.Figure()
        color_index = 0
        if "总分" in selected_series and "总分" in series_map:
            fig.add_trace(
                go.Scatter(
                    x=trend.exam_order,
                    y=series_map["总分"],
                    mode="lines+markers",
                    name="总分",
                    line=dict(color=palette[color_index % len(palette)], width=3),
                )
            )
            color_index += 1

        for subject, values in series_map.items():
            if subject == "总分":
                continue
            if subject not in selected_series:
                continue
            color = palette[color_index % len(palette)]
            width = 3
            fig.add_trace(
                go.Scatter(
                    x=trend.exam_order,
                    y=values,
                    mode="lines+markers",
                    name=subject,
                    line=dict(color=color, width=width),
                )
            )
            color_index += 1
        fig.update_layout(template="plotly_white", height=360, margin=dict(l=24, r=24, t=24, b=24))
        st.plotly_chart(fig, use_container_width=True)

    with tab_settings:
        st.checkbox("紧凑展示", value=True, key="trend_compact")
        st.selectbox("图表指标", ["赋分", "原始", "名次"], key="trend_metric")
        st.multiselect("图表显示内容", series_options_all, default=selected_all, key="trend_series")
        if st.session_state.get("trend_metric") == "名次" and "总分" in st.session_state.get("trend_series", ["总分"]):
            st.selectbox("总分名次来源", ["年级排名", "班级排名"], key="trend_total_rank_source")
        st.checkbox("自动播放", key="trend_autoplay")
        st.slider("播放间隔（秒）", 3, 15, key="trend_interval")
        st.caption("设置只影响当前页面展示。")


def render_report_export(store: DataStore, class_id: int) -> None:
    st.subheader("导出分析报告")
    students = store.get_students(class_id)
    if not students:
        st.info("暂无学生数据。")
        return

    aliases = load_subject_aliases(store)
    metric = st.selectbox("导出图表指标", ["赋分", "原始", "名次"], index=0)
    subject_options = ["总分"] + [apply_subject_alias(s, aliases) for s in store.list_subjects(class_id) if s != "总分"]
    subject_options = list(dict.fromkeys(subject_options))
    selected_series = st.multiselect("导出图表包含内容", subject_options, default=["总分"])
    total_rank_source = st.session_state.get("export_total_rank_source", "年级排名")
    if metric == "名次" and "总分" in selected_series:
        total_rank_source = st.selectbox("总分名次来源", ["年级排名", "班级排名"], index=0, key="export_total_rank_source")

    selected = st.multiselect("选择学生（默认全部）", [s.name for s in students])
    combined = st.checkbox("合并为单个文件", value=True)
    if st.button("生成并导出", type="primary"):
        target_students = [s for s in students if not selected or s.name in selected]
        report_map: Dict[str, Dict[str, List[float]]] = {}
        exams_map: Dict[str, List[str]] = {}
        for student in target_students:
            rows = store.get_scores_by_student(student.id)
            df = to_dataframe(rows)
            if df.empty:
                continue
            df["subject"] = df["subject"].apply(lambda s: apply_subject_alias(s, aliases))
            df = normalize_forced_subjects(df)

            exams = df["exam_name"].drop_duplicates().tolist()

            def series_for_subject(subject: str, value_col: str) -> List[Optional[float]]:
                values = []
                for exam in exams:
                    exam_df = df[(df["exam_name"] == exam) & (df["subject"] == subject)]
                    val = exam_df[value_col].dropna()
                    values.append(val.iloc[0] if not val.empty else None)
                return values

            series_map: Dict[str, List[Optional[float]]] = {}
            if metric == "名次":
                series_map["总分"] = []
                for exam in exams:
                    exam_df = df[df["exam_name"] == exam]
                    rank_col = "grade_rank" if total_rank_source == "年级排名" else "class_rank"
                    val = exam_df[rank_col].dropna()
                    series_map["总分"].append(val.iloc[0] if not val.empty else None)
                for subject in sorted(df["subject"].dropna().unique().tolist()):
                    if subject == "总分":
                        continue
                    series_map[subject] = series_for_subject(subject, "rank")
            elif metric == "原始":
                series_map["总分"] = []
                for exam in exams:
                    exam_df = df[df["exam_name"] == exam]
                    val = exam_df["total_raw"].dropna()
                    series_map["总分"].append(val.iloc[0] if not val.empty else None)
                for subject in sorted(df["subject"].dropna().unique().tolist()):
                    if subject == "总分":
                        continue
                    series_map[subject] = series_for_subject(subject, "score_raw")
            else:
                series_map["总分"] = []
                for exam in exams:
                    exam_df = df[df["exam_name"] == exam]
                    val = exam_df["total_score"].dropna()
                    if not val.empty:
                        series_map["总分"].append(val.iloc[0])
                    else:
                        total_row = exam_df[exam_df["subject"] == "总分"]["score"].dropna()
                        series_map["总分"].append(total_row.iloc[0] if not total_row.empty else None)
                for subject in sorted(df["subject"].dropna().unique().tolist()):
                    if subject == "总分":
                        continue
                    series_map[subject] = series_for_subject(subject, "score")

            if selected_series:
                series_map = {k: v for k, v in series_map.items() if k in selected_series}

            report_map[student.name] = series_map
            exams_map[student.name] = exams

        if not report_map:
            st.warning("没有可导出的数据。")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = EXPORT_DIR / f"报告_{timestamp}"
        out_path = export_student_reports(
            out_dir,
            report_map,
            exams_map,
            combined=combined,
            project_name=APP_NAME,
            copyright_text=APP_COPYRIGHT,
        )
        st.success(f"导出完成：{out_path.name}")
        with open(out_path, "rb") as f:
            st.download_button("下载报告", f, file_name=out_path.name)


def render_compare(store: DataStore, class_id: int) -> None:
    st.subheader("学生对比")
    students = store.get_students(class_id)
    if len(students) < 2:
        st.info("至少需要 2 位学生。")
        return

    aliases = load_subject_aliases(store)

    names = [s.name for s in students]
    selected = st.multiselect("选择 2-6 位学生", names, default=names[:2])
    if len(selected) < 2 or len(selected) > 6:
        st.warning("请选择 2-6 位学生。")
        return

    mode = st.radio("对比模式", ["历次成绩", "单次考试"])
    rows = store.get_scores_for_students([s.id for s in students if s.name in selected])
    df = to_dataframe(rows)
    if df.empty:
        st.info("暂无数据。")
        return

    df["subject"] = df["subject"].apply(lambda s: apply_subject_alias(s, aliases))
    df = normalize_forced_subjects(df)

    if mode == "单次考试":
        exams = df["exam_name"].drop_duplicates().tolist()
        exam_selected = st.selectbox("选择考试", exams)
        df = df[df["exam_name"] == exam_selected]

    subject_list = sorted([s for s in df["subject"].dropna().unique().tolist() if s != "总分"])
    subject_choice = st.multiselect("对比科目", ["总分"] + subject_list, default=["总分"])
    metric_choice = st.multiselect("对比指标", ["赋分", "原始", "名次"], default=["赋分"])
    total_rank_source = st.session_state.get("compare_total_rank_source", "年级排名")
    if "总分" in subject_choice and "名次" in metric_choice:
        total_rank_source = st.selectbox("总分名次来源", ["年级排名", "班级排名"], index=0, key="compare_total_rank_source")

    long_rows = []
    if "总分" in subject_choice:
        total_df = df.drop_duplicates(subset=["student_name", "exam_name", "total_score", "total_raw", "class_rank"])
        for _, row in total_df.iterrows():
            for metric in metric_choice:
                if metric == "赋分":
                    value = row.get("total_score")
                elif metric == "原始":
                    value = row.get("total_raw")
                else:
                    value = row.get("grade_rank") if total_rank_source == "年级排名" else row.get("class_rank")
                long_rows.append(
                    {
                        "student_name": row.get("student_name"),
                        "exam_name": row.get("exam_name"),
                        "subject": "总分",
                        "metric": metric,
                        "value": value,
                    }
                )

    df_subject = df[df["subject"].isin([s for s in subject_choice if s != "总分"])]
    for _, row in df_subject.iterrows():
        subject = row.get("subject")
        for metric in metric_choice:
            if metric == "赋分":
                value = row.get("score")
            elif metric == "原始":
                value = row.get("score_raw")
            else:
                value = row.get("rank")
            long_rows.append(
                {
                    "student_name": row.get("student_name"),
                    "exam_name": row.get("exam_name"),
                    "subject": subject,
                    "metric": metric,
                    "value": value,
                }
            )

    long_df = pd.DataFrame(long_rows)
    if not long_df.empty:
        palette = [
            "#2F6FED",
            "#F2994A",
            "#27AE60",
            "#9B51E0",
            "#EB5757",
            "#219653",
            "#56CCF2",
            "#BB6BD9",
        ]
        student_colors = {name: palette[i % len(palette)] for i, name in enumerate(sorted(long_df["student_name"].unique()))}
        line_styles = ["solid", "dash", "dot", "dashdot", "longdash", "longdashdot"]
        subject_styles = {name: line_styles[i % len(line_styles)] for i, name in enumerate(sorted(long_df["subject"].unique()))}

        fig = go.Figure()
        for (student, subject, metric), group in long_df.groupby(["student_name", "subject", "metric"]):
            label = f"{student}-{subject}-{metric}"
            fig.add_trace(
                go.Scatter(
                    x=group["exam_name"],
                    y=group["value"],
                    mode="lines+markers",
                    name=label,
                    line=dict(color=student_colors.get(student, "#2F6FED"), dash=subject_styles.get(subject, "solid"), width=2),
                )
            )
        fig.update_layout(template="plotly_white", height=420, title="学生对比")
        st.plotly_chart(fig, use_container_width=True)


def render_stats(store: DataStore, class_id: int) -> None:
    st.subheader("统计分析")
    aliases = load_subject_aliases(store)
    scope = st.radio("统计范围", ["班级", "年级"], horizontal=True)
    if scope == "班级":
        rows = store.get_scores_for_students([s.id for s in store.get_students(class_id)])
    else:
        rows = store.get_all_scores()
    df = to_dataframe(rows)
    if df.empty:
        st.info("暂无数据。")
        return

    df["subject"] = df["subject"].apply(lambda s: apply_subject_alias(s, aliases))
    df = normalize_forced_subjects(df)

    subject_list = sorted([s for s in df["subject"].dropna().unique().tolist() if s != "总分"])
    subject = st.selectbox("选择科目", ["总分"] + subject_list)
    metric = st.selectbox("指标", ["赋分", "原始", "名次"])
    mode = st.radio("统计方式", ["两次考试对比", "单次考试条件筛选"], horizontal=True)
    total_rank_source = st.session_state.get("stats_total_rank_source", "年级排名")
    if subject == "总分" and metric == "名次":
        total_rank_source = st.selectbox("总分名次来源", ["年级排名", "班级排名"], index=0, key="stats_total_rank_source")

    percentile = 0
    use_percentile = False
    if mode == "单次考试条件筛选":
        use_percentile = st.checkbox("按指定百分比划线", value=False)
        if use_percentile:
            percentile = st.slider("选择百分比（前 n%）", 1, 100, 10)
        else:
            operator = st.selectbox("条件", [">=", "<="]) 
            threshold = st.number_input("阈值", value=0.0)
    else:
        operator = st.selectbox("条件", [">=", "<="]) 
        threshold = st.number_input("阈值", value=0.0)

    exams = df["exam_name"].drop_duplicates().tolist()
    if not exams:
        st.info("暂无考试数据。")
        return
    exam_a = None
    exam_b = None
    exam_single = None
    if mode == "两次考试对比":
        if len(exams) < 2:
            st.info("至少需要两次考试数据。")
            return
        exam_a = st.selectbox("基准考试", exams, index=0)
        exam_b = st.selectbox("对比考试", exams, index=min(1, len(exams) - 1))
    else:
        exam_single = st.selectbox("选择考试", exams, index=0)

    if subject == "总分":
        base_df = df.drop_duplicates(subset=["student_name", "exam_name", "total_score", "total_raw", "class_rank", "grade_rank"])
        if metric == "赋分":
            value_col = "total_score"
        elif metric == "原始":
            value_col = "total_raw"
        else:
            value_col = "grade_rank" if total_rank_source == "年级排名" else "class_rank"
    else:
        base_df = df[df["subject"] == subject]
        if metric == "赋分":
            value_col = "score"
        elif metric == "原始":
            value_col = "score_raw"
        else:
            value_col = "rank"

    if mode == "两次考试对比":
        left = base_df[base_df["exam_name"] == exam_a][["student_name", value_col]].rename(columns={value_col: "基准"})
        right = base_df[base_df["exam_name"] == exam_b][["student_name", value_col]].rename(columns={value_col: "对比"})
        merged = pd.merge(left, right, on="student_name", how="inner")
        merged["变化"] = merged["对比"] - merged["基准"]
        result_source = merged
    else:
        result_source = base_df[base_df["exam_name"] == exam_single][
            ["student_name", "exam_name", value_col]
        ].rename(columns={value_col: "值"})

    if mode == "两次考试对比":
        if operator == ">=":
            result = result_source[result_source["变化"] >= threshold]
        else:
            result = result_source[result_source["变化"] <= threshold]
    else:
        if percentile > 0:
            if metric == "名次" and value_col in ("rank", "class_rank", "grade_rank"):
                result = result_source.sort_values(by=["值"]).head(max(1, int(len(result_source) * percentile / 100)))
            else:
                result = result_source.sort_values(by=["值"], ascending=False).head(max(1, int(len(result_source) * percentile / 100)))
        else:
            if operator == ">=":
                result = result_source[result_source["值"] >= threshold]
            else:
                result = result_source[result_source["值"] <= threshold]

    if percentile > 0:
        if metric == "名次" and value_col in ("rank", "class_rank", "grade_rank"):
            line_value = result_source["值"].quantile(percentile / 100)
        else:
            line_value = result_source["值"].quantile(1 - percentile / 100)
        st.info(f"百分位线：{line_value:.2f}")

    st.dataframe(result)

    if st.button("导出统计为Excel"):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = export_stats_excel(result, EXPORT_DIR / f"统计_{timestamp}.xlsx")
        with open(out_path, "rb") as f:
            st.download_button("下载统计Excel", f, file_name=out_path.name)

    if st.button("导出统计为图片") and not result.empty:
        y_col = "变化" if mode == "两次考试对比" else "值"
        title = "统计结果（变化）" if mode == "两次考试对比" else "统计结果"
        chart = px.bar(
            result,
            x="student_name" if "student_name" in result.columns else result.index,
            y=y_col,
            color="student_name" if "student_name" in result.columns else None,
            title=title,
        )
        chart.update_layout(template="plotly_white", height=420)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = EXPORT_DIR / f"统计_{timestamp}.png"
        chart.write_image(out_path)
        with open(out_path, "rb") as f:
            st.download_button("下载统计图片", f, file_name=out_path.name)

def render_overall_stats(store: DataStore, class_id: int) -> None:
    st.subheader("整体数据")
    rows = store.get_scores_for_students([s.id for s in store.get_students(class_id)])
    df = to_dataframe(rows)
    if df.empty:
        st.info("暂无数据。")
        return

    aliases = load_subject_aliases(store)
    df["subject"] = df["subject"].apply(lambda s: apply_subject_alias(s, aliases))
    df = normalize_forced_subjects(df)

    exams = df["exam_name"].drop_duplicates().tolist()
    exam_selected = st.selectbox("选择考试", exams)

    subject_list = sorted([s for s in df["subject"].dropna().unique().tolist() if s != "总分"])
    subject = st.selectbox("选择科目", ["总分"] + subject_list)
    metric = st.selectbox("统计指标", ["赋分", "原始"], index=0)

    if subject == "总分":
        value_col = "total_score" if metric == "赋分" else "total_raw"
        base_df = df.drop_duplicates(subset=["student_name", "exam_name", "total_score", "total_raw"])
    else:
        value_col = "score" if metric == "赋分" else "score_raw"
        base_df = df[df["subject"] == subject]

    exam_df = base_df[base_df["exam_name"] == exam_selected]
    values = exam_df[value_col].dropna()
    if values.empty:
        st.warning("该考试暂无有效数据。")
        return

    stats = {
        "平均分": values.mean(),
        "中位数": values.median(),
        "最大值": values.max(),
        "最小值": values.min(),
        "方差": values.var(),
    }
    stats_df = pd.DataFrame([stats])
    st.dataframe(stats_df)

    min_val = values.min()
    max_val = values.max()
    bin_start = (min_val // 10) * 10
    bin_end = ((max_val // 10) + 1) * 10
    bins = list(range(int(bin_start), int(bin_end) + 10, 10))

    hist = px.histogram(
        exam_df,
        x=value_col,
        nbins=len(bins) - 1,
        title="频数分布直方图（组距 10）",
    )
    hist.update_layout(template="plotly_white", height=360)
    st.plotly_chart(hist, use_container_width=True)


def render_data_manage(store: DataStore, class_id: int) -> None:
    st.subheader("数据管理")
    st.markdown("**学生列表**")
    students = store.get_students(class_id)
    for s in students:
        col1, col2 = st.columns([4, 1])
        with col1:
            st.write(s.name)
        with col2:
            if st.button("删除", key=f"del_student_{s.id}"):
                store.delete_student(s.id)
                st.rerun()

    st.markdown("**考试列表**")
    exams = store.get_exams(class_id)
    for e in exams:
        col1, col2 = st.columns([4, 1])
        with col1:
            st.write(e.name)
        with col2:
            if st.button("删除", key=f"del_exam_{e.id}"):
                store.delete_exam(e.id)
                st.rerun()

    st.markdown("**手动修正成绩**")
    if students:
        selected = st.selectbox("选择学生", [s.name for s in students])
        student = next(s for s in students if s.name == selected)
        rows = store.get_scores_by_student(student.id)
        df = to_dataframe(rows)
        if not df.empty:
            edit_df = df[["id", "exam_name", "subject", "score", "score_raw", "rank", "total_score", "total_raw", "class_rank", "grade_rank"]]
            edited = st.data_editor(edit_df, num_rows="fixed")
            if st.button("保存修改"):
                for _, row in edited.iterrows():
                    store.update_score_item(
                        int(row["id"]),
                        float(row["score"]) if pd.notna(row["score"]) else None,
                        float(row["score_raw"]) if pd.notna(row["score_raw"]) else None,
                        int(row["rank"]) if pd.notna(row["rank"]) else None,
                        float(row["total_score"]) if pd.notna(row["total_score"]) else None,
                        float(row["total_raw"]) if pd.notna(row["total_raw"]) else None,
                        int(row["grade_rank"]) if pd.notna(row["grade_rank"]) else None,
                        int(row["class_rank"]) if pd.notna(row["class_rank"]) else None,
                    )
                st.success("已保存。")

    st.markdown("**导出全量数据**")
    if st.button("导出（本软件可读）"):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = EXPORT_DIR / f"全量数据_{timestamp}.db"
        store.export_database(out_path)
        with open(out_path, "rb") as f:
            st.download_button("下载数据库", f, file_name=out_path.name)

    st.markdown("**导入全量数据**")
    upload_db = st.file_uploader("上传数据库文件（.db）", type=["db"], key="import_db")
    if upload_db is not None:
        if st.button("导入并覆盖现有数据"):
            temp_path = EXPORT_DIR / f"_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            with open(temp_path, "wb") as f:
                f.write(upload_db.getbuffer())
            store.import_database(temp_path)
            st.success("导入完成，已覆盖当前数据。")
            st.rerun()


def render_settings(store: DataStore, class_id: int) -> None:
    st.subheader("设置")
    st.markdown("**班级管理**")
    classes = store.get_classes()
    if classes:
        class_names = [c.name for c in classes]
        new_class = st.text_input("新增班级名称", value="新班级")
        if st.button("保存班级") and new_class:
            store.add_class(new_class)
            st.rerun()

        rename_target = st.selectbox("选择班级", class_names, key="rename_class")
        new_name = st.text_input("新名称", value=rename_target)
        if st.button("重命名") and new_name:
            target_id = next(c.id for c in classes if c.name == rename_target)
            store.rename_class(target_id, new_name)
            st.rerun()

        delete_target = st.selectbox("删除班级", class_names, key="delete_class")
        if st.button("删除选中班级"):
            target_id = next(c.id for c in classes if c.name == delete_target)
            store.delete_class(target_id)
            st.rerun()
    else:
        st.info("暂无班级。")

    st.markdown("---")
    st.markdown("**科目标准名称映射**")
    existing_aliases = load_subject_aliases(store)
    subjects = store.list_subjects(class_id)
    rows = []
    used = set()
    for subject in subjects:
        if subject in used:
            continue
        used.add(subject)
        rows.append({"导入科目": subject, "标准名称": existing_aliases.get(subject, "")})
    if not rows and existing_aliases:
        rows = [{"导入科目": k, "标准名称": v} for k, v in existing_aliases.items()]

    alias_df = pd.DataFrame(rows or [{"导入科目": "", "标准名称": ""}])
    edited_aliases = st.data_editor(alias_df, num_rows="dynamic", key="subject_aliases")
    if st.button("保存科目映射"):
        alias_map = {}
        for _, row in edited_aliases.iterrows():
            raw_name = str(row.get("导入科目", "")).strip()
            std_name = str(row.get("标准名称", "")).strip()
            if raw_name:
                alias_map[raw_name] = std_name or raw_name
        store.set_setting("subject_aliases", json.dumps(alias_map, ensure_ascii=False))
        st.success("科目映射已保存。")

    st.markdown("---")
    enabled = store.get_setting("password_enabled") == "1"
    st.write("开启打开密码保护：", "已开启" if enabled else "未开启")

    new_pwd = st.text_input("设置/修改密码", type="password")
    confirm = st.text_input("确认密码", type="password")
    if st.button("保存密码"):
        if not new_pwd or new_pwd != confirm:
            st.error("两次输入不一致。")
        else:
            hashed = bcrypt.hashpw(new_pwd.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            store.set_setting("password_hash", hashed)
            store.set_setting("password_enabled", "1")
            st.success("密码已保存并启用。")

    if st.button("关闭密码保护"):
        store.set_setting("password_enabled", "0")
        st.success("已关闭密码保护。")


def main() -> None:
    init_page()
    store = DataStore()
    check_password(store)

    class_id = select_class(store)


    page = st.sidebar.radio(
        "功能导航",
        ["导入成绩", "成绩展示", "导出报告", "学生对比", "统计分析", "整体数据", "数据管理", "设置"],
        index=1,
    )

    if page == "导入成绩":
        tab1, tab2 = st.tabs(["输入成绩单", "通常成绩单"])
        with tab1:
            import_input_sheet(store, class_id)
        with tab2:
            import_usual_sheet(store, class_id)
    elif page == "成绩展示":
        render_trend_page(store, class_id)
    elif page == "导出报告":
        render_report_export(store, class_id)
    elif page == "学生对比":
        render_compare(store, class_id)
    elif page == "统计分析":
        render_stats(store, class_id)
    elif page == "整体数据":
        render_overall_stats(store, class_id)
    elif page == "数据管理":
        render_data_manage(store, class_id)
    else:
        render_settings(store, class_id)


if __name__ == "__main__":
    main()
