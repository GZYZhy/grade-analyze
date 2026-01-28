from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import plotly.graph_objects as go


def _line_chart(title: str, x: List[str], series: Dict[str, List[float]]) -> go.Figure:
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
    for name, values in series.items():
        color = palette[color_index % len(palette)]
        color_index += 1
        fig.add_trace(go.Scatter(x=x, y=values, mode="lines+markers", name=name, line=dict(color=color, width=3)))
    fig.update_layout(title=title, template="plotly_white", height=360, margin=dict(l=24, r=24, t=48, b=24))
    return fig


def render_student_report(student_name: str, series_map: Dict[str, List[float]], exams: List[str]) -> str:
    fig = _line_chart(f"{student_name} 成绩趋势", exams, series_map)
    return fig.to_html(include_plotlyjs="cdn", full_html=False)


def _render_report_header(project_name: str, copyright_text: Optional[str]) -> str:
    copyright_html = f"<span class='meta'>版权 {copyright_text}</span>" if copyright_text else ""
    return (
        "<header class='report-header'>"
        f"<div class='title'>{project_name}</div>"
        f"{copyright_html}"
        "</header>"
    )


def export_student_reports(
    output_dir: Path,
    report_map: Dict[str, Dict[str, List[float]]],
    exams_map: Dict[str, List[str]],
    combined: bool,
    project_name: str,
    copyright_text: Optional[str] = None,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    style = """
    <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif; color: #111827; margin: 24px; }
    .report-header { display: flex; align-items: baseline; gap: 12px; border-bottom: 1px solid #e5e7eb; padding-bottom: 12px; margin-bottom: 20px; }
    .report-header .title { font-size: 22px; font-weight: 700; }
    .report-header .meta { font-size: 14px; color: #6b7280; }
    h2 { margin-top: 24px; }
    </style>
    """
    header_html = _render_report_header(project_name, copyright_text)
    if combined:
        parts = ["<html><head><meta charset='utf-8'>", style, "</head><body>", header_html]
        for name, trend in report_map.items():
            parts.append(f"<h2>{name}</h2>")
            parts.append(render_student_report(name, trend, exams_map[name]))
        parts.append("</body></html>")
        out_path = output_dir / "成绩分析报告.html"
        out_path.write_text("\n".join(parts), encoding="utf-8")
        return out_path

    index_lines = [
        "<html><head><meta charset='utf-8'>",
        style,
        "</head><body>",
        header_html,
        "<h1>成绩分析报告</h1><ul>",
    ]
    for name, trend in report_map.items():
        content = "<html><head><meta charset='utf-8'>" + style + "</head><body>" + header_html
        content += f"<h2>{name}</h2>"
        content += render_student_report(name, trend, exams_map[name])
        content += "</body></html>"
        out_path = output_dir / f"{name}.html"
        out_path.write_text(content, encoding="utf-8")
        index_lines.append(f"<li><a href='{out_path.name}'>{name}</a></li>")
    index_lines.append("</ul></body></html>")
    index_path = output_dir / "成绩分析报告索引.html"
    index_path.write_text("\n".join(index_lines), encoding="utf-8")
    return index_path


def export_stats_excel(df: pd.DataFrame, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return output_path
