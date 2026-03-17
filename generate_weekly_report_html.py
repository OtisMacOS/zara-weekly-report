from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio

from search_quality_report import (
    DEFAULT_PATHS,
    CATEGORIES,
    load_data,
    weekly_from_daily,
    contribution_summary,
    type_weekly_summary,
    category_scatter,
    fmt_num,
    fmt_pct,
    safe_pct_change,
)


def build_core_figures(wk, contrib, by_type):
    # 图1：CTR/ATC/CVR 本周 vs 上周
    rate_colors = {"CTR": "#1f77b4", "ATC": "#2ca02c", "CVR": "#d62728"}
    fig_rate = go.Figure()
    for metric in ["CTR", "ATC", "CVR"]:
        fig_rate.add_trace(
            go.Scatter(
                x=wk["cur"]["date"],
                y=wk["cur"][metric],
                mode="lines+markers+text",
                name=f"{metric}-本周",
                line=dict(color=rate_colors[metric], width=2),
                text=[f"{v:.1%}" if pd.notna(v) else "-" for v in wk["cur"][metric]],
                textposition="top center",
            )
        )
        fig_rate.add_trace(
            go.Scatter(
                x=wk["pre"]["date"],
                y=wk["pre"][metric],
                mode="lines+markers+text",
                name=f"{metric}-上周",
                line=dict(color=rate_colors[metric], dash="dash", width=2),
                text=[f"{v:.1%}" if pd.notna(v) else "-" for v in wk["pre"][metric]],
                textposition="bottom center",
            )
        )
    fig_rate.update_layout(
        height=380, margin=dict(l=20, r=20, t=20, b=20), title="整体转化率（本周 vs 上周）"
    )
    fig_rate.update_yaxes(tickformat=".0%")

    # 图2：搜索对大盘贡献（金额 / 人数）
    fig_share = go.Figure()
    for metric, color in [("金额占比", "#1f77b4"), ("人数占比", "#ff7f0e")]:
        fig_share.add_trace(
            go.Scatter(
                x=contrib["daily_cur"]["date"],
                y=contrib["daily_cur"][metric],
                mode="lines+markers+text",
                name=f"{metric}-本周",
                line=dict(color=color),
                text=[
                    f"{v:.1%}" if pd.notna(v) else "-"
                    for v in contrib["daily_cur"][metric]
                ],
                textposition="top center",
            )
        )
        fig_share.add_trace(
            go.Scatter(
                x=contrib["daily_pre"]["date"],
                y=contrib["daily_pre"][metric],
                mode="lines+markers+text",
                name=f"{metric}-上周",
                line=dict(color=color, dash="dash"),
                text=[
                    f"{v:.1%}" if pd.notna(v) else "-"
                    for v in contrib["daily_pre"][metric]
                ],
                textposition="bottom center",
            )
        )
    fig_share.update_layout(
        height=380,
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis=dict(title="占比", tickformat=".0%"),
        title="搜索对大盘贡献（金额 / 人数）",
    )

    # 图3：UV_VALUE 本周 vs 上周
    fig_uvv = go.Figure()
    fig_uvv.add_trace(
        go.Scatter(
            x=contrib["daily_cur"]["date"],
            y=contrib["daily_cur"]["UV_VALUE"],
            mode="lines+markers+text",
            name="UV_VALUE-本周",
            line=dict(color="#2ca02c"),
            text=[
                f"{v:.1f}" if pd.notna(v) else "-"
                for v in contrib["daily_cur"]["UV_VALUE"]
            ],
            textposition="top center",
        )
    )
    fig_uvv.add_trace(
        go.Scatter(
            x=contrib["daily_pre"]["date"],
            y=contrib["daily_pre"]["UV_VALUE"],
            mode="lines+markers+text",
            name="UV_VALUE-上周",
            line=dict(color="#2ca02c", dash="dash"),
            text=[
                f"{v:.1f}" if pd.notna(v) else "-"
                for v in contrib["daily_pre"]["UV_VALUE"]
            ],
            textposition="bottom center",
        )
    )
    fig_uvv.update_layout(
        height=320,
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis=dict(title="UV_VALUE"),
        title="UV_VALUE（本周 vs 上周）",
    )

    # 图4：搜索类型周环比（CTR / ATC / CVR）
    metrics_cfg = [
        ("CTR", "#9ecae9", "#1f77b4"),
        ("ATC", "#a1d99b", "#2ca02c"),
        ("CVR", "#fdae6b", "#d62728"),
    ]
    fig_t2 = make_subplots(
        rows=1, cols=3, subplot_titles=["CTR", "ATC", "CVR"]
    )
    for idx, (m, c_cur, c_pre) in enumerate(metrics_cfg, start=1):
        fig_t2.add_trace(
            go.Bar(
                x=by_type["操作类型"],
                y=by_type[f"{m}_本周"],
                name=f"{m}-本周",
                marker_color=c_cur,
                text=[
                    f"{v:.1%}" if pd.notna(v) else "-"
                    for v in by_type[f"{m}_本周"]
                ],
                textposition="outside",
                showlegend=(idx == 1),
            ),
            row=1,
            col=idx,
        )
        fig_t2.add_trace(
            go.Bar(
                x=by_type["操作类型"],
                y=by_type[f"{m}_上周"],
                name=f"{m}-上周",
                marker_color=c_pre,
                text=[
                    f"{v:.1%}" if pd.notna(v) else "-"
                    for v in by_type[f"{m}_上周"]
                ],
                textposition="outside",
                showlegend=(idx == 1),
            ),
            row=1,
            col=idx,
        )
        fig_t2.update_yaxes(tickformat=".0%", row=1, col=idx)
    fig_t2.update_layout(
        height=430,
        barmode="group",
        margin=dict(l=20, r=20, t=50, b=20),
        title="搜索类型周环比（CTR / ATC / CVR）",
    )

    return fig_rate, fig_share, fig_uvv, fig_t2


def generate_html(output_path: Path):
    # 使用默认路径（指向 zara周报数据源）
    paths = DEFAULT_PATHS

    mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words = load_data(
        paths
    )
    wk = weekly_from_daily(zara_daily_cur, zara_daily_pre)
    contrib = contribution_summary(wk, mini)
    by_type = type_weekly_summary(zara_by_type_cur, zara_by_type_pre)

    fig_rate, fig_share, fig_uvv, fig_t2 = build_core_figures(wk, contrib, by_type)

    # 各品类热词 / 自然词散点图（可选）
    cate_figs = []
    for cate in CATEGORIES:
        fig_hot, _ = category_scatter(
            hotwords, cate, f"{cate} 热词：CTR vs CVR（气泡=搜索PV）"
        )
        fig_nat, _ = category_scatter(
            natural_words, cate, f"{cate} 自然词：CTR vs CVR（气泡=搜索PV）"
        )
        if fig_hot is not None:
            cate_figs.append((f"{cate} 热词分析", fig_hot))
        if fig_nat is not None:
            cate_figs.append((f"{cate} 自然词分析", fig_nat))

    # 生成 HTML
    html_parts = []

    # 和 Streamlit 页面保持一致的标题
    title = "搜索引擎质量周报（自动化版）"
    period_cur = wk["period"]["cur"]
    period_pre = wk["period"]["pre"]

    # 样式
    html_parts.append(
        f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <title>{title}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans", sans-serif; margin: 16px 40px; }}
    h1 {{ font-size: 26px; margin-bottom: 4px; }}
    h2 {{ margin-top: 32px; border-left: 4px solid #4c78a8; padding-left: 8px; }}
    .kpi-grid {{ display: flex; flex-wrap: wrap; gap: 12px; margin-top: 8px; }}
    .kpi-card {{ border: 1px solid #e0e0e0; border-radius: 6px; padding: 8px 12px; min-width: 120px; background: #fafafa; }}
    .kpi-label {{ font-size: 12px; color: #666; }}
    .kpi-value {{ font-size: 16px; font-weight: 600; }}
    .kpi-delta {{ font-size: 11px; color: #888; }}
    .section-desc {{ color: #666; font-size: 13px; margin-bottom: 4px; }}
    .small {{ font-size: 12px; color: #777; }}
  </style>
</head>
<body>
  <h1>{title}</h1>
  <div class="small">本周：{period_cur} ｜ 上周：{period_pre}</div>
"""
    )

    # 1) 搜索引擎整体周环比（仅 zara日度数据.xlsx）
    html_parts.append('<h2>1) 搜索引擎整体周环比（仅 zara日度数据.xlsx）</h2>')
    html_parts.append('<div class="kpi-grid">')
    for name in ["搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]:
        cur, pre = wk["metrics"][name]
        if name in ["CTR", "ATC", "CVR"]:
            val_str = fmt_pct(cur)
            delta_str = fmt_pct(safe_pct_change(cur, pre))
        else:
            val_str = fmt_num(cur)
            delta_str = fmt_pct(safe_pct_change(cur, pre))
        html_parts.append(
            f"""
    <div class="kpi-card">
      <div class="kpi-label">{name}</div>
      <div class="kpi-value">{val_str}</div>
      <div class="kpi-delta">环比：{delta_str}</div>
    </div>
"""
        )
    html_parts.append("</div>")

    # 嵌入核心图表
    fig_html_rate = pio.to_html(
        fig_rate, include_plotlyjs="cdn", full_html=False
    )
    fig_html_share = pio.to_html(
        fig_share, include_plotlyjs=False, full_html=False
    )
    fig_html_uvv = pio.to_html(
        fig_uvv, include_plotlyjs=False, full_html=False
    )
    fig_html_t2 = pio.to_html(
        fig_t2, include_plotlyjs=False, full_html=False
    )

    # 2) 搜索对大盘贡献（包含整体 CTR/ATC/CVR 曲线、占比、UV_VALUE）
    html_parts.append('<h2>2) 搜索对大盘贡献</h2>')
    html_parts.append(fig_html_rate)
    html_parts.append(fig_html_share)
    html_parts.append(fig_html_uvv)

    # 3) 搜索类型周环比（全链路）
    html_parts.append('<h2>3) 搜索类型周环比（全链路）</h2>')
    html_parts.append(fig_html_t2)

    # 品类散点图
    if cate_figs:
        # 4) 热词与自然词分析
        html_parts.append('<h2>4) 热词与自然词分析</h2>')
        for title_cate, fig in cate_figs:
            html_parts.append(f"<h3>{title_cate}</h3>")
            html_parts.append(
                pio.to_html(fig, include_plotlyjs=False, full_html=False)
            )

    html_parts.append("</body></html>")

    output_path.write_text("".join(html_parts), encoding="utf-8")
    return output_path


if __name__ == "__main__":
    # 文件名里带上本周结束日期，便于区分
    # 为了复用逻辑，先算一次数据拿到 cur_end
    mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words = load_data(
        DEFAULT_PATHS
    )
    wk_tmp = weekly_from_daily(zara_daily_cur, zara_daily_pre)
    cur_end = wk_tmp["cur"]["date"].max().date()
    out_file = Path(__file__).parent / f"search_quality_report_{cur_end}.html"
    generate_html(out_file)
    print(f"HTML 报告已生成: {out_file}")

