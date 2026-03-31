import uuid
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio

from search_quality_report import (
    DEFAULT_PATHS,
    CATEGORIES,
    NATURAL_WOMEN_PAGE_SIZE,
    load_data,
    weekly_from_daily,
    contribution_summary,
    type_weekly_summary,
    category_scatter,
    women_natural_top_pool,
    display_keyword_table,
    style_change_cell,
    split_keyword_perf_groups,
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

    # 图4：搜索量占比（本周/上周）
    fig_t0 = go.Figure()
    total_uv_cur = by_type["搜索UV_本周"].sum()
    total_uv_pre = by_type["搜索UV_上周"].sum()
    uv_share_cur = by_type["搜索UV_本周"] / total_uv_cur if total_uv_cur else pd.Series([float("nan")] * len(by_type))
    uv_share_pre = by_type["搜索UV_上周"] / total_uv_pre if total_uv_pre else pd.Series([float("nan")] * len(by_type))
    text_uv_cur = [f"{v:,.0f}\n({s:.1%})" if pd.notna(s) else f"{v:,.0f}" for v, s in zip(by_type["搜索UV_本周"], uv_share_cur)]
    text_uv_pre = [f"{v:,.0f}\n({s:.1%})" if pd.notna(s) else f"{v:,.0f}" for v, s in zip(by_type["搜索UV_上周"], uv_share_pre)]
    fig_t0.add_trace(go.Bar(x=by_type["操作类型"], y=uv_share_cur, name="本周", marker_color="#4c78a8", text=text_uv_cur, textposition="outside"))
    fig_t0.add_trace(go.Bar(x=by_type["操作类型"], y=uv_share_pre, name="上周", marker_color="#9ecae9", text=text_uv_pre, textposition="outside"))
    fig_t0.update_layout(height=360, barmode="group", margin=dict(l=20, r=20, t=20, b=20), yaxis=dict(title="搜索UV占比", tickformat=".0%"), title="按搜索类型的搜索量占比")

    # 图5：金额占比（本周/上周）
    fig_t1 = go.Figure()
    total_amt_pre = by_type["购买总金额_上周"].sum()
    pre_share = by_type["购买总金额_上周"] / total_amt_pre if total_amt_pre else pd.Series([float("nan")] * len(by_type))
    text_cur = [f"{v:,.0f}\n({s:.1%})" for v, s in zip(by_type["购买总金额_本周"], by_type["金额占比_本周"])]
    text_pre = [f"{v:,.0f}\n({s:.1%})" if pd.notna(s) else f"{v:,.0f}" for v, s in zip(by_type["购买总金额_上周"], pre_share)]
    fig_t1.add_trace(go.Bar(x=by_type["操作类型"], y=by_type["金额占比_本周"], name="本周", marker_color="#4c78a8", text=text_cur, textposition="outside"))
    fig_t1.add_trace(go.Bar(x=by_type["操作类型"], y=pre_share, name="上周", marker_color="#9ecae9", text=text_pre, textposition="outside"))
    fig_t1.update_layout(height=380, barmode="group", margin=dict(l=20, r=20, t=20, b=20), yaxis=dict(title="金额占比", tickformat=".0%"))

    # 图6：搜索类型周环比（CTR / ATC / CVR）
    metrics_cfg = [
        ("CTR", "#1f77b4", "#9ecae9"),
        ("ATC", "#2ca02c", "#a1d99b"),
        ("CVR", "#d62728", "#fdae6b"),
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

    return fig_rate, fig_share, fig_uvv, fig_t0, fig_t1, fig_t2


def table_html(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "<p class='small'>暂无符合条件的词。</p>"
    tdf, styled_cols = display_keyword_table(df)
    uid = uuid.uuid4().hex[:12]
    styler = tdf.style.set_table_attributes(f'class="kw-table" id="kw_{uid}"')
    if hasattr(styler, "map"):
        styler = styler.map(style_change_cell, subset=styled_cols)
    else:
        styler = styler.applymap(style_change_cell, subset=styled_cols)
    inner = styler.to_html(table_uuid=uid, exclude_styles=False)
    return f'<div class="table-scroll">{inner}</div>'


def generate_html(output_path: Path):
    # 使用默认路径（指向 zara周报数据源）
    paths = DEFAULT_PATHS

    mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words = load_data(
        paths
    )
    wk = weekly_from_daily(zara_daily_cur, zara_daily_pre)
    contrib = contribution_summary(wk, mini)
    by_type = type_weekly_summary(zara_by_type_cur, zara_by_type_pre)

    fig_rate, fig_share, fig_uvv, fig_t0, fig_t1, fig_t2 = build_core_figures(wk, contrib, by_type)

    # 各品类热词 / 自然词散点图 + 分组表格（女士自然词：Top100 拆 5 段 × 20 词）
    cate_blocks = []
    for cate in CATEGORIES:
        fig_hot, data_hot = category_scatter(
            hotwords, cate, f"{cate} 热词：CTR vs CVR（气泡=搜索PV）"
        )
        nat_parts = []
        if cate == "女士":
            pool_w = women_natural_top_pool(natural_words)
            if pool_w.empty:
                nat_parts.append((None, None, None, 0))
            else:
                n_pool = len(pool_w)
                for i in range(0, n_pool, NATURAL_WOMEN_PAGE_SIZE):
                    end = min(i + NATURAL_WOMEN_PAGE_SIZE, n_pool)
                    label = f"{i + 1}-{end}"
                    page_df = pool_w.iloc[i:end].copy()
                    title = f"{cate} 自然词：CTR vs CVR（气泡=搜索PV）｜{label} / Top{n_pool}"
                    fig_n, data_n = category_scatter(
                        natural_words,
                        cate,
                        title,
                        plot_df=page_df,
                        ref_pool_for_avg=pool_w,
                        bubble_max_df=pool_w,
                    )
                    nat_parts.append((fig_n, data_n, label, n_pool))
        else:
            fig_nat, data_nat = category_scatter(
                natural_words, cate, f"{cate} 自然词：CTR vs CVR（气泡=搜索PV）"
            )
            nat_parts.append((fig_nat, data_nat, None, 0))
        cate_blocks.append((cate, fig_hot, data_hot, nat_parts))

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
    .kpi-delta.kpi-up {{ color: #09ab3b; font-weight: 600; }}
    .kpi-delta.kpi-down {{ color: #ff4b4b; font-weight: 600; }}
    .section-desc {{ color: #666; font-size: 13px; margin-bottom: 4px; }}
    .small {{ font-size: 12px; color: #777; }}
    .table-scroll {{ max-height: 420px; overflow: auto; border: 1px solid #e6eaf0; border-radius: 8px; margin-top: 8px; background: #fff; }}
    .table-scroll table.dataframe {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
    .table-scroll th {{ position: sticky; top: 0; background: #f5f7fa; z-index: 1; }}
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
        else:
            val_str = fmt_num(cur)
        delta = safe_pct_change(cur, pre)
        delta_str = fmt_pct(delta)
        if pd.isna(delta):
            delta_cls = "kpi-delta"
        elif delta > 0:
            delta_cls = "kpi-delta kpi-up"
        elif delta < 0:
            delta_cls = "kpi-delta kpi-down"
        else:
            delta_cls = "kpi-delta"
        html_parts.append(
            f"""
    <div class="kpi-card">
      <div class="kpi-label">{name}</div>
      <div class="kpi-value">{val_str}</div>
      <div class="{delta_cls}">环比：{delta_str}</div>
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
    fig_html_t0 = pio.to_html(
        fig_t0, include_plotlyjs=False, full_html=False
    )
    fig_html_t1 = pio.to_html(
        fig_t1, include_plotlyjs=False, full_html=False
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
    html_parts.append(fig_html_t0)
    html_parts.append(fig_html_t1)
    html_parts.append(fig_html_t2)

    # 4) 热词与自然词分析
    html_parts.append('<h2>4) 热词与自然词分析</h2>')
    for cate, fig_hot, data_hot, nat_parts in cate_blocks:
        html_parts.append(f"<h3>{cate}</h3>")

        html_parts.append(f"<h4>{cate} - 热词分析</h4>")
        if fig_hot is None:
            html_parts.append("<p class='small'>该品类暂无热词数据。</p>")
        else:
            html_parts.append(pio.to_html(fig_hot, include_plotlyjs=False, full_html=False))
            good_hot, bad_hot, basis_hot = split_keyword_perf_groups(data_hot)
            if basis_hot == "无环比":
                html_parts.append(table_html(data_hot))
            else:
                html_parts.append(
                    f"""<div style="display:flex;gap:16px;align-items:flex-start;">
<div style="flex:1;min-width:0;"><h5>表现上升（{basis_hot} ≥ 0）</h5>{table_html(good_hot)}</div>
<div style="flex:1;min-width:0;"><h5>表现下降（{basis_hot} < 0）</h5>{table_html(bad_hot)}</div>
</div>"""
                )

        html_parts.append(f"<h4>{cate} - 自然词分析</h4>")
        for fig_nat, data_nat, seg_label, pool_n in nat_parts:
            if fig_nat is None:
                if cate == "女士" and pool_n == 0:
                    html_parts.append("<p class='small'>该品类暂无自然词数据（按末尾品类提取后为空）。</p>")
                elif cate != "女士":
                    html_parts.append("<p class='small'>该品类暂无自然词数据（按末尾品类提取后为空）。</p>")
                continue
            if seg_label and pool_n:
                html_parts.append(
                    f"<h5>自然词分段 {seg_label} / Top{pool_n}（均值线=全 Top{pool_n}）</h5>"
                )
            html_parts.append(pio.to_html(fig_nat, include_plotlyjs=False, full_html=False))
            good_nat, bad_nat, basis_nat = split_keyword_perf_groups(data_nat)
            if basis_nat == "无环比":
                html_parts.append(table_html(data_nat))
            else:
                html_parts.append(
                    f"""<div style="display:flex;gap:16px;align-items:flex-start;">
<div style="flex:1;min-width:0;"><h5>表现上升（{basis_nat} ≥ 0）</h5>{table_html(good_nat)}</div>
<div style="flex:1;min-width:0;"><h5>表现下降（{basis_nat} < 0）</h5>{table_html(bad_nat)}</div>
</div>"""
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

