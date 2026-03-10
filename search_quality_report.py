import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from matplotlib import pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from plotly.subplots import make_subplots

DEFAULT_BASE_DIR_CANDIDATES = [
    Path(__file__).parent / "zara周报数据源",  # 相对路径（Streamlit Cloud）
    Path("/Users/otis/Downloads/GIT/zara周报自动化/zara周报数据源"),
    Path("/Users/otis/Downloads/zara周报数据源"),
    Path("/Users/otis/Downloads"),
]


def find_latest_periods(base_dir: Path) -> tuple:
    """扫描日度数据文件夹找到最新和次新的数据周期"""
    zara_dir = base_dir / "日度数据整体"
    if not zara_dir.exists():
        return None, None
    xlsx_files = sorted(zara_dir.glob("*.xlsx"))
    if len(xlsx_files) >= 2:
        # 取最后两个文件的周期
        latest = xlsx_files[-1].stem.split("_")[0]
        prev = xlsx_files[-2].stem.split("_")[0]
        return latest, prev
    elif len(xlsx_files) == 1:
        return xlsx_files[0].stem.split("_")[0], None
    return None, None


def find_hotword_periods(base_dir: Path) -> tuple:
    """扫描热词文件夹找到最新和次新的热词周期"""
    hot_dir = base_dir / "热词部分"
    if not hot_dir.exists():
        return None, None
    # 查找所有以"热词"结尾的子目录
    hot_folders = sorted([d for d in hot_dir.iterdir() if d.is_dir() and d.name.endswith("热词")])
    if len(hot_folders) >= 2:
        latest = hot_folders[-1].name.replace("热词", "")
        prev = hot_folders[-2].name.replace("热词", "")
        return latest, prev
    elif len(hot_folders) == 1:
        return hot_folders[0].name.replace("热词", ""), None
    return None, None


def find_natural_word_periods(base_dir: Path) -> tuple:
    """扫描自然搜索词文件夹找到最新和次新的周期"""
    nat_dir = base_dir / "自然搜索词部分"
    if not nat_dir.exists():
        return None, None
    xlsx_files = sorted([f for f in nat_dir.glob("*.xlsx") if "自然搜索词" in f.name])
    if len(xlsx_files) >= 2:
        latest = xlsx_files[-1].stem.replace("自然搜索词", "")
        prev = xlsx_files[-2].stem.replace("自然搜索词", "")
        return latest, prev
    elif len(xlsx_files) == 1:
        return xlsx_files[0].stem.replace("自然搜索词", ""), None
    return None, None


def build_default_paths() -> dict:
    base_dir = next((p for p in DEFAULT_BASE_DIR_CANDIDATES if p.exists()), DEFAULT_BASE_DIR_CANDIDATES[-1])
    latest_period, prev_period = find_latest_periods(base_dir)
    hot_latest, hot_prev = find_hotword_periods(base_dir)
    nat_latest, nat_prev = find_natural_word_periods(base_dir)
    
    paths = {
        "mini": str(base_dir / "小程序大盘部分" / "小程序大盘数据-近30天.xlsx"),
        # 当前周和上周数据分别来自不同文件
        "zara_daily_cur": str(base_dir / "日度数据整体" / f"{latest_period}_zara日度数据.xlsx") if latest_period else "",
        "zara_daily_pre": str(base_dir / "日度数据整体" / f"{prev_period}_zara日度数据.xlsx") if prev_period else "",
        "zara_by_type_cur": str(base_dir / "日度数据-by搜索类型" / f"{latest_period}_zara日度数据-by搜索类型.xlsx") if latest_period else "",
        "zara_by_type_pre": str(base_dir / "日度数据-by搜索类型" / f"{prev_period}_zara日度数据-by搜索类型.xlsx") if prev_period else "",
        # 热词：当前周（使用热词文件夹的周期）
        "hot_women_cur": str(base_dir / "热词部分" / f"{hot_latest}热词" / "女士.xlsx") if hot_latest else "",
        "hot_men_cur": str(base_dir / "热词部分" / f"{hot_latest}热词" / "男士.xlsx") if hot_latest else "",
        "hot_kids_cur": str(base_dir / "热词部分" / f"{hot_latest}热词" / "儿童.xlsx") if hot_latest else "",
        "hot_home_cur": str(base_dir / "热词部分" / f"{hot_latest}热词" / "家居.xlsx") if hot_latest else "",
        # 热词：上周
        "hot_women_pre": str(base_dir / "热词部分" / f"{hot_prev}热词" / "女士.xlsx") if hot_prev else "",
        "hot_men_pre": str(base_dir / "热词部分" / f"{hot_prev}热词" / "男士.xlsx") if hot_prev else "",
        "hot_kids_pre": str(base_dir / "热词部分" / f"{hot_prev}热词" / "儿童.xlsx") if hot_prev else "",
        "hot_home_pre": str(base_dir / "热词部分" / f"{hot_prev}热词" / "家居.xlsx") if hot_prev else "",
        # 自然词（使用自然词文件的周期）
        "natural_words_cur": str(base_dir / "自然搜索词部分" / f"{nat_latest}自然搜索词.xlsx") if nat_latest else "",
        "natural_words_pre": str(base_dir / "自然搜索词部分" / f"{nat_prev}自然搜索词.xlsx") if nat_prev else "",
        "latest_period": latest_period,
        "prev_period": prev_period,
    }
    
    # 兼容旧的文件名格式（当前周）
    for key in ["hot_women_cur", "hot_men_cur", "hot_kids_cur", "hot_home_cur"]:
        fpath = Path(paths[key])
        if not fpath.exists():
            category_map = {"hot_women_cur": "女士", "hot_men_cur": "男士", "hot_kids_cur": "儿童", "hot_home_cur": "家居"}
            alt_path = fpath.parent / f"{category_map[key]}品类热词.xlsx"
            if alt_path.exists():
                paths[key] = str(alt_path)
    
    # 兼容旧的文件名格式（上周）
    for key in ["hot_women_pre", "hot_men_pre", "hot_kids_pre", "hot_home_pre"]:
        fpath = Path(paths[key])
        if not fpath.exists():  # 文件不存在时尝试旧格式
            category_map = {"hot_women_pre": "女士", "hot_men_pre": "男士", "hot_kids_pre": "儿童", "hot_home_pre": "家居"}
            alt_path = fpath.parent / f"{category_map[key]}品类热词.xlsx"
            if alt_path.exists():
                paths[key] = str(alt_path)
    
    return paths


DEFAULT_PATHS = build_default_paths()

TYPE_VALUE_COLS = ["搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]
CATEGORIES = ["女士", "男士", "儿童", "家居"]


def to_num(df: pd.DataFrame, cols):
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def fmt_with_change(value, change):
    """格式化数值和环比信息"""
    if pd.isna(value):
        return "-"
    
    # 根据列类型判断是百分比还是数字
    if value >= 0 and value <= 1 and change is not None and not pd.isna(change):
        # 百分比列
        main_str = f"{value:.2%}"
    else:
        # 数字列
        main_str = f"{value:,.0f}" if isinstance(value, (int, float)) and value > 100 else f"{value:.0f}"
    
    # 添加环比
    if change is not None and not pd.isna(change):
        change_str = f"{change:.1%}"
        # 根据涨跌显示不同颜色
        if change > 0:
            color = "green"
            symbol = "↑"
        elif change < 0:
            color = "red"
            symbol = "↓"
        else:
            color = "gray"
            symbol = "→"
        return f"{main_str}\n<span style='color:{color};font-size:0.8em'>{symbol}{abs(change):.1%}</span>"
    
    return main_str


def format_pct_change(val):
    """格式化环比百分比值"""
    if pd.isna(val):
        return ""
    pct_str = f"{abs(val):.1%}"
    return f"{'↑' if val > 0 else '↓'}{pct_str}" if val != 0 else "→0%"


def format_cell_with_change(value, change, is_pct=False):
    """格式化单元格：主值 + 带颜色的环比"""
    if pd.isna(value):
        return "-"
    
    # 主值格式化
    if is_pct:
        main_str = f"{value:.2%}"
    else:
        main_str = f"{value:,.0f}"
    
    # 环比部分
    if change is not None and not pd.isna(change):
        if change > 0:
            return f"{main_str} ↑{abs(change):.1%}"
        elif change < 0:
            return f"{main_str} ↓{abs(change):.1%}"
        else:
            return f"{main_str} →0%"
    return main_str


def style_change_cell(val):
    """为包含环比的单元格添加颜色样式"""
    if isinstance(val, str):
        if '↑' in val:
            return 'color: green'
        elif '↓' in val:
            return 'color: red'
    return ''


def display_keyword_table(data: pd.DataFrame):
    """准备包含环比信息的表格数据，环比显示在同一单元格内"""
    if data.empty:
        return data, []
    
    output_data = data.copy()
    styled_cols = []
    
    # 为每个指标创建合并显示列（主值+环比）
    metrics_config = [
        ("搜索PV", False),
        ("搜索UV", False),
        ("CTR", True),
        ("ATC", True),
        ("CVR", True),
    ]
    
    for col, is_pct in metrics_config:
        change_col = f"{col}_change"
        if col in output_data.columns and change_col in output_data.columns:
            output_data[col] = output_data.apply(
                lambda row: format_cell_with_change(row[col], row[change_col], is_pct),
                axis=1
            )
            styled_cols.append(col)
        elif col in output_data.columns:
            # 没有环比数据，只格式化主值
            if is_pct:
                output_data[col] = output_data[col].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "-")
            else:
                output_data[col] = output_data[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "-")
    
    # 只保留需要显示的列
    display_cols = ["关键词", "搜索PV", "搜索UV", "CTR", "ATC", "CVR"]
    result = output_data[[c for c in display_cols if c in output_data.columns]]
    
    return result, styled_cols


def safe_pct_change(cur, pre):
    if pre is None or pd.isna(pre) or pre == 0:
        return np.nan
    return (cur - pre) / pre


def fmt_num(val):
    """格式化数字"""
    if pd.isna(val):
        return "-"
    return f"{val:,.0f}"


def fmt_pct(val):
    """格式化百分比"""
    if pd.isna(val):
        return "-"
    return f"{val:.2%}"


def uv_rates(df: pd.DataFrame):
    out = df.copy()
    out["CTR"] = out["点击UV"] / out["搜索UV"].replace(0, np.nan)
    out["ATC"] = out["加购UV"] / out["搜索UV"].replace(0, np.nan)
    out["CVR"] = out["购买人数"] / out["搜索UV"].replace(0, np.nan)
    out["UV_VALUE"] = out["购买总金额"] / out["搜索UV"].replace(0, np.nan)
    return out


def period_from_max(df: pd.DataFrame, date_col: str):
    max_date = df[date_col].max()
    cur_end = max_date
    cur_start = max_date - pd.Timedelta(days=6)
    pre_end = cur_start - pd.Timedelta(days=1)
    pre_start = pre_end - pd.Timedelta(days=6)
    return cur_start, cur_end, pre_start, pre_end


def find_col(df: pd.DataFrame, cands):
    for c in cands:
        if c in df.columns:
            return c
    return None


@st.cache_data(show_spinner=False)
def load_data(paths: dict):
    mini = pd.read_excel(paths["mini"]).copy()
    mini["date"] = pd.to_datetime(mini["日期"].astype(str), errors="coerce")
    mini = to_num(mini, ["成交金额", "成交人数", "UV", "UV价值"]).dropna(subset=["date"]).sort_values("date")

    # 加载当前周和上周的日度数据
    zara_daily_cur = pd.read_excel(paths["zara_daily_cur"], header=2).copy()
    zara_daily_cur = zara_daily_cur.dropna(subset=["Date"])
    zara_daily_cur["date"] = pd.to_datetime(zara_daily_cur["Date"], errors="coerce")
    zara_daily_cur = to_num(zara_daily_cur, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"])
    zara_daily_cur = zara_daily_cur.dropna(subset=["date"]).sort_values("date")
    zara_daily_cur = uv_rates(zara_daily_cur)

    zara_daily_pre = None
    if paths.get("zara_daily_pre") and Path(paths["zara_daily_pre"]).exists():
        zara_daily_pre = pd.read_excel(paths["zara_daily_pre"], header=2).copy()
        zara_daily_pre = zara_daily_pre.dropna(subset=["Date"])
        zara_daily_pre["date"] = pd.to_datetime(zara_daily_pre["Date"], errors="coerce")
        zara_daily_pre = to_num(zara_daily_pre, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"])
        zara_daily_pre = zara_daily_pre.dropna(subset=["date"]).sort_values("date")
        zara_daily_pre = uv_rates(zara_daily_pre)

    zara_by_type_cur = pd.read_excel(paths["zara_by_type_cur"], header=2).copy()
    zara_by_type_cur = zara_by_type_cur.dropna(subset=["Date", "操作类型"])
    zara_by_type_cur["date"] = pd.to_datetime(zara_by_type_cur["Date"], errors="coerce")
    zara_by_type_cur = to_num(zara_by_type_cur, ["搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额", "搜索PV"])
    zara_by_type_cur = zara_by_type_cur.dropna(subset=["date"]).sort_values("date")
    zara_by_type_cur = uv_rates(zara_by_type_cur)

    zara_by_type_pre = None
    if paths.get("zara_by_type_pre") and Path(paths["zara_by_type_pre"]).exists():
        zara_by_type_pre = pd.read_excel(paths["zara_by_type_pre"], header=2).copy()
        zara_by_type_pre = zara_by_type_pre.dropna(subset=["Date", "操作类型"])
        zara_by_type_pre["date"] = pd.to_datetime(zara_by_type_pre["Date"], errors="coerce")
        zara_by_type_pre = to_num(zara_by_type_pre, ["搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额", "搜索PV"])
        zara_by_type_pre = zara_by_type_pre.dropna(subset=["date"]).sort_values("date")
        zara_by_type_pre = uv_rates(zara_by_type_pre)

    hot_cfg = [
        ("女士", "hot_women_cur", "hot_women_pre", "女士热词分类"),
        ("男士", "hot_men_cur", "hot_men_pre", "男士热词分类"),
        ("儿童", "hot_kids_cur", "hot_kids_pre", "儿童热词分类"),
        ("家居", "hot_home_cur", "hot_home_pre", "家居热词分类"),
    ]
    hot_frames = []
    for category, fp_cur, fp_pre, kw_col in hot_cfg:
        # 当前周热词
        if fp_cur in paths and Path(paths[fp_cur]).exists():
            df_cur = pd.read_excel(paths[fp_cur], header=2).copy()
            if kw_col in df_cur.columns:
                df_cur = df_cur.dropna(subset=[kw_col]).copy()
                df_cur = df_cur.rename(columns={kw_col: "关键词", "购买UV": "购买人数"})
                df_cur["品类"] = category
                df_cur = to_num(df_cur, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
                df_cur["购买总金额"] = np.nan
                df_cur = uv_rates(df_cur)
                
                # 加载上周热词
                df_pre = None
                if fp_pre in paths and Path(paths[fp_pre]).exists():
                    df_pre = pd.read_excel(paths[fp_pre], header=2).copy()
                    if kw_col in df_pre.columns:
                        df_pre = df_pre.dropna(subset=[kw_col]).copy()
                        df_pre = df_pre.rename(columns={kw_col: "关键词", "购买UV": "购买人数"})
                        df_pre = to_num(df_pre, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
                        df_pre["购买总金额"] = np.nan
                        df_pre = uv_rates(df_pre)
                
                # 合并两周数据计算环比
                if df_pre is not None and not df_pre.empty:
                    merged = df_cur.merge(
                        df_pre[["关键词", "搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]],
                        on="关键词",
                        how="left",
                        suffixes=("_cur", "_pre"),
                    )
                    # 计算环比
                    for col in ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]:
                        merged[f"{col}_change"] = (merged[f"{col}_cur"] - merged[f"{col}_pre"]) / merged[f"{col}_pre"].replace(0, np.nan)
                    df_cur = merged
                else:
                    # 即使没有前周数据，也要添加_cur后缀和_change列（设为NaN）
                    for col in ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]:
                        df_cur[f"{col}_cur"] = df_cur[col]
                        df_cur[f"{col}_change"] = np.nan
                
                hot_frames.append(df_cur)
    
    hotwords = pd.concat(hot_frames, ignore_index=True) if hot_frames else pd.DataFrame()

    natural_words = pd.DataFrame()
    npath_cur = Path(paths.get("natural_words_cur", ""))
    npath_pre = Path(paths.get("natural_words_pre", ""))
    if npath_cur.exists():
        raw_cur = pd.read_excel(npath_cur, header=2).copy()
        kw_col = find_col(raw_cur, ["自然搜索词", "关键词", "搜索词", "query", "Query", "word", "词"])
        pv_col = find_col(raw_cur, ["搜索PV", "pv", "PV"])
        uv_col = find_col(raw_cur, ["搜索UV", "uv", "UV"])
        click_uv_col = find_col(raw_cur, ["点击UV", "click_uv", "点击uv"])
        cart_uv_col = find_col(raw_cur, ["加购UV", "cart_uv", "加购uv"])
        buy_uv_col = find_col(raw_cur, ["购买UV", "购买人数", "buy_uv", "支付UV"])

        if kw_col and pv_col and uv_col and click_uv_col and buy_uv_col:
            df_cur = raw_cur.copy()
            df_cur = df_cur.rename(columns={
                kw_col: "关键词",
                pv_col: "搜索PV",
                uv_col: "搜索UV",
                click_uv_col: "点击UV",
                buy_uv_col: "购买人数",
            })
            if cart_uv_col:
                df_cur = df_cur.rename(columns={cart_uv_col: "加购UV"})
            else:
                df_cur["加购UV"] = np.nan
            df_cur = to_num(df_cur, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
            df_cur = df_cur.dropna(subset=["关键词"]).copy()
            s = df_cur["关键词"].astype(str).str.strip()
            # 优先按词尾抽取固定品类，失败时回退到右侧两个字
            df_cur["品类"] = s.str.extract(r"(女士|男士|儿童|家居)$", expand=False)
            df_cur["品类"] = df_cur["品类"].fillna(s.str[-2:])
            df_cur = df_cur[df_cur["品类"].isin(CATEGORIES)].copy()
            df_cur["购买总金额"] = np.nan
            df_cur = uv_rates(df_cur)
            
            # 加载上周自然词
            df_pre = None
            if npath_pre.exists():
                raw_pre = pd.read_excel(npath_pre, header=2).copy()
                if kw_col and pv_col and uv_col and click_uv_col and buy_uv_col:
                    df_pre = raw_pre.copy()
                    df_pre = df_pre.rename(columns={
                        kw_col: "关键词",
                        pv_col: "搜索PV",
                        uv_col: "搜索UV",
                        click_uv_col: "点击UV",
                        buy_uv_col: "购买人数",
                    })
                    if cart_uv_col:
                        df_pre = df_pre.rename(columns={cart_uv_col: "加购UV"})
                    else:
                        df_pre["加购UV"] = np.nan
                    df_pre = to_num(df_pre, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
                    df_pre = df_pre.dropna(subset=["关键词"]).copy()
                    df_pre["购买总金额"] = np.nan
                    df_pre = uv_rates(df_pre)
            
            # 合并两周数据计算环比
            if df_pre is not None and not df_pre.empty:
                merged = df_cur.merge(
                    df_pre[["关键词", "搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]],
                    on="关键词",
                    how="left",
                    suffixes=("_cur", "_pre"),
                )
                # 计算环比
                for col in ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]:
                    if f"{col}_pre" in merged.columns:
                        merged[f"{col}_change"] = (merged[f"{col}_cur"] - merged[f"{col}_pre"]) / merged[f"{col}_pre"].replace(0, np.nan)
                df_cur = merged
            
            # 每个品类只保留top30个自然搜索词
            # 注意：merge后列名可能变成 搜索PV_cur，需要判断
            pv_sort_col = "搜索PV_cur" if "搜索PV_cur" in df_cur.columns else "搜索PV"
            natural_words_list = []
            for cat in CATEGORIES:
                cat_data = df_cur[df_cur["品类"] == cat].nlargest(30, pv_sort_col)
                natural_words_list.append(cat_data)
            natural_words = pd.concat(natural_words_list, ignore_index=True) if natural_words_list else df_cur

    return mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words


def weekly_from_daily(zara_daily_cur: pd.DataFrame, zara_daily_pre: pd.DataFrame = None):
    """从分离的当前周和上周数据计算周环比"""
    # 当前周数据
    cur = zara_daily_cur.copy()
    cur_total = cur[["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]].sum(numeric_only=True)
    cur_start = cur["date"].min()
    cur_end = cur["date"].max()

    # 上周数据
    if zara_daily_pre is not None and not zara_daily_pre.empty:
        pre = zara_daily_pre.copy()
        pre_total = pre[["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]].sum(numeric_only=True)
        pre_start = pre["date"].min()
        pre_end = pre["date"].max()
    else:
        pre = pd.DataFrame()
        pre_total = pd.Series({col: np.nan for col in ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]})
        pre_start = cur_start - pd.Timedelta(days=7)
        pre_end = cur_start - pd.Timedelta(days=1)

    def ratio(s, n, d):
        den = s.get(d, np.nan)
        if den == 0 or pd.isna(den):
            return np.nan
        return s.get(n, np.nan) / den

    metrics = {
        "搜索UV": (cur_total["搜索UV"], pre_total["搜索UV"]),
        "点击UV": (cur_total["点击UV"], pre_total["点击UV"]),
        "加购UV": (cur_total["加购UV"], pre_total["加购UV"]),
        "购买人数": (cur_total["购买人数"], pre_total["购买人数"]),
        "购买总金额": (cur_total["购买总金额"], pre_total["购买总金额"]),
        "CTR": (ratio(cur_total, "点击UV", "搜索UV"), ratio(pre_total, "点击UV", "搜索UV")),
        "ATC": (ratio(cur_total, "加购UV", "搜索UV"), ratio(pre_total, "加购UV", "搜索UV")),
        "CVR": (ratio(cur_total, "购买人数", "搜索UV"), ratio(pre_total, "购买人数", "搜索UV")),
        "UV_VALUE": (ratio(cur_total, "购买总金额", "搜索UV"), ratio(pre_total, "购买总金额", "搜索UV")),
    }

    return {
        "cur": cur,
        "pre": pre,
        "cur_total": cur_total,
        "pre_total": pre_total,
        "metrics": metrics,
        "period": {"cur": f"{cur_start.date()} ~ {cur_end.date()}", "pre": f"{pre_start.date()} ~ {pre_end.date()}"},
    }


def contribution_summary(wk: dict, mini: pd.DataFrame):
    cur_start = wk["cur"]["date"].min()
    cur_end = wk["cur"]["date"].max()
    pre_start = cur_start - pd.Timedelta(days=7)
    pre_end = cur_start - pd.Timedelta(days=1)

    mini_cur = mini[(mini["date"] >= cur_start) & (mini["date"] <= cur_end)]
    mini_pre = mini[(mini["date"] >= pre_start) & (mini["date"] <= pre_end)]

    cur_amt_share = wk["cur_total"]["购买总金额"] / mini_cur["成交金额"].sum() if mini_cur["成交金额"].sum() else np.nan
    pre_amt_share = wk["pre_total"]["购买总金额"] / mini_pre["成交金额"].sum() if mini_pre["成交金额"].sum() else np.nan

    cur_buyer_share = wk["cur_total"]["购买人数"] / mini_cur["成交人数"].sum() if mini_cur["成交人数"].sum() else np.nan
    pre_buyer_share = wk["pre_total"]["购买人数"] / mini_pre["成交人数"].sum() if mini_pre["成交人数"].sum() else np.nan

    cur_uv_value = wk["metrics"]["UV_VALUE"][0]
    pre_uv_value = wk["metrics"]["UV_VALUE"][1]

    daily_cur = wk["cur"][["date", "购买总金额", "购买人数", "搜索UV", "CTR", "ATC", "CVR", "UV_VALUE"]].merge(
        mini[["date", "成交金额", "成交人数"]], on="date", how="left"
    )
    daily_pre = wk["pre"][["date", "购买总金额", "购买人数", "搜索UV", "CTR", "ATC", "CVR", "UV_VALUE"]].merge(
        mini[["date", "成交金额", "成交人数"]], on="date", how="left"
    )

    for d in [daily_cur, daily_pre]:
        d["金额占比"] = d["购买总金额"] / d["成交金额"].replace(0, np.nan)
        d["人数占比"] = d["购买人数"] / d["成交人数"].replace(0, np.nan)

    return {
        "shares": {"金额占比": (cur_amt_share, pre_amt_share), "人数占比": (cur_buyer_share, pre_buyer_share), "UV_VALUE": (cur_uv_value, pre_uv_value)},
        "daily_cur": daily_cur,
        "daily_pre": daily_pre,
    }


def type_weekly_summary(zara_by_type_cur: pd.DataFrame, zara_by_type_pre: pd.DataFrame = None):
    """从分离的当前周和上周数据汇总搜索类型周环比"""
    gcols = ["搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]
    
    cur_g = zara_by_type_cur.groupby("操作类型", as_index=False)[gcols].sum()
    
    if zara_by_type_pre is not None and not zara_by_type_pre.empty:
        pre_g = zara_by_type_pre.groupby("操作类型", as_index=False)[gcols].sum()
    else:
        pre_g = pd.DataFrame(columns=["操作类型"] + gcols)
    
    merged = cur_g.merge(pre_g, on="操作类型", how="outer", suffixes=("_本周", "_上周")).fillna(0)

    merged["CTR_本周"] = merged["点击UV_本周"] / merged["搜索UV_本周"].replace(0, np.nan)
    merged["ATC_本周"] = merged["加购UV_本周"] / merged["搜索UV_本周"].replace(0, np.nan)
    merged["CVR_本周"] = merged["购买人数_本周"] / merged["搜索UV_本周"].replace(0, np.nan)
    merged["CTR_上周"] = merged["点击UV_上周"] / merged["搜索UV_上周"].replace(0, np.nan)
    merged["ATC_上周"] = merged["加购UV_上周"] / merged["搜索UV_上周"].replace(0, np.nan)
    merged["CVR_上周"] = merged["购买人数_上周"] / merged["搜索UV_上周"].replace(0, np.nan)

    total_amt_cur = merged["购买总金额_本周"].sum()
    merged["金额占比_本周"] = merged["购买总金额_本周"] / total_amt_cur if total_amt_cur else np.nan
    return merged


def category_scatter(df: pd.DataFrame, cate: str, title: str):
    if df is None or df.empty or "品类" not in df.columns:
        return None, None
    sub = df[df["品类"] == cate].copy()
    if sub.empty:
        return None, None
    sub = sub.fillna(0)
    
    # 确定搜索PV列名（merge后可能是搜索PV_cur）
    pv_col = "搜索PV_cur" if "搜索PV_cur" in sub.columns else "搜索PV"
    
    # 只显示该品类的top30个关键词（按搜索PV排序）
    sub = sub.nlargest(30, pv_col).copy()
    
    # 在该品类内部进行气泡大小排序（搜索PV最大的词圆圈最大）
    max_pv = sub[pv_col].max()
    if max_pv == 0:
        sub["bubble_size"] = 20
    else:
        sub["bubble_size"] = np.clip(sub[pv_col] / max_pv * 50 + 8, 8, 60)
    
    # 确定其他指标列名
    uv_col = "搜索UV_cur" if "搜索UV_cur" in sub.columns else "搜索UV"
    ctr_col = "CTR_cur" if "CTR_cur" in sub.columns else "CTR"
    cvr_col = "CVR_cur" if "CVR_cur" in sub.columns else "CVR"
    atc_col = "ATC_cur" if "ATC_cur" in sub.columns else "ATC"
    
    avg_ctr = sub[ctr_col].mean()
    avg_cvr = sub[cvr_col].mean()
    word_count = len(sub)
    
    # 显示全部词标签，大小缩小40%
    sub["标签"] = sub["关键词"]
    
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=sub[ctr_col],
            y=sub[cvr_col],
            mode="markers+text",
            text=sub["标签"],
            textposition="top center",
            textfont=dict(size=8),
            marker=dict(size=sub["bubble_size"], sizemode="diameter", opacity=0.65),
            customdata=np.stack([sub["关键词"], sub[pv_col], sub[uv_col]], axis=1),
            hovertemplate="关键词: %{customdata[0]}<br>搜索PV: %{customdata[1]:,.0f}<br>搜索UV: %{customdata[2]:,.0f}<br>CTR: %{x:.2%}<br>CVR: %{y:.2%}<extra></extra>",
        )
    )
    
    # 添加平均值虚线和注释
    if not np.isnan(avg_ctr):
        fig.add_vline(x=avg_ctr, line_dash="dash", line_color="green", opacity=0.6)
        fig.add_annotation(
            x=avg_ctr,
            y=sub[cvr_col].max() if len(sub[cvr_col]) > 0 else 0,
            text=f"平均CTR: {avg_ctr:.2%}",
            showarrow=False,
            bgcolor="rgba(144, 238, 144, 0.7)",
            font=dict(size=10, color="darkgreen"),
            xanchor="center",
            yanchor="bottom",
        )
    
    if not np.isnan(avg_cvr):
        fig.add_hline(y=avg_cvr, line_dash="dash", line_color="green", opacity=0.6)
        fig.add_annotation(
            x=sub[ctr_col].max() if len(sub[ctr_col]) > 0 else 0,
            y=avg_cvr,
            text=f"平均CVR: {avg_cvr:.2%}",
            showarrow=False,
            bgcolor="rgba(144, 238, 144, 0.7)",
            font=dict(size=10, color="darkgreen"),
            xanchor="right",
            yanchor="middle",
        )
    
    fig.add_annotation(
        text=f"图表中词数: {word_count}/全量数据",
        xref="paper",
        yref="paper",
        x=0.02,
        y=0.98,
        showarrow=False,
        bgcolor="rgba(200, 220, 255, 0.7)",
        font=dict(size=9, color="darkblue"),
        xanchor="left",
        yanchor="top",
    )
    
    fig.update_layout(title=title, height=420, margin=dict(l=20, r=20, t=60, b=20), xaxis_title="CTR(点击UV/搜索UV)", yaxis_title="CVR(购买人数/搜索UV)")
    fig.update_xaxes(tickformat=".0%")
    fig.update_yaxes(tickformat=".0%")
    
    # 返回图表和数据表格数据（按搜索PV降序）
    # 选择合适的列组合（有_cur后缀表示有环比数据，否则只有当前周数据）
    if "搜索PV_cur" in sub.columns:
        # 有环比数据
        table_data = sub[["关键词", "搜索PV_cur", "搜索UV_cur", "CTR_cur", "ATC_cur", "CVR_cur", 
                          "搜索PV_change", "搜索UV_change", "CTR_change", "ATC_change", "CVR_change"]].copy()
        table_data.columns = ["关键词", "搜索PV", "搜索UV", "CTR", "ATC", "CVR", 
                             "搜索PV_change", "搜索UV_change", "CTR_change", "ATC_change", "CVR_change"]
    else:
        # 无环比数据
        table_data = sub[["关键词", "搜索PV", "搜索UV", "CTR", "ATC", "CVR"]].copy()
    
    table_data = table_data.sort_values("搜索PV", ascending=False).reset_index(drop=True)
    
    return fig, table_data


def build_pdf_bytes(wk: dict, contrib: dict, by_type: pd.DataFrame):
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        fig = plt.figure(figsize=(11.69, 8.27))
        ax = fig.add_subplot(111)
        ax.axis("off")
        lines = [
            "Search Quality Weekly Report",
            f"Current Week: {wk['period']['cur']}",
            f"Previous Week: {wk['period']['pre']}",
            "",
            f"Search UV: {wk['cur_total']['搜索UV']:.0f} (WoW {safe_pct_change(wk['cur_total']['搜索UV'], wk['pre_total']['搜索UV']):.2%})",
            f"CTR: {wk['metrics']['CTR'][0]:.2%} (WoW {safe_pct_change(wk['metrics']['CTR'][0], wk['metrics']['CTR'][1]):.2%})",
            f"ATC: {wk['metrics']['ATC'][0]:.2%} (WoW {safe_pct_change(wk['metrics']['ATC'][0], wk['metrics']['ATC'][1]):.2%})",
            f"CVR: {wk['metrics']['CVR'][0]:.2%} (WoW {safe_pct_change(wk['metrics']['CVR'][0], wk['metrics']['CVR'][1]):.2%})",
            f"UV Value: {wk['metrics']['UV_VALUE'][0]:.2f} (WoW {safe_pct_change(wk['metrics']['UV_VALUE'][0], wk['metrics']['UV_VALUE'][1]):.2%})",
            "",
            f"Share(Amount): {contrib['shares']['金额占比'][0]:.2%}",
            f"Share(Buyer): {contrib['shares']['人数占比'][0]:.2%}",
            "",
            "By Type CVR (Current Week):",
        ]
        for _, r in by_type.iterrows():
            lines.append(f"- {r['操作类型']}: {r['CVR_本周']:.2%}")
        ax.text(0.02, 0.98, "\n".join(lines), va="top", fontsize=10)
        pdf.savefig(fig)
        plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render():
    st.set_page_config(page_title="搜索引擎质量周报", layout="wide")
    st.title("搜索引擎质量周报（自动化版）")

    with st.sidebar:
        st.subheader("数据文件路径")
        # 过滤掉元数据字段，只显示文件路径
        file_keys = [k for k in DEFAULT_PATHS.keys() if not k.startswith("latest_") and not k.startswith("prev_")]
        display_paths = {k: st.text_input(k, value=DEFAULT_PATHS.get(k, "")) for k in file_keys}
        # 保留元数据字段
        paths = display_paths.copy()
        paths.update({k: v for k, v in DEFAULT_PATHS.items() if k.startswith("latest_") or k.startswith("prev_")})

    required = ["mini", "zara_daily_cur", "zara_by_type_cur", "hot_women_cur", "hot_men_cur", "hot_kids_cur", "hot_home_cur"]
    if not all(Path(paths[k]).exists() for k in required):
        st.error("必需文件路径无效，请检查左侧。")
        st.stop()

    mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words = load_data(paths)
    wk = weekly_from_daily(zara_daily_cur, zara_daily_pre)
    contrib = contribution_summary(wk, mini)

    by_type = type_weekly_summary(zara_by_type_cur, zara_by_type_pre)

    st.caption(f"本周：{wk['period']['cur']} | 上周：{wk['period']['pre']}")

    st.subheader("1) 搜索引擎整体周环比（仅 zara日度数据.xlsx）")
    kpi = ["搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]
    cols = st.columns(len(kpi))
    for i, name in enumerate(kpi):
        cur, pre = wk["metrics"][name]
        if name in ["CTR", "ATC", "CVR"]:
            cols[i].metric(name, fmt_pct(cur), fmt_pct(safe_pct_change(cur, pre)))
        else:
            cols[i].metric(name, fmt_num(cur), fmt_pct(safe_pct_change(cur, pre)))

    # 图1-2: 本周+上周转化率，同色虚线
    rate_colors = {"CTR": "#1f77b4", "ATC": "#2ca02c", "CVR": "#d62728"}
    fig_rate = go.Figure()
    for metric in ["CTR", "ATC", "CVR"]:
        fig_rate.add_trace(go.Scatter(x=wk["cur"]["date"], y=wk["cur"][metric], mode="lines+markers+text", name=f"{metric}-本周", line=dict(color=rate_colors[metric], width=2), text=[f"{v:.1%}" if pd.notna(v) else "-" for v in wk["cur"][metric]], textposition="top center"))
        fig_rate.add_trace(go.Scatter(x=wk["pre"]["date"], y=wk["pre"][metric], mode="lines+markers+text", name=f"{metric}-上周", line=dict(color=rate_colors[metric], dash="dash", width=2), text=[f"{v:.1%}" if pd.notna(v) else "-" for v in wk["pre"][metric]], textposition="bottom center"))
    fig_rate.update_layout(height=380, margin=dict(l=20, r=20, t=20, b=20))
    fig_rate.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig_rate, use_container_width=True)

    st.subheader("2) 搜索对大盘贡献")
    s1, s2, s3 = st.columns(3)
    for i, name in enumerate(["金额占比", "人数占比", "UV_VALUE"]):
        cur, pre = contrib["shares"][name]
        label = "搜索UV价值" if name == "UV_VALUE" else f"搜索{name}"
        target = [s1, s2, s3][i]
        if name == "UV_VALUE":
            target.metric(label, f"{cur:.2f}" if pd.notna(cur) else "-", fmt_pct(safe_pct_change(cur, pre)))
        else:
            target.metric(label, fmt_pct(cur), fmt_pct(safe_pct_change(cur, pre)))

    # 图2-1: 占比（本周实线，上周虚线）
    fig_share = go.Figure()
    for metric, color in [("金额占比", "#1f77b4"), ("人数占比", "#ff7f0e")]:
        fig_share.add_trace(go.Scatter(x=contrib["daily_cur"]["date"], y=contrib["daily_cur"][metric], mode="lines+markers+text", name=f"{metric}-本周", line=dict(color=color), text=[f"{v:.1%}" if pd.notna(v) else "-" for v in contrib["daily_cur"][metric]], textposition="top center"))
        fig_share.add_trace(go.Scatter(x=contrib["daily_pre"]["date"], y=contrib["daily_pre"][metric], mode="lines+markers+text", name=f"{metric}-上周", line=dict(color=color, dash="dash"), text=[f"{v:.1%}" if pd.notna(v) else "-" for v in contrib["daily_pre"][metric]], textposition="bottom center"))
    fig_share.update_layout(height=380, margin=dict(l=20, r=20, t=20, b=20), yaxis=dict(title="占比", tickformat=".0%"))
    st.plotly_chart(fig_share, use_container_width=True)

    # 图2-2: UV_VALUE（本周实线，上周虚线）
    fig_uvv = go.Figure()
    fig_uvv.add_trace(go.Scatter(x=contrib["daily_cur"]["date"], y=contrib["daily_cur"]["UV_VALUE"], mode="lines+markers+text", name="UV_VALUE-本周", line=dict(color="#2ca02c"), text=[f"{v:.1f}" if pd.notna(v) else "-" for v in contrib["daily_cur"]["UV_VALUE"]], textposition="top center"))
    fig_uvv.add_trace(go.Scatter(x=contrib["daily_pre"]["date"], y=contrib["daily_pre"]["UV_VALUE"], mode="lines+markers+text", name="UV_VALUE-上周", line=dict(color="#2ca02c", dash="dash"), text=[f"{v:.1f}" if pd.notna(v) else "-" for v in contrib["daily_pre"]["UV_VALUE"]], textposition="bottom center"))
    fig_uvv.update_layout(height=320, margin=dict(l=20, r=20, t=20, b=20), yaxis=dict(title="UV_VALUE"))
    st.plotly_chart(fig_uvv, use_container_width=True)

    st.subheader("3) 搜索类型周环比（全链路）")
    # 图3-1: 购买总金额+占比标签
    fig_t1 = go.Figure()
    total_amt_pre = by_type["购买总金额_上周"].sum()
    if total_amt_pre and not pd.isna(total_amt_pre) and total_amt_pre != 0:
        pre_share = by_type["购买总金额_上周"] / total_amt_pre
    else:
        pre_share = pd.Series([np.nan] * len(by_type), index=by_type.index)
    text_cur = [f"{v:,.0f}\n({s:.1%})" for v, s in zip(by_type["购买总金额_本周"], by_type["金额占比_本周"])]
    text_pre = [f"{v:,.0f}\n({s:.1%})" if pd.notna(s) else f"{v:,.0f}" for v, s in zip(by_type["购买总金额_上周"], pre_share)]
    fig_t1.add_trace(go.Bar(x=by_type["操作类型"], y=by_type["购买总金额_本周"], name="本周", marker_color="#4c78a8", text=text_cur, textposition="outside"))
    fig_t1.add_trace(go.Bar(x=by_type["操作类型"], y=by_type["购买总金额_上周"], name="上周", marker_color="#9ecae9", text=text_pre, textposition="outside"))
    fig_t1.update_layout(height=380, barmode="group", margin=dict(l=20, r=20, t=20, b=20))
    st.plotly_chart(fig_t1, use_container_width=True)

    # 图3-2: 同指标一起，本周浅色、上周深色
    metrics_cfg = [
        ("CTR", "#9ecae9", "#1f77b4"),
        ("ATC", "#a1d99b", "#2ca02c"),
        ("CVR", "#fdae6b", "#d62728"),
    ]
    fig_t2 = make_subplots(rows=1, cols=3, subplot_titles=["CTR", "ATC", "CVR"])
    for idx, (m, c_cur, c_pre) in enumerate(metrics_cfg, start=1):
        fig_t2.add_trace(go.Bar(x=by_type["操作类型"], y=by_type[f"{m}_本周"], name=f"{m}-本周", marker_color=c_cur, text=[f"{v:.1%}" if pd.notna(v) else "-" for v in by_type[f"{m}_本周"]], textposition="outside", showlegend=(idx == 1)), row=1, col=idx)
        fig_t2.add_trace(go.Bar(x=by_type["操作类型"], y=by_type[f"{m}_上周"], name=f"{m}-上周", marker_color=c_pre, text=[f"{v:.1%}" if pd.notna(v) else "-" for v in by_type[f"{m}_上周"]], textposition="outside", showlegend=(idx == 1)), row=1, col=idx)
        fig_t2.update_yaxes(tickformat=".0%", row=1, col=idx)
    fig_t2.update_layout(height=430, barmode="group", margin=dict(l=20, r=20, t=50, b=20))
    st.plotly_chart(fig_t2, use_container_width=True)

    st.dataframe(
        by_type[["操作类型", "搜索UV_本周", "点击UV_本周", "加购UV_本周", "购买人数_本周", "CTR_本周", "ATC_本周", "CVR_本周", "金额占比_本周"]].style.format(
            {
                "搜索UV_本周": "{:,.0f}", "点击UV_本周": "{:,.0f}", "加购UV_本周": "{:,.0f}", "购买人数_本周": "{:,.0f}",
                "CTR_本周": "{:.2%}", "ATC_本周": "{:.2%}", "CVR_本周": "{:.2%}", "金额占比_本周": "{:.2%}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("4) 热词与自然词分析")
    
    for cate in CATEGORIES:
        st.markdown(f"### {cate}")
        st.markdown(f"**{cate} - 热词分析**")
        fig_hot, data_hot = category_scatter(hotwords, cate, f"{cate} 热词：CTR vs CVR（气泡=搜索PV）")
        if fig_hot is None:
            st.info("该品类暂无热词数据。")
        else:
            st.plotly_chart(fig_hot, use_container_width=True)
            # 显示热词数据表格（带环比信息，环比显示在同一单元格内）
            table_data, styled_cols = display_keyword_table(data_hot)
            styled_df = table_data.style.applymap(style_change_cell, subset=styled_cols)
            st.dataframe(
                styled_df,
                use_container_width=True,
                height=400,
                hide_index=True,
            )

        st.markdown(f"**{cate} - 自然词分析**")
        fig_nat, data_nat = category_scatter(natural_words, cate, f"{cate} 自然词：CTR vs CVR（气泡=搜索PV）")
        if fig_nat is None:
            st.info("该品类暂无自然词数据（按末尾品类提取后为空）。")
        else:
            st.plotly_chart(fig_nat, use_container_width=True)
            # 显示自然词数据表格（带环比信息，环比显示在同一单元格内）
            table_data_nat, styled_cols_nat = display_keyword_table(data_nat)
            styled_df_nat = table_data_nat.style.applymap(style_change_cell, subset=styled_cols_nat)
            st.dataframe(
                styled_df_nat,
                use_container_width=True,
                height=400,
                hide_index=True,
            )

    st.subheader("5) 导出 PDF")
    pdf_bytes = build_pdf_bytes(wk, contrib, by_type)
    st.download_button(
        label="生成并下载PDF文件",
        data=pdf_bytes,
        file_name=f"search_quality_report_{wk['cur']['date'].max().date()}.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    render()
