import io
import json
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from matplotlib import pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from plotly.subplots import make_subplots

# 固定数据源：优先使用真实数据目录「zara周报数据源」，无则回退到「演示」
DEFAULT_BASE_DIR_CANDIDATES = [
    Path(__file__).parent / "zara周报数据源",
    Path(__file__).parent / "演示",
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


def find_homepage_config_periods(base_dir: Path) -> tuple:
    """扫描首页配置词文件夹找到最新和次新的周期（目录名以「配置词」结尾）。"""
    cfg_dir = base_dir / "首页配置词部分"
    if not cfg_dir.exists():
        return None, None
    folders = sorted([d for d in cfg_dir.iterdir() if d.is_dir() and d.name.endswith("配置词")])
    if len(folders) >= 2:
        latest = folders[-1].name.replace("配置词", "")
        prev = folders[-2].name.replace("配置词", "")
        return latest, prev
    if len(folders) == 1:
        return folders[0].name.replace("配置词", ""), None
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
    home_cfg_latest, home_cfg_prev = find_homepage_config_periods(base_dir)
    nat_latest, nat_prev = find_natural_word_periods(base_dir)
    
    # 小程序大盘：自动从文件夹中选取最新的 Excel 文件
    mini_dir = base_dir / "小程序大盘部分"
    mini_files = sorted(mini_dir.glob("*.xlsx")) if mini_dir.exists() else []
    mini_path = str(mini_files[-1]) if mini_files else ""

    paths = {
        "mini": mini_path,
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
        # 首页配置词（独立周期；缺失时不影响整站启动）
        "home_cfg_women_cur": str(base_dir / "首页配置词部分" / f"{home_cfg_latest}配置词" / "女士配置.xlsx") if home_cfg_latest else "",
        "home_cfg_men_cur": str(base_dir / "首页配置词部分" / f"{home_cfg_latest}配置词" / "男士配置.xlsx") if home_cfg_latest else "",
        "home_cfg_kids_cur": str(base_dir / "首页配置词部分" / f"{home_cfg_latest}配置词" / "儿童配置.xlsx") if home_cfg_latest else "",
        "home_cfg_home_cur": str(base_dir / "首页配置词部分" / f"{home_cfg_latest}配置词" / "家居配置.xlsx") if home_cfg_latest else "",
        "home_cfg_women_pre": str(base_dir / "首页配置词部分" / f"{home_cfg_prev}配置词" / "女士配置.xlsx") if home_cfg_prev else "",
        "home_cfg_men_pre": str(base_dir / "首页配置词部分" / f"{home_cfg_prev}配置词" / "男士配置.xlsx") if home_cfg_prev else "",
        "home_cfg_kids_pre": str(base_dir / "首页配置词部分" / f"{home_cfg_prev}配置词" / "儿童配置.xlsx") if home_cfg_prev else "",
        "home_cfg_home_pre": str(base_dir / "首页配置词部分" / f"{home_cfg_prev}配置词" / "家居配置.xlsx") if home_cfg_prev else "",
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

# 自然搜索词：女士 Top100（图表按每页 20 词分页）；其余品类仍 Top30
NATURAL_TOPN_BY_CATE = {"女士": 100, "男士": 30, "儿童": 30, "家居": 30}
NATURAL_WOMEN_PAGE_SIZE = 20


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
    # 仅保留整数百分比（带正负号）
    return f"{val:.0%}"


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
        # 旧格式：带箭头符号
        if '↑' in val:
            return 'color: green'
        elif '↓' in val:
            return 'color: red'
        # 新格式：纯百分比，如 "5%"、"-12%"
        if val.endswith("%"):
            try:
                num = float(val.replace("%", ""))
                frac = num / 100.0
                # 仅当绝对值 >= 30% 时着色
                if frac >= 0.3:
                    return 'color: green'
                elif frac <= -0.3:
                    return 'color: red'
            except ValueError:
                pass
    return ''


def apply_styler_change(styler, subset_cols):
    """Streamlit Cloud 等环境多为 pandas≥2.1，Styler.applymap 已移除，须用 map；旧版仍用 applymap。"""
    if hasattr(styler, "map"):
        return styler.map(style_change_cell, subset=subset_cols)
    return styler.applymap(style_change_cell, subset=subset_cols)


def display_keyword_table(data: pd.DataFrame):
    """准备包含环比信息的表格数据，环比显示在同一单元格内"""
    if data.empty:
        return data, []
    
    output_data = data.copy()
    styled_cols = []
    
    # 计算综合得分：CTR*5 + ATC*4 + CVR*3
    # 注意：这里按照「去掉百分号后的数值」来算，
    # 即 57% 作为 57 而不是 0.57，所以要先乘以 100
    score_cur = None
    score_pre = None
    if all(col in data.columns for col in ["CTR", "ATC", "CVR"]):
        # 当前周得分
        score_cur = (
            data["CTR"].fillna(0) * 100 * 5
            + data["ATC"].fillna(0) * 100 * 4
            + data["CVR"].fillna(0) * 100 * 3
        )
        output_data["得分"] = score_cur.round(0).astype(int)

        # 如果有 CTR/ATC/CVR 的环比列，则反推上一周得分并计算得分环比
        if all(col in data.columns for col in ["CTR_change", "ATC_change", "CVR_change"]):
            # pre = cur / (1 + change)，其中 change = (cur - pre) / pre
            with np.errstate(divide="ignore", invalid="ignore"):
                ctr_pre = data["CTR"] / (1 + data["CTR_change"])
                atc_pre = data["ATC"] / (1 + data["ATC_change"])
                cvr_pre = data["CVR"] / (1 + data["CVR_change"])

            score_pre = (
                ctr_pre.fillna(0) * 100 * 5
                + atc_pre.fillna(0) * 100 * 4
                + cvr_pre.fillna(0) * 100 * 3
            )
            with np.errstate(divide="ignore", invalid="ignore"):
                score_change = (score_cur - score_pre) / score_pre.replace(0, np.nan)
            output_data["得分_change"] = score_change
    
    # 上架天数（热词表常见字段）：仅展示当前值，不参与环比
    if "上架天数" in output_data.columns:
        output_data["上架天数"] = output_data["上架天数"].apply(
            lambda x: f"{x:,.0f}" if pd.notna(x) else "-"
        )

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
            # 主列只显示当前值，新增一列单独显示环比
            if is_pct:
                # 百分比主值精度：CVR 保留一位小数，其它（CTR/ATC）取整
                if col == "CVR":
                    output_data[col] = output_data[col].apply(
                        lambda x: f"{x:.1%}" if pd.notna(x) else "-"
                    )
                else:
                    output_data[col] = output_data[col].apply(
                        lambda x: f"{x:.0%}" if pd.notna(x) else "-"
                    )
            else:
                output_data[col] = output_data[col].apply(
                    lambda x: f"{x:,.0f}" if pd.notna(x) else "-"
                )
            change_display_col = f"{col}_环比"
            # 环比列按指标区分精度：CVR 保留一位小数，其它（CTR/ATC）保留整数
            if col == "CVR":
                output_data[change_display_col] = output_data[change_col].apply(
                    lambda v: "" if pd.isna(v) else f"{v:.1%}"
                )
            else:
                output_data[change_display_col] = output_data[change_col].apply(
                    lambda v: "" if pd.isna(v) else f"{v:.0%}"
                )
            styled_cols.append(change_display_col)
        elif col in output_data.columns:
            # 没有环比数据，只格式化主值
            if is_pct:
                if col == "CVR":
                    output_data[col] = output_data[col].apply(
                        lambda x: f"{x:.1%}" if pd.notna(x) else "-"
                    )
                else:
                    output_data[col] = output_data[col].apply(
                        lambda x: f"{x:.0%}" if pd.notna(x) else "-"
                    )
            else:
                output_data[col] = output_data[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "-")
    
    # 只保留需要显示的列：指标值 + 对应环比列（如果存在）
    display_cols = ["关键词"]
    if "上架天数" in output_data.columns:
        display_cols.append("上架天数")
    for col, _ in metrics_config:
        if col in output_data.columns:
            display_cols.append(col)
        change_display_col = f"{col}_环比"
        if change_display_col in output_data.columns:
            display_cols.append(change_display_col)
    if "得分" in output_data.columns:
        display_cols.append("得分")
        if "得分_change" in output_data.columns:
            # 得分环比列，统一用整数百分比
            output_data["得分_环比"] = output_data["得分_change"].apply(
                lambda v: "" if pd.isna(v) else f"{v:.0%}"
            )
            display_cols.append("得分_环比")
            styled_cols.append("得分_环比")
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


def read_excel_checked(path: str, label: str, header=0):
    """读取 Excel；对 .xlsx 校验 ZIP 文件头（PK），避免把接口错误 JSON 等误存为 .xlsx 导致 pandas 报「无法判断格式」。"""
    p = Path(path)
    if not path or not str(path).strip():
        raise FileNotFoundError(f"{label}：路径为空，请检查侧栏「mini」或数据源「小程序大盘部分」下是否有有效 .xlsx。")
    if not p.is_file():
        raise FileNotFoundError(f"{label}：文件不存在：{p}")
    suf = p.suffix.lower()
    if suf not in (".xlsx", ".xls"):
        raise ValueError(f"{label}：需要 .xlsx 或 .xls，当前为 {suf!r}：{p}")
    with p.open("rb") as f:
        sig = f.read(4)
    if suf == ".xlsx":
        if not sig.startswith(b"PK"):
            raise ValueError(
                f"{label}：内容不是有效 xlsx（应以 PK 开头）。常见原因：导出失败，文件实为 JSON/HTML 错误页。"
                f"请重新下载或从 Excel 另存为真实表格。\n路径：{p}"
            )
        return pd.read_excel(p, engine="openpyxl", header=header)
    if not sig.startswith(b"\xd0\xcf\x11\xe0"):
        raise ValueError(f"{label}：内容不是有效 .xls：{p}")
    return pd.read_excel(p, engine="xlrd", header=header)


@st.cache_data(show_spinner=False)
def load_data(paths: dict):
    mini = read_excel_checked(paths["mini"], "小程序大盘").copy()
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
        # 当前周热词（路径须非空：Path(\"\") 在部分环境下会解析为 cwd 导致误读）
        if fp_cur in paths and paths[fp_cur] and Path(paths[fp_cur]).exists():
            df_cur = pd.read_excel(paths[fp_cur], header=2).copy()
            if kw_col in df_cur.columns:
                df_cur = df_cur.dropna(subset=[kw_col]).copy()
                df_cur = df_cur.rename(columns={kw_col: "关键词", "购买UV": "购买人数"})
                days_col = find_col(df_cur, ["上架天数", "上架天数(天)"])
                if days_col and days_col != "上架天数":
                    df_cur = df_cur.rename(columns={days_col: "上架天数"})
                if "上架天数" in df_cur.columns:
                    df_cur["上架天数"] = pd.to_numeric(df_cur["上架天数"], errors="coerce")
                df_cur["品类"] = category
                df_cur = to_num(df_cur, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
                df_cur["购买总金额"] = np.nan
                df_cur = uv_rates(df_cur)
                
                # 加载上周热词
                df_pre = None
                if fp_pre in paths and paths[fp_pre] and Path(paths[fp_pre]).exists():
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

    home_cfg_keys = [
        ("女士", "home_cfg_women_cur", "home_cfg_women_pre"),
        ("男士", "home_cfg_men_cur", "home_cfg_men_pre"),
        ("儿童", "home_cfg_kids_cur", "home_cfg_kids_pre"),
        ("家居", "home_cfg_home_cur", "home_cfg_home_pre"),
    ]
    home_frames = []
    for category, fp_cur, fp_pre in home_cfg_keys:
        if fp_cur in paths and paths[fp_cur] and Path(paths[fp_cur]).exists():
            df_cur = pd.read_excel(paths[fp_cur], header=2).copy()
            kw_col = find_col(
                df_cur,
                ["自然搜索词", "关键词", f"{category}配置词分类", f"{category}热词分类", "配置词"],
            )
            if not kw_col:
                continue
            df_cur = df_cur.dropna(subset=[kw_col]).copy()
            df_cur = df_cur.rename(columns={kw_col: "关键词", "购买UV": "购买人数"})
            days_col = find_col(df_cur, ["上架天数", "上架天数(天)"])
            if days_col and days_col != "上架天数":
                df_cur = df_cur.rename(columns={days_col: "上架天数"})
            if "上架天数" in df_cur.columns:
                df_cur["上架天数"] = pd.to_numeric(df_cur["上架天数"], errors="coerce")
            df_cur["品类"] = category
            df_cur = to_num(df_cur, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
            df_cur["购买总金额"] = np.nan
            df_cur = uv_rates(df_cur)

            df_pre = None
            if fp_pre in paths and paths[fp_pre] and Path(paths[fp_pre]).exists():
                df_pre = pd.read_excel(paths[fp_pre], header=2).copy()
                kw_pre = find_col(
                    df_pre,
                    ["自然搜索词", "关键词", f"{category}配置词分类", f"{category}热词分类", "配置词"],
                )
                if kw_pre:
                    df_pre = df_pre.dropna(subset=[kw_pre]).copy()
                    df_pre = df_pre.rename(columns={kw_pre: "关键词", "购买UV": "购买人数"})
                    df_pre = to_num(df_pre, ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数"])
                    df_pre["购买总金额"] = np.nan
                    df_pre = uv_rates(df_pre)

            if df_pre is not None and not df_pre.empty:
                merged = df_cur.merge(
                    df_pre[["关键词", "搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]],
                    on="关键词",
                    how="left",
                    suffixes=("_cur", "_pre"),
                )
                for col in ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]:
                    merged[f"{col}_change"] = (merged[f"{col}_cur"] - merged[f"{col}_pre"]) / merged[f"{col}_pre"].replace(0, np.nan)
                df_cur = merged
            else:
                for col in ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "CTR", "ATC", "CVR"]:
                    df_cur[f"{col}_cur"] = df_cur[col]
                    df_cur[f"{col}_change"] = np.nan

            home_frames.append(df_cur)

    home_config_words = pd.concat(home_frames, ignore_index=True) if home_frames else pd.DataFrame()

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
            
            # 每个品类保留 TopN 个自然搜索词（女士 100，其余 30）
            # 注意：merge后列名可能变成 搜索PV_cur，需要判断
            pv_sort_col = "搜索PV_cur" if "搜索PV_cur" in df_cur.columns else "搜索PV"
            natural_words_list = []
            for cat in CATEGORIES:
                topn = NATURAL_TOPN_BY_CATE.get(cat, 30)
                cat_data = df_cur[df_cur["品类"] == cat].nlargest(topn, pv_sort_col)
                natural_words_list.append(cat_data)
            natural_words = pd.concat(natural_words_list, ignore_index=True) if natural_words_list else df_cur

    return mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words, home_config_words


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


def attach_keyword_score_change(df: pd.DataFrame) -> pd.DataFrame:
    """为关键词明细补充「得分_change」，供散点分色与表格分栏（与 display_keyword_table 口径一致）。"""
    if df is None or df.empty:
        return df
    out = df.copy()
    if "得分_change" in out.columns:
        return out
    ctr_c = "CTR_cur" if "CTR_cur" in out.columns else "CTR"
    atc_c = "ATC_cur" if "ATC_cur" in out.columns else "ATC"
    cvr_c = "CVR_cur" if "CVR_cur" in out.columns else "CVR"
    if not all(c in out.columns for c in [ctr_c, atc_c, cvr_c]):
        return out
    score_cur = (
        out[ctr_c].fillna(0) * 100 * 5
        + out[atc_c].fillna(0) * 100 * 4
        + out[cvr_c].fillna(0) * 100 * 3
    )
    if all(c in out.columns for c in ["CTR_change", "ATC_change", "CVR_change"]):
        with np.errstate(divide="ignore", invalid="ignore"):
            ctr_pre = out[ctr_c] / (1 + out["CTR_change"])
            atc_pre = out[atc_c] / (1 + out["ATC_change"])
            cvr_pre = out[cvr_c] / (1 + out["CVR_change"])
        score_pre = (
            ctr_pre.fillna(0) * 100 * 5
            + atc_pre.fillna(0) * 100 * 4
            + cvr_pre.fillna(0) * 100 * 3
        )
        with np.errstate(divide="ignore", invalid="ignore"):
            out["得分_change"] = (score_cur - score_pre) / score_pre.replace(0, np.nan)
    return out


def split_keyword_perf_groups(df: pd.DataFrame):
    """按得分环比拆分好/差词；无得分环比时退回 CTR 环比。环比为 NaN 的归入「较好」一侧以便仍能展示。"""
    if df is None or df.empty:
        return pd.DataFrame(), pd.DataFrame(), "无环比"
    d = attach_keyword_score_change(df.copy())
    if "得分_change" in d.columns:
        good = d[(d["得分_change"] >= 0) | d["得分_change"].isna()]
        bad = d[d["得分_change"] < 0]
        return good, bad, "得分环比"
    if "CTR_change" in d.columns:
        good = d[(d["CTR_change"] >= 0) | d["CTR_change"].isna()]
        bad = d[d["CTR_change"] < 0]
        return good, bad, "CTR环比"
    return d, pd.DataFrame(), "无环比"


def type_weekly_summary(zara_by_type_cur: pd.DataFrame, zara_by_type_pre: pd.DataFrame = None):
    """从分离的当前周和上周数据汇总搜索类型周环比。

    口径（与「先日均再平均」不同）：每个操作类型在本周/上周内，先把约 7 天的 UV 等指标按列求和，
    再计算比率——即加权口径：
    - CTR = sum(点击UV) / sum(搜索UV)
    - ATC = sum(加购UV) / sum(搜索UV)
    - CVR = sum(购买人数) / sum(搜索UV)
    金额占比等为 sum(购买总金额) 在类型间的占比，同理基于周汇总而非日均平均。
    """
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


def category_scatter(
    df: pd.DataFrame,
    cate: str,
    title: str,
    top_n=30,
    plot_df=None,
    ref_pool_for_avg=None,
    bubble_max_df=None,
):
    """
    散点图：默认取该品类 TopN 词。若传入 plot_df，则只绘制该子集；
    ref_pool_for_avg 用于平均 CTR/CVR 虚线（如女士 Top100 全量）；
    bubble_max_df 若指定则用其 PV 最大值归一气泡（如女士分页时传 Top100 全池，各页气泡与全榜可比）；默认 None 表示用当前子集。
    """
    if plot_df is not None:
        sub = plot_df.copy()
        if sub.empty:
            return None, None
    else:
        if df is None or df.empty or "品类" not in df.columns:
            return None, None
        sub = df[df["品类"] == cate].copy()
        if sub.empty:
            return None, None
        pv_pick = "搜索PV_cur" if "搜索PV_cur" in sub.columns else "搜索PV"
        sub = sub.nlargest(top_n, pv_pick).copy()

    sub = sub.fillna(0)

    pv_col = "搜索PV_cur" if "搜索PV_cur" in sub.columns else "搜索PV"

    pool_avg = ref_pool_for_avg.copy() if ref_pool_for_avg is not None else sub.copy()
    pool_avg = pool_avg.fillna(0)

    max_src = bubble_max_df.copy() if bubble_max_df is not None else sub.copy()
    max_src = max_src.fillna(0)
    if len(max_src) and pv_col in max_src.columns:
        max_pv = max_src[pv_col].max()
    else:
        max_pv = 0
    if max_pv == 0 and len(sub) and pv_col in sub.columns:
        max_pv = sub[pv_col].max()

    if max_pv == 0:
        sub["bubble_size"] = 20
    else:
        sub["bubble_size"] = np.clip(sub[pv_col] / max_pv * 50 + 8, 8, 60)

    uv_col = "搜索UV_cur" if "搜索UV_cur" in sub.columns else "搜索UV"
    ctr_col = "CTR_cur" if "CTR_cur" in sub.columns else "CTR"
    cvr_col = "CVR_cur" if "CVR_cur" in sub.columns else "CVR"
    atc_col = "ATC_cur" if "ATC_cur" in sub.columns else "ATC"

    avg_ctr = pool_avg[ctr_col].mean() if len(pool_avg) else np.nan
    avg_cvr = pool_avg[cvr_col].mean() if len(pool_avg) else np.nan
    word_count = len(sub)
    pool_total = len(pool_avg) if ref_pool_for_avg is not None else word_count
    
    # 显示全部词标签，大小缩小40%
    sub["标签"] = sub["关键词"]
    sub = attach_keyword_score_change(sub)

    fig = go.Figure()

    # 根据「得分环比」区分表现好/差的词，使用不同颜色（若无得分环比，则退化为 CTR 环比）
    has_score_change = "得分_change" in sub.columns
    has_ctr_change = "CTR_change" in sub.columns
    if has_score_change or has_ctr_change:
        if has_score_change:
            good_mask = sub["得分_change"] >= 0
            bad_mask = sub["得分_change"] < 0
        else:
            good_mask = sub["CTR_change"] >= 0
            bad_mask = sub["CTR_change"] < 0

        good_sub = sub[good_mask]
        bad_sub = sub[bad_mask]

        if not good_sub.empty:
            fig.add_trace(
                go.Scatter(
                    x=good_sub[ctr_col],
                    y=good_sub[cvr_col],
                    mode="markers+text",
                    name="表现上升（得分环比 ≥ 0）",
                    text=good_sub["标签"],
                    textposition="top center",
                    textfont=dict(size=8),
                    marker=dict(
                        size=good_sub["bubble_size"],
                        sizemode="diameter",
                        opacity=0.7,
                        color="#1f77b4",  # 蓝色：好
                    ),
                    customdata=np.stack(
                        [good_sub["关键词"], good_sub[pv_col], good_sub[uv_col]], axis=1
                    ),
                    hovertemplate="关键词: %{customdata[0]}<br>搜索PV: %{customdata[1]:,.0f}<br>搜索UV: %{customdata[2]:,.0f}<br>CTR: %{x:.2%}<br>CVR: %{y:.2%}<extra></extra>",
                )
            )

        if not bad_sub.empty:
            fig.add_trace(
                go.Scatter(
                    x=bad_sub[ctr_col],
                    y=bad_sub[cvr_col],
                    mode="markers+text",
                    name="表现下降（得分环比 < 0）",
                    text=bad_sub["标签"],
                    textposition="top center",
                    textfont=dict(size=8),
                    marker=dict(
                        size=bad_sub["bubble_size"],
                        sizemode="diameter",
                        opacity=0.7,
                        color="#d62728",  # 红色：差
                    ),
                    customdata=np.stack(
                        [bad_sub["关键词"], bad_sub[pv_col], bad_sub[uv_col]], axis=1
                    ),
                    hovertemplate="关键词: %{customdata[0]}<br>搜索PV: %{customdata[1]:,.0f}<br>搜索UV: %{customdata[2]:,.0f}<br>CTR: %{x:.2%}<br>CVR: %{y:.2%}<extra></extra>",
                )
            )

        # 如果某一类为空（全部好或全部差），避免图例缺失时看不懂颜色含义
        if good_sub.empty or bad_sub.empty:
            fig.update_layout(showlegend=True)
    else:
        # 无环比数据时，保持统一颜色
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
    
    if ref_pool_for_avg is not None:
        ann_note = f"本页词数: {word_count}/{pool_total}（均值线=Top{pool_total} 平均CTR/CVR）"
    else:
        ann_note = f"图表中词数: {word_count}/全量数据"
    fig.add_annotation(
        text=ann_note,
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
        cols = [
            "关键词", "搜索PV_cur", "搜索UV_cur", "CTR_cur", "ATC_cur", "CVR_cur",
            "搜索PV_change", "搜索UV_change", "CTR_change", "ATC_change", "CVR_change",
        ]
        if "上架天数" in sub.columns:
            cols.insert(1, "上架天数")
        if "得分_change" in sub.columns:
            cols.append("得分_change")
        table_data = sub[cols].copy()
        rename = {
            "关键词": "关键词",
            "搜索PV_cur": "搜索PV",
            "搜索UV_cur": "搜索UV",
            "CTR_cur": "CTR",
            "ATC_cur": "ATC",
            "CVR_cur": "CVR",
            "搜索PV_change": "搜索PV_change",
            "搜索UV_change": "搜索UV_change",
            "CTR_change": "CTR_change",
            "ATC_change": "ATC_change",
            "CVR_change": "CVR_change",
        }
        if "得分_change" in table_data.columns:
            rename["得分_change"] = "得分_change"
        table_data = table_data.rename(columns=rename)
    else:
        # 无环比数据
        base_cols = ["关键词", "搜索PV", "搜索UV", "CTR", "ATC", "CVR"]
        if "上架天数" in sub.columns:
            base_cols.insert(1, "上架天数")
        table_data = sub[base_cols].copy()
    
    table_data = table_data.sort_values("搜索PV", ascending=False).reset_index(drop=True)
    
    return fig, table_data


def women_natural_top_pool(natural_words: pd.DataFrame) -> pd.DataFrame:
    """女士自然词 TopN（与 load_data 一致，按搜索PV），供分页图表与离线 HTML。"""
    if natural_words is None or natural_words.empty:
        return pd.DataFrame()
    w = natural_words[natural_words["品类"] == "女士"].copy()
    if w.empty:
        return w
    pv = "搜索PV_cur" if "搜索PV_cur" in w.columns else "搜索PV"
    n = NATURAL_TOPN_BY_CATE.get("女士", 100)
    return w.nlargest(n, pv).reset_index(drop=True)


# ---------- 数据校验 ----------
REL_TOL = 1e-5
ABS_TOL = 1e-8
BASELINE_REL_TOL = 0.001
BASELINE_PATH = Path(__file__).parent / "演示" / "search_quality_baseline.json"


def _num_close(a, b, rel_tol=REL_TOL, abs_tol=ABS_TOL):
    """数值近似相等（支持 NaN）。"""
    if pd.isna(a) and pd.isna(b):
        return True
    if pd.isna(a) or pd.isna(b):
        return False
    if a == 0 and b == 0:
        return True
    return np.isclose(float(a), float(b), rtol=rel_tol, atol=abs_tol)


def run_display_consistency_checks(wk, contrib, by_type, zara_daily_cur, zara_daily_pre, mini, zara_by_type_cur, zara_by_type_pre):
    """校验页面展示的数值与从源数据重新计算的结果一致。返回 [(check_name, passed, detail_str), ...]"""
    results = []
    gcols = ["搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]

    # 1) 整体周环比：cur_total / pre_total
    cur_total_recalc = zara_daily_cur[gcols].sum()
    for col in gcols:
        ok = _num_close(wk["cur_total"][col], cur_total_recalc[col])
        results.append((f"本周合计-{col}", ok, f"展示={wk['cur_total'][col]}, 重算={cur_total_recalc[col]}"))

    if zara_daily_pre is not None and not zara_daily_pre.empty:
        pre_total_recalc = zara_daily_pre[gcols].sum()
        for col in gcols:
            ok = _num_close(wk["pre_total"][col], pre_total_recalc[col])
            results.append((f"上周合计-{col}", ok, f"展示={wk['pre_total'][col]}, 重算={pre_total_recalc[col]}"))
    else:
        results.append(("上周合计", True, "无上周数据，跳过"))

    # 2) 周汇总比率：CTR/ATC/CVR/UV_VALUE
    def ratio(s, num, den):
        d = s.get(den, np.nan)
        if d == 0 or pd.isna(d):
            return np.nan
        return s.get(num, np.nan) / d
    cur_total = wk["cur_total"]
    pre_total = wk["pre_total"]
    for name, num, den in [("CTR", "点击UV", "搜索UV"), ("ATC", "加购UV", "搜索UV"), ("CVR", "购买人数", "搜索UV"), ("UV_VALUE", "购买总金额", "搜索UV")]:
        cur_r = ratio(cur_total, num, den)
        pre_r = ratio(pre_total, num, den)
        ok_cur = _num_close(wk["metrics"][name][0], cur_r)
        ok_pre = _num_close(wk["metrics"][name][1], pre_r)
        results.append((f"metrics-{name}-本周", ok_cur, f"展示={wk['metrics'][name][0]}, 重算={cur_r}"))
        results.append((f"metrics-{name}-上周", ok_pre, f"展示={wk['metrics'][name][1]}, 重算={pre_r}"))

    # 3) 贡献占比
    cur_start = wk["cur"]["date"].min()
    cur_end = wk["cur"]["date"].max()
    pre_start = cur_start - pd.Timedelta(days=7)
    pre_end = cur_start - pd.Timedelta(days=1)
    mini_cur = mini[(mini["date"] >= cur_start) & (mini["date"] <= cur_end)]
    mini_pre = mini[(mini["date"] >= pre_start) & (mini["date"] <= pre_end)]
    cur_amt_sum = mini_cur["成交金额"].sum()
    pre_amt_sum = mini_pre["成交金额"].sum()
    cur_buyer_sum = mini_cur["成交人数"].sum()
    pre_buyer_sum = mini_pre["成交人数"].sum()
    cur_amt_share_recalc = (wk["cur_total"]["购买总金额"] / cur_amt_sum) if cur_amt_sum else np.nan
    pre_amt_share_recalc = (wk["pre_total"]["购买总金额"] / pre_amt_sum) if pre_amt_sum else np.nan
    cur_buyer_share_recalc = (wk["cur_total"]["购买人数"] / cur_buyer_sum) if cur_buyer_sum else np.nan
    pre_buyer_share_recalc = (wk["pre_total"]["购买人数"] / pre_buyer_sum) if pre_buyer_sum else np.nan
    results.append(("贡献-金额占比-本周", _num_close(contrib["shares"]["金额占比"][0], cur_amt_share_recalc), ""))
    results.append(("贡献-金额占比-上周", _num_close(contrib["shares"]["金额占比"][1], pre_amt_share_recalc), ""))
    results.append(("贡献-人数占比-本周", _num_close(contrib["shares"]["人数占比"][0], cur_buyer_share_recalc), ""))
    results.append(("贡献-人数占比-上周", _num_close(contrib["shares"]["人数占比"][1], pre_buyer_share_recalc), ""))

    # 4) 按类型：by_type 与 zara_by_type_cur groupby 一致
    cur_g = zara_by_type_cur.groupby("操作类型", as_index=False)[gcols].sum()
    total_amt_cur = cur_g["购买总金额"].sum()
    for _, row in cur_g.iterrows():
        op = row["操作类型"]
        bt_row = by_type[by_type["操作类型"] == op]
        if bt_row.empty:
            results.append((f"按类型-{op}-存在", False, "当前结果中缺失该操作类型"))
            continue
        bt_row = bt_row.iloc[0]
        for col in gcols:
            lbl = f"{col}_本周"
            if lbl in bt_row.index:
                ok = _num_close(bt_row[lbl], row[col])
                results.append((f"按类型-{op}-{col}", ok, f"展示={bt_row[lbl]}, 重算={row[col]}"))
        share_recalc = (row["购买总金额"] / total_amt_cur) if total_amt_cur else np.nan
        ok_share = _num_close(bt_row.get("金额占比_本周"), share_recalc)
        results.append((f"按类型-{op}-金额占比", ok_share, ""))

    return results


def run_formula_checks(wk, contrib, by_type, zara_daily_cur, zara_by_type_cur):
    """校验 CTR/ATC/CVR/UV_VALUE 及环比公式。返回 [(check_name, passed, detail_str), ...]"""
    results = []

    # 日度：每行 CTR = 点击UV/搜索UV 等
    for rate_name, num_col, den_col in [
        ("CTR", "点击UV", "搜索UV"),
        ("ATC", "加购UV", "搜索UV"),
        ("CVR", "购买人数", "搜索UV"),
        ("UV_VALUE", "购买总金额", "搜索UV"),
    ]:
        if den_col not in zara_daily_cur.columns or num_col not in zara_daily_cur.columns or rate_name not in zara_daily_cur.columns:
            continue
        den = zara_daily_cur[den_col]
        num = zara_daily_cur[num_col]
        expected = num / den.replace(0, np.nan)
        actual = zara_daily_cur[rate_name]
        mask = den.notna() & (den != 0)
        if mask.any():
            ok = np.isclose(actual[mask].astype(float), expected[mask].astype(float), rtol=REL_TOL, atol=ABS_TOL, equal_nan=True).all()
            bad_idx = np.where(~np.isclose(actual[mask].astype(float), expected[mask].astype(float), rtol=REL_TOL, atol=ABS_TOL, equal_nan=True))[0]
            detail = f"首行不匹配: idx={bad_idx[0]}" if not ok and len(bad_idx) else ""
            results.append((f"日度-{rate_name}", ok, detail))

    # 周汇总 metrics 与 cur_total 重算一致（已在 display 中覆盖，这里只做公式层面）
    cur_total = wk["cur_total"]
    def ratio(s, n, d):
        den = s.get(d, np.nan)
        if den == 0 or pd.isna(den):
            return np.nan
        return s.get(n, np.nan) / den
    for name, num, den in [("CTR", "点击UV", "搜索UV"), ("ATC", "加购UV", "搜索UV"), ("CVR", "购买人数", "搜索UV"), ("UV_VALUE", "购买总金额", "搜索UV")]:
        r = ratio(cur_total, num, den)
        results.append((f"周汇总-{name}", _num_close(wk["metrics"][name][0], r), f"cur_total 重算={r}"))

    # 按类型：CTR_本周 = 点击UV_本周/搜索UV_本周
    for _, row in by_type.iterrows():
        op = row["操作类型"]
        su = row.get("搜索UV_本周", 0) or 0
        if su == 0:
            continue
        ctr_exp = row["点击UV_本周"] / su
        atc_exp = row["加购UV_本周"] / su
        cvr_exp = row["购买人数_本周"] / su
        results.append((f"按类型-{op}-CTR公式", _num_close(row["CTR_本周"], ctr_exp), ""))
        results.append((f"按类型-{op}-ATC公式", _num_close(row["ATC_本周"], atc_exp), ""))
        results.append((f"按类型-{op}-CVR公式", _num_close(row["CVR_本周"], cvr_exp), ""))

    return results


def run_data_quality_checks(mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words, home_config_words):
    """数据质量：缺失、异常值、日期等。返回 [(check_name, passed, message), ...]"""
    results = []

    # mini
    if mini is not None and not mini.empty:
        has_date = "date" in mini.columns or "日期" in mini.columns
        results.append(("mini-日期列存在", has_date, "缺少日期列" if not has_date else ""))
        if "成交金额" in mini.columns:
            neg = (mini["成交金额"] < 0).any()
            results.append(("mini-成交金额无负值", not neg, "存在负值" if neg else ""))
    else:
        results.append(("mini-有数据", False, "mini 为空"))

    # zara_daily_cur
    required_daily = ["Date", "搜索PV", "搜索UV", "点击UV", "加购UV", "购买人数", "购买总金额"]
    for df, label in [(zara_daily_cur, "日度-本周"), (zara_daily_pre, "日度-上周")]:
        if df is None or df.empty:
            results.append((f"{label}-有数据", True, "无数据跳过"))
            continue
        for c in required_daily:
            if c == "Date" and "date" in df.columns:
                c = "date"
            if c not in df.columns:
                results.append((f"{label}-列{c}", False, f"缺少列 {c}"))
        if "date" in df.columns:
            n_days = df["date"].nunique()
            results.append((f"{label}-天数", n_days == 7 or label == "日度-上周", f"天数={n_days}"))

    # zara_by_type
    for df, label in [(zara_by_type_cur, "按类型-本周"), (zara_by_type_pre, "按类型-上周")]:
        if df is None or df.empty:
            results.append((f"{label}-有数据", True, "无数据跳过"))
            continue
        results.append((f"{label}-操作类型列", "操作类型" in df.columns, ""))
        results.append((f"{label}-日期列", "date" in df.columns, ""))

    # hotwords
    if hotwords is not None and not hotwords.empty:
        has_kw = "关键词" in hotwords.columns
        results.append(("热词-关键词列", has_kw, ""))
        if "搜索PV" in hotwords.columns and hotwords["搜索PV"].notna().any():
            neg_pv = (hotwords["搜索PV"] < 0).any()
            results.append(("热词-搜索PV无负值", not neg_pv, "存在负值" if neg_pv else ""))
    else:
        results.append(("热词-有数据", True, "无热词数据"))

    # home_config_words
    if home_config_words is not None and not home_config_words.empty:
        has_kw_h = "关键词" in home_config_words.columns
        results.append(("首页配置词-关键词列", has_kw_h, ""))
        pv_h = "搜索PV_cur" if "搜索PV_cur" in home_config_words.columns else "搜索PV"
        if pv_h in home_config_words.columns and home_config_words[pv_h].notna().any():
            neg_pv_h = (home_config_words[pv_h] < 0).any()
            results.append(("首页配置词-搜索PV无负值", not neg_pv_h, "存在负值" if neg_pv_h else ""))
    else:
        results.append(("首页配置词-有数据", True, "无首页配置词数据"))

    # natural_words
    if natural_words is not None and not natural_words.empty:
        results.append(("自然词-关键词列", "关键词" in natural_words.columns, ""))
        if "品类" in natural_words.columns:
            bad_cat = ~natural_words["品类"].isin(CATEGORIES)
            results.append(("自然词-品类在CATEGORIES内", not bad_cat.any(), f"异常品类: {natural_words.loc[bad_cat, '品类'].unique().tolist()}" if bad_cat.any() else ""))
    else:
        results.append(("自然词-有数据", True, "无自然词数据"))

    return results


def _json_safe(val):
    if pd.isna(val) or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (np.integer, np.int64)):
        return int(val)
    if isinstance(val, (np.floating, np.float64)):
        return float(val)
    if isinstance(val, (np.bool_, bool)):
        return bool(val)
    return val


def build_baseline_from_run(wk, contrib, by_type):
    """从当前运行结果生成基准 dict，可序列化为 JSON。"""
    def ser_df(df):
        if df is None or df.empty:
            return []
        cols = [c for c in ["操作类型", "搜索UV_本周", "点击UV_本周", "加购UV_本周", "购买人数_本周", "购买总金额_本周", "CTR_本周", "ATC_本周", "CVR_本周", "金额占比_本周"] if c in df.columns]
        rows = []
        for _, r in df[cols].iterrows():
            rows.append({k: _json_safe(r[k]) for k in cols})
        return rows
    return {
        "generated_at": datetime.now().isoformat(),
        "period": dict(wk["period"]),
        "cur_total": {k: (float(v) if pd.notna(v) else None) for k, v in wk["cur_total"].items()},
        "pre_total": {k: (float(v) if pd.notna(v) else None) for k, v in wk["pre_total"].items()},
        "metrics": {k: [float(x) if pd.notna(x) else None for x in v] for k, v in wk["metrics"].items()},
        "shares": {k: [float(x) if pd.notna(x) else None for x in v] for k, v in contrib["shares"].items()},
        "by_type": ser_df(by_type),
    }


def compare_against_baseline(wk, contrib, by_type, baseline_path):
    """与基准 JSON 对比。返回 (all_passed, list_of_diffs)。"""
    if not baseline_path or not Path(baseline_path).exists():
        return None, []  # 无基准
    try:
        with open(baseline_path, "r", encoding="utf-8") as f:
            base = json.load(f)
    except Exception as e:
        return False, [f"读取基准失败: {e}"]
    diffs = []

    def cmp_val(name, cur, ref, rel_tol=BASELINE_REL_TOL):
        if ref is None and (cur is None or (isinstance(cur, float) and np.isnan(cur))):
            return True
        if ref is None or cur is None:
            if ref != cur:
                diffs.append(f"{name}: 基准={ref}, 当前={cur}")
            return ref == cur
        try:
            a, b = float(cur), float(ref)
            if np.isnan(a) and np.isnan(b):
                return True
            if np.isnan(a) or np.isnan(b):
                diffs.append(f"{name}: 基准={ref}, 当前={cur}")
                return False
            den = max(abs(b), 1e-9)
            if abs(a - b) / den > rel_tol:
                diffs.append(f"{name}: 基准={b}, 当前={a}, 差异={(a-b)/den:.2%}")
                return False
        except (TypeError, ValueError):
            if cur != ref:
                diffs.append(f"{name}: 基准={ref}, 当前={cur}")
            return cur == ref
        return True

    # cur_total
    for k, v in base.get("cur_total", {}).items():
        cur_v = wk["cur_total"].get(k)
        cmp_val(f"cur_total.{k}", cur_v, v)
    # pre_total
    for k, v in base.get("pre_total", {}).items():
        cur_v = wk["pre_total"].get(k)
        cmp_val(f"pre_total.{k}", cur_v, v)
    # metrics
    for k, pair in base.get("metrics", {}).items():
        cur_pair = wk["metrics"].get(k, (None, None))
        cmp_val(f"metrics.{k}[0]", cur_pair[0], pair[0] if pair else None)
        cmp_val(f"metrics.{k}[1]", cur_pair[1], pair[1] if pair and len(pair) > 1 else None)
    # shares
    for k, pair in base.get("shares", {}).items():
        cur_pair = contrib["shares"].get(k, (None, None))
        cmp_val(f"shares.{k}[0]", cur_pair[0], pair[0] if pair else None)
        cmp_val(f"shares.{k}[1]", cur_pair[1], pair[1] if pair and len(pair) > 1 else None)
    # by_type: 按操作类型对齐比较
    base_bt = {r["操作类型"]: r for r in base.get("by_type", []) if isinstance(r, dict) and "操作类型" in r}
    for _, row in by_type.iterrows():
        op = row["操作类型"]
        if op not in base_bt:
            diffs.append(f"by_type.{op}: 当前存在，基准中无")
            continue
        ref_row = base_bt[op]
        for col in ["搜索UV_本周", "购买总金额_本周", "CTR_本周", "CVR_本周", "金额占比_本周"]:
            if col in row.index and col in ref_row:
                cmp_val(f"by_type.{op}.{col}", row[col], ref_row.get(col))
    for op in base_bt:
        if op not in by_type["操作类型"].values:
            diffs.append(f"by_type.{op}: 基准中存在，当前无")

    all_passed = len(diffs) == 0
    return all_passed, diffs


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

    mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words, home_config_words = load_data(paths)
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
    # 图3-0: 搜索量占比柱状图（柱高为搜索UV占比）
    fig_t0 = go.Figure()
    total_uv_cur = by_type["搜索UV_本周"].sum()
    total_uv_pre = by_type["搜索UV_上周"].sum()
    uv_share_cur = by_type["搜索UV_本周"] / total_uv_cur if total_uv_cur else np.nan
    uv_share_pre = by_type["搜索UV_上周"] / total_uv_pre if total_uv_pre else np.nan
    text_uv_cur = [f"{v:,.0f}\n({s:.1%})" if not pd.isna(s) else f"{v:,.0f}" for v, s in zip(by_type["搜索UV_本周"], uv_share_cur)]
    text_uv_pre = [f"{v:,.0f}\n({s:.1%})" if not pd.isna(s) else f"{v:,.0f}" for v, s in zip(by_type["搜索UV_上周"], uv_share_pre)]
    fig_t0.add_trace(
        go.Bar(
            x=by_type["操作类型"],
            y=uv_share_cur,
            name="本周",
            marker_color="#4c78a8",
            text=text_uv_cur,
            textposition="outside",
        )
    )
    fig_t0.add_trace(
        go.Bar(
            x=by_type["操作类型"],
            y=uv_share_pre,
            name="上周",
            marker_color="#9ecae9",
            text=text_uv_pre,
            textposition="outside",
        )
    )
    fig_t0.update_layout(
        height=360,
        barmode="group",
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis=dict(title="搜索UV占比", tickformat=".0%"),
        title="按搜索类型的搜索量占比",
    )
    st.plotly_chart(fig_t0, use_container_width=True)

    # 图3-1: 金额占比柱状图（柱高为占比，标签显示金额+占比）
    fig_t1 = go.Figure()
    total_amt_pre = by_type["购买总金额_上周"].sum()
    if total_amt_pre and not pd.isna(total_amt_pre) and total_amt_pre != 0:
        pre_share = by_type["购买总金额_上周"] / total_amt_pre
    else:
        pre_share = pd.Series([np.nan] * len(by_type), index=by_type.index)
    text_cur = [f"{v:,.0f}\n({s:.1%})" for v, s in zip(by_type["购买总金额_本周"], by_type["金额占比_本周"])]
    text_pre = [f"{v:,.0f}\n({s:.1%})" if pd.notna(s) else f"{v:,.0f}" for v, s in zip(by_type["购买总金额_上周"], pre_share)]
    # y 轴使用金额占比，而不是金额本身
    fig_t1.add_trace(go.Bar(x=by_type["操作类型"], y=by_type["金额占比_本周"], name="本周", marker_color="#4c78a8", text=text_cur, textposition="outside"))
    fig_t1.add_trace(go.Bar(x=by_type["操作类型"], y=pre_share, name="上周", marker_color="#9ecae9", text=text_pre, textposition="outside"))
    fig_t1.update_layout(height=380, barmode="group", margin=dict(l=20, r=20, t=20, b=20), yaxis=dict(title="金额占比", tickformat=".0%"))
    st.plotly_chart(fig_t1, use_container_width=True)

    # 图3-2: 同指标一起（深色=本周，浅色=上周）
    metrics_cfg = [
        ("CTR", "#1f77b4", "#9ecae9"),   # 本周深蓝，上周浅蓝
        ("ATC", "#2ca02c", "#a1d99b"),   # 本周深绿，上周浅绿
        ("CVR", "#d62728", "#fdae6b"),   # 本周深橙/红，上周浅橙
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

    st.subheader("4) 搜索词分析")
    
    for cate in CATEGORIES:
        st.markdown(f"### {cate}")
        st.markdown(f"**{cate} - 热词分析**")
        fig_hot, data_hot = category_scatter(hotwords, cate, f"{cate} 热词：CTR vs CVR（气泡=搜索PV）")
        if fig_hot is None:
            st.info("该品类暂无热词数据。")
        else:
            st.plotly_chart(fig_hot, use_container_width=True)
            # 显示热词数据表格：按得分环比（或 CTR 环比）区分表现好/差
            good_df, bad_df, perf_basis = split_keyword_perf_groups(data_hot)
            if perf_basis != "无环比":
                col_good, col_bad = st.columns(2)
                with col_good:
                    st.markdown(f"表现上升（{perf_basis} ≥ 0）")
                    if good_df.empty:
                        st.caption("暂无符合条件的词。")
                    else:
                        table_data_g, styled_cols_g = display_keyword_table(good_df)
                        styled_df_g = apply_styler_change(table_data_g.style, styled_cols_g)
                        st.dataframe(
                            styled_df_g,
                            use_container_width=True,
                            height=380,
                            hide_index=True,
                        )

                with col_bad:
                    st.markdown(f"表现下降（{perf_basis} < 0）")
                    if bad_df.empty:
                        st.caption("暂无符合条件的词。")
                    else:
                        table_data_b, styled_cols_b = display_keyword_table(bad_df)
                        styled_df_b = apply_styler_change(table_data_b.style, styled_cols_b)
                        st.dataframe(
                            styled_df_b,
                            use_container_width=True,
                            height=380,
                            hide_index=True,
                        )
            else:
                # 无环比数据时，保持原有单表展示
                table_data, styled_cols = display_keyword_table(data_hot)
                styled_df = apply_styler_change(table_data.style, styled_cols)
                st.dataframe(
                    styled_df,
                    use_container_width=True,
                    height=400,
                    hide_index=True,
                )

        st.markdown(f"**{cate} - 首页配置词分析**")
        fig_home, data_home = category_scatter(home_config_words, cate, f"{cate} 首页配置词：CTR vs CVR（气泡=搜索PV）")
        if fig_home is None:
            st.info("该品类暂无首页配置词数据。")
        else:
            st.plotly_chart(fig_home, use_container_width=True)
            good_h, bad_h, perf_basis_h = split_keyword_perf_groups(data_home)
            if perf_basis_h != "无环比":
                col_gh, col_bh = st.columns(2)
                with col_gh:
                    st.markdown(f"表现上升（{perf_basis_h} ≥ 0）")
                    if good_h.empty:
                        st.caption("暂无符合条件的词。")
                    else:
                        table_g_h, styled_cols_g_h = display_keyword_table(good_h)
                        styled_df_g_h = apply_styler_change(table_g_h.style, styled_cols_g_h)
                        st.dataframe(
                            styled_df_g_h,
                            use_container_width=True,
                            height=380,
                            hide_index=True,
                        )
                with col_bh:
                    st.markdown(f"表现下降（{perf_basis_h} < 0）")
                    if bad_h.empty:
                        st.caption("暂无符合条件的词。")
                    else:
                        table_b_h, styled_cols_b_h = display_keyword_table(bad_h)
                        styled_df_b_h = apply_styler_change(table_b_h.style, styled_cols_b_h)
                        st.dataframe(
                            styled_df_b_h,
                            use_container_width=True,
                            height=380,
                            hide_index=True,
                        )
            else:
                table_h, styled_cols_h = display_keyword_table(data_home)
                styled_df_h = apply_styler_change(table_h.style, styled_cols_h)
                st.dataframe(
                    styled_df_h,
                    use_container_width=True,
                    height=400,
                    hide_index=True,
                )

        st.markdown(f"**{cate} - 自然词分析**")
        fig_nat, data_nat = None, None
        women_natural_skip_tables = False
        if cate == "女士":
            pool_w = women_natural_top_pool(natural_words)
            if pool_w.empty:
                st.info("该品类暂无自然词数据（按末尾品类提取后为空）。")
                women_natural_skip_tables = True
            else:
                n_pool = len(pool_w)
                page_labels = [
                    f"{i + 1}-{min(i + NATURAL_WOMEN_PAGE_SIZE, n_pool)}"
                    for i in range(0, n_pool, NATURAL_WOMEN_PAGE_SIZE)
                ]
                seg = st.radio(
                    "女士自然词分段（按搜索PV Top100，每页 20 词；均值线与气泡大小均相对全部 Top 词）",
                    page_labels,
                    horizontal=True,
                    key="women_natural_segment",
                )
                lo = page_labels.index(seg) * NATURAL_WOMEN_PAGE_SIZE
                hi = min(lo + NATURAL_WOMEN_PAGE_SIZE, n_pool)
                page_df = pool_w.iloc[lo:hi].copy()
                title_nat = f"{cate} 自然词：CTR vs CVR（气泡=搜索PV）｜{seg} / Top{n_pool}"
                fig_nat, data_nat = category_scatter(
                    natural_words,
                    cate,
                    title_nat,
                    plot_df=page_df,
                    ref_pool_for_avg=pool_w,
                    bubble_max_df=pool_w,
                )
        else:
            fig_nat, data_nat = category_scatter(
                natural_words, cate, f"{cate} 自然词：CTR vs CVR（气泡=搜索PV）"
            )

        if not women_natural_skip_tables:
            if fig_nat is None:
                st.info("该品类暂无自然词数据（按末尾品类提取后为空）。")
            else:
                st.plotly_chart(fig_nat, use_container_width=True)
                # 显示自然词数据表格：按得分环比（或 CTR 环比）区分表现好/差（与当前页/当前图一致）
                good_nat, bad_nat, perf_basis_n = split_keyword_perf_groups(data_nat)
                if perf_basis_n != "无环比":
                    col_good_n, col_bad_n = st.columns(2)
                    with col_good_n:
                        st.markdown(f"表现上升（{perf_basis_n} ≥ 0）")
                        if good_nat.empty:
                            st.caption("暂无符合条件的词。")
                        else:
                            table_data_gn, styled_cols_gn = display_keyword_table(good_nat)
                            styled_df_gn = apply_styler_change(table_data_gn.style, styled_cols_gn)
                            st.dataframe(
                                styled_df_gn,
                                use_container_width=True,
                                height=380,
                                hide_index=True,
                            )

                    with col_bad_n:
                        st.markdown(f"表现下降（{perf_basis_n} < 0）")
                        if bad_nat.empty:
                            st.caption("暂无符合条件的词。")
                        else:
                            table_data_bn, styled_cols_bn = display_keyword_table(bad_nat)
                            styled_df_bn = apply_styler_change(table_data_bn.style, styled_cols_bn)
                            st.dataframe(
                                styled_df_bn,
                                use_container_width=True,
                                height=380,
                                hide_index=True,
                            )
                else:
                    table_data_nat, styled_cols_nat = display_keyword_table(data_nat)
                    styled_df_nat = apply_styler_change(table_data_nat.style, styled_cols_nat)
                    st.dataframe(
                        styled_df_nat,
                        use_container_width=True,
                        height=400,
                        hide_index=True,
                    )

    st.subheader("5) 数据校验")
    disp_checks = run_display_consistency_checks(wk, contrib, by_type, zara_daily_cur, zara_daily_pre, mini, zara_by_type_cur, zara_by_type_pre)
    formula_checks = run_formula_checks(wk, contrib, by_type, zara_daily_cur, zara_by_type_cur)
    quality_checks = run_data_quality_checks(mini, zara_daily_cur, zara_daily_pre, zara_by_type_cur, zara_by_type_pre, hotwords, natural_words, home_config_words)
    baseline_passed, baseline_diffs = compare_against_baseline(wk, contrib, by_type, str(BASELINE_PATH))

    def render_check_list(items, title):
        passed = sum(1 for _, ok, _ in items if ok)
        total = len(items)
        status = "通过" if passed == total else f"{passed}/{total} 通过"
        with st.expander(f"{title} — {status}", expanded=(passed != total)):
            for name, ok, detail in items:
                icon = "✓" if ok else "✗"
                color = "green" if ok else "red"
                st.markdown(f"- <span style='color:{color}'>**{icon}**</span> {name}" + (f" — {detail}" if detail else ""), unsafe_allow_html=True)

    render_check_list(disp_checks, "源数据与展示一致")
    render_check_list(formula_checks, "公式正确性")
    render_check_list(quality_checks, "数据质量")

    if baseline_passed is None:
        st.info("未设置基准，可先点击「保存当前结果为基准」再对比。")
    else:
        status = "通过" if baseline_passed else "存在差异"
        with st.expander(f"基准对比 — {status}", expanded=not baseline_passed):
            if baseline_passed:
                st.success("当前结果与基准一致。")
            else:
                for d in baseline_diffs:
                    st.markdown(f"- **✗** {d}")

    col_save, col_dl = st.columns(2)
    with col_save:
        if st.button("保存当前结果为基准"):
            bl = build_baseline_from_run(wk, contrib, by_type)
            try:
                BASELINE_PATH.parent.mkdir(parents=True, exist_ok=True)
                with open(BASELINE_PATH, "w", encoding="utf-8") as f:
                    json.dump(bl, f, ensure_ascii=False, indent=2)
                st.success(f"已保存至 {BASELINE_PATH}")
            except Exception as e:
                st.error(f"保存失败: {e}")
    with col_dl:
        report_lines = [
            "# 数据校验报告",
            f"生成时间: {datetime.now().isoformat()}",
            f"周期: {wk['period']['cur']} | {wk['period']['pre']}",
            "",
            "## 1. 源数据与展示一致",
        ]
        for name, ok, detail in disp_checks:
            report_lines.append(f"- {'✓' if ok else '✗'} {name}" + (f" — {detail}" if detail else ""))
        report_lines.extend(["", "## 2. 公式正确性"])
        for name, ok, detail in formula_checks:
            report_lines.append(f"- {'✓' if ok else '✗'} {name}" + (f" — {detail}" if detail else ""))
        report_lines.extend(["", "## 3. 数据质量"])
        for name, ok, detail in quality_checks:
            report_lines.append(f"- {'✓' if ok else '✗'} {name}" + (f" — {detail}" if detail else ""))
        report_lines.extend(["", "## 4. 基准对比"])
        if baseline_passed is None:
            report_lines.append("- 未设置基准")
        elif baseline_passed:
            report_lines.append("- ✓ 与基准一致")
        else:
            for d in baseline_diffs:
                report_lines.append(f"- ✗ {d}")
        report_txt = "\n".join(report_lines)
        st.download_button(
            label="下载校验报告",
            data=report_txt,
            file_name=f"validation_report_{wk['cur']['date'].max().date()}.md",
            mime="text/markdown",
        )

    st.subheader("6) 导出 PDF")
    pdf_bytes = build_pdf_bytes(wk, contrib, by_type)
    st.download_button(
        label="生成并下载PDF文件",
        data=pdf_bytes,
        file_name=f"search_quality_report_{wk['cur']['date'].max().date()}.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    render()
