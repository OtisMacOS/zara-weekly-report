"""Read-only Zara Spark data access for the weekly dashboard."""

from __future__ import annotations

import os
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pymysql


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_ENV_FILES = (
    PROJECT_ROOT / ".env",
    PROJECT_ROOT.parent / "zara 热词追踪" / ".env",
)
KNOWN_SECTIONS = ("女士", "男士", "儿童", "家居")


def _strip_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def _load_env_files() -> None:
    override = os.environ.get("ZARA_SPARK_ENV_FILE", "").strip()
    candidates = (Path(override).expanduser(),) if override else DEFAULT_ENV_FILES
    for path in candidates:
        if not path.is_file():
            continue
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), _strip_quotes(value))


def _connection():
    _load_env_files()
    required = {
        "host": os.environ.get("SPARK_MYSQL_HOST", "").strip(),
        "user": os.environ.get("SPARK_MYSQL_USER", "").strip(),
        "password": (
            os.environ.get("SPARK_MYSQL_PASSWORD")
            or os.environ.get("SPARK_MYSQL_PASS")
            or ""
        ),
        "database": os.environ.get("SPARK_MYSQL_DATABASE", "").strip(),
    }
    missing = [key for key, value in required.items() if not value]
    if missing:
        raise RuntimeError(
            "缺少 Spark 配置："
            + ", ".join(f"SPARK_MYSQL_{key.upper()}" for key in missing)
            + "。请在项目 .env 中配置，或通过 ZARA_SPARK_ENV_FILE 指定配置文件。"
        )
    return pymysql.connect(
        host=required["host"],
        port=int(os.environ.get("SPARK_MYSQL_PORT", "3306")),
        user=required["user"],
        password=required["password"],
        database=required["database"],
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=10,
        read_timeout=120,
    )


def _fetch_frame(sql: str, params: tuple) -> pd.DataFrame:
    conn = _connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            rows = cursor.fetchall()
        return pd.DataFrame(rows)
    finally:
        conn.close()


def latest_complete_day(today: date | None = None) -> date:
    """Return the newest available natural day, excluding today's partial data."""
    today = today or date.today()
    df = _fetch_frame(
        """
        SELECT MAX(user_date) AS max_date
        FROM dwd_user_behavior_daily
        WHERE user_date < %s
        """,
        (today,),
    )
    if df.empty or pd.isna(df.iloc[0]["max_date"]):
        raise RuntimeError("Spark 中没有昨天及以前的完整日用户数据。")
    return pd.Timestamp(df.iloc[0]["max_date"]).date()


def complete_week_windows(today: date | None = None) -> tuple[date, date, date, date]:
    cur_end = latest_complete_day(today)
    cur_start = cur_end - timedelta(days=6)
    pre_end = cur_start - timedelta(days=1)
    pre_start = pre_end - timedelta(days=6)
    return cur_start, cur_end, pre_start, pre_end


def fetch_daily_funnel(start_date: date, end_date: date) -> pd.DataFrame:
    """Fetch daily search UV funnel using the user-level Spark table."""
    sql = """
        WITH SearchMetrics AS (
            SELECT
                user_date,
                SUM(user_search_pv) AS search_pv,
                COUNT(DISTINCT user_uid) AS search_uv
            FROM dwd_user_behavior_daily
            WHERE user_search_pv > 0
              AND user_date BETWEEN %s AND %s
            GROUP BY user_date
        ),
        ClickMetrics AS (
            SELECT
                user_date,
                COUNT(DISTINCT CASE
                    WHEN user_search_pv_pdp_view > 0 OR user_search_pv_click > 0
                    THEN user_uid
                END) AS click_uv
            FROM dwd_user_behavior_daily
            WHERE (user_search_pv_pdp_view > 0 OR user_search_pv_click > 0)
              AND user_date BETWEEN %s AND %s
            GROUP BY user_date
        ),
        ATCMetrics AS (
            SELECT
                user_date,
                COUNT(DISTINCT CASE WHEN user_search_pv_atc > 0 THEN user_uid END) AS atc_uv
            FROM dwd_user_behavior_daily
            WHERE user_search_pv_atc > 0
              AND user_date BETWEEN %s AND %s
            GROUP BY user_date
        ),
        PayMetrics AS (
            SELECT
                user_date,
                COUNT(DISTINCT CASE WHEN user_search_pv_pay > 0 THEN user_uid END) AS purchase_uv,
                SUM(user_pay_amount) AS purchase_amount
            FROM dwd_user_behavior_daily
            WHERE user_search_pv_pay > 0
              AND user_date BETWEEN %s AND %s
            GROUP BY user_date
        )
        SELECT
            s.user_date AS Date,
            COALESCE(s.search_pv, 0) AS `搜索PV`,
            COALESCE(s.search_uv, 0) AS `搜索UV`,
            COALESCE(c.click_uv, 0) AS `点击UV`,
            COALESCE(a.atc_uv, 0) AS `加购UV`,
            COALESCE(p.purchase_uv, 0) AS `购买人数`,
            COALESCE(p.purchase_amount, 0) AS `购买总金额`
        FROM SearchMetrics s
        LEFT JOIN ClickMetrics c ON s.user_date = c.user_date
        LEFT JOIN ATCMetrics a ON s.user_date = a.user_date
        LEFT JOIN PayMetrics p ON s.user_date = p.user_date
        ORDER BY s.user_date
    """
    params = (
        start_date,
        end_date,
        start_date,
        end_date,
        start_date,
        end_date,
        start_date,
        end_date,
    )
    return _fetch_frame(sql, params)


def fetch_hotwords(start_date: date, end_date: date) -> pd.DataFrame:
    """Fetch the complete hot-query pool and aggregate metrics for one period."""
    section_values = ", ".join(f"'{section}'" for section in KNOWN_SECTIONS)
    tail = "TRIM(SUBSTRING_INDEX(query, ' ', -1))"
    sql = f"""
        WITH parsed AS (
            SELECT
                CASE
                    WHEN {tail} IN ({section_values})
                    THEN TRIM(LEFT(query, CHAR_LENGTH(query) - CHAR_LENGTH({tail}) - 1))
                    ELSE query
                END AS keyword,
                CASE
                    WHEN {tail} IN ({section_values}) THEN {tail}
                    ELSE '未知'
                END AS section,
                date,
                query_search_pv,
                query_search_uv,
                query_search_click_uv,
                query_search_atc_uv,
                query_search_pay_uv,
                query_search_pay_amount
            FROM dwd_query_behavior_daily
            WHERE query_handle_type = '203'
              AND date BETWEEN %s AND %s
        )
        SELECT
            keyword AS `关键词`,
            section AS `品类`,
            COUNT(DISTINCT date) AS `上架天数`,
            SUM(query_search_pv) AS `搜索PV`,
            SUM(query_search_uv) AS `搜索UV`,
            SUM(query_search_click_uv) AS `点击UV`,
            SUM(query_search_atc_uv) AS `加购UV`,
            SUM(query_search_pay_uv) AS `购买人数`,
            SUM(query_search_pay_amount) AS `购买总金额`
        FROM parsed
        WHERE keyword <> ''
          AND section IN ({section_values})
        GROUP BY keyword, section
        ORDER BY section, `搜索PV` DESC, keyword
    """
    return _fetch_frame(sql, (start_date, end_date))


def fetch_by_type_funnel(start_date: date, end_date: date) -> pd.DataFrame:
    """按 ``query_handle_type`` 汇总搜索类型周报所需的漏斗指标。

    Spark 没有主库 Excel 中的 ``操作类型`` 字段，因此这里保留底表编码口径：
    200/400 合并为自然搜索，203 为热词搜索，其余类型以编码透明展示。
    """
    sql = """
        SELECT
            CASE
                WHEN query_handle_type IN ('200', '400') THEN '自然搜索'
                WHEN query_handle_type = '203' THEN '热词搜索'
                ELSE CONCAT('其他搜索（', query_handle_type, '）')
            END AS `操作类型`,
            date AS Date,
            SUM(query_search_pv) AS `搜索PV`,
            SUM(query_search_uv) AS `搜索UV`,
            SUM(query_search_click_uv) AS `点击UV`,
            SUM(query_search_atc_uv) AS `加购UV`,
            SUM(query_search_pay_uv) AS `购买人数`,
            SUM(query_search_pay_amount) AS `购买总金额`
        FROM dwd_query_behavior_daily
        WHERE date BETWEEN %s AND %s
        GROUP BY query_handle_type, date
        ORDER BY date, `操作类型`
    """
    return _fetch_frame(sql, (start_date, end_date))


def fetch_natural_words(start_date: date, end_date: date) -> pd.DataFrame:
    """汇总 Spark 自然词（``query_handle_type`` 200/400）词级漏斗。

    品类沿用热词/自然词项目的约定，从 query 末尾的女士、男士、儿童、家居解析。
    未带品类后缀的 query 不纳入周报品类词表，避免把跨品类词错误归类。
    """
    section_values = ", ".join(f"'{section}'" for section in KNOWN_SECTIONS)
    tail = "TRIM(SUBSTRING_INDEX(query, ' ', -1))"
    sql = f"""
        WITH parsed AS (
            SELECT
                CASE
                    WHEN {tail} IN ({section_values})
                    THEN TRIM(LEFT(query, CHAR_LENGTH(query) - CHAR_LENGTH({tail}) - 1))
                    ELSE query
                END AS keyword,
                CASE
                    WHEN {tail} IN ({section_values}) THEN {tail}
                    ELSE '未知'
                END AS section,
                date,
                query_search_pv,
                query_search_uv,
                query_search_click_uv,
                query_search_atc_uv,
                query_search_pay_uv,
                query_search_pay_amount
            FROM dwd_query_behavior_daily
            WHERE query_handle_type IN ('200', '400')
              AND date BETWEEN %s AND %s
        )
        SELECT
            keyword AS `关键词`,
            section AS `品类`,
            COUNT(DISTINCT date) AS `上架天数`,
            SUM(query_search_pv) AS `搜索PV`,
            SUM(query_search_uv) AS `搜索UV`,
            SUM(query_search_click_uv) AS `点击UV`,
            SUM(query_search_atc_uv) AS `加购UV`,
            SUM(query_search_pay_uv) AS `购买人数`,
            SUM(query_search_pay_amount) AS `购买总金额`
        FROM parsed
        WHERE keyword <> ''
          AND section IN ({section_values})
        GROUP BY keyword, section
        ORDER BY section, `搜索PV` DESC, keyword
    """
    return _fetch_frame(sql, (start_date, end_date))
