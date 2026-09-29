# 搜索引擎质量报告

## 1) 交互看板（Streamlit）
在项目根目录运行：

```bash
streamlit run search_quality_report.py
```

说明：

- 侧栏的「非大盘数据源」可在「原始 Excel」和「Zara Spark」之间完整切换，默认使用 Zara Spark。
- 小程序大盘固定读取 `zara周报数据源` 下的 Excel，不受选择器影响。
- Spark 模式下，日度搜索汇总读取 `dwd_user_behavior_daily`；搜索类型、热词和自然词统一读取 `dwd_query_behavior_daily`。
- Spark 热词限定 `query_handle_type = 203`；自然词限定 `query_handle_type IN (200, 400)`；搜索类型保留底表编码口径。
- Spark 没有首页配置词的等价字段，该板块在 Spark 模式明确显示不可用，不回退 Excel。
- 最近完整日向前 7 天为本周，再向前 7 天为上周；当天部分日数据不进入周报。
- Spark 配置从项目根目录 `.env` 读取；也可用 `ZARA_SPARK_ENV_FILE` 指向其他配置文件。

## 2) 静态周报（Markdown）
已生成文件：

- `report/weekly_report_2026-03-02.md`

后续可按同样逻辑定时生成每周版本。
