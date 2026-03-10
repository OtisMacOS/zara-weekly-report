# 搜索引擎质量报告

## 1) 交互看板（Streamlit）
在项目根目录运行：

```bash
streamlit run report/search_quality_report.py
```

说明：
- 默认读取以下文件：
  - `/Users/otis/Downloads/zara周报数据源/小程序大盘数据-近30天.xlsx`
  - `/Users/otis/Downloads/zara周报数据源/zara日度数据.xlsx`
  - `/Users/otis/Downloads/zara周报数据源/zara日度数据-by搜索类型.xlsx`
  - `/Users/otis/Downloads/zara周报数据源/本周女士品类热词.xlsx`
  - `/Users/otis/Downloads/zara周报数据源/本周男士品类热词.xlsx`
  - `/Users/otis/Downloads/zara周报数据源/本周儿童品类热词.xlsx`
  - `/Users/otis/Downloads/zara周报数据源/本周家居品类热词.xlsx`
  - 可在左侧输入框改路径。
- 周维度窗口自动按 `zara日度数据-by搜索类型.xlsx` 的最大日期回溯 7 天生成本周，并向前 7 天作为上周。

## 2) 静态周报（Markdown）
已生成文件：

- `report/weekly_report_2026-03-02.md`

后续可按同样逻辑定时生成每周版本。
