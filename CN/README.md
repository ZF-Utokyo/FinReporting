# CN Market Pipeline (A-Share)

## 1. 数据源设计（推荐）
- Primary: `cninfo` 年报披露接口 + 年报 PDF 下载。
- Exchange-aware: 按股票自动切换交易所参数。
  - 上交所：`plate=sh`, `column=sse`
  - 深交所：`plate=sz`, `column=szse`
- Fallback: 若 cninfo 接口不可用，再走上交所/深交所公告页直链。

说明：`cninfo` 已聚合沪深上市公司公告，作为统一入口最稳妥；交易所直链作为兜底。

## 2. 端到端流程
1. 输入 `symbol`（如 `300750`）。
2. 通过 cninfo 股票列表拿 `orgId`。
3. 查询近 N 年“年度报告”公告，选择最新完整版 PDF（排除摘要/英文版）。
4. 下载 PDF 到 `CN/raw_pdfs/`。
5. 调用 `extract_consolidated_statements.py` 抽取合并报表（IS/BS/CF）。
6. 映射到 canonical 三表宽表并导出 Excel：
   - `CN_FIN_IS`
   - `CN_FIN_BS`
   - `CN_FIN_CF`
7. 同时保留原始映射结果：
   - `RAW_CN_FIN_IS_GEN`
   - `RAW_CN_FIN_BS_GEN`
   - `RAW_CN_FIN_CF_GEN`

## 3. 运行命令

```bash
./venv/bin/python CN/export_three_statements_excel_cn.py \
  --symbol 300750 \
  --out CN/300750_3statements_from_web.xlsx
```

可选：本地 PDF 模式（不走 cninfo 下载）

```bash
python CN/export_three_statements_excel_cn.py \
  --symbol 300750 \
  --pdf "/path/to/annual_report.pdf" \
  --schema-file schemas/CN_Schemas.xlsx \
  --out CN/300750_3statements_from_local_pdf.xlsx
```

## 4. 合作者可读版（external）
```bash
./venv/bin/python CN/convert_three_statements_readable.py \
  --in CN/300750_3statements_from_web.xlsx \
  --out CN/300750_3statements_for_collab_zh_external.xlsx
```

默认只输出 3 个可读 sheet（利润表/资产负债表/现金流量表）。
如需附带 raw sheet，可加 `--include-raw`。

## 5. Anomaly-aware 建议
- 对每张报表做最小完整性检查（核心字段缺失率）。
- 对数值做合理性检查（如 `total_revenue`、`net_income` 异常小/空）。
- 命中规则时置 `anomaly_flag=1`，并写入 `anomaly_type/anomaly_detail`，用于 demo 中提示“需要人工复核”。
