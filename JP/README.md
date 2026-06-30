# JP EDINET 三表抽取

从 EDINET 抽取日本公司年度三表（IS/BS/CF），导出一个 3-sheet Excel。

## 前置条件

支持两种模式：

- API 模式：EDINET API v2，需要有效 `Subscription-Key`
- 本地模式：直接读取你从 EDINET 网页手动下载的 ZIP（无需 key）

API 模式下需要有效 `Subscription-Key`。

可以通过环境变量设置：

```bash
export EDINET_API_KEY=
```

## 生成 Toyota 示例（7203）

```bash
./venv/bin/python JP/export_three_statements_excel_jp.py \
  --symbol 7203 \
  --company-name "トヨタ" \
  --as-of-date 2026-02-21 \
  --lookback-days 400 \
  --out JP/toyota_3statements.xlsx
```

## 无 key：本地 ZIP 模式

先在 EDINET 网页手动下载 type=1 ZIP，然后执行：

```bash
./venv/bin/python JP/export_three_statements_excel_jp.py \
  --xbrl-zip "JP/toyota_edinet.zip" \
  --symbol 7203 \
  --company-name "トヨタ" \
  --out JP/toyota_3statements.xlsx
```

如自动识别报告日期不准确，可补充：

```bash
--report-date 2025-03-31 --filing-date 2025-06-20
```

## 无 key：自动下载 ZIP（网页模式）

先安装 Playwright（仅首次）：

```bash
./venv/bin/pip install playwright
./venv/bin/playwright install chromium
```

下载（优先选年报 ASR）：

```bash
./venv/bin/python JP/download_edinet_zip_no_key.py \
  --keyword 7203 \
  --prefer-asr \
  --out JP/toyota_asr.zip
```
