# 涨价节奏监控系统

内部使用的 Excel 导入、价格阶段检查和提醒系统。

## 功能

- 价格统计表：自动识别 SKU 表头行，支持手动确认平台价格列，清洗 `$260.00`、`260`、`260.00`，将宽表转为平台价格长表。
- 涨价计划表：自动识别 SKU 表头行，按计划年份解析 `6.25-`、`8.06-10.04`、`10.05后价格` 这类阶段列。
- 价格判断：默认 `MINIMUM_PRICE` 模式，平台当前价低于目标价 0.01 以上判为异常。
- 异常类型：支持 `BELOW_BASE_PRICE`、`BELOW_STAGE_TARGET`、`PLATFORM_PRICE_MISSING`、`PLAN_MISSING`、`PRICE_PARSE_ERROR`、`NORMAL`。
- 提醒中心：基于阶段开始日期输出 30 天、15 天、7 天提醒和已开始仍异常的高风险提醒。

## 启动

```powershell
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\uvicorn backend.app.main:app --reload --host 127.0.0.1 --port 8000
```

打开 http://127.0.0.1:8000 访问页面。

## 数据库

SQLite 文件默认生成在：

```text
data/price_monitor.sqlite3
```

核心表：

- `price_snapshots`
- `price_plan_items`
- `price_plan_stages`

## API

- `POST /api/price-statistics/preview`
- `POST /api/price-statistics/import`
- `POST /api/price-plan/preview`
- `POST /api/price-plan/import`
- `GET /api/dashboard`
- `GET /api/checks`
- `GET /api/reminders`
