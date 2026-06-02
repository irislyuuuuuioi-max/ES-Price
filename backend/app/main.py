from __future__ import annotations

import math
import re
import sqlite3
import tempfile
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles


ROOT = Path(__file__).resolve().parents[2]
DB_PATH = ROOT / "data" / "price_monitor.sqlite3"
FRONTEND_DIR = ROOT / "frontend"

SKU_ALIASES = {"sku", "skuid", "商品编码", "产品编码", "sku编码"}
CATEGORY_1_ALIASES = {"一级类目", "一级分类", "category_1", "category1"}
CATEGORY_2_ALIASES = {"二级类目", "二级分类", "category_2", "category2"}
OLD_PRICE_ALIASES = {"协同价（旧）", "协同价(旧)", "协同价旧", "协同价"}
NEW_PRICE_ALIASES = {"协同价（新）", "协同价(新)", "协同价新", "协同价新参考", "新参考"}
FINAL_EXEC_ALIASES = {"最终执行时间", "执行时间", "effective_date"}
NON_PLATFORM_HEADERS = (
    SKU_ALIASES
    | CATEGORY_1_ALIASES
    | CATEGORY_2_ALIASES
    | OLD_PRICE_ALIASES
    | NEW_PRICE_ALIASES
    | FINAL_EXEC_ALIASES
    | {"新旧差价", "差价", "备注", "图片", "链接", "asin", "spu"}
)
NON_PLATFORM_HEADER_KEYS = {header.lower().replace(" ", "") for header in NON_PLATFORM_HEADERS}


app = FastAPI(title="涨价节奏监控系统")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db() -> None:
    with db() as con:
        con.executescript(
            """
            create table if not exists price_snapshots (
              id integer primary key,
              sku text not null,
              category_1 text,
              category_2 text,
              snapshot_date text not null,
              platform text not null,
              current_price real,
              raw_value text,
              price_status text not null,
              imported_at text not null
            );
            create table if not exists price_plan_items (
              id integer primary key,
              sku text not null unique,
              category_1 text,
              category_2 text,
              base_price real,
              new_reference_price real,
              price_diff real,
              final_effective_date text,
              final_effective_note text,
              plan_year integer not null,
              imported_at text not null
            );
            create table if not exists price_plan_stages (
              id integer primary key,
              sku text not null,
              stage_name text not null,
              start_date text,
              end_date text,
              target_price real,
              raw_value text,
              price_status text not null,
              sort_order integer not null,
              imported_at text not null
            );
            create table if not exists import_batches (
              id integer primary key,
              import_type text not null,
              filename text,
              imported_at text not null,
              row_count integer not null,
              sku_count integer not null,
              detail_count integer not null,
              snapshot_date text,
              plan_year integer,
              selected_columns text,
              note text
            );
            create table if not exists reminder_statuses (
              reminder_key text primary key,
              status text not null,
              updated_at text not null
            );
            create index if not exists idx_snapshots_sku_date on price_snapshots(sku, snapshot_date);
            create index if not exists idx_stages_sku_dates on price_plan_stages(sku, start_date, end_date);
            create index if not exists idx_import_batches_type_time on import_batches(import_type, imported_at);
            """
        )
        ensure_column(con, "price_snapshots", "batch_id", "integer")
        ensure_column(con, "price_plan_items", "batch_id", "integer")
        ensure_column(con, "price_plan_stages", "batch_id", "integer")


def ensure_column(con: sqlite3.Connection, table: str, column: str, column_type: str) -> None:
    columns = {row["name"] for row in con.execute(f"pragma table_info({table})").fetchall()}
    if column not in columns:
        con.execute(f"alter table {table} add column {column} {column_type}")


@app.on_event("startup")
def startup() -> None:
    init_db()


def clean_header(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalized(value: Any) -> str:
    return clean_header(value).lower().replace(" ", "")


def read_sheet(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, header=None, dtype=object, engine="openpyxl")


def save_upload(upload: UploadFile) -> Path:
    suffix = Path(upload.filename or "upload.xlsx").suffix or ".xlsx"
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp.write(upload.file.read())
    temp.close()
    return Path(temp.name)


def find_header_row(raw: pd.DataFrame) -> int:
    best_row = -1
    best_score = 0
    for idx in range(min(len(raw), 50)):
        cells = [normalized(v) for v in raw.iloc[idx].tolist()]
        score = 0
        if any(v in SKU_ALIASES for v in cells):
            score += 5
        if any(v in CATEGORY_1_ALIASES for v in cells):
            score += 2
        if any(v in CATEGORY_2_ALIASES for v in cells):
            score += 2
        if any("协同价" in v for v in cells):
            score += 2
        non_empty = sum(1 for v in cells if v)
        if non_empty >= 5:
            score += 1
        if score > best_score:
            best_score = score
            best_row = idx
    if best_row < 0 or best_score < 5:
        raise HTTPException(400, "未找到包含 SKU 的表头行")
    return best_row


def frame_from_header(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    headers: list[str] = []
    seen: dict[str, int] = {}
    for i, value in enumerate(raw.iloc[header_row].tolist()):
        name = clean_header(value) or f"unnamed_{i + 1}"
        seen[name] = seen.get(name, 0) + 1
        headers.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = headers
    return data.dropna(how="all")


def find_column(columns: list[str], aliases: set[str]) -> str | None:
    alias_norm = {a.lower().replace(" ", "") for a in aliases}
    for col in columns:
        if normalized(col) in alias_norm:
            return col
    return None


def parse_price(value: Any) -> tuple[float | None, str]:
    if value is None:
        return None, "MISSING"
    if isinstance(value, float) and math.isnan(value):
        return None, "MISSING"
    text = str(value).strip()
    if text == "":
        return None, "MISSING"
    if text.upper() in {"#N/A", "N/A", "NA", "NULL"}:
        return None, "MISSING"
    if text in {"/", "-", "--"}:
        return None, "NOT_APPLICABLE"
    cleaned = text.replace("$", "").replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        return None, "PARSE_ERROR"
    try:
        return float(Decimal(match.group(0))), "OK"
    except InvalidOperation:
        return None, "PARSE_ERROR"


def extract_snapshot_date(raw: pd.DataFrame, header_row: int) -> str | None:
    pattern = re.compile(r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})")
    for row in range(header_row):
        for value in raw.iloc[row].dropna().tolist():
            match = pattern.search(str(value))
            if match:
                y, m, d = map(int, match.groups())
                return date(y, m, d).isoformat()
    return None


def detect_price_columns(columns: list[str]) -> list[str]:
    result = []
    for col in columns:
        n = normalized(col)
        if not n or n in NON_PLATFORM_HEADER_KEYS:
            continue
        if "price" in n or "价格" in n or "-" in col or re.search(r"[A-Z]{2,}", col):
            result.append(col)
    return result


def detect_stage_columns(columns: list[str]) -> list[str]:
    result = []
    for col in columns:
        n = normalized(col)
        if not n or n in NON_PLATFORM_HEADER_KEYS:
            continue
        if re.search(r"\d{1,2}[./]\d{1,2}|后价格|phase|red white blue", col, re.I):
            result.append(col)
    return result


def parse_stage_dates(name: str, year: int) -> tuple[str | None, str | None]:
    text = name.strip()
    rng = re.search(r"(\d{1,2})[./](\d{1,2})\s*[-~至]\s*(\d{1,2})[./](\d{1,2})", text)
    if rng:
        sm, sd, em, ed = map(int, rng.groups())
        return date(year, sm, sd).isoformat(), date(year, em, ed).isoformat()
    start_open = re.search(r"(\d{1,2})[./](\d{1,2})\s*[-~至]?\s*(?:后|起|开始)?", text)
    if start_open:
        sm, sd = map(int, start_open.groups())
        return date(year, sm, sd).isoformat(), None
    return None, None


def parse_final_effective(value: Any) -> tuple[str | None, str | None]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None, None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date().isoformat(), None
    if isinstance(value, date):
        return value.isoformat(), None
    text = str(value).strip()
    if not text:
        return None, None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat(), None
        except ValueError:
            pass
    match = re.search(r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})", text)
    if match:
        y, m, d = map(int, match.groups())
        return date(y, m, d).isoformat(), None
    return None, text


def rows_to_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    return df.where(pd.notnull(df), None).head(20).to_dict("records")


def insert_import_batch(
    con: sqlite3.Connection,
    *,
    import_type: str,
    filename: str | None,
    imported_at: str,
    row_count: int,
    sku_count: int,
    detail_count: int,
    snapshot_date: str | None = None,
    plan_year: int | None = None,
    selected_columns: list[str] | None = None,
    note: str | None = None,
) -> int:
    cursor = con.execute(
        """
        insert into import_batches
        (import_type, filename, imported_at, row_count, sku_count, detail_count, snapshot_date, plan_year, selected_columns, note)
        values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            import_type,
            filename,
            imported_at,
            row_count,
            sku_count,
            detail_count,
            snapshot_date,
            plan_year,
            ",".join(selected_columns or []),
            note,
        ),
    )
    return int(cursor.lastrowid)


@app.get("/api/import-history")
def import_history(import_type: str | None = None) -> list[dict[str, Any]]:
    query = "select * from import_batches"
    params: tuple[Any, ...] = ()
    if import_type:
        query += " where import_type = ?"
        params = (import_type,)
    query += " order by imported_at desc, id desc limit 50"
    with db() as con:
        rows = con.execute(query, params).fetchall()
    return [dict(row) for row in rows]


def latest_batch(con: sqlite3.Connection, import_type: str) -> sqlite3.Row | None:
    return con.execute(
        "select * from import_batches where import_type = ? order by imported_at desc, id desc limit 1",
        (import_type,),
    ).fetchone()


@app.post("/api/price-statistics/preview")
def preview_price_statistics(file: UploadFile = File(...), snapshot_date: str | None = Form(None)) -> dict[str, Any]:
    path = save_upload(file)
    raw = read_sheet(path)
    header_row = find_header_row(raw)
    df = frame_from_header(raw, header_row)
    columns = list(df.columns)
    inferred_date = snapshot_date or extract_snapshot_date(raw, header_row)
    return {
        "header_row": header_row + 1,
        "snapshot_date": inferred_date,
        "columns": columns,
        "detected_platform_columns": detect_price_columns(columns),
        "preview_rows": rows_to_records(df),
    }


@app.post("/api/price-statistics/import")
def import_price_statistics(
    file: UploadFile = File(...),
    snapshot_date: str = Form(...),
    platform_columns: str = Form(...),
) -> dict[str, Any]:
    path = save_upload(file)
    raw = read_sheet(path)
    df = frame_from_header(raw, find_header_row(raw))
    columns = list(df.columns)
    sku_col = find_column(columns, SKU_ALIASES)
    if not sku_col:
        raise HTTPException(400, "缺少 SKU 列")
    c1 = find_column(columns, CATEGORY_1_ALIASES)
    c2 = find_column(columns, CATEGORY_2_ALIASES)
    platforms = [
        p.strip()
        for p in platform_columns.split(",")
        if p.strip() in columns and normalized(p.strip()) not in NON_PLATFORM_HEADER_KEYS
    ]
    imported_at = datetime.now().isoformat(timespec="seconds")
    records = []
    skus = set()
    for _, row in df.iterrows():
        sku = clean_header(row.get(sku_col))
        if not sku:
            continue
        skus.add(sku)
        for platform in platforms:
            price, status = parse_price(row.get(platform))
            records.append(
                (
                    sku,
                    clean_header(row.get(c1)) if c1 else None,
                    clean_header(row.get(c2)) if c2 else None,
                    snapshot_date,
                    platform,
                    price,
                    clean_header(row.get(platform)),
                    status,
                    imported_at,
                )
            )
    with db() as con:
        batch_id = insert_import_batch(
            con,
            import_type="price_statistics",
            filename=file.filename,
            imported_at=imported_at,
            row_count=len(df),
            sku_count=len(skus),
            detail_count=len(records),
            snapshot_date=snapshot_date,
            selected_columns=platforms,
        )
        records = [record + (batch_id,) for record in records]
        con.executemany(
            """
            insert into price_snapshots
            (sku, category_1, category_2, snapshot_date, platform, current_price, raw_value, price_status, imported_at, batch_id)
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            records,
        )
    return {"batch_id": batch_id, "imported": len(records)}


@app.post("/api/price-plan/preview")
def preview_price_plan(file: UploadFile = File(...), plan_year: int = Form(...)) -> dict[str, Any]:
    path = save_upload(file)
    raw = read_sheet(path)
    header_row = find_header_row(raw)
    df = frame_from_header(raw, header_row)
    columns = list(df.columns)
    stages = [
        {"column": col, "start_date": parse_stage_dates(col, plan_year)[0], "end_date": parse_stage_dates(col, plan_year)[1]}
        for col in detect_stage_columns(columns)
    ]
    return {"header_row": header_row + 1, "columns": columns, "detected_stage_columns": stages, "preview_rows": rows_to_records(df)}


@app.post("/api/price-plan/import")
def import_price_plan(
    file: UploadFile = File(...),
    plan_year: int = Form(...),
    stage_columns: str = Form(...),
) -> dict[str, Any]:
    path = save_upload(file)
    raw = read_sheet(path)
    df = frame_from_header(raw, find_header_row(raw))
    columns = list(df.columns)
    sku_col = find_column(columns, SKU_ALIASES)
    if not sku_col:
        raise HTTPException(400, "缺少 SKU 列")
    c1 = find_column(columns, CATEGORY_1_ALIASES)
    c2 = find_column(columns, CATEGORY_2_ALIASES)
    base_col = find_column(columns, OLD_PRICE_ALIASES)
    new_col = find_column(columns, NEW_PRICE_ALIASES)
    final_col = find_column(columns, FINAL_EXEC_ALIASES)
    diff_col = next((c for c in columns if normalized(c) in {"新旧差价", "差价"}), None)
    stage_cols = [
        p.strip()
        for p in stage_columns.split(",")
        if p.strip() in columns and normalized(p.strip()) not in NON_PLATFORM_HEADER_KEYS
    ]
    parsed_stages = []
    for idx, col in enumerate(stage_cols):
        start, end = parse_stage_dates(col, plan_year)
        parsed_stages.append({"column": col, "start": start, "end": end, "sort": idx})
    for idx, stage in enumerate(parsed_stages[:-1]):
        if stage["start"] and not stage["end"] and parsed_stages[idx + 1]["start"]:
            stage["end"] = (date.fromisoformat(parsed_stages[idx + 1]["start"]) - timedelta(days=1)).isoformat()
    imported_at = datetime.now().isoformat(timespec="seconds")
    items = []
    stages = []
    skus = set()
    for _, row in df.iterrows():
        sku = clean_header(row.get(sku_col))
        if not sku:
            continue
        skus.add(sku)
        base_price, _ = parse_price(row.get(base_col)) if base_col else (None, "MISSING")
        new_price, _ = parse_price(row.get(new_col)) if new_col else (None, "MISSING")
        diff, _ = parse_price(row.get(diff_col)) if diff_col else (None, "MISSING")
        effective_date, note = parse_final_effective(row.get(final_col)) if final_col else (None, None)
        items.append((sku, clean_header(row.get(c1)) if c1 else None, clean_header(row.get(c2)) if c2 else None, base_price, new_price, diff, effective_date, note, plan_year, imported_at))
        for stage in parsed_stages:
            price, status = parse_price(row.get(stage["column"]))
            stages.append((sku, stage["column"], stage["start"], stage["end"], price, clean_header(row.get(stage["column"])), status, stage["sort"], imported_at))
    with db() as con:
        batch_id = insert_import_batch(
            con,
            import_type="price_plan",
            filename=file.filename,
            imported_at=imported_at,
            row_count=len(df),
            sku_count=len(skus),
            detail_count=len(stages),
            plan_year=plan_year,
            selected_columns=stage_cols,
        )
        items = [item + (batch_id,) for item in items]
        stages = [stage + (batch_id,) for stage in stages]
        con.executemany(
            """
            insert into price_plan_items
            (sku, category_1, category_2, base_price, new_reference_price, price_diff, final_effective_date, final_effective_note, plan_year, imported_at, batch_id)
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            on conflict(sku) do update set
              category_1=excluded.category_1, category_2=excluded.category_2,
              base_price=excluded.base_price, new_reference_price=excluded.new_reference_price,
              price_diff=excluded.price_diff, final_effective_date=excluded.final_effective_date,
              final_effective_note=excluded.final_effective_note, plan_year=excluded.plan_year,
              imported_at=excluded.imported_at, batch_id=excluded.batch_id
            """,
            items,
        )
        con.executemany(
            """
            insert into price_plan_stages
            (sku, stage_name, start_date, end_date, target_price, raw_value, price_status, sort_order, imported_at, batch_id)
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            stages,
        )
    return {"batch_id": batch_id, "imported_items": len(items), "imported_stages": len(stages)}


def get_latest_snapshot_date(con: sqlite3.Connection) -> str | None:
    row = con.execute(
        """
        select snapshot_date as d
        from price_snapshots
        where batch_id = (
          select id from import_batches where import_type = 'price_statistics' order by imported_at desc, id desc limit 1
        )
        limit 1
        """
    ).fetchone()
    if row and row["d"]:
        return row["d"]
    row = con.execute("select max(snapshot_date) as d from price_snapshots").fetchone()
    return row["d"] if row else None


def load_checks(audit_date: str | None = None, snapshot_batch_id: int | None = None, plan_batch_id: int | None = None) -> list[dict[str, Any]]:
    with db() as con:
        latest_price_batch = latest_batch(con, "price_statistics")
        latest_plan_batch = latest_batch(con, "price_plan")
        price_batch_id = snapshot_batch_id or (latest_price_batch["id"] if latest_price_batch else None)
        active_plan_batch_id = plan_batch_id or (latest_plan_batch["id"] if latest_plan_batch else None)
        snapshot_date = get_latest_snapshot_date(con) or date.today().isoformat()
        if price_batch_id:
            row = con.execute("select snapshot_date from price_snapshots where batch_id = ? limit 1", (price_batch_id,)).fetchone()
            if row and row["snapshot_date"]:
                snapshot_date = row["snapshot_date"]
        check_date = audit_date or date.today().isoformat()
        price_filter = "s.batch_id = ?" if price_batch_id is not None else "s.snapshot_date = ?"
        price_param = price_batch_id if price_batch_id is not None else snapshot_date
        rows = con.execute(
            f"""
            select s.*, i.base_price, i.new_reference_price, i.final_effective_date, i.final_effective_note
            from price_snapshots s
            left join price_plan_items i on i.sku = s.sku
            where {price_filter}
            """,
            (price_param,),
        ).fetchall()
        stages_by_sku: dict[str, list[sqlite3.Row]] = {}
        if active_plan_batch_id:
            stage_rows = con.execute("select * from price_plan_stages where batch_id = ? order by sku, sort_order", (active_plan_batch_id,)).fetchall()
        else:
            stage_rows = con.execute("select * from price_plan_stages order by sku, sort_order").fetchall()
        for st in stage_rows:
            stages_by_sku.setdefault(st["sku"], []).append(st)
        snapshot_skus = {row["sku"] for row in rows}
        plan_items = con.execute(
            "select * from price_plan_items where batch_id = ?" if active_plan_batch_id else "select * from price_plan_items",
            (active_plan_batch_id,) if active_plan_batch_id else (),
        ).fetchall()
    checks = []
    day = date.fromisoformat(check_date)
    for row in rows:
        stages = stages_by_sku.get(row["sku"], [])
        current_stage = None
        for st in stages:
            if st["start_date"] and date.fromisoformat(st["start_date"]) <= day and (not st["end_date"] or day <= date.fromisoformat(st["end_date"])):
                current_stage = st
                break
        target = row["base_price"]
        issue = "NORMAL"
        stage_name = "协同价"
        if not stages and row["base_price"] is None:
            issue = "PLAN_MISSING"
        if current_stage:
            target = current_stage["target_price"]
            stage_name = current_stage["stage_name"]
        elif stages and stages[0]["start_date"] and day < date.fromisoformat(stages[0]["start_date"]):
            target = row["base_price"]
        elif stages:
            latest = stages[-1]
            target = latest["target_price"]
            stage_name = latest["stage_name"]
        if row["price_status"] == "MISSING":
            issue = "PLATFORM_PRICE_MISSING"
        elif row["price_status"] == "NOT_APPLICABLE":
            issue = "PLATFORM_NOT_APPLICABLE"
        elif row["price_status"] == "PARSE_ERROR":
            issue = "PRICE_PARSE_ERROR"
        elif issue == "NORMAL" and target is not None and row["current_price"] is not None and row["current_price"] < target - 0.01:
            if row["final_effective_date"] and date.fromisoformat(row["final_effective_date"]) <= day:
                issue = "NOT_UPDATED_AFTER_EFFECTIVE_DATE"
            else:
                issue = "BELOW_STAGE_TARGET" if current_stage else "BELOW_BASE_PRICE"
        diff = None if target is None or row["current_price"] is None else round(row["current_price"] - target, 2)
        checks.append(
            {
                "snapshot_batch_id": row["batch_id"],
                "plan_batch_id": active_plan_batch_id,
                "audit_date": check_date,
                "snapshot_date": snapshot_date,
                "sku": row["sku"],
                "category_1": row["category_1"],
                "category_2": row["category_2"],
                "platform": row["platform"],
                "current_price": row["current_price"],
                "target_price": target,
                "diff": diff,
                "current_stage": stage_name,
                "issue_type": issue,
                "severity": severity_for_issue(issue),
                "suggested_action": suggestion_for_issue(issue, row["platform"], target),
            }
        )
    for item in plan_items:
        if item["sku"] in snapshot_skus:
            continue
        checks.append(
            {
                "snapshot_batch_id": price_batch_id,
                "plan_batch_id": active_plan_batch_id,
                "audit_date": check_date,
                "snapshot_date": snapshot_date,
                "sku": item["sku"],
                "category_1": item["category_1"],
                "category_2": item["category_2"],
                "platform": None,
                "current_price": None,
                "target_price": item["base_price"],
                "diff": None,
                "current_stage": "价格统计表缺失",
                "issue_type": "SNAPSHOT_SKU_MISSING",
                "severity": "MEDIUM",
                "suggested_action": "补充该 SKU 的价格统计数据",
            }
        )
    return checks


def severity_for_issue(issue: str) -> str:
    if issue in {"BELOW_STAGE_TARGET", "NOT_UPDATED_AFTER_EFFECTIVE_DATE", "PRICE_PARSE_ERROR"}:
        return "HIGH"
    if issue in {"BELOW_BASE_PRICE", "PLATFORM_PRICE_MISSING", "PLAN_MISSING", "SNAPSHOT_SKU_MISSING"}:
        return "MEDIUM"
    if issue == "PLATFORM_NOT_APPLICABLE":
        return "LOW"
    return "NORMAL"


def suggestion_for_issue(issue: str, platform: str | None, target: float | None) -> str:
    if issue in {"BELOW_BASE_PRICE", "BELOW_STAGE_TARGET", "NOT_UPDATED_AFTER_EFFECTIVE_DATE"}:
        return f"核查 {platform or '平台'} 并调整至 {target:.2f} 以上" if target is not None else "核查并调整平台价"
    if issue == "PLATFORM_PRICE_MISSING":
        return f"补充 {platform or '平台'} 价格数据"
    if issue == "PLATFORM_NOT_APPLICABLE":
        return "平台不适用，确认无需监控"
    if issue == "PLAN_MISSING":
        return "补充涨价计划"
    if issue == "SNAPSHOT_SKU_MISSING":
        return "补充价格统计表 SKU"
    if issue == "PRICE_PARSE_ERROR":
        return "修正无法解析的价格字段"
    return "无需处理"


@app.get("/api/checks")
def checks(
    issue_type: str | None = None,
    sku: str | None = None,
    platform: str | None = None,
    audit_date: str | None = None,
    snapshot_batch_id: int | None = None,
    plan_batch_id: int | None = None,
) -> list[dict[str, Any]]:
    data = load_checks(audit_date=audit_date, snapshot_batch_id=snapshot_batch_id, plan_batch_id=plan_batch_id)
    if issue_type:
        data = [x for x in data if x["issue_type"] == issue_type]
    if sku:
        data = [x for x in data if sku.lower() in x["sku"].lower()]
    if platform:
        data = [x for x in data if x["platform"] == platform]
    return data


@app.get("/api/dashboard")
def dashboard(audit_date: str | None = None) -> dict[str, Any]:
    today = date.fromisoformat(audit_date) if audit_date else date.today()
    checks_data = load_checks(audit_date=today.isoformat())
    with db() as con:
        stages = con.execute("select count(*) c from price_plan_stages where start_date between ? and ?", (today.isoformat(), (today + timedelta(days=30)).isoformat())).fetchone()["c"]
        stages15 = con.execute("select count(*) c from price_plan_stages where start_date between ? and ?", (today.isoformat(), (today + timedelta(days=15)).isoformat())).fetchone()["c"]
        recent_batches = import_history()
    platform_distribution: dict[str, dict[str, int]] = {}
    for item in checks_data:
        platform = item["platform"] or "未匹配平台"
        bucket = platform_distribution.setdefault(platform, {"below_target": 0, "missing": 0, "normal": 0, "other": 0})
        if item["issue_type"] in {"BELOW_BASE_PRICE", "BELOW_STAGE_TARGET", "NOT_UPDATED_AFTER_EFFECTIVE_DATE"}:
            bucket["below_target"] += 1
        elif item["issue_type"] == "PLATFORM_PRICE_MISSING":
            bucket["missing"] += 1
        elif item["issue_type"] == "NORMAL":
            bucket["normal"] += 1
        else:
            bucket["other"] += 1
    return {
        "audit_date": today.isoformat(),
        "upcoming_30_days": stages,
        "upcoming_15_days": stages15,
        "below_target_count": sum(1 for x in checks_data if x["issue_type"] in {"BELOW_BASE_PRICE", "BELOW_STAGE_TARGET", "NOT_UPDATED_AFTER_EFFECTIVE_DATE"}),
        "plan_missing_sku_count": len({x["sku"] for x in checks_data if x["issue_type"] == "PLAN_MISSING"}),
        "platform_price_missing_count": sum(1 for x in checks_data if x["issue_type"] == "PLATFORM_PRICE_MISSING"),
        "recent_batches": recent_batches[:6],
        "platform_distribution": [{"platform": k, **v} for k, v in sorted(platform_distribution.items())],
    }


@app.get("/api/reminders")
def reminders(audit_date: str | None = None) -> list[dict[str, Any]]:
    today = date.fromisoformat(audit_date) if audit_date else date.today()
    checks_data = load_checks(audit_date=today.isoformat())
    risk_checks = [x for x in checks_data if x["issue_type"] in {"BELOW_STAGE_TARGET", "BELOW_BASE_PRICE", "NOT_UPDATED_AFTER_EFFECTIVE_DATE"}]
    risk = {(x["sku"], x["current_stage"]): x for x in risk_checks}
    out = []
    with db() as con:
        statuses = {row["reminder_key"]: row["status"] for row in con.execute("select * from reminder_statuses").fetchall()}
        for row in con.execute("select * from price_plan_stages where start_date is not null order by start_date, sku").fetchall():
            start = date.fromisoformat(row["start_date"])
            days = (start - today).days
            key = f"{row['sku']}|{row['stage_name']}|{row['start_date']}"
            risk_item = risk.get((row["sku"], row["stage_name"]))
            reminder_type = "HIGH_RISK" if days < 0 and risk_item else ("D30" if days <= 30 and days > 15 else "D15" if days <= 15 and days > 7 else "D7")
            status = statuses.get(key, "未处理")
            if days in {30, 15} or 0 <= days <= 7 or risk_item:
                out.append(
                    {
                        "reminder_key": key,
                        "sku": row["sku"],
                        "platform": risk_item["platform"] if risk_item else None,
                        "stage_name": row["stage_name"],
                        "start_date": row["start_date"],
                        "end_date": row["end_date"],
                        "target_price": row["target_price"],
                        "current_price": risk_item["current_price"] if risk_item else None,
                        "diff": risk_item["diff"] if risk_item else None,
                        "days_until_start": days,
                        "reminder_type": reminder_type,
                        "severity": "HIGH" if reminder_type == "HIGH_RISK" else "MEDIUM",
                        "status": status,
                        "suggested_action": risk_item["suggested_action"] if risk_item else "关注阶段开始日期并提前确认平台价格",
                    }
                )
    return out


@app.post("/api/reminders/status")
def update_reminder_status(reminder_key: str = Form(...), status: str = Form(...)) -> dict[str, Any]:
    if status not in {"未处理", "处理中", "已处理", "忽略"}:
        raise HTTPException(400, "不支持的提醒状态")
    with db() as con:
        con.execute(
            """
            insert into reminder_statuses (reminder_key, status, updated_at)
            values (?, ?, ?)
            on conflict(reminder_key) do update set status=excluded.status, updated_at=excluded.updated_at
            """,
            (reminder_key, status, datetime.now().isoformat(timespec="seconds")),
        )
    return {"ok": True}


@app.get("/api/checks/export")
def export_checks(audit_date: str | None = None, issue_type: str | None = None, sku: str | None = None, platform: str | None = None) -> FileResponse:
    data = checks(issue_type=issue_type, sku=sku, platform=platform, audit_date=audit_date)
    df = pd.DataFrame(data)
    path = Path(tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name)
    df.to_excel(path, index=False, engine="openpyxl")
    return FileResponse(path, filename=f"price_checks_{audit_date or date.today().isoformat()}.xlsx")


app.mount("/assets", StaticFiles(directory=FRONTEND_DIR), name="assets")


@app.get("/")
def index() -> FileResponse:
    return FileResponse(FRONTEND_DIR / "index.html")
