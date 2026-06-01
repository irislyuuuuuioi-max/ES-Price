import io
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from backend.dev_server import _UploadFile
from backend.app import main


def run():
    main.startup()
    price = io.BytesIO()
    pd.DataFrame(
        [
            ["价格快照日期：2026-06-26", "", "", "", "", "", ""],
            ["SKU", "一级类目", "二级类目", "协同价（旧）", "协同价（新）", "OS-ES", "HD-GJ"],
            ["A1", "Bath", "Tub", 260, 280, "$279.00", "#N/A"],
            ["A2", "Bath", "Sink", 100, 120, "/", "90"],
        ]
    ).to_excel(price, index=False, header=False)
    price.seek(0)
    preview = main.preview_price_statistics(_UploadFile("price.xlsx", price), None)
    assert preview["header_row"] == 2
    assert preview["snapshot_date"] == "2026-06-26"
    assert "OS-ES" in preview["detected_platform_columns"]
    price.seek(0)
    assert main.import_price_statistics(_UploadFile("price.xlsx", price), preview["snapshot_date"], "OS-ES,HD-GJ")["imported"] == 4

    plan = io.BytesIO()
    pd.DataFrame(
        [
            ["备注"],
            ["SKU", "一级类目", "二级类目", "协同价", "协同价新参考", "新旧差价", "最终执行时间", "Red White Blue（6.25-）", "Summer Bath Phase 2（8.06-10.04）", "10.05后价格"],
            ["A1", "Bath", "Tub", 260, 280, 20, "2026-06-25", 280, 300, 320],
            ["A2", "Bath", "Sink", 100, 120, 20, "5月，所有平台不低于协同价", 120, 130, 150],
        ]
    ).to_excel(plan, index=False, header=False)
    plan.seek(0)
    plan_preview = main.preview_price_plan(_UploadFile("plan.xlsx", plan), 2026)
    assert plan_preview["header_row"] == 2
    assert plan_preview["detected_stage_columns"][1]["start_date"] == "2026-08-06"
    assert plan_preview["detected_stage_columns"][1]["end_date"] == "2026-10-04"
    plan.seek(0)
    imported = main.import_price_plan(
        _UploadFile("plan.xlsx", plan),
        2026,
        "Red White Blue（6.25-）,Summer Bath Phase 2（8.06-10.04）,10.05后价格",
    )
    assert imported["imported_items"] == 2
    assert imported["imported_stages"] == 6
    assert isinstance(main.dashboard(), dict)
    assert len(main.checks()) >= 4
    print("smoke ok")


if __name__ == "__main__":
    run()
