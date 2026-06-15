#!/usr/bin/env python3
"""
翊鼎國際廣告 · 資料更新腳本
用法：python update_data.py <Excel檔路徑>
每次有新投組月報 Excel 時執行，產生 db.json 後 git push 即可更新網頁。
"""
import sys, json, re, pathlib, openpyxl
from datetime import datetime

SCRIPT_DIR = pathlib.Path(__file__).parent
DB_JSON    = SCRIPT_DIR / "db.json"
MONTH_PAT  = re.compile(r'^1\d\d\.\d{2}$')
VOL_PAT    = re.compile(r'月成交總量', re.IGNORECASE)

def num(v):
    if v is None: return None
    try:    return round(float(v), 2)
    except: return None

def parse_month_sheet(ws):
    HEADER = 7
    col = {}
    for c in range(1, 30):
        v = str(ws.cell(row=HEADER, column=c).value or "").strip()
        if v: col[v] = c
    C = {
        "agency":  col.get("承接公司名", 1), "name": col.get("工地名稱", 3),
        "units":   col.get("戶", 4),         "parking": col.get("車", 5),
        "license": col.get("取得使照", 6),   "entry": col.get("進場日期", 7),
        "volume":  col.get("案量", 8),       "month_pnl": col.get("當月合計", 14),
        "cum_pnl": col.get("總合計", 15),    "remain_u": col.get("剩餘戶數", 19),
        "remain_p":col.get("剩餘車位", 20),  "oh": col.get("公司管銷", 21),
    }
    rows = []
    for r in range(HEADER + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=C["name"]).value
        if not name: continue
        name = str(name).strip()
        if not name or "合計" in name: break
        rows.append({
            "project_name":   name,
            "agency_name":    str(ws.cell(row=r, column=C["agency"]).value or "").strip(),
            "units":          num(ws.cell(row=r, column=C["units"]).value),
            "parking":        num(ws.cell(row=r, column=C["parking"]).value),
            "license_status": str(ws.cell(row=r, column=C["license"]).value or "").strip(),
            "entry_date":     str(ws.cell(row=r, column=C["entry"]).value or "").strip(),
            "total_volume":   num(ws.cell(row=r, column=C["volume"]).value),
            "month_pnl":      num(ws.cell(row=r, column=C["month_pnl"]).value),
            "cumulative_pnl": num(ws.cell(row=r, column=C["cum_pnl"]).value),
            "remain_units":   num(ws.cell(row=r, column=C["remain_u"]).value),
            "remain_parking": num(ws.cell(row=r, column=C["remain_p"]).value),
            "company_oh":     num(ws.cell(row=r, column=C["oh"]).value),
            "weeks":          [num(ws.cell(row=r, column=c).value) for c in range(9, 14)],
        })
    return rows

def parse_vol_sheet(ws):
    col_vis = col_deals = col_buy = None
    hdr = 3
    for r in range(1, 8):
        for c in range(1, 50):
            v = str(ws.cell(row=r, column=c).value or "").strip()
            if v == "來人":   col_vis   = c
            if v == "月成交": col_deals = c
            if v == "月買回": col_buy   = c
        if col_vis: hdr = r; break
    result = {}
    for r in range(hdr + 1, ws.max_row + 1):
        nm = ws.cell(row=r, column=1).value
        if not nm: continue
        nm = str(nm).strip()
        if "合計" in nm: break
        result[nm] = {
            "visitors": num(ws.cell(row=r, column=col_vis).value)   if col_vis   else None,
            "deals":    num(ws.cell(row=r, column=col_deals).value) if col_deals else None,
        }
    return result

def parse_annual(wb):
    candidates = [s for s in wb.sheetnames if "銷售量" in s]
    if not candidates: return []
    ws = wb[candidates[0]]
    rows = []
    for r in range(3, 20):
        mn = ws.cell(row=r, column=1).value
        if not mn: continue
        rows.append({
            "year_month":   str(mn).strip(),
            "visitors":     num(ws.cell(row=r, column=2).value),
            "units_sold":   num(ws.cell(row=r, column=3).value),
            "parking_sold": num(ws.cell(row=r, column=4).value),
            "pnl":          num(ws.cell(row=r, column=6).value),
            "deals":        num(ws.cell(row=r, column=8).value),
        })
    return [x for x in rows if x["year_month"] and x["pnl"] is not None]

def load_existing():
    if DB_JSON.exists():
        return json.loads(DB_JSON.read_text(encoding="utf-8"))
    return {"months": [], "portfolio": {}, "weekly": {}, "rollup": [], "updated_at": ""}

def merge_and_save(existing, new_portfolio, new_weekly, new_annual, new_months):
    existing["portfolio"].update(new_portfolio)
    existing["weekly"].update(new_weekly)

    # merge months list (deduplicated, sorted)
    all_months = sorted(set(existing.get("months", []) + new_months))
    existing["months"] = all_months

    # merge rollup (upsert by year_month)
    rollup_map = {r["year_month"]: r for r in existing.get("rollup", [])}
    for r in new_annual:
        rollup_map[r["year_month"]] = r
    existing["rollup"] = [rollup_map[k] for k in sorted(rollup_map)]

    existing["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    DB_JSON.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
    return existing

def run(filepath):
    print(f"📂 讀取：{filepath}")
    wb = openpyxl.load_workbook(filepath, data_only=True)

    month_sheets = [s for s in wb.sheetnames if MONTH_PAT.match(s.strip())]
    if not month_sheets:
        print("❌ 找不到月報工作表（格式需為 1XX.XX，如 114.05）")
        sys.exit(1)

    print(f"✅ 偵測到月份：{month_sheets}")

    new_portfolio, new_weekly = {}, {}

    for ym in month_sheets:
        rows = parse_month_sheet(wb[ym])
        vol_sn = next((s for s in wb.sheetnames if ym in s and VOL_PAT.search(s)), None)
        vol    = parse_vol_sheet(wb[vol_sn]) if vol_sn else {}

        for r in rows:
            v = vol.get(r["project_name"], {})
            r["visitors"] = v.get("visitors")
            r["deals"]    = v.get("deals")
            # weekly: key by project_name
            k = r["project_name"]
            for i, pnl in enumerate(r.pop("weeks")):
                new_weekly.setdefault(k, [])
                entry = {"year_month": ym, "week_label": f"W{i+1}", "pnl": pnl}
                # remove old entry for same month+week, then add
                new_weekly[k] = [
                    x for x in new_weekly[k]
                    if not (x["year_month"] == ym and x["week_label"] == f"W{i+1}")
                ] + [entry]
        new_portfolio[ym] = rows

    annual = parse_annual(wb)

    existing = load_existing()
    # merge weekly with existing
    for k, entries in new_weekly.items():
        existing_wk = existing.get("weekly", {}).get(k, [])
        # remove entries for months being updated
        existing_wk = [x for x in existing_wk if x["year_month"] not in month_sheets]
        new_weekly[k] = existing_wk + entries

    result = merge_and_save(existing, new_portfolio, new_weekly, annual, month_sheets)

    total_projects = len(set(
        r["project_name"]
        for m in result["portfolio"].values()
        for r in m
    ))
    print(f"✅ db.json 已更新")
    print(f"   月份：{result['months']}")
    print(f"   建案：{total_projects} 個")
    print(f"   更新時間：{result['updated_at']}")
    print()
    print("📌 下一步：")
    print("   git add db.json")
    ts = result['updated_at']
    print(f'   git commit -m "更新資料 {ts}"')
    print("   git push")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python update_data.py <Excel檔路徑>")
        print("範例：python update_data.py ~/Downloads/5131公司2025每周盈虧.xlsx")
        sys.exit(1)
    run(sys.argv[1])
