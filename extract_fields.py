import sys, openpyxl, json, re
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EXCEL_PATH = Path(r"C:\Users\jiancui\Desktop\表名理解\所有样本表表头汇总.xlsx")
OUT_PATH   = Path(__file__).parent / "fields.json"

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱"

TABLE_COLORS = {
    "dw_checkout_trans":  "#1a6eb5",
    "dw_lstg_item":       "#2e7d32",
    "dw_cal_dt":          "#6a1b9a",
    "Cate_ENG":           "#e65100",
    "LEAF_ENG":           "#e65100",
    "CBT_SLR_vertical":   "#00695c",
    "Category_analysis":  "#00695c",
    "Top_Seller_analysis":"#00695c",
    "AMS_ERS_NAME":       "#c62828",
    "SLR_BYR_MAPPING":    "#c62828",
}

wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws = wb.active

# merged cell map
merged_top = {}
for rng in ws.merged_cells.ranges:
    for r in range(rng.min_row, rng.max_row + 1):
        for c in range(rng.min_col, rng.max_col + 1):
            if not (r == rng.min_row and c == rng.min_col):
                merged_top[(r, c)] = (rng.min_row, rng.min_col)

def cell(row, col):
    key = (row, col)
    if key in merged_top:
        r0, c0 = merged_top[key]
        return ws.cell(row=r0, column=c0).value
    return ws.cell(row=row, column=col).value

# Pass 1: synonym groups
groups = {}
cur = None
for rn in range(2, 58):
    e = ws.cell(row=rn, column=5).value
    if not e:
        continue
    e = str(e).strip()
    if e and e[0] in CIRCLED:
        f_raw = cell(rn, 6)
        ver = str(f_raw).strip() if f_raw else ""
        cur = e
        groups[cur] = {"name": e, "entries": [], "verification": ver}
    elif cur and "→" in e:
        left, _, right = e.partition("→")
        raw_fields = [x.strip() for x in left.strip().split("/") if x.strip()]
        cleaned = []
        for f in raw_fields:
            f = re.split(r"[（(]", f)[0].strip()
            if len(f) > 1:
                cleaned.append(f)
        groups[cur]["entries"].append({
            "fields": cleaned,
            "tables": right.strip(),
        })

# Pass 2: field → group mapping
field_to_group = {}
group_synonyms = {}
for gname, gdata in groups.items():
    syns = []
    for entry in gdata["entries"]:
        for f in entry["fields"]:
            field_to_group[f.lower()] = gname
            syns.append({"field": f, "tables": entry["tables"]})
    group_synonyms[gname] = syns

# Pass 3: main field rows
fields = []
for rn in range(2, ws.max_row + 1):
    t_raw = ws.cell(row=rn, column=1).value
    fn    = ws.cell(row=rn, column=3).value
    desc  = ws.cell(row=rn, column=4).value
    if not fn:
        continue
    t_raw = str(t_raw).strip() if t_raw else ""
    parts = t_raw.split(" ", 1)
    table_en = parts[0]
    table_cn = parts[1] if len(parts) > 1 else ""
    fn = str(fn).strip()
    gname = field_to_group.get(fn.lower())
    synonyms = []
    if gname:
        synonyms = [s for s in group_synonyms[gname] if s["field"].lower() != fn.lower()]
    fields.append({
        "table":        table_en,
        "table_cn":     table_cn,
        "table_color":  TABLE_COLORS.get(table_en, "#546e7a"),
        "field":        fn,
        "description":  str(desc).strip() if desc else "",
        "group":        gname or "",
        "verification": groups[gname]["verification"] if gname else "",
        "synonyms":     synonyms,
    })

wb.close()

OUT_PATH.write_text(json.dumps(fields, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
print(f"Done: {len(fields)} fields saved to {OUT_PATH}")
