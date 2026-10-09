import sys, io, re
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import win32com.client

acad = win32com.client.GetActiveObject('AutoCAD.Application')
doc  = acad.ActiveDocument
ms   = doc.ModelSpace

# ── 依 Y 座標定義樓層區間 ──────────────────────────────────────
FLOORS = [
    ('B3F (地下三層)',  -9999,  1800),
    ('B2F (地下二層)',   1800,  5000),
    ('B1F (地下一層)',   5000,  8500),
    ('1F  (一層)',       8500, 12000),
    ('2F  (二層)',      12000, 15500),
    ('3F-5F (三至五層)',15500, 19500),
    ('6F  (六層)',      19500, 22200),
    ('7F  (七層)',      22200, 25500),
    ('RF1 (屋突一層)',  25500, 28700),
    ('RF2 (屋突二層)',  28700, 30000),
    ('RFT (屋突頂層)', 30000, 99999),
]

# 大樑圖層（含標籤與幾何）
MAJOR_BEAM_LAYERS = {
    'S-BEAM-CONC', 'S-BEAM-STEL', 'S-BEAM-STEL-RIGD',
    'S-BEAM-LABL-HORZ', 'S-BEAM-LABL-VERT',
    'S-BEAM-HTCH', 'S-BEAM-DIMS', 'S-BEAM-ELEV',
}
# 小樑圖層
MINOR_BEAM_LAYERS = {
    'S-BEAM-CONC-SECD', 'S-BEAM-STEL-SECD',
    'S-BEAM-LABL-SECD', 'S-BEAM-PRLN',
}
# 版/樓板圖層
SLAB_LAYERS = {
    'S-SLAB-CONC', 'S-SLAB-DECK', 'S-SLAB-EDGE',
    'S-SLAB-BNDY', 'S-SLAB_BNDY', 'S-SLAB-LABL',
    'S-SLAB-ELEV', 'S-SLAB-HTCH', 'S-SLAB-OPNG', 'S-SLAB-OPNX',
}

ALL_TARGET = MAJOR_BEAM_LAYERS | MINOR_BEAM_LAYERS | SLAB_LAYERS

# 尺寸表標題關鍵字 → 對應類別
TABLE_KEYWORDS = {
    '大梁': '大樑尺寸表',
    '小梁': '小樑尺寸表',
    '版':   '版尺寸表',
}

def fix_text(raw):
    for enc in ('big5', 'gbk', 'cp950'):
        try:
            return raw.encode('latin-1').decode(enc)
        except:
            pass
    return raw

def clean(txt):
    txt = re.sub(r'\\[A-Za-z][^;]*;', '', txt)
    txt = re.sub(r'\{[^}]*\}', '', txt)
    return txt.strip()

def get_floor(y):
    for name, ymin, ymax in FLOORS:
        if ymin <= y < ymax:
            return name
    return '未分類'

# ── (1) 從 AcDbTable 讀尺寸表 ────────────────────────────────
# dim_tables: {'大樑尺寸表': [('編號群', '尺寸'), ...], ...}
dim_tables = {}

for layout in doc.Layouts:
    blk = layout.Block
    for entity in blk:
        try:
            if entity.ObjectName != 'AcDbTable':
                continue
            rows = entity.Rows
            cols = entity.Columns
            if rows < 2 or cols < 2:
                continue

            # 第一列：標題
            title = fix_text(entity.GetText(0, 0)).strip()
            title = re.sub(r'\s+', '', title)  # 移除全形空白
            matched_key = None
            for kw, label in TABLE_KEYWORDS.items():
                if kw in title:
                    matched_key = label
                    break
            if not matched_key:
                continue

            entries = []
            for r in range(2, rows):  # 跳過標題列和欄位列
                id_cell  = fix_text(entity.GetText(r, 0)).strip()
                dim_cell = fix_text(entity.GetText(r, 1)).strip()
                if id_cell and dim_cell:
                    entries.append((id_cell, dim_cell))

            if entries:
                dim_tables[matched_key] = entries
        except:
            pass

# ── (2) 從文字圖層收集各樓層標籤 ─────────────────────────────
data      = {f[0]: {'大樑': set(), '小樑': set(), '版': set()} for f in FLOORS}
geo_count = {f[0]: {'大樑': 0,    '小樑': 0,     '版': 0}     for f in FLOORS}

for entity in ms:
    try:
        layer = entity.Layer
        if layer not in ALL_TARGET:
            continue
        pt    = entity.InsertionPoint
        floor = get_floor(pt[1])
        if floor == '未分類':
            continue

        if layer in MAJOR_BEAM_LAYERS:
            cat = '大樑'
        elif layer in MINOR_BEAM_LAYERS:
            cat = '小樑'
        else:
            cat = '版'

        obj_name = entity.ObjectName
        if obj_name in ('AcDbText', 'AcDbMText'):
            txt = clean(fix_text(entity.TextString))
            if txt and len(txt) >= 2:
                data[floor][cat].add(txt)
        else:
            geo_count[floor][cat] += 1
    except:
        pass

# ── (3) 輸出 ──────────────────────────────────────────────────
output_lines = []
output_lines.append("=" * 70)
output_lines.append("  各樓層 樑、版 圖層資料彙整表（供 Revit 族群建立使用）")
output_lines.append("=" * 70)

for floor_name, ymin, ymax in FLOORS:
    d = data[floor_name]
    g = geo_count[floor_name]
    total = g['大樑'] + g['小樑'] + g['版'] + len(d['大樑']) + len(d['小樑']) + len(d['版'])
    if total == 0:
        continue

    output_lines.append(f"\n{'─'*70}")
    output_lines.append(f"【{floor_name}】")
    output_lines.append(f"{'─'*70}")

    for cat, layer_set, table_key in [
        ('大樑',     MAJOR_BEAM_LAYERS, '大樑尺寸表'),
        ('小樑',     MINOR_BEAM_LAYERS, '小樑尺寸表'),
        ('版（樓板）', SLAB_LAYERS,     '版尺寸表'),
    ]:
        key    = cat if cat != '版（樓板）' else '版'
        labels = sorted(d[key])
        cnt    = g[key]
        output_lines.append(f"\n  ▌{cat}")
        output_lines.append(f"  　使用圖層：{', '.join(sorted(layer_set))}")
        output_lines.append(f"  　幾何物件數：{cnt}")
        if labels:
            output_lines.append(f"  　標籤清單（{len(labels)} 種）：")
            for lb in labels:
                output_lines.append(f"  　　• {lb}")
        else:
            output_lines.append(f"  　標籤清單：（無文字標籤）")

        # 附上尺寸表（若有）
        if table_key in dim_tables:
            output_lines.append(f"  　{table_key}：")
            for id_cell, dim_cell in dim_tables[table_key]:
                output_lines.append(f"  　　  {id_cell:<30} {dim_cell}")

output_lines.append(f"\n{'='*70}")

# 附錄：印出所有尺寸表（集中顯示）
if dim_tables:
    output_lines.append("\n\n" + "=" * 70)
    output_lines.append("  圖面尺寸表（AcDbTable 擷取）")
    output_lines.append("=" * 70)
    for tbl_name, entries in dim_tables.items():
        output_lines.append(f"\n  ▌{tbl_name}")
        output_lines.append(f"  　{'編號':<32} 尺寸(cm)")
        output_lines.append(f"  　{'─'*46}")
        for id_cell, dim_cell in entries:
            output_lines.append(f"  　  {id_cell:<30} {dim_cell}")

output_lines.append(f"\n{'='*70}")

result = "\n".join(output_lines)
print(result)

REPORT_PATH = Path(__file__).parent / 'beams_slabs_report.txt'
with open(REPORT_PATH, 'w', encoding='utf-8') as f:
    f.write(result)

print(f"\n✔ 已儲存至 {REPORT_PATH}")
