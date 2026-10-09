"""
generate_excel.py
從 cad_dim_tables.txt 與 cad_labels_xy.txt 產生結構構件尺寸表 Excel。

用法：
  python generate_excel.py \
    --dim-tables <path_to_cad_dim_tables.txt> \
    --labels-xy  <path_to_cad_labels_xy.txt> \
    --output     <output_excel_path.xlsx>
"""

import argparse
import re
import sys
from collections import defaultdict

def _natural_key(s):
    """自然排序鍵：b1 < b2 < b10（數字部分以整數比較）"""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', str(s))]

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not found. Run: pip install openpyxl")
    sys.exit(1)


# ── 顏色設定 ────────────────────────────────────────────────
TYPE_COLORS = {
    "大梁": "DDEEFF",
    "小梁": "DDFFDD",
    "版":   "FFF3CC",
    "柱":   "F0E6FF",
    "鋼梁": "FFE4E1",
    "鋼柱": "FFD9E8",
}
HEADER_BG    = "2F4F8F"
HEADER_FG    = "FFFFFF"
COL_WIDTHS   = [10, 24, 20]
HEADERS      = ["類型", "編號", "尺寸 (cm)"]


# ── 解析 cad_dim_tables.txt ──────────────────────────────────
def parse_dim_tables(path):
    """
    回傳：
      tables   = [ { "title": str, "floor": str, "type": str,
                     "rows": [{"id": str, "size": str}],
                     "default_size": str or None } ]
    """
    tables = []
    current = None

    with _open_any_encoding(path) as f:
        for line in f:
            line = line.rstrip()
            if line.startswith("=== TABLE"):
                if current:
                    tables.append(current)
                current = {"title": line, "floor": "", "type": "",
                           "rows": [], "default_size": None}
            elif current is not None and line.startswith("["):
                cells = re.findall(r'\[([^\]]*)\]', line)
                if len(cells) < 2:
                    continue
                raw_id   = _clean(cells[0])
                raw_size = _clean(cells[1])

                # 偵測標題列：第二格空白且第一格含「尺寸表」，或第一格是表頭「編號」
                if "尺寸表" in raw_id or (not raw_size and re.search(r'FL|fl', raw_id)):
                    _parse_title(current, raw_id)
                    continue
                if re.match(r'^編[\s　]*號$', raw_id) or raw_id in ("編號",):
                    continue  # 跳過欄位標題列
                if not raw_id or not raw_size:
                    continue

                if "未特別註明" in raw_id:
                    current["default_size"] = raw_size
                else:
                    # 展開範圍（如 b3~b5 → b3, b4, b5）
                    for id_ in _expand_range(raw_id):
                        current["rows"].append({"id": id_, "size": raw_size})

    if current:
        tables.append(current)

    # 補充：從 === TABLE N Layer:... === 行之後的第一個有意義行判斷樓層
    for t in tables:
        if not t["floor"] and t["rows"]:
            _infer_floor_from_title(t)

    return tables


def _clean(s):
    # 移除 AutoCAD %%U（底線）、%%O（上線）等格式碼
    s = re.sub(r'%%[UuOoLlXx]', '', s)
    # 移除 MText 格式碼 \xxx; 但 \F 字型名稱例外（保留其後的文字）
    s = re.sub(r'\\[A-Za-z][^;]*;', '', s)
    # 將 {content} 展開為 content（保留逗號分隔符）
    s = re.sub(r'\{([^}]*)\}', r'\1', s)
    return s.strip()


def _parse_title(table, text):
    """從表格第一列文字推斷樓層與構件類型"""
    text = _clean(text)
    # 去掉空白後比對，以應對 "小  梁  尺  寸  表" 的寬字元間距
    compact = re.sub(r'\s+', '', text)
    for kw in ["大梁", "小梁", "鋼梁", "鋼柱", "版", "柱", "結構牆"]:
        if kw in compact:
            table["type"] = kw
            break
    # 樓層：先嘗試 B3FL~B2FL 或 2FL~5FL 格式，再嘗試單個 FL
    m = re.search(r'([A-Z0-9]+FL[~～][A-Z0-9]+FL)', compact)
    if not m:
        m = re.search(r'([A-Z0-9]+FL)', compact)
    if m:
        table["floor"] = m.group(1)


def _infer_floor_from_title(table):
    """從 TABLE 標題列（title 欄位）推斷樓層"""
    text = table["title"]
    m = re.search(r'([BRr0-9~]+FL[\w~\-]*)', text, re.IGNORECASE)
    if m:
        table["floor"] = m.group(1).upper()


def _expand_range(s):
    """
    將 b3~b5 展開為 [b3, b4, b5]；
    B1~B4 展開為 [B1, B2, B3, B4]；
    否則直接回傳 [s]（可能是逗號分隔）
    """
    results = []
    for part in re.split(r'[、,，]', s):
        part = part.strip()
        m = re.match(r'^([A-Za-z]+)(\d+)[~～]([A-Za-z]+)(\d+)$', part)
        if m and m.group(1) == m.group(3):
            prefix = m.group(1)
            start, end = int(m.group(2)), int(m.group(4))
            for i in range(start, end + 1):
                results.append(f"{prefix}{i}")
        else:
            if part:
                results.append(part)
    return results if results else [s]


# ── 解析 cad_labels_xy.txt ──────────────────────────────────
LABEL_LAYERS = {
    "S-BEAM-LABL-HORZ": "大梁",
    "S-BEAM-LABL-VERT": "大梁",
    "S-BEAM-LABL-SECD": "小梁",
    "S-SLAB-LABL":      "版",
    "S-COLS-LABL":      "柱",
}
IGNORE_LABELS = {
    "S-BEAM-LABL-HORZ", "S-BEAM-LABL-VERT", "S-BEAM-LABL-SECD",
    "S-SLAB-LABL", "S-COLS-LABL",
    "水平向大梁編號", "垂直向大梁編號", "小梁編號", "版編號", "柱編號",
}

# 精確樓層平面圖標題 → FL 代碼（字典 key 為去除空白後的標題全文）
# 只識別「個別樓層結構平面圖」，不匹配「X至Y層結構平面圖」等合訂標題
# 三至五層結構平面圖 = 一張涵蓋三到五層的平面圖 → 視為一個平面圖
EXACT_FLOOR_TITLES = {
    '地下三層結構平面圖':   'B3FL',
    '地下二層結構平面圖':   'B2FL',
    '地下一層結構平面圖':   'B1FL',
    '一層結構平面圖':       '1FL',
    '二層結構平面圖':       '2FL',
    '三至五層結構平面圖':   '3FL~5FL',   # 一張涵蓋三到五層
    '二層至六層結構平面圖': '2FL~6FL',   # 一張涵蓋二到六層
    '六層結構平面圖':       '6FL',
    '七層結構平面圖':       '7FL',
    '屋突一層結構平面圖':   'R1FL',
    '屋突二層結構平面圖':   'R2FL',
    '屋突頂層結構平面圖':   'PRFL',
}

# 平面圖 FL → 尺寸表 FL（查表 fallback 用，不影響分頁名稱）
FL_TO_TABLE_FL = {
    # 前一專案 (251015)
    '2FL':      '2FL~5FL',
    '3FL~5FL':  '2FL~5FL',
    '6FL':      '6FL~7FL',
    '7FL':      '6FL~7FL',
    'PRFL':     'R2FL~PRFL',
    # 新專案 (1141120)：大梁/版用 2FL~R1FL；小梁用 2FL~PRFL
    '2FL~6FL':  '2FL~R1FL',
    'R1FL':     '2FL~R1FL',
    # 小梁表 2FL~PRFL 涵蓋 R1FL / R2FL / PRFL（動態 fallback 無法自動比對，需明確對應）
    'R2FL':     '2FL~PRFL',
    'PRFL':     '2FL~PRFL',
}

# 平面圖 FL → Excel 工作表顯示名稱（一張圖一個分頁）
FL_DISPLAY_NAME = {
    'B3FL':    'B3FL',
    'B2FL':    'B2FL',
    'B1FL':    'B1FL',
    '1FL':     '1FL',
    '2FL':     '2FL',
    '3FL~5FL': '3FL-5FL',
    '2FL~6FL': '2FL-6FL',
    '6FL':     '6FL',
    '7FL':     '7FL',
    'R1FL':    'R1FL',
    'R2FL':    'R2FL',
    'PRFL':    'PRFL',
    '鋼構件':  '鋼構件',
}

# 合併尺寸表 FL → 對應的個別平面圖 FL 清單（補完步驟用）
TABLE_FL_TO_FLOORS = {
    # 前一專案 (251015)
    'B3FL~B2FL':  ['B3FL', 'B2FL'],
    '2FL~5FL':    ['2FL', '3FL~5FL'],
    '6FL~7FL':    ['6FL', '7FL'],
    'R2FL~PRFL':  ['R2FL', 'PRFL'],
    # 新專案 (1141120)
    'B1FL~1FL':   ['B1FL', '1FL'],
    '2FL~R1FL':   ['2FL~6FL', 'R1FL'],
    '2FL~PRFL':   ['2FL~6FL', 'R1FL', 'R2FL', 'PRFL'],
}

# Excel 工作表固定順序（涵蓋兩個專案的所有可能分頁）
SHEET_ORDER = ['B3FL', 'B2FL', 'B1FL', '1FL',
               '2FL', '3FL-5FL', '2FL-6FL',
               '6FL', '7FL', 'R1FL', 'R2FL', 'PRFL', '鋼構件']

# 構件類型在同一分頁內的排序：大梁 → 小梁 → 版 → 柱 → 鋼樑 → 鋼柱
TYPE_ORDER = ['大梁', '小梁', '版', '柱', '鋼梁', '鋼柱', '構件']

def _detect_floor_fl(text):
    """
    從中文樓層標題精確偵測 FL 代碼（純字典比對，不用子字串搜尋）。
    避免 '地下三層至地下二層結構平面圖' 誤判為 B2FL 等問題。
    """
    compact = re.sub(r'[\s　]+', '', text)
    return EXACT_FLOOR_TITLES.get(compact)

def _preprocess_big5(data: bytes) -> bytes:
    """
    預先處理 AutoCAD 輸出的 Big5 位元組流中 Python cp950 無法解碼的特殊字元：
      0xA1 0x42 → b','  （Big5 頓號 `、` → ASCII 逗號，作為分隔符）
      0xA1 0xD1 → b'x'  （Big5 乘號 `×` → ASCII x，保留可讀性）
    """
    data = data.replace(b'\xa1\x42', b',')
    data = data.replace(b'\xa1\xd1', b'x')
    return data

def _open_any_encoding(path):
    """
    讀取檔案，先做 Big5 特殊位元組預處理，再依序嘗試 UTF-8、cp950 解碼。
    回傳 io.StringIO，行為與 file object 相同。
    """
    import io
    raw = open(path, 'rb').read()
    # 嘗試判斷是否含 Big5 特徵位元組
    if b'\xa1' in raw:
        raw = _preprocess_big5(raw)
    for enc in ("utf-8", "utf-8-sig", "cp950", "gbk"):
        try:
            text = raw.decode(enc)
            return io.StringIO(text)
        except (UnicodeDecodeError, UnicodeError):
            pass
    return io.StringIO(raw.decode("utf-8", errors="ignore"))


def parse_labels_xy(path):
    """
    回傳：
      floor_titles = [ {"text": str, "x": float, "y": float, "fl": str} ]
      labels       = [ {"layer": str, "type": str, "label": str,
                         "x": float, "y": float,
                         "bracket_size": str or None} ]

    樓層標題偵測邏輯（先找圖！）：
      - 偵測個別樓層平面圖標題（精確比對），記錄 X/Y 座標
      - X 座標用於同 Y 水平並排版面的區分
    """
    floor_titles = []
    labels = []

    with _open_any_encoding(path) as f:
        for line in f:
            line = line.rstrip()
            parts = line.split("||")
            if len(parts) < 3:
                continue
            layer, coord, text = parts[0], parts[1], parts[2]
            text = _clean(text)
            if not text:
                continue
            try:
                x = float(coord.split(",")[0])
                y = float(coord.split(",")[1])
            except Exception:
                continue

            # ── 先找圖：偵測個別樓層平面圖標題（含 X 座標）──────────
            if '結構平面圖' in text and not re.search(r'S1000\d-\d+_', text):
                fl = _detect_floor_fl(text)
                if fl:
                    floor_titles.append({"text": text, "x": x, "y": y, "fl": fl})
                    continue

            # S10000-N_ 大圖號：跳過
            if re.search(r'S1000\d-\d+_', text):
                continue

            ltype = LABEL_LAYERS.get(layer)
            if not ltype:
                continue
            if text in IGNORE_LABELS:
                continue

            # 提取括號尺寸
            bm = re.search(r'\((\d+[×xX×]\d+(?:[×xX×]\d+)*)\)', text)
            bracket_size = bm.group(1) if bm else None

            # 移除括號部份得到純編號
            clean_label = re.sub(r'\(.*?\)', '', text).strip()

            labels.append({
                "layer":        layer,
                "type":         ltype,
                "label":        text,
                "clean_label":  clean_label,
                "bracket_size": bracket_size,
                "x":            x,
                "y":            y,
            })

    # 去除同 FL 同 Y 同 X 區域的重複圖名
    seen = set()
    unique_titles = []
    for t in floor_titles:
        key = (t["fl"], round(t["y"] / 100), round(t["x"] / 500))
        if key not in seen:
            seen.add(key)
            unique_titles.append(t)

    return unique_titles, labels


# ── 計算樓層 Y 邊界 ──────────────────────────────────────────

# 樓層高低排序（Y 相同時，較低樓層排在前，確保捕捉到較低 Y 的內容）
FLOOR_LEVEL = {
    'B3FL': 0, 'B2FL': 1, 'B1FL': 2,
    '1FL': 3,
    '2FL': 4, '3FL~5FL': 5, '2FL~6FL': 5,
    '6FL': 6, '7FL': 7,
    'R1FL': 8, 'R2FL': 9, 'PRFL': 10,
}


_SAME_Y_TOLERANCE = 200   # Y 差距在此範圍內視為「同 Y」（水平並排）


def build_floor_boundaries(floor_titles):
    """
    先找圖：依個別樓層平面圖標題建立二維（X+Y）邊界。

    ● 垂直堆疊版面（各樓層 Y 不同）→ 以 Y 中點切分，xmin/xmax = ±∞
    ● 水平並排版面（同 Y 多樓層）→ 同 Y 群組內改以 X 中點切分，
      各樓層在 Y 方向共用相同範圍，X 方向各自有獨立區段
    ● X 欄繼承（關鍵修正）：水平並排後，上方單一樓層若標題 X 落在
      某已知 X 欄內，則繼承該欄的 xmin/xmax，避免全寬誤抓鄰欄標籤。

    floor_titles 各項需包含 {"x": float, "y": float, "fl": str}。
    """
    if not floor_titles:
        return []

    # ── Step 1：依 Y 分組（Y 差距 ≤ SAME_Y_TOLERANCE 視為同組）────
    sorted_by_y = sorted(floor_titles, key=lambda t: t["y"])
    y_groups = []   # [ [title, ...], [title, ...], ... ]
    cur_group = [sorted_by_y[0]]
    for t in sorted_by_y[1:]:
        if abs(t["y"] - cur_group[0]["y"]) <= _SAME_Y_TOLERANCE:
            cur_group.append(t)   # 同 Y 群組
        else:
            y_groups.append(cur_group)
            cur_group = [t]
    y_groups.append(cur_group)

    # ── Step 2：計算每個 Y 群組的 Y 邊界（以群組代表 Y 為基準）───────
    def _group_y(grp):
        return sum(t["y"] for t in grp) / len(grp)   # 群組代表 Y（平均）

    group_ys = [_group_y(g) for g in y_groups]

    # ── Step 3：收集所有水平並排群組建立的 X 欄定義 ─────────────────
    # x_columns = [ (xmin, xmax), ... ] — 每次遇到多標題群組就累積
    x_columns = []

    def _find_column(title_x):
        """找出 title_x 落在哪個已知 X 欄，回傳 (xmin, xmax) 或 None"""
        for xmin, xmax in x_columns:
            if xmin <= title_x < xmax:
                return xmin, xmax
        return None

    boundaries = []
    for gi, group in enumerate(y_groups):
        ymin = (group_ys[gi-1] + group_ys[gi]) / 2 if gi > 0 else -1e9
        # ymax 用「下一組標題 Y」而非中點：因為標題在圖面下方，
        # 圖面內容向上延伸，兩層圖的真實分界就在下一層標題位置。
        # 這樣 R1FL 才能涵蓋其實際延伸到 Y≈16400 的內容，
        # 不會被過低的中點切掉後誤歸到 R2FL。
        ymax = group_ys[gi+1] if gi < len(y_groups)-1 else 1e9

        if len(group) == 1:
            t = group[0]
            tx = t.get("x", 0)
            # ── X 欄繼承：若標題 X 落在已知欄內，繼承該欄邊界（mode=xy）
            col = _find_column(tx)
            if col:
                xmin_c, xmax_c = col
                boundaries.append({
                    "fl":    t["fl"],
                    "ymin":  ymin, "ymax":  ymax,
                    "xmin":  xmin_c, "xmax": xmax_c,
                    "title_y": t["y"],
                    "title_x": tx,
                    "mode":  "xy",
                })
            else:
                # 全 X 範圍（無水平並排歷史）
                boundaries.append({
                    "fl":    t["fl"],
                    "ymin":  ymin, "ymax":  ymax,
                    "xmin": -1e9,  "xmax":  1e9,
                    "title_y": t["y"],
                    "title_x": tx,
                    "mode":  "y_only",
                })
        else:
            # ── 水平並排：同 Y 群組以 X 中點再切分 ──────────────────
            sorted_group_by_x = sorted(group, key=lambda t: t.get("x", 0))

            # 若所有標題都落在已知 X 欄內（且各對應不同欄），繼承現有欄邊界
            # 避免 R2FL/PRFL 等「同圖兩份」因中點過寬而納入鄰欄標籤
            cols_for_group = [_find_column(t.get("x", 0)) for t in sorted_group_by_x]
            use_inherited = (
                all(c is not None for c in cols_for_group) and
                len(set(cols_for_group)) == len(sorted_group_by_x)
            )

            if use_inherited:
                for t, col in zip(sorted_group_by_x, cols_for_group):
                    xmin_c, xmax_c = col
                    boundaries.append({
                        "fl":    t["fl"],
                        "ymin":  ymin, "ymax":  ymax,
                        "xmin":  xmin_c, "xmax": xmax_c,
                        "title_y": t["y"],
                        "title_x": t.get("x", 0),
                        "mode":  "xy",
                    })
            else:
                title_xs = [t.get("x", 0) for t in sorted_group_by_x]
                for xi, t in enumerate(sorted_group_by_x):
                    xmin = (title_xs[xi-1] + title_xs[xi]) / 2 if xi > 0 else -1e9
                    xmax = (title_xs[xi] + title_xs[xi+1]) / 2 if xi < len(title_xs)-1 else 1e9
                    boundaries.append({
                        "fl":    t["fl"],
                        "ymin":  ymin, "ymax":  ymax,
                        "xmin":  xmin, "xmax":  xmax,
                        "title_y": t["y"],
                        "title_x": t.get("x", 0),
                        "mode":  "xy",
                    })
                    # 累積 X 欄定義（供後續樓層繼承）
                    if (xmin, xmax) not in x_columns:
                        x_columns.append((xmin, xmax))

    return boundaries


def get_floor_for_y(y, boundaries, x=None):
    """
    回傳 (fl_code, is_boundary_edge)。

    ● 垂直模式（mode='y_only'）：只用 Y 判斷
    ● 水平並排模式（mode='xy'）：X+Y 同時判斷
      若 x=None（舊呼叫方式），water-parallel 模式退化為 Y-only（容錯）
    """
    for b in boundaries:
        if not (b["ymin"] <= y < b["ymax"]):
            continue
        if b.get("mode") == "xy" and x is not None:
            if not (b["xmin"] <= x < b["xmax"]):
                continue
        is_edge = (abs(y - b["ymin"]) < 500 or abs(y - b["ymax"]) < 500)
        if b.get("mode") == "xy" and x is not None:
            is_edge = is_edge or abs(x - b["xmin"]) < 500 or abs(x - b["xmax"]) < 500
        return b["fl"], is_edge
    return "未分類", False


# ── 主流程 ───────────────────────────────────────────────────
def build_sheet_data(tables, floor_titles, labels):
    """
    先找圖！再查表！
    1. 先以個別樓層平面圖標題建立精確的 Y 邊界 → 每個標籤取得正確 FL 代碼
    2. 再以 FL 代碼查詢對應尺寸表，取得構件尺寸
    3. 括號尺寸直接使用，無括號且尺寸表無定義則套用「未特別註明者」

    回傳 { fl_code: [ (type, id, size), ... ] }
    """
    boundaries = build_floor_boundaries(floor_titles)
    sheet_data = defaultdict(list)   # fl_code -> rows
    needs_review = []                # 需人工確認的項目

    # 預建尺寸表查詢結構：fl_code → { type → { id_lower → size } }
    table_lookup = defaultdict(lambda: defaultdict(dict))
    default_sizes = {}   # (fl_code, type) -> default_size

    for t in tables:
        fl = _normalize_floor(t["floor"])
        if not fl:
            continue
        ctype = t["type"] if t["type"] else "構件"
        if t["default_size"]:
            default_sizes[(fl, ctype)] = t["default_size"]
        for row in t["rows"]:
            table_lookup[fl][ctype][row["id"].lower()] = row["size"]

    def _find_size(fl, ltype, label_lower):
        """
        依序嘗試：
        1. 精確 fl
        2. FL_TO_TABLE_FL 合併 fl（如 2FL→2FL~5FL）
        3. table_lookup 中任何包含此 fl 字串的 FL（如 B3FL→B3FL~B2FL）
        回傳 (size, resolved_fl) 或 (None, fl)
        """
        # 1. 精確
        s = table_lookup[fl].get(ltype, {}).get(label_lower)
        if s: return s, fl
        # 2. 已知合併
        tfl = FL_TO_TABLE_FL.get(fl, fl)
        if tfl != fl:
            s = table_lookup[tfl].get(ltype, {}).get(label_lower)
            if s: return s, tfl
        # 3. 動態：找 table_lookup 中含 fl 字串的鍵（如 B3FL→B3FL~B2FL）
        for tfl2, by_type in table_lookup.items():
            if tfl2 != fl and fl in tfl2:
                s = by_type.get(ltype, {}).get(label_lower)
                if s: return s, tfl2
        # 4. 前綴匹配：fl 含 ~ 時拆成子段，找 table_lookup 中以子段開頭的鍵
        #    用途：2FL~6FL 小梁 → 找到 2FL~PRFL 小梁表
        if '~' in fl:
            for part in fl.split('~'):
                if not part: continue
                for tfl2, by_type in table_lookup.items():
                    if tfl2 == fl or tfl2 == tfl: continue
                    if tfl2.startswith(part + '~') or tfl2 == part:
                        s = by_type.get(ltype, {}).get(label_lower)
                        if s: return s, tfl2
        # 5. FL_TO_TABLE_FL 的所有 value 逐一嘗試（捕捉跨越多層的合併尺寸表）
        #    用途：R1FL 小梁 → R1FL 未在 FL_TO_TABLE_FL 小梁表中，但 R2FL→2FL~PRFL
        #    也可捕捉：R1FL 小梁 → 查遍所有已知合併表（2FL~PRFL）
        checked = {fl, tfl}
        for candidate_tfl in set(FL_TO_TABLE_FL.values()):
            if candidate_tfl in checked: continue
            s = table_lookup[candidate_tfl].get(ltype, {}).get(label_lower)
            if s: return s, candidate_tfl
        return None, fl

    # ── 先找圖：依 X+Y 座標取得每個標籤所屬的 FL 代碼，再查表 ──
    for lbl in labels:
        if boundaries:
            fl, is_edge = get_floor_for_y(lbl["y"], boundaries, x=lbl.get("x"))
        else:
            fl, is_edge = "未分類", False

        ltype = lbl["type"]

        # ── 分頁分配：依圖面座標的 FL（結構平面圖 = 一張圖一個分頁）──
        # table_fl 僅用於尺寸表查詢，不影響分頁名稱
        table_fl = FL_TO_TABLE_FL.get(fl, fl)
        sheet_name = FL_DISPLAY_NAME.get(fl, fl)   # 直接用個別樓層 FL

        if lbl["bracket_size"]:
            # 括號尺寸 → 直接使用括號內尺寸，不查表
            entry = (ltype, lbl["label"], lbl["bracket_size"])
            if entry not in sheet_data[sheet_name]:
                sheet_data[sheet_name].append(entry)
            if is_edge:
                needs_review.append(
                    f"⚠️  樓層邊界附近：{lbl['label']} (Y={lbl['y']:.0f})，歸入「{sheet_name}」請確認")
        else:
            # 無括號 → 查表取尺寸（含動態 fallback）
            clean_lower = lbl["clean_label"].lower()
            size, _resolved_fl = _find_size(fl, ltype, clean_lower)
            if size:
                entry = (ltype, lbl["clean_label"], size)
                if entry not in sheet_data[sheet_name]:
                    sheet_data[sheet_name].append(entry)
            else:
                # 查不到 → 套用「未特別註明者」
                ds = default_sizes.get((fl, ltype)) or default_sizes.get((table_fl, ltype))
                if ds:
                    entry = (ltype, lbl["clean_label"], ds)
                    if entry not in sheet_data[sheet_name]:
                        sheet_data[sheet_name].append(entry)
                    needs_review.append(
                        f"📌 套用未特別註明者：{lbl['clean_label']} → {ds} cm（{sheet_name} {ltype}）")
                else:
                    needs_review.append(
                        f"❓ 尺寸表未定義：{lbl['clean_label']} (Y={lbl['y']:.0f})，樓層={fl}，類型={ltype}")

    # 尺寸表補完步驟已移除。
    # 規則：先找圖，再查表——只有結構平面圖中出現的標籤才能進入 Excel。
    # 尺寸表僅用於查詢尺寸，不用於決定構件是否存在。
    # 若尺寸表與平面圖不一致（如 B3FL 無 S1 但尺寸表有 BS1,S1），
    # 以平面圖為準，不自動補入。

    if False:  # SHARED_PLAN_FL 已移除，保留佔位符避免 indentation 錯誤
     for target_fl, source_fl in {}.items():
        target_table_fl = FL_TO_TABLE_FL.get(target_fl, target_fl)
        target_sheet = FL_DISPLAY_NAME.get(target_fl, target_fl)
        source_sheet  = FL_DISPLAY_NAME.get(source_fl,
                        FL_DISPLAY_NAME.get(FL_TO_TABLE_FL.get(source_fl, source_fl),
                        source_fl))

        # 取得來源分頁已偵測到的 label 型別/ID（排除來自 table fallback 的純表格項）
        # 這裡使用 labels 列表重新過濾，找到被分配到 source_fl 的標籤
        src_labels = []
        if boundaries:
            for lbl in labels:
                fl_hit, _ = get_floor_for_y(lbl["y"], boundaries)
                if fl_hit == source_fl:
                    src_labels.append(lbl)

        if not src_labels:
            continue

        # 取得目標分頁已有的 base ID 集合（去括號後 lowercase）
        existing_base = {re.sub(r'\(.*?\)', '', e[1]).strip().lower()
                         for e in sheet_data[target_sheet]}

        for lbl in src_labels:
            ltype = lbl["type"]
            if lbl["bracket_size"]:
                entry = (ltype, lbl["label"], lbl["bracket_size"])
                base = re.sub(r'\(.*?\)', '', lbl["label"]).strip().lower()
                if base not in existing_base:
                    sheet_data[target_sheet].append(entry)
                    existing_base.add(base)
            else:
                clean_lower = lbl["clean_label"].lower()
                if clean_lower in existing_base:
                    continue
                size, _ = _find_size(target_fl, ltype, clean_lower)
                if not size:
                    ds = (default_sizes.get((target_fl, ltype))
                          or default_sizes.get((target_table_fl, ltype)))
                    if ds:
                        size = ds
                        needs_review.append(
                            f"📌 共用圖面套用預設：{lbl['clean_label']} → {ds} cm"
                            f"（{target_sheet} {ltype}，來源 {source_sheet}）")
                if size:
                    entry = (ltype, lbl["clean_label"], size)
                    if entry not in sheet_data[target_sheet]:
                        sheet_data[target_sheet].append(entry)
                        existing_base.add(clean_lower)

    return dict(sheet_data), needs_review


def _normalize_floor(s):
    """將圖面上的樓層名稱標準化，方便作為工作表名稱"""
    if not s:
        return ""
    s = s.strip()
    s = re.sub(r'\s+', '', s)
    # 去掉尾部的「結構平面圖」等
    s = re.sub(r'(結構平面圖|各層|至)', lambda m: {"至":"~"}.get(m.group(), ""), s)
    return s


# ── 寫入 Excel ───────────────────────────────────────────────
def write_excel(sheet_data, output_path):
    wb = Workbook()
    thin   = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill("solid", start_color=HEADER_BG, end_color=HEADER_BG)
    h_font = Font(name="Arial", bold=True, color=HEADER_FG, size=11)

    first = True
    # 工作表依 SHEET_ORDER 排序，未在清單內的排最後
    known = {s: i for i, s in enumerate(SHEET_ORDER)}
    ordered = sorted(sheet_data.keys(),
                     key=lambda k: (known.get(k, 999), k))

    for sheet_name in ordered:
        rows = sheet_data[sheet_name]
        if not rows:
            continue
        ws = wb.active if first else wb.create_sheet()
        ws.title = sheet_name[:31]   # Excel 工作表名稱最多 31 字
        first = False

        # 標題列
        for ci, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
            cell = ws.cell(row=1, column=ci, value=h)
            cell.font      = h_font
            cell.fill      = h_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = border
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 22

        # 資料列依構件類型排序（大梁→小梁→版→柱→鋼梁→鋼柱），去重複
        # 同類型內按自然排序（b1 < b2 < b10，非 b1 < b10 < b11）
        type_rank = {t: i for i, t in enumerate(TYPE_ORDER)}
        rows_sorted = sorted(rows, key=lambda r: (type_rank.get(r[0], 99), _natural_key(r[1])))
        seen = set()
        ri = 2
        for (t, id_, size) in rows_sorted:
            key = (t, id_, size)
            if key in seen:
                continue
            seen.add(key)
            bg = TYPE_COLORS.get(t, "FFFFFF")
            disp_size = re.sub(r'(?<=\d)x(?=\d)', '×', size) if size else size
            for ci, val in enumerate([t, id_, disp_size], 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font      = Font(name="Arial", size=10)
                cell.fill      = PatternFill("solid", start_color=bg, end_color=bg)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border    = border
            ws.row_dimensions[ri].height = 18
            ri += 1

    wb.save(output_path)
    return wb


# ── 差異分析分頁 ──────────────────────────────────────────────
def build_diff_sheet(wb, tables, sheet_data):
    """
    新增「差異分析」工作表，逐樓層比對：
      ❶ 尺寸表有定義、但輸出中不存在 → 可能漏列（紅底）
      ❷ 輸出有、但尺寸表無明確定義（來自括號或預設）→ 標示來源（橘底）
      ❸ 某樓層/類型完全無圖面標籤（僅預設，無法自動補） → 黃底警示
    """
    ws = wb.create_sheet("差異分析")

    HEADER_DIFF = ["分頁", "類型", "編號", "尺寸(表)", "尺寸(輸出)", "狀態"]
    COL_W = [12, 8, 18, 16, 16, 30]
    thin   = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill("solid", start_color=HEADER_BG, end_color=HEADER_BG)
    h_font = Font(name="Arial", bold=True, color=HEADER_FG, size=10)

    # 標題列
    for ci, (h, w) in enumerate(zip(HEADER_DIFF, COL_W), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = h_font; cell.fill = h_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 22

    # 顏色
    RED_FILL    = PatternFill("solid", start_color="FFCCCC", end_color="FFCCCC")  # 表有 輸出無
    ORANGE_FILL = PatternFill("solid", start_color="FFE0B2", end_color="FFE0B2")  # 括號/預設
    YELLOW_FILL = PatternFill("solid", start_color="FFFDE7", end_color="FFFDE7")  # 無圖面標籤
    GREEN_FILL  = PatternFill("solid", start_color="E8F5E9", end_color="E8F5E9")  # 正確對應

    # 建立「尺寸表明確定義」的查詢結構
    # {display_name: {(type, id_lower): size}}
    table_explicit = defaultdict(dict)
    for t in tables:
        fl = _normalize_floor(t.get("floor", ""))
        if not fl: continue
        ctype = t["type"] if t["type"] else "構件"
        # 解析目標分頁（與補完步驟一致）
        if fl in TABLE_FL_TO_FLOORS:
            displays = [FL_DISPLAY_NAME.get(f, f) for f in TABLE_FL_TO_FLOORS[fl]]
        elif fl in FL_DISPLAY_NAME:
            displays = [FL_DISPLAY_NAME[fl]]
        else:
            displays = [fl]
        for display in displays:
            for row in t["rows"]:
                if row["size"] and "未特別" not in row["size"]:
                    table_explicit[display][(ctype, row["id"].lower())] = row["size"]

    # 建立輸出資料查詢
    # {display_name: {(type, base_id_lower): size}}
    output_index = defaultdict(dict)
    for sheet_name, rows in sheet_data.items():
        for (ctype, id_, size) in rows:
            base = re.sub(r'\(.*?\)', '', id_).strip().lower()
            key = (ctype, base)
            if key not in output_index[sheet_name]:
                output_index[sheet_name][key] = (id_, size)

    ri = 2
    all_sheets = sorted(set(list(table_explicit.keys()) + list(output_index.keys())),
                        key=lambda k: (SHEET_ORDER.index(k) if k in SHEET_ORDER else 999, k))

    for sname in all_sheets:
        tbl = table_explicit.get(sname, {})
        out = output_index.get(sname, {})
        all_keys = sorted(set(list(tbl.keys()) + list(out.keys())),
                          key=lambda k: (TYPE_ORDER.index(k[0]) if k[0] in TYPE_ORDER else 99,
                                         _natural_key(k[1])))

        for key in all_keys:
            ctype, id_lower = key
            tbl_size = tbl.get(key)
            out_val  = out.get(key)
            out_id   = out_val[0] if out_val else None
            out_size = out_val[1] if out_val else None

            if tbl_size and not out_val:
                status = "❌ 尺寸表有定義，輸出中缺漏"
                fill = RED_FILL
            elif out_val and not tbl_size:
                if out_id and '(' in out_id:
                    status = "🔷 括號尺寸（圖面標註）"
                else:
                    status = "📌 套用未特別註明者預設"
                fill = ORANGE_FILL
            elif tbl_size and out_val:
                tbl_disp = re.sub(r'(?<=\d)x(?=\d)', '×', tbl_size)
                out_disp = re.sub(r'(?<=\d)x(?=\d)', '×', out_size) if out_size else ''
                if tbl_disp == out_disp:
                    status = "✅ 正確對應"
                else:
                    status = f"⚠️ 尺寸不符（表={tbl_disp}）"
                fill = GREEN_FILL
            else:
                continue

            disp_tbl = re.sub(r'(?<=\d)x(?=\d)', '×', tbl_size) if tbl_size else "—"
            disp_out = re.sub(r'(?<=\d)x(?=\d)', '×', out_size) if out_size else "—"
            vals = [sname, ctype, out_id or id_lower, disp_tbl, disp_out, status]
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
            ws.row_dimensions[ri].height = 16
            ri += 1

        # 檢查是否有類型完全無輸出（尺寸表只有預設，無圖面標籤）
        for ctype in ['大梁', '小梁', '版', '柱']:
            tbl_has = any(k[0] == ctype for k in tbl)
            out_has = any(k[0] == ctype for k in out)
            if not tbl_has and not out_has: continue  # 不在本表
            # 若輸出無此類型但尺寸表也無明確定義 → 警示
            if not out_has and not tbl_has:
                vals = [sname, ctype, "—", "（僅預設）", "—",
                        "⚠️ 尺寸表僅預設值，圖面無標籤，需人工確認"]
                for ci, val in enumerate(vals, 1):
                    cell = ws.cell(row=ri, column=ci, value=val)
                    cell.font = Font(name="Arial", size=9)
                    cell.fill = YELLOW_FILL
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border
                ws.row_dimensions[ri].height = 16
                ri += 1

    # 凍結標題列
    ws.freeze_panes = "A2"

    # 圖例（底部）
    ri += 1
    legends = [
        ("❌ 紅色", "尺寸表有明確定義，但輸出中缺漏（可能是 block attribute 未捕捉）"),
        ("🔷 橘色", "輸出有此項，但尺寸表無明確定義（括號尺寸或套用預設）"),
        ("⚠️ 黃色", "尺寸表僅預設值，且圖面無文字標籤（6FL-7FL大梁等）需人工補入"),
        ("✅ 綠色", "尺寸表與輸出完全對應"),
    ]
    for leg_key, leg_desc in legends:
        ws.cell(row=ri, column=1, value=leg_key).font = Font(name="Arial", size=9, bold=True)
        ws.cell(row=ri, column=2, value=leg_desc).font = Font(name="Arial", size=9)
        ws.merge_cells(start_row=ri, start_column=2, end_row=ri, end_column=6)
        ri += 1


# ── 摘要報告 ─────────────────────────────────────────────────
def print_summary(sheet_data, needs_review, output_path):
    print(f"\n✅ Excel 已輸出：{output_path}\n")
    print("📊 各工作表筆數：")
    for sheet, rows in sheet_data.items():
        counts = defaultdict(int)
        for (t, _, _) in rows:
            counts[t] += 1
        detail = "、".join(f"{t}{n}筆" for t, n in counts.items())
        print(f"   {sheet}：{detail}")

    if needs_review:
        print("\n🔍 需人工確認的項目：")
        for item in needs_review:
            print(f"   {item}")
    else:
        print("\n✅ 無需特別人工確認的項目。")


# ── 入口 ─────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dim-tables", required=True)
    parser.add_argument("--labels-xy",  required=True)
    parser.add_argument("--output",     required=True)
    args = parser.parse_args()

    print("🔍 解析尺寸表...")
    tables = parse_dim_tables(args.dim_tables)
    print(f"   找到 {len(tables)} 張尺寸表")

    print("🔍 解析標籤座標...")
    floor_titles, labels = parse_labels_xy(args.labels_xy)
    print(f"   找到 {len(floor_titles)} 個樓層標題、{len(labels)} 個構件標籤")

    print("🔧 整合資料...")
    sheet_data, needs_review = build_sheet_data(tables, floor_titles, labels)

    print("📝 寫入 Excel...")
    wb = write_excel(sheet_data, args.output)

    print("📊 產生差異分析分頁...")
    build_diff_sheet(wb, tables, sheet_data)
    wb.save(args.output)

    print_summary(sheet_data, needs_review, args.output)


if __name__ == "__main__":
    main()
