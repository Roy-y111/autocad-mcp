# -*- coding: utf-8 -*-
"""
從目前 AutoCAD 已開啟的門窗圖(如 A6)抓取門窗型號主檔：
編號 + 寬x高尺寸 + 防火時效(若有)。

用法：先用 AutoCAD MCP 的 open_drawing 開啟門窗圖，再執行：
    python extract_master.py --out master_dims.json

抓取邏輯：
- _LCNo001 / _LDNO001 / _LOPN001 圖塊的 NO 或 U 屬性 = 編號
- D 屬性 = 防火時效(door 用)
- 該圖塊插入點附近最近的「數字x數字」格式文字 = 寬x高（cm）
"""
import sys, io, re, math, json, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import win32com.client


def fix_text(raw):
    if raw is None:
        return ""
    for enc in ('big5', 'gbk', 'cp950'):
        try:
            return raw.encode('latin-1').decode(enc)
        except Exception:
            pass
    return raw


def category_of(no):
    if no.startswith("SD"):
        return "防火門"
    if no.startswith("DW"):
        return "落地窗"
    if no.startswith("WB"):
        return "小窗/通風窗"
    if no.startswith("D"):
        return "門"
    if no.startswith("W"):
        return "窗"
    return "其他"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="master_dims.json", help="輸出 JSON 路徑")
    ap.add_argument("--max-dist", type=float, default=400, help="編號圖塊與尺寸文字的最大配對距離")
    ap.add_argument("--block-names", default="_LCNo001,_LDNO001,_LOPN001",
                     help="門窗編號圖塊名稱清單(逗號分隔)，不同專案的圖塊命名可能不同")
    args = ap.parse_args()
    block_names = tuple(b.strip() for b in args.block_names.split(","))

    acad = win32com.client.GetActiveObject('AutoCAD.Application')
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print(f"Doc: {doc.Name}, entities: {ms.Count}")

    SIZE_RE = re.compile(r'^\d+[xXＸ×]\d+$')
    no_blocks = []   # (no, x, y, fire)
    size_texts = []  # (text, x, y)

    for e in ms:
        try:
            name = e.ObjectName
        except Exception:
            continue

        if name == 'AcDbBlockReference':
            try:
                bname = e.EffectiveName if hasattr(e, 'EffectiveName') else e.Name
            except Exception:
                bname = "?"
            if bname in block_names:
                pt = e.InsertionPoint
                no_val, fire = "", ""
                try:
                    if e.HasAttributes:
                        for att in e.GetAttributes():
                            tag = fix_text(att.TagString).upper()
                            val = fix_text(att.TextString)
                            if tag in ('NO', 'U'):
                                no_val = val.strip()
                            elif tag == 'D':
                                fire = val.strip()
                except Exception as ex:
                    no_val = f"ERR:{ex}"
                if no_val:
                    no_blocks.append((no_val, pt[0], pt[1], fire))

        elif name in ('AcDbText', 'AcDbMText'):
            try:
                raw = e.TextString
                txt = fix_text(raw).strip()
                txt = re.sub(r'\\[A-Za-z][^;]*;', '', txt)
                txt = re.sub(r'\{[^}]*\}', '', txt)
                txt = txt.strip().replace(' ', '')
                if SIZE_RE.match(txt):
                    pt = e.InsertionPoint
                    size_texts.append((txt, pt[0], pt[1]))
            except Exception:
                pass

    print(f"NO blocks: {len(no_blocks)}, size texts: {len(size_texts)}")

    def nearest_size(x, y):
        best, bd = None, args.max_dist
        for txt, tx, ty in size_texts:
            d = math.hypot(tx - x, ty - y)
            if d < bd:
                bd, best = d, txt
        return best

    master = {}
    unmatched = []
    for no_val, x, y, fire in no_blocks:
        size_str = nearest_size(x, y)
        if not size_str:
            unmatched.append(no_val)
            continue
        w, h = size_str.replace('X', 'x').replace('Ｘ', 'x').replace('×', 'x').split('x')
        entry = {"category": category_of(no_val), "w": int(w), "h": int(h), "fire": fire}
        if no_val in master and master[no_val] != entry:
            print(f"⚠️ 編號重複但尺寸不同: {no_val}  舊={master[no_val]}  新={entry}（採用後者，請至CAD確認）")
        master[no_val] = entry

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=1)

    print(f"Saved: {args.out}  (共 {len(master)} 個型號)")
    if unmatched:
        print(f"⚠️ 有 {len(unmatched)} 個編號找不到附近尺寸文字（距離 > {args.max_dist}）：{unmatched}")


if __name__ == "__main__":
    main()
