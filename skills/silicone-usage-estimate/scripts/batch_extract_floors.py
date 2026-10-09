# -*- coding: utf-8 -*-
"""
批次開啟多個樓層平面圖 .dwg，統計每個樓層各門窗編號出現次數。
會用 win32com 直連目前執行中的 AutoCAD，逐一開啟→掃描→關閉。

用法：
    python batch_extract_floors.py --dir "D:/專案/CAD底圖(建築)" \
        --pattern "{floor}_建築平面圖.dwg" \
        --floors B5F,B4F,B3F,B2F,B1F,1F,2F,3F,...,R2F \
        --out floor_counts.json

--pattern 用 {floor} 佔位符表示檔名中樓層代碼的位置。
"""
import sys, io, json, argparse
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True, help="樓層平面圖所在資料夾")
    ap.add_argument("--pattern", default="{floor}_建築平面圖.dwg", help="檔名樣板，{floor} 為樓層代碼佔位符")
    ap.add_argument("--floors", required=True, help="樓層代碼清單，逗號分隔，依你想要的 Excel 分頁順序排列")
    ap.add_argument("--out", default="floor_counts.json")
    ap.add_argument("--block-names", default="_LCNo001,_LDNO001,_LOPN001")
    args = ap.parse_args()

    floor_list = [f.strip() for f in args.floors.split(",") if f.strip()]
    block_names = tuple(b.strip() for b in args.block_names.split(","))

    acad = win32com.client.GetActiveObject('AutoCAD.Application')
    acad.Visible = True

    results = {}

    for floor in floor_list:
        path = f"{args.dir.rstrip(chr(92)).rstrip('/')}\\{args.pattern.format(floor=floor)}"
        print(f"=== {floor}: {path}")

        doc = None
        for d in acad.Documents:
            try:
                if d.FullName.lower() == path.lower():
                    doc = d
                    break
            except Exception:
                pass
        already_open = doc is not None
        if doc is None:
            try:
                doc = acad.Documents.Open(path)
            except Exception as e:
                print(f"  FAILED to open: {e}")
                results[floor] = {"error": str(e)}
                continue

        ms = doc.ModelSpace
        counts = {}
        try:
            for e in ms:
                try:
                    name = e.ObjectName
                except Exception:
                    continue
                if name != 'AcDbBlockReference':
                    continue
                try:
                    bname = e.EffectiveName if hasattr(e, 'EffectiveName') else e.Name
                except Exception:
                    bname = "?"
                if bname not in block_names:
                    continue
                no_val = ""
                try:
                    if e.HasAttributes:
                        for att in e.GetAttributes():
                            tag = fix_text(att.TagString).upper()
                            val = fix_text(att.TextString)
                            if tag in ('NO', 'U'):
                                no_val = val.strip().replace(" ", "")
                except Exception as ex:
                    no_val = f"ERR:{ex}"
                if not no_val:
                    continue
                counts[no_val] = counts.get(no_val, 0) + 1
        except Exception as e:
            print(f"  scan error: {e}")

        results[floor] = counts
        print(f"  {sum(counts.values())} tagged items, {len(counts)} distinct codes")

        if not already_open:
            try:
                doc.Close(False)
            except Exception as e:
                print(f"  close error: {e}")

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print("Saved:", args.out)


if __name__ == "__main__":
    main()
