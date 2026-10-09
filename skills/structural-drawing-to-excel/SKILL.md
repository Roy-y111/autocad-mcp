---
name: structural-drawing-to-excel
description: >
  從已開啟的 AutoCAD 結構平面圖（.dwg）自動讀取所有尺寸表，整理成各樓層分頁的
  Excel 族群清單，供 Revit/Dynamo 建立族群類型使用。
  此 skill 需要 AutoCAD MCP 連線才能運作，僅在 CAD 工作情境下使用。
  當使用者說「讀取結構圖尺寸表」、「輸出結構 Excel」、「產生 Revit 族群清單」、
  「從 CAD 讀尺寸表」、「建立 Revit 族群」、「結構圖轉 Excel」、「產出結構尺寸表」，
  且 AutoCAD MCP 已連線時，使用此 skill。
---

# 結構平面圖 → Excel 族群清單

從已開啟的 AutoCAD 結構平面圖中，自動提取大梁、小梁、版、柱的尺寸表，
輸出**依結構平面圖分頁**的 Excel（一張圖一個分頁），每列為「單一編號對應單一尺寸」，
適合 Dynamo 直接讀取。末頁附「差異分析」對照尺寸表與圖面差異。

---

## ⚠️ 前置注意事項

1. **清圖：刪除上視梁**（必做）：圖面中若有「上視梁」（俯視可見的下層梁線及文字），
   其標籤圖層與正常梁相同，LISP 無法自動區分，會混入 Excel 造成誤判。
   **執行 LISP 前，請先手動選取並刪除這些上視梁標籤。**
2. **圖塊問題**：若某樓層平面圖被做成圖塊（Block），LISP 無法讀取其內部文字。
   請先在 AutoCAD 對該圖塊執行 `EXPLODE`（分解）後再執行本 skill。
3. **openpyxl**：確認已安裝 → `python -c "import openpyxl"`
4. **使用者名稱**：以下 LISP 中的 `{USER}` 替換為實際 Windows 使用者名稱（執行 `mcp__autocad__acad_status` 取得）。

---

## Step 1：同時提取尺寸表 + 標籤（單一 LISP，一次完成）

```lisp
(setq f (open "C:/Users/{USER}/Desktop/cad_dim_tables.txt" "w"))
(setq doc (vla-get-ActiveDocument (vlax-get-acad-object)) ms (vla-get-ModelSpace doc) tcount 0)
(vlax-for obj ms
  (if (= (vla-get-ObjectName obj) "AcDbTable")
    (progn (setq tcount (1+ tcount))
      (write-line (strcat "=== TABLE " (itoa tcount) " Layer:" (vla-get-Layer obj) " ===") f)
      (setq rows (vla-get-Rows obj) cols (vla-get-Columns obj) r 0)
      (while (< r rows) (setq c 0 rowstr "")
        (while (< c cols) (setq rowstr (strcat rowstr "[" (vla-GetText obj r c) "]")) (setq c (1+ c)))
        (write-line rowstr f) (setq r (1+ r))))))
(close f)
(setq f (open "C:/Users/{USER}/Desktop/cad_labels_xy.txt" "w"))
(setq ms (vla-get-ModelSpace doc))
(vlax-for obj ms
  (if (or (= (vla-get-ObjectName obj) "AcDbText") (= (vla-get-ObjectName obj) "AcDbMText"))
    (progn (setq pt (vla-get-InsertionPoint obj))
      (setq x (car (vlax-safearray->list (vlax-variant-value pt))))
      (setq y (cadr (vlax-safearray->list (vlax-variant-value pt))))
      (write-line (strcat (vla-get-Layer obj) "||" (rtos x 2 0) "," (rtos y 2 0) "||" (vla-get-TextString obj)) f))))
(close f)
(princ (strcat "Tables:" (itoa tcount) " done"))
```

> ⚠️ 若 `Tables:0`，表示尺寸表非 AcDbTable（純線條繪製），無法自動讀取，告知使用者。

---

## Step 2：執行 generate_excel.py

腳本位於 `CL_CAD/skills/structural-drawing-to-excel/scripts/generate_excel.py`

```
python -c "
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.argv = ['generate_excel.py',
  '--dim-tables', 'C:/Users/{USER}/Desktop/cad_dim_tables.txt',
  '--labels-xy',  'C:/Users/{USER}/Desktop/cad_labels_xy.txt',
  '--output',     '{輸出路徑}/{圖檔名稱}_結構構件尺寸表.xlsx']
exec(open('C:/Users/{USER}/Desktop/CL_CAD/skills/structural-drawing-to-excel/scripts/generate_excel.py', encoding='utf-8').read())
"
```

---

## Step 3：清理暫存檔

```powershell
Remove-Item "C:\Users\{USER}\Desktop\cad_dim_tables.txt","C:\Users\{USER}\Desktop\cad_labels_xy.txt" -ErrorAction SilentlyContinue
```

---

## 腳本核心邏輯（generate_excel.py）

### 先找圖！再查表！

| 步驟 | 說明 |
|------|------|
| **① 偵測樓層邊界** | 精確比對個別樓層標題（如「六層結構平面圖」），用 Y 座標切分邊界。排除「六至七層結構平面圖」等合訂標題（避免誤判）。 |
| **② 標籤分樓層** | 每個構件標籤依 Y 座標歸入對應樓層（一張圖一個分頁，6FL、7FL 分開）。 |
| **③ 查詢尺寸** | 以樓層 FL 代碼 + 動態 fallback 查詢尺寸表（如 6FL 找 6FL~7FL 表）。括號尺寸直接使用。 |
| **④ 尺寸表補完** | 尺寸表中有定義但 LISP 未捕捉的項目（如 block attribute、CS1a）自動補入。 |
| **⑤ 差異分析分頁** | 最後一個分頁列出：❌ 表有圖無（紅）、🔷 圖有表無（橘）、⚠️ 僅預設無標籤（黃）、✅ 完全對應（綠）。 |

### 分頁規則

- **一張結構平面圖 = 一個 Excel 分頁**
- 若多層共用同一尺寸表（如 6FL~7FL 表），各層仍分開（6FL 一頁、7FL 一頁），查表時自動對應合併表
- 分頁順序：B3FL → B2FL → B1FL → 1FL → 2FL → 3FL-5FL → 6FL → 7FL → R1FL → R2FL → PRFL → 差異分析

### 版面配置自動偵測（支援水平並排）

| 版面類型 | 說明 | 邊界模式 |
|----------|------|----------|
| 垂直堆疊 | 各樓層標題 Y 座標不同（如大型基地案） | Y 中點切分 |
| **水平並排** | 同一行有多個樓層標題（如小基地案，B3FL+B2FL 並排） | **X+Y 二維切分** |
| 混合 | 部分垂直、部分水平 | 自動偵測，分段處理 |

判斷條件：同一 Y（差距 ≤ 200 單位）出現多個標題 → 自動切換 X 中點模式。

### 編碼處理

- Big5 特殊字元：`0xA1 0x42`（頓號分隔符）→ `,`；`0xA1 0xD1`（乘號）→ `x`（輸出時還原為 `×`）
- AutoCAD 格式碼：`%%U`（底線）、`\A1;`（對齊）、`{...}` 自動剝除
- 排序：自然排序（b1, b2...b9, b10, b11）

---

## 回報格式

向使用者說明：
1. Excel 輸出路徑
2. 各分頁（樓層）筆數摘要（大梁 N 筆、小梁 N 筆…）
3. ⚠️ 需人工確認的項目（邊界附近、套用預設值等）
4. 提醒查看「差異分析」分頁確認完整性

---

## 常見問題與排查

| 問題 | 原因 | 處理 |
|------|------|------|
| 某樓層標籤全部找不到 | 平面圖被做成圖塊（Block） | 在 AutoCAD 執行 `EXPLODE` 分解後重跑 |
| `Tables found: 0` | 尺寸表為純線條繪製，非 AcDbTable | 告知使用者，無法自動讀取 |
| 某項目出現在錯誤樓層 | 標籤 Y 座標靠近樓層邊界（±500 單位內） | 查看 ⚠️ 邊界警示，手動確認歸屬 |
| 項目在差異分析標記為 ❌ | 可能是 block attribute（如 CS1a）LISP 無法讀取 | 尺寸表補完步驟會自動補入，確認無誤即可 |
| 尺寸顯示 `x` 而非 `×` | 大檔案中間有異常 | 正常情況下輸出已自動還原為 `×`，若未還原請回報 |
| openpyxl 未安裝 | 缺少套件 | `pip install openpyxl` |
