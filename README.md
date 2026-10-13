# AutoCAD MCP Server
**繁體中文** | [English](README.en.md)

讓 Claude 等 MCP 用戶端透過 Windows COM 介面直接操作執行中的 AutoCAD：開關圖檔、繪製圖元、管理圖層、下指令。另附一支讀取結構平面圖梁／版標籤與尺寸表的腳本，以及三個 Claude skill（結構尺寸表轉 Excel、DWG 圖面比較、矽利康用量估算）。

## 系統需求

- **Windows**（使用 `pywin32` 的 COM 自動化，無法在 macOS / Linux 執行）
- 已安裝 **AutoCAD**（桌面版，需先啟動）
- Python 3.10 以上（程式碼使用 `match` 語法）

## 安裝

```bash
git clone https://github.com/Roy-y111/autocad-mcp.git
cd autocad-mcp
pip install -r requirements.txt
```

## 設定 MCP 用戶端

複製 `claude_mcp_config.example.json`，把 `<專案資料夾完整路徑>` 換成你本機 `main.py` 所在的資料夾，再貼到 MCP 用戶端的設定檔（例如 Claude Desktop 的 `claude_desktop_config.json`）：

```json
{
  "mcpServers": {
    "autocad": {
      "command": "python",
      "args": ["C:\\path\\to\\autocad-mcp\\main.py"],
      "env": {}
    }
  }
}
```

注意：JSON 內的反斜線要寫成 `\\`。設定完成後重新啟動 MCP 用戶端，並確認 AutoCAD 已開啟。

## 提供的工具

| 類別 | 工具 |
|---|---|
| 狀態與檔案 | `acad_status`、`new_drawing`、`open_drawing`、`save_drawing` |
| 繪圖 | `draw_line`、`draw_circle`、`draw_rectangle`、`draw_polyline`、`add_text`、`add_mtext` |
| 圖層與顏色 | `list_layers`、`set_layer`、`rename_layer`、`set_color` |
| 檢視與統計 | `zoom_extents`、`count_entities` |
| 進階（**預設停用**） | `run_command`（在 AutoCAD 命令列執行任意指令，需手動啟用，見下方） |

## 輔助腳本

`read_beams_slabs.py`：連線到目前開啟的結構平面圖，依樓層區間讀出梁、版標籤與 AcDbTable 尺寸表，輸出 `beams_slabs_report.txt`（與腳本同一資料夾）。樓層的 Y 座標區間寫在腳本開頭的 `FLOORS`，請依你的圖面調整。

## Claude Skills

`skills/` 內附三個搭配 AutoCAD 使用的 Claude skill，各自附 `SKILL.md`（觸發條件與流程）與 `scripts/`：

| Skill | 用途 | 需要 AutoCAD MCP |
|---|---|---|
| [structural-drawing-to-excel](skills/structural-drawing-to-excel/SKILL.md) | 讀取結構平面圖的尺寸表，輸出各樓層分頁的 Excel 構件清單（大梁、小梁、版、柱），供 Revit / Dynamo 建立族群類型使用 | 是 |
| [compare-dwg-revisions](skills/compare-dwg-revisions/SKILL.md) | 比較新舊兩版 DWG，忽略僅位移的內容，整理成修改清單 Excel，並產生帶編號標註的 AutoCAD 比較圖 | 否（使用 AutoCAD Core Console） |
| [silicone-usage-estimate](skills/silicone-usage-estimate/SKILL.md) | 由門窗圖取得各型號尺寸，統計各樓層門窗數量，估算門窗外框與玻璃的矽利康用量 | 是（或 win32com 直連） |

使用方式：把要用的 skill 資料夾複製到 `~/.claude/skills/`（或專案的 `.claude/skills/`），重新啟動 Claude 後，說出 `SKILL.md` 裡列的觸發語句即可。

注意事項：
- 腳本會讀寫中間檔，預設放在 `<專案根目錄>/outputs/`（`silicone-usage-estimate` 的 `.lsp` 可用環境變數 `SILICONE_OUT_DIR` 改位置，詳見該 skill 的 `SKILL.md`）。
- `structural-drawing-to-excel` 與 `silicone-usage-estimate` 的範例指令含 `{USER}` 或示意路徑，需換成你自己的路徑。
- `.lsp` 與 PowerShell 腳本僅在作者環境測試過，AutoCAD 版本不同時請先在測試圖上試跑。

## ⚠️ 安全注意事項

此伺服器等於把你的 AutoCAD 與本機檔案存取權交給 AI，使用前請了解：

- **`run_command` 可執行任意 AutoCAD 指令**，包含載入 LISP、執行腳本等，能力近似於在你電腦上執行程式。**此工具預設停用**，不會出現在工具清單中。確定需要時，在 MCP 設定的 `env` 加上 `"AUTOCAD_MCP_ENABLE_RUN_COMMAND": "1"` 並重新啟動：

  ```json
  "env": { "AUTOCAD_MCP_ENABLE_RUN_COMMAND": "1" }
  ```
- **`open_drawing`、`save_drawing` 接受任意路徑**，可讀取或覆寫你有權限的任何位置。
- **間接提示注入（prompt injection）**：DWG 內的文字、圖層名稱或屬性可能被 AI 當成指令。**請只開啟你信任來源的圖檔。**
- 建議在執行高風險操作前保持 MCP 用戶端的「逐項核准」設定，並先備份重要圖檔。
- 本伺服器僅在本機以 stdio 執行，不開放網路連線；請勿自行包裝成對外網路服務。

## 授權

[MIT License](LICENSE)
