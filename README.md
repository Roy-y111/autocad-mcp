# AutoCAD MCP Server

讓 Claude 等 MCP 用戶端透過 Windows COM 介面直接操作執行中的 AutoCAD：開關圖檔、繪製圖元、管理圖層、下指令。另附一支讀取結構平面圖梁／版標籤與尺寸表的腳本。

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
| 進階 | `run_command`（在 AutoCAD 命令列執行任意指令） |

## 輔助腳本

`read_beams_slabs.py`：連線到目前開啟的結構平面圖，依樓層區間讀出梁、版標籤與 AcDbTable 尺寸表，輸出 `beams_slabs_report.txt`（與腳本同一資料夾）。樓層的 Y 座標區間寫在腳本開頭的 `FLOORS`，請依你的圖面調整。

## ⚠️ 安全注意事項

此伺服器等於把你的 AutoCAD 與本機檔案存取權交給 AI，使用前請了解：

- **`run_command` 可執行任意 AutoCAD 指令**，包含載入 LISP、執行腳本等，能力近似於在你電腦上執行程式。
- **`open_drawing`、`save_drawing` 接受任意路徑**，可讀取或覆寫你有權限的任何位置。
- **間接提示注入（prompt injection）**：DWG 內的文字、圖層名稱或屬性可能被 AI 當成指令。**請只開啟你信任來源的圖檔。**
- 建議在執行高風險操作前保持 MCP 用戶端的「逐項核准」設定，並先備份重要圖檔。
- 本伺服器僅在本機以 stdio 執行，不開放網路連線；請勿自行包裝成對外網路服務。

## 授權

尚未指定授權條款。在加入 LICENSE 檔之前，依預設保留所有權利。
