"""AutoCAD MCP 伺服器主體"""
import json
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

from . import tools

app = Server("autocad-mcp")


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="acad_status",
            description="取得 AutoCAD 目前狀態（版本、文件名稱、儲存狀態）",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="new_drawing",
            description="在 AutoCAD 中開啟一個新的空白圖檔",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="open_drawing",
            description="開啟指定路徑的 AutoCAD DWG 檔案",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "DWG 檔案的完整路徑"},
                },
                "required": ["file_path"],
            },
        ),
        types.Tool(
            name="save_drawing",
            description="儲存目前圖檔，可選擇性指定另存路徑",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "另存路徑（留空則直接儲存）"},
                },
                "required": [],
            },
        ),
        types.Tool(
            name="draw_line",
            description="在 AutoCAD 模型空間繪製一條直線",
            inputSchema={
                "type": "object",
                "properties": {
                    "x1": {"type": "number"}, "y1": {"type": "number"},
                    "x2": {"type": "number"}, "y2": {"type": "number"},
                },
                "required": ["x1", "y1", "x2", "y2"],
            },
        ),
        types.Tool(
            name="draw_circle",
            description="在 AutoCAD 模型空間繪製一個圓",
            inputSchema={
                "type": "object",
                "properties": {
                    "cx": {"type": "number", "description": "圓心 X"},
                    "cy": {"type": "number", "description": "圓心 Y"},
                    "radius": {"type": "number", "description": "半徑"},
                },
                "required": ["cx", "cy", "radius"],
            },
        ),
        types.Tool(
            name="draw_rectangle",
            description="在 AutoCAD 模型空間繪製一個矩形",
            inputSchema={
                "type": "object",
                "properties": {
                    "x1": {"type": "number"}, "y1": {"type": "number"},
                    "x2": {"type": "number"}, "y2": {"type": "number"},
                },
                "required": ["x1", "y1", "x2", "y2"],
            },
        ),
        types.Tool(
            name="draw_polyline",
            description="繪製多段線，傳入頂點座標陣列",
            inputSchema={
                "type": "object",
                "properties": {
                    "points": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "number"}, "minItems": 2, "maxItems": 2},
                        "description": "[[x1,y1],[x2,y2],...] 格式的頂點列表",
                    },
                },
                "required": ["points"],
            },
        ),
        types.Tool(
            name="add_text",
            description="在指定座標新增單行文字",
            inputSchema={
                "type": "object",
                "properties": {
                    "x": {"type": "number"}, "y": {"type": "number"},
                    "text": {"type": "string", "description": "文字內容"},
                    "height": {"type": "number", "description": "文字高度（預設 2.5）"},
                },
                "required": ["x", "y", "text"],
            },
        ),
        types.Tool(
            name="add_mtext",
            description="在指定座標新增多行文字（MText）",
            inputSchema={
                "type": "object",
                "properties": {
                    "x": {"type": "number"}, "y": {"type": "number"},
                    "text": {"type": "string"},
                    "width": {"type": "number", "description": "文字方塊寬度（預設 100）"},
                    "height": {"type": "number", "description": "文字高度（預設 2.5）"},
                },
                "required": ["x", "y", "text"],
            },
        ),
        types.Tool(
            name="run_command",
            description="在 AutoCAD 命令列執行任意指令（如 ZOOM、REGEN、LAYER 等）",
            inputSchema={
                "type": "object",
                "properties": {
                    "command": {"type": "string", "description": "AutoCAD 指令字串"},
                },
                "required": ["command"],
            },
        ),
        types.Tool(
            name="zoom_extents",
            description="縮放視圖以顯示所有物件（ZOOM E）",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="list_layers",
            description="列出目前圖檔中所有圖層名稱",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="set_layer",
            description="切換至指定圖層（不存在則自動建立）",
            inputSchema={
                "type": "object",
                "properties": {
                    "layer_name": {"type": "string", "description": "圖層名稱"},
                },
                "required": ["layer_name"],
            },
        ),
        types.Tool(
            name="rename_layer",
            description="重新命名圖層（例如將中文圖層名稱改為英文，避免 Dynamo 等節點不支援中文圖層名稱）",
            inputSchema={
                "type": "object",
                "properties": {
                    "old_name": {"type": "string", "description": "目前的圖層名稱"},
                    "new_name": {"type": "string", "description": "新的圖層名稱"},
                },
                "required": ["old_name", "new_name"],
            },
        ),
        types.Tool(
            name="set_color",
            description="設定目前繪圖顏色（ACI 色碼 1-255，1=紅 2=黃 3=綠 5=藍 7=白）",
            inputSchema={
                "type": "object",
                "properties": {
                    "color_index": {"type": "integer", "minimum": 1, "maximum": 255},
                },
                "required": ["color_index"],
            },
        ),
        types.Tool(
            name="count_entities",
            description="統計模型空間中各類型物件的數量",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    try:
        match name:
            case "acad_status":
                result = tools.acad_status()
            case "new_drawing":
                result = tools.new_drawing()
            case "open_drawing":
                result = tools.open_drawing(arguments["file_path"])
            case "save_drawing":
                result = tools.save_drawing(arguments.get("file_path"))
            case "draw_line":
                result = tools.draw_line(
                    arguments["x1"], arguments["y1"],
                    arguments["x2"], arguments["y2"],
                )
            case "draw_circle":
                result = tools.draw_circle(
                    arguments["cx"], arguments["cy"], arguments["radius"]
                )
            case "draw_rectangle":
                result = tools.draw_rectangle(
                    arguments["x1"], arguments["y1"],
                    arguments["x2"], arguments["y2"],
                )
            case "draw_polyline":
                result = tools.draw_polyline(arguments["points"])
            case "add_text":
                result = tools.add_text(
                    arguments["x"], arguments["y"],
                    arguments["text"],
                    arguments.get("height", 2.5),
                )
            case "add_mtext":
                result = tools.add_mtext(
                    arguments["x"], arguments["y"],
                    arguments["text"],
                    arguments.get("width", 100.0),
                    arguments.get("height", 2.5),
                )
            case "run_command":
                result = tools.run_command(arguments["command"])
            case "zoom_extents":
                result = tools.zoom_extents()
            case "list_layers":
                result = tools.list_layers()
            case "set_layer":
                result = tools.set_layer(arguments["layer_name"])
            case "rename_layer":
                result = tools.rename_layer(arguments["old_name"], arguments["new_name"])
            case "set_color":
                result = tools.set_color(arguments["color_index"])
            case "count_entities":
                result = tools.count_entities()
            case _:
                result = f"未知工具：{name}"

        if isinstance(result, (dict, list)):
            text = json.dumps(result, ensure_ascii=False, indent=2)
        else:
            text = str(result)

    except Exception as e:
        text = f"錯誤：{e}"

    return [types.TextContent(type="text", text=text)]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
