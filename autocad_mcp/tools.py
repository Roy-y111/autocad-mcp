"""AutoCAD COM 操作工具集"""
import win32com.client
from typing import Optional


def get_acad():
    """取得或啟動 AutoCAD 應用程式實例"""
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True
    return acad


def get_model_space():
    """取得目前文件的模型空間"""
    acad = get_acad()
    doc = acad.ActiveDocument
    return doc.ModelSpace, doc


def make_point(x: float, y: float, z: float = 0.0):
    """建立 AutoCAD 點陣列"""
    import win32com.client as wc
    pt = wc.VARIANT(
        wc.pythoncom.VT_ARRAY | wc.pythoncom.VT_R8, (x, y, z)
    )
    return pt


# ---------- 工具函式 ----------

def acad_status() -> dict:
    """回傳 AutoCAD 狀態資訊"""
    acad = get_acad()
    doc = acad.ActiveDocument
    return {
        "version": acad.Version,
        "document": doc.Name,
        "saved": doc.Saved,
        "full_path": doc.FullName,
    }


def new_drawing() -> str:
    """開啟一個新的空白圖檔"""
    acad = get_acad()
    acad.Documents.Add()
    return "已開啟新圖檔"


def open_drawing(file_path: str) -> str:
    """開啟指定路徑的 DWG 檔案"""
    acad = get_acad()
    acad.Documents.Open(file_path)
    return f"已開啟：{file_path}"


def save_drawing(file_path: Optional[str] = None) -> str:
    """儲存目前圖檔，可指定另存路徑"""
    _, doc = get_model_space()
    if file_path:
        doc.SaveAs(file_path)
        return f"已另存為：{file_path}"
    else:
        doc.Save()
        return "已儲存"


def draw_line(x1: float, y1: float, x2: float, y2: float) -> str:
    """繪製直線"""
    ms, _ = get_model_space()
    start = make_point(x1, y1)
    end = make_point(x2, y2)
    ms.AddLine(start, end)
    return f"已繪製直線：({x1},{y1}) → ({x2},{y2})"


def draw_circle(cx: float, cy: float, radius: float) -> str:
    """繪製圓形"""
    ms, _ = get_model_space()
    center = make_point(cx, cy)
    ms.AddCircle(center, radius)
    return f"已繪製圓：圓心({cx},{cy}) 半徑{radius}"


def draw_rectangle(x1: float, y1: float, x2: float, y2: float) -> str:
    """繪製矩形（四條直線）"""
    ms, _ = get_model_space()
    corners = [
        (x1, y1, x1, y2),
        (x1, y2, x2, y2),
        (x2, y2, x2, y1),
        (x2, y1, x1, y1),
    ]
    for lx1, ly1, lx2, ly2 in corners:
        ms.AddLine(make_point(lx1, ly1), make_point(lx2, ly2))
    return f"已繪製矩形：({x1},{y1}) 到 ({x2},{y2})"


def draw_polyline(points: list[list[float]]) -> str:
    """
    繪製多段線
    points: [[x1,y1],[x2,y2],...] 至少需要兩點
    """
    import win32com.client as wc
    if len(points) < 2:
        return "錯誤：至少需要兩個點"
    flat = []
    for p in points:
        flat.extend([float(p[0]), float(p[1])])
    arr = wc.VARIANT(wc.pythoncom.VT_ARRAY | wc.pythoncom.VT_R8, flat)
    ms, _ = get_model_space()
    ms.AddLightWeightPolyline(arr)
    return f"已繪製多段線，共 {len(points)} 個頂點"


def add_text(x: float, y: float, text: str, height: float = 2.5) -> str:
    """在指定位置新增單行文字"""
    ms, _ = get_model_space()
    insert = make_point(x, y)
    ms.AddText(text, insert, height)
    return f"已新增文字「{text}」於 ({x},{y})"


def add_mtext(x: float, y: float, text: str, width: float = 100.0, height: float = 2.5) -> str:
    """新增多行文字（MText）"""
    ms, _ = get_model_space()
    insert = make_point(x, y)
    mtext = ms.AddMText(insert, width, text)
    mtext.Height = height
    return f"已新增多行文字於 ({x},{y})"


def run_command(command: str) -> str:
    """
    在 AutoCAD 命令列執行指令
    例如：ZOOM, PAN, REGEN, LAYER 等
    """
    acad = get_acad()
    doc = acad.ActiveDocument
    doc.SendCommand(command + "\n")
    return f"已執行命令：{command}"


def zoom_extents() -> str:
    """縮放至顯示所有物件"""
    return run_command("ZOOM E")


def list_layers() -> list[str]:
    """列出所有圖層名稱"""
    acad = get_acad()
    doc = acad.ActiveDocument
    layers = doc.Layers
    return [layer.Name for layer in layers]


def set_layer(layer_name: str) -> str:
    """切換目前作用圖層，若不存在則建立"""
    acad = get_acad()
    doc = acad.ActiveDocument
    layers = doc.Layers
    # 檢查是否已存在
    exists = any(l.Name == layer_name for l in layers)
    if not exists:
        layers.Add(layer_name)
    doc.ActiveLayer = doc.Layers.Item(layer_name)
    return f"已切換至圖層：{layer_name}"


def set_color(color_index: int) -> str:
    """
    設定目前繪圖顏色（ACI 色碼 1-255）
    常用：1=紅 2=黃 3=綠 4=青 5=藍 6=洋紅 7=白
    """
    acad = get_acad()
    doc = acad.ActiveDocument
    doc.ActiveColor.ColorIndex = color_index
    return f"已設定顏色索引：{color_index}"


def count_entities() -> dict:
    """統計模型空間中各類型物件數量"""
    ms, _ = get_model_space()
    counts: dict[str, int] = {}
    for entity in ms:
        etype = entity.EntityType
        name = entity.ObjectName
        counts[name] = counts.get(name, 0) + 1
    return counts
