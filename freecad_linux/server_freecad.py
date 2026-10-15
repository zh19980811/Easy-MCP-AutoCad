"""FreeCAD 版 AutoCAD MCP 服务器 — 工具名/参数与原 Easy-MCP-AutoCad 1:1，后端换成 FreeCAD bridge。

运行：~/venvs/autocad/bin/python server_freecad.py（stdio MCP）
依赖：mcp 1.x（FastMCP）、SQLite（标准库）；bridge 见 bridge.py（127.0.0.1:8765）
"""
from mcp.server.fastmcp import FastMCP, Context
from typing import Optional, List, Dict, Any
import sqlite3, json, socket, re

BRIDGE_HOST, BRIDGE_PORT = "127.0.0.1", 8765
DB_FILE = "autocad_data.db"

mcp = FastMCP("FreeCAD-DB-Server")

# ---------------- bridge 客户端 ----------------

def bridge_call(cmd: str, params: Optional[dict] = None, timeout: int = 60) -> str:
    """调用 FreeCAD bridge，成功返回 result 字符串，失败抛 RuntimeError"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(timeout)
        s.connect((BRIDGE_HOST, BRIDGE_PORT))
        s.sendall((json.dumps({"id": 1, "cmd": cmd, "params": params or {}}) + "\n").encode("utf-8"))
        resp = json.loads(s.makefile().readline())
        s.close()
    except (socket.error, OSError):
        raise RuntimeError("无法连接 FreeCAD bridge（127.0.0.1:8765），请先启动 bridge")
    if not resp.get("ok"):
        raise RuntimeError(resp.get("error", "bridge 返回未知错误"))
    return resp.get("result", "")

def _clean_params(**kw):
    """去掉值为 None 的可选参数"""
    return {k: v for k, v in kw.items() if v is not None}

# ---------------- 数据库 ----------------

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS cad_elements (
        id INTEGER PRIMARY KEY, handle TEXT UNIQUE, name TEXT NOT NULL,
        type TEXT NOT NULL, layer TEXT, properties TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS text_patterns (
        id INTEGER PRIMARY KEY, pattern TEXT UNIQUE,
        count INTEGER DEFAULT 0, drawing TEXT)""")
    conn.commit()
    conn.close()

init_db()

def _db_insert(handle, name, type_, layer, props: dict):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    # FreeCAD 对象名按文档重复，同名句柄以新代旧（DB 镜像当前图纸）
    cur.execute("INSERT OR REPLACE INTO cad_elements (handle, name, type, layer, properties) VALUES (?,?,?,?,?)",
                (handle, name, type_, layer or "0", json.dumps(props, ensure_ascii=False)))
    conn.commit()
    conn.close()

# ---------------- 图纸/绘制工具 ----------------

@mcp.tool()
def create_new_drawing(ctx: Context, template: Optional[str] = None) -> str:
    try:
        name = bridge_call("new_document", {"name": template} if template else {})
        # 新图纸 = 新的实体集合，清空实体表（text_patterns 保留统计历史）
        conn = sqlite3.connect(DB_FILE)
        conn.execute("DELETE FROM cad_elements")
        conn.commit()
        conn.close()
        return f"成功创建新图纸 {name}"
    except Exception as e:
        return f"创建图纸失败: {str(e)}"

@mcp.tool()
def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float,
              layer: Optional[str] = None) -> str:
    try:
        handle = bridge_call("draw_line", _clean_params(
            start_x=start_x, start_y=start_y, end_x=end_x, end_y=end_y, layer=layer))
        _db_insert(handle, "Line", "Part::Line", layer,
                   {"start_point": [start_x, start_y, 0], "end_point": [end_x, end_y, 0]})
        return f"已创建直线，Handle: {handle}, 图层: {layer or '0'}"
    except Exception as e:
        return f"创建直线失败: {str(e)}"

@mcp.tool()
def draw_circle(ctx: Context, center_x: float, center_y: float, radius: float,
                layer: Optional[str] = None) -> str:
    try:
        handle = bridge_call("draw_circle", _clean_params(
            center_x=center_x, center_y=center_y, radius=radius, layer=layer))
        _db_insert(handle, "Circle", "Part::Circle", layer,
                   {"center_point": [center_x, center_y, 0], "radius": radius})
        return f"已创建圆，Handle: {handle}, 半径: {radius}, 图层: {layer or '0'}"
    except Exception as e:
        return f"创建圆失败: {str(e)}"

@mcp.tool()
def draw_polyline(ctx: Context, points: List[float], closed: bool = False,
                  layer: Optional[str] = None) -> str:
    try:
        handle = bridge_call("draw_polyline", _clean_params(
            points=points, closed=closed, layer=layer))
        _db_insert(handle, "Polyline", "Part::Feature", layer,
                   {"points": points, "closed": closed})
        return f"已创建多段线，Handle: {handle}"
    except Exception as e:
        return f"创建多段线失败: {str(e)}"

@mcp.tool()
def draw_rectangle(ctx: Context, x1: float, y1: float, x2: float, y2: float,
                   layer: Optional[str] = None) -> str:
    try:
        handle = bridge_call("draw_rectangle", _clean_params(
            x1=x1, y1=y1, x2=x2, y2=y2, layer=layer))
        _db_insert(handle, "Rectangle", "Part::Feature", layer,
                   {"corner1": [x1, y1, 0], "corner2": [x2, y2, 0]})
        return f"已创建矩形，Handle: {handle}"
    except Exception as e:
        return f"创建矩形失败: {str(e)}"

@mcp.tool()
def draw_text(ctx: Context, text_string: str, insert_x: float, insert_y: float,
              height: float = 2.5, rotation: float = 0, layer: Optional[str] = None) -> str:
    try:
        handle = bridge_call("draw_text", _clean_params(
            text=text_string, insert_x=insert_x, insert_y=insert_y,
            height=height, rotation=rotation, layer=layer))
        _db_insert(handle, "Text", "App::Annotation", layer,
                   {"text": text_string, "position": [insert_x, insert_y, 0],
                    "height": height, "rotation": rotation})
        return f"已创建文字 '{text_string}'，Handle: {handle}"
    except Exception as e:
        return f"创建文字失败: {str(e)}"

@mcp.tool()
def create_layer(ctx: Context, name: str, color_index: int = 7) -> str:
    try:
        bridge_call("create_layer", {"name": name, "color_index": color_index})
        return f"图层 '{name}' 已设置，颜色: {color_index}"
    except Exception as e:
        return f"设置图层失败: {str(e)}"

# ---------------- 扫描 / 高亮 / 文本 ----------------

@mcp.tool()
def scan_all_entities(ctx: Context) -> str:
    try:
        raw = bridge_call("scan_all")
        entities = json.loads(raw)
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("DELETE FROM cad_elements")
        type_count = {}
        for e in entities:
            tid = e["type"]
            type_count[tid] = type_count.get(tid, 0) + 1
            cur.execute(
                "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?,?,?,?,?)",
                (e["name"], e["type"].replace("Part::", "").replace("App::", ""),
                 tid, e.get("layer", "0"), e["properties"]))
        conn.commit()
        conn.close()
        summary = "\n".join(f"{t}: {c}" for t, c in type_count.items())
        return f"已扫描并保存 {len(entities)} 个实体到数据库。\n\n实体类型统计:\n{summary}"
    except Exception as e:
        return f"扫描实体失败: {str(e)}"

@mcp.tool()
def highlight_entity(ctx: Context, handle: str, color: int = 1) -> str:
    try:
        bridge_call("highlight", {"handle": handle, "color": color})
        return f"已高亮实体 {handle}，颜色码: {color}"
    except Exception as e:
        return f"高亮实体失败: {str(e)}"

@mcp.tool()
def count_text_patterns(ctx: Context, pattern: str = "PMC-3M") -> str:
    try:
        raw = bridge_call("count_text", {"pattern": pattern})
        matches = json.loads(raw)
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("INSERT OR REPLACE INTO text_patterns (pattern, count, drawing) VALUES (?,?,?)",
                    (pattern, len(matches), "FreeCAD"))
        conn.commit()
        conn.close()
        result = f"在图纸中找到 {len(matches)} 处匹配模式 '{pattern}' 的文本。"
        if matches:
            details = "\n".join(
                f"{i+1}. 文本: '{m['text']}', 图层: {m.get('layer', '0')}, Handle: {m['name']}"
                for i, m in enumerate(matches[:10]))
            if len(matches) > 10:
                details += f"\n... 以及其他 {len(matches) - 10} 个匹配项"
            result += "\n\n匹配详情：\n" + details
        return result
    except Exception as e:
        return f"统计文本模式失败: {str(e)}"

@mcp.tool()
def highlight_text_matches(ctx: Context, pattern: str = "PMC-3M", color: int = 1) -> str:
    try:
        bridge_call("highlight_text_matches", {"pattern": pattern, "color": color})
        n = len(json.loads(bridge_call("count_text", {"pattern": pattern})))
        return f"已高亮显示 {n} 个包含 '{pattern}' 的文本实体"
    except Exception as e:
        return f"高亮文本匹配失败: {str(e)}"

# ---------------- 数据库查询 ----------------

@mcp.tool()
def get_all_tables(ctx: Context) -> str:
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [r[0] for r in cur.fetchall()]
        conn.close()
        return json.dumps(tables, indent=2)
    except Exception as e:
        return f"获取表列表失败: {str(e)}"

@mcp.tool()
def get_table_schema(ctx: Context, table_name: str) -> str:
    try:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", table_name):
            return f"获取表结构失败: 非法表名 {table_name}"
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute(f"PRAGMA table_info({table_name})")
        rows = cur.fetchall()
        conn.close()
        schema = [{"cid": c[0], "name": c[1], "type": c[2], "notnull": c[3],
                   "default_value": c[4], "pk": c[5]} for c in rows]
        return json.dumps(schema, indent=2)
    except Exception as e:
        return f"获取表结构失败: {str(e)}"

@mcp.tool()
def execute_query(ctx: Context, query: str) -> str:
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute(query)
        if query.strip().upper().startswith("SELECT"):
            columns = [d[0] for d in cur.description]
            rows = [dict(zip(columns, row)) for row in cur.fetchall()]
            conn.close()
            return json.dumps(rows, indent=2)
        conn.commit()
        affected = cur.rowcount
        conn.close()
        return f"执行成功，影响了 {affected} 行"
    except Exception as e:
        return f"执行查询失败: {str(e)}"

@mcp.tool()
def query_and_highlight(ctx: Context, sql_query: str, highlight_color: int = 1) -> str:
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute(sql_query)
        rows = cur.fetchall()
        col_names = [d[0] for d in cur.description]
        conn.close()
        if not rows:
            return "查询未返回任何结果"
        handle_idx = next((i for i, c in enumerate(col_names) if c.lower() == "handle"), -1)
        if handle_idx == -1:
            return "查询结果中未找到handle列"
        handles = [r[handle_idx] for r in rows]
        bridge_call("query_highlight", {"handles": handles, "color": highlight_color})
        return f"已高亮显示 {len(handles)} 个实体（共 {len(handles)} 个结果）"
    except Exception as e:
        return f"查询并高亮失败: {str(e)}"

# ---------------- 设备连接 / 编辑 / 存盘 ----------------

@mcp.tool()
def draw_device_connection(ctx: Context, start_device: str, end_device: str,
                           start_x: Optional[float] = None, start_y: Optional[float] = None,
                           end_x: Optional[float] = None, end_y: Optional[float] = None,
                           layer: Optional[str] = None) -> str:
    try:
        if start_x is None or start_y is None or end_x is None or end_y is None:
            conn = sqlite3.connect(DB_FILE)
            cur = conn.cursor()
            cur.execute("SELECT properties FROM cad_elements WHERE type='CustomDevice'")
            rows = cur.fetchall()
            conn.close()
            found = {}
            for (props_json,) in rows:
                try:
                    props = json.loads(props_json)
                except Exception:
                    continue
                label = props.get("label")
                if label:
                    found[label] = props.get("position", [0, 0])
            if start_device not in found:
                return f"未找到标签为 {start_device} 的设备"
            if end_device not in found:
                return f"未找到标签为 {end_device} 的设备"
            sp, ep = found[start_device], found[end_device]
            start_x, start_y, end_x, end_y = sp[0], sp[1], ep[0], ep[1]
        # 与原版相同的三段折线：起始水平 + 垂直 + 结束水平
        p = _clean_params(layer=layer)
        bridge_call("draw_line", {**p, "start_x": start_x, "start_y": start_y,
                                  "end_x": start_x - 10, "end_y": start_y})
        bridge_call("draw_line", {**p, "start_x": start_x - 10, "start_y": start_y,
                                  "end_x": start_x - 10, "end_y": end_y})
        bridge_call("draw_line", {**p, "start_x": start_x - 10, "start_y": end_y,
                                  "end_x": end_x, "end_y": end_y})
        return f"已创建从 {start_device} 到 {end_device} 的连接线"
    except Exception as e:
        return f"创建连接线失败: {str(e)}"

@mcp.tool()
def move_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
    try:
        bridge_call("move", {"handle": handle, "start_point": start_point, "end_point": end_point})
        return f"已移动实体 {handle}"
    except Exception as e:
        return f"移动实体失败: {str(e)}"

@mcp.tool()
def rotate_entity(ctx: Context, handle: str, base_point: List[float], angle: float) -> str:
    try:
        bridge_call("rotate", {"handle": handle, "base_point": base_point, "angle": angle})
        return f"已旋转实体 {handle} {angle} 度"
    except Exception as e:
        return f"旋转实体失败: {str(e)}"

@mcp.tool()
def copy_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
    try:
        new_handle = bridge_call("copy", {"handle": handle,
                                          "start_point": start_point, "end_point": end_point})
        return f"已复制实体，新Handle: {new_handle}"
    except Exception as e:
        return f"复制实体失败: {str(e)}"

@mcp.tool()
def save_drawing(ctx: Context, path: str) -> str:
    try:
        return bridge_call("save", {"path": path})
    except Exception as e:
        return f"保存图纸失败: {str(e)}"

@mcp.tool()
def export_dxf(ctx: Context, path: str) -> str:
    try:
        return bridge_call("export_dxf", {"path": path})
    except Exception as e:
        return f"导出DXF失败: {str(e)}"

if __name__ == "__main__":
    mcp.run()
