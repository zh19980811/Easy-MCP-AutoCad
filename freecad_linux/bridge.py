"""FreeCAD TCP 桥接服务器 — 运行在 FreeCADCmd（无头）或 FreeCAD GUI 宏内。

协议：每行一个 JSON 请求 {"id": 1, "cmd": "draw_line", "params": {...}}
     每行回复 JSON：{"id": 1, "ok": true, "result": "..."} 或 {"id": 1, "ok": false, "error": "..."}
"""
import FreeCAD as App
import Part
import threading, socket, json, math, sys

COLORS = {1: (1.0, 0.0, 0.0), 2: (1.0, 1.0, 0.0), 3: (0.0, 1.0, 0.0),
          4: (0.0, 1.0, 1.0), 5: (0.0, 0.0, 1.0), 6: (1.0, 0.0, 1.0), 7: (1.0, 1.0, 1.0)}

def get_doc():
    if App.ActiveDocument is None:
        App.newDocument("Unnamed")
    return App.ActiveDocument

def set_layer(obj, layer):
    """给对象附加 Layer 字符串属性（FreeCAD 无原生图层，用附加属性模拟）"""
    if not hasattr(obj, "Layer"):
        obj.addProperty("App::PropertyString", "Layer", "Base")
    obj.Layer = layer or "0"
    try:
        grp = get_layer_group(obj.Layer)
        if grp is not None and obj not in grp.Group:
            grp.addObject(obj)
    except Exception:
        pass

def get_layer_group(name):
    layer_name = "LAYER_" + (name or "0")
    doc = get_doc()
    grp = doc.getObject(layer_name)
    if grp is None:
        grp = doc.addObject("App::DocumentObjectGroup", layer_name)
        if not hasattr(grp, "ColorIndex"):
            grp.addProperty("App::PropertyString", "ColorIndex", "Base")
        grp.ColorIndex = "7"
    return grp

def gui_up():
    try:
        import FreeCADGui
        return True
    except Exception:
        return False

def set_color(obj, color_index):
    """GUI 模式下设置对象颜色；无头模式返回 False"""
    if not gui_up():
        return False
    vo = obj.ViewObject
    if vo is None:
        return False
    vo.ShapeColor = COLORS.get(color_index, COLORS[7])
    return True

# ---------------- 命令实现：全部返回纯字符串 result，错误抛 ValueError ----------------

def cmd_ping(params):
    return "pong"

def cmd_new_document(params):
    doc = App.newDocument(params.get("name", "Unnamed"))
    return doc.Name

def cmd_draw_line(params):
    line = get_doc().addObject("Part::Line", "Line")
    line.X1 = float(params["start_x"]); line.Y1 = float(params["start_y"])
    line.X2 = float(params["end_x"]);   line.Y2 = float(params["end_y"])
    set_layer(line, params.get("layer", "0"))
    get_doc().recompute()
    return line.Name

def cmd_draw_circle(params):
    circle = get_doc().addObject("Part::Circle", "Circle")
    circle.Radius = float(params["radius"])
    circle.Placement = App.Placement(
        App.Vector(float(params["center_x"]), float(params["center_y"]), 0), App.Rotation())
    set_layer(circle, params.get("layer", "0"))
    get_doc().recompute()
    return circle.Name

def cmd_draw_polyline(params):
    pts = [App.Vector(x, y, 0) for x, y in zip(params["points"][::2], params["points"][1::2])]
    if len(pts) < 2:
        raise ValueError("points 至少需要 2 个点（4 个数）")
    closed = bool(params.get("closed", False))
    if closed:
        pts = pts + [pts[0]]
    obj = Part.show(Part.makePolygon(pts), "Polyline")
    set_layer(obj, params.get("layer", "0"))
    get_doc().recompute()
    return obj.Name

def cmd_draw_rectangle(params):
    x1, y1 = float(params["x1"]), float(params["y1"])
    x2, y2 = float(params["x2"]), float(params["y2"])
    pts = [App.Vector(x1, y1, 0), App.Vector(x2, y1, 0),
           App.Vector(x2, y2, 0), App.Vector(x1, y2, 0), App.Vector(x1, y1, 0)]
    obj = Part.show(Part.makePolygon(pts), "Rectangle")
    set_layer(obj, params.get("layer", "0"))
    get_doc().recompute()
    return obj.Name

def _text_of(obj):
    """App::Annotation 的 LabelText 在 FreeCAD 1.x 是字符串列表，统一取回纯文本"""
    t = obj.LabelText
    if isinstance(t, (list, tuple)):
        return "\n".join(t)
    return t or ""

def cmd_draw_text(params):
    obj = get_doc().addObject("App::Annotation", "Text")
    obj.LabelText = [params["text"]]
    obj.Position = App.Vector(float(params["insert_x"]), float(params["insert_y"]), 0)
    if hasattr(obj, "FontSize"):  # 1.0+ 已移除 FontSize
        obj.FontSize = float(params.get("height", 2.5))
    if not hasattr(obj, "Rotation"):
        obj.addProperty("App::PropertyAngle", "Rotation", "Base")
    obj.Rotation = float(params.get("rotation", 0))
    set_layer(obj, params.get("layer", "0"))
    return obj.Name

def cmd_create_layer(params):
    name = params["name"]
    grp = get_layer_group(name)
    if not hasattr(grp, "ColorIndex"):
        grp.addProperty("App::PropertyString", "ColorIndex", "Base")
    grp.ColorIndex = str(params.get("color_index", 7))
    return name

def _shape_points(shape, per_edge=8):
    pts = []
    try:
        for e in shape.Edges:
            try:
                for p in e.discretize(Number=per_edge):
                    pts.append([round(p.x, 4), round(p.y, 4), round(p.z, 4)])
            except Exception:
                pass
    except Exception:
        pass
    return pts

def cmd_scan_all(params):
    result = []
    for obj in get_doc().Objects:
        tid = obj.TypeId
        if tid == "App::DocumentObjectGroup":
            continue  # 图层组不是几何对象
        layer = getattr(obj, "Layer", "0")
        entry = {"name": obj.Name, "label": obj.Label, "type": tid, "layer": layer, "properties": "{}"}
        if tid == "Part::Line":
            p = obj.Placement
            s = p.multVec(App.Vector(float(obj.X1), float(obj.Y1), 0))
            e = p.multVec(App.Vector(float(obj.X2), float(obj.Y2), 0))
            entry["properties"] = json.dumps({"start_point": [round(s.x, 4), round(s.y, 4), round(s.z, 4)],
                                              "end_point": [round(e.x, 4), round(e.y, 4), round(e.z, 4)]})
        elif tid == "Part::Circle":
            entry["properties"] = json.dumps({"center_point": [obj.Placement.Base.x, obj.Placement.Base.y, 0],
                                              "radius": float(obj.Radius)})
        elif tid == "App::Annotation":
            entry["properties"] = json.dumps({"text": _text_of(obj),
                                              "position": [obj.Position.x, obj.Position.y, 0]},
                                             ensure_ascii=False)
        elif hasattr(obj, "Shape") and obj.Shape is not None:
            entry["properties"] = json.dumps({"points": _shape_points(obj.Shape)})
        else:
            continue  # 图层组等非几何对象
        result.append(entry)
    return json.dumps(result, ensure_ascii=False)

def _get_obj(handle):
    obj = get_doc().getObject(handle)
    if obj is None:
        raise ValueError("未找到句柄为 %s 的实体" % handle)
    return obj

def cmd_highlight(params):
    obj = _get_obj(params["handle"])
    if not set_color(obj, params.get("color", 1)):
        return "headless: 无视图对象（无头模式不显示颜色）"
    return "已高亮 %s" % obj.Name

def cmd_count_text(params):
    pattern = params.get("pattern", "")
    found = []
    for obj in get_doc().Objects:
        if obj.TypeId == "App::Annotation" and pattern in _text_of(obj):
            found.append({"name": obj.Name, "text": _text_of(obj), "layer": getattr(obj, "Layer", "0")})
    return json.dumps(found, ensure_ascii=False)

def cmd_highlight_text_matches(params):
    pattern = params.get("pattern", "")
    color = params.get("color", 1)
    n = 0
    for obj in get_doc().Objects:
        if obj.TypeId == "App::Annotation" and pattern in _text_of(obj):
            if set_color(obj, color):
                n += 1
    return "已高亮 %d 个文本" % n

def cmd_move(params):
    obj = _get_obj(params["handle"])
    s, e = params["start_point"], params["end_point"]
    obj.Placement.move(App.Vector(e[0] - s[0], e[1] - s[1], e[2] - s[2]))
    get_doc().recompute()
    return "已移动 %s" % obj.Name

def cmd_rotate(params):
    obj = _get_obj(params["handle"])
    base = params["base_point"]
    angle = float(params["angle"])
    bv = App.Vector(base[0], base[1], base[2])
    P = obj.Placement
    R = App.Rotation(App.Vector(0, 0, 1), angle)  # 绕 z 轴，单位度
    obj.Placement.Base = bv + R.multVec(P.Base - bv)
    obj.Placement.Rotation = R.multiply(P.Rotation)
    get_doc().recompute()
    return "已旋转 %s %g 度" % (obj.Name, angle)

def cmd_copy(params):
    obj = _get_obj(params["handle"])
    new_obj = get_doc().copyObject(obj, False)
    s, e = params["start_point"], params["end_point"]
    new_obj.Placement.move(App.Vector(e[0] - s[0], e[1] - s[1], e[2] - s[2]))
    get_doc().recompute()
    return new_obj.Name

def cmd_query_highlight(params):
    color = params.get("color", 1)
    n = 0
    for h in params.get("handles", []):
        obj = get_doc().getObject(h)
        if obj is not None and set_color(obj, color):
            n += 1
    return "已高亮 %d 个实体" % n

def inject_guidoc(path):
    """无头保存的 FCStd 缺 GuiDocument.xml（视图状态），GUI 打开时所有对象 Visibility=False。
    按 templates/ 里的视图模板注入 Visibility=true 的描述，让 GUI 打开即见。"""
    import zipfile, os
    if not os.path.exists(path):
        return
    with zipfile.ZipFile(path) as z:
        if "GuiDocument.xml" in z.namelist():
            return  # 已有视图状态（GUI 保存的），不动
    # exec() 环境下没有 __file__，回退到工作目录（启动脚本已 cd 到本目录）
    tpl_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates") \
        if "__file__" in globals() else os.path.join(os.getcwd(), "templates")

    def tpl(name):
        try:
            with open(os.path.join(tpl_dir, name), encoding="utf-8") as f:
                return f.read()
        except OSError:
            return None

    blocks = []
    for obj in get_doc().Objects:
        body = tpl("vp_" + obj.TypeId.replace("::", "__").replace(" ", "_") + ".xml")
        if body is None:
            body = tpl("vp_fallback.xml")
        if body is None:
            continue
        blocks.append(body.replace("{NAME}", obj.Name))
    head = tpl("guidoc_head.txt") or "<Document SchemaVersion=\"1\" HasExpansion=\"1\">\n"
    tail = tpl("guidoc_tail.txt") or "    <Camera settings=\"\"/>\n"
    gui = (head + "    <ViewProviderData Count=\"" + str(len(blocks)) + "\">\n"
           + "".join(blocks) + "    </ViewProviderData>\n" + tail + "</Document>\n")
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("GuiDocument.xml", gui)
    os.replace(tmp, path)

def cmd_save(params):
    path = params["path"]
    if not path.lower().endswith(".fcstd"):
        path += ".FCStd"
    get_doc().saveAs(path)
    inject_guidoc(path)
    return "已保存 %s" % path

def cmd_export_dxf(params):
    path = params["path"]
    if not path.lower().endswith(".dxf"):
        path += ".dxf"
    try:
        import importers.importDXF as importDXF
    except Exception:
        try:
            import importDXF
        except Exception:
            raise ValueError("FreeCAD 无 importDXF 模块，无法导出 DXF")
    importDXF.export(get_doc().Objects, path)
    return "已导出 %s" % path

def cmd_close_document(params):
    doc = get_doc()
    name = doc.Name
    App.closeDocument(name)
    return "已关闭 %s" % name

commands = {
    "ping": cmd_ping, "new_document": cmd_new_document, "draw_line": cmd_draw_line,
    "draw_circle": cmd_draw_circle, "draw_polyline": cmd_draw_polyline,
    "draw_rectangle": cmd_draw_rectangle, "draw_text": cmd_draw_text,
    "create_layer": cmd_create_layer, "scan_all": cmd_scan_all, "highlight": cmd_highlight,
    "count_text": cmd_count_text, "highlight_text_matches": cmd_highlight_text_matches,
    "move": cmd_move, "rotate": cmd_rotate, "copy": cmd_copy,
    "query_highlight": cmd_query_highlight, "save": cmd_save,
    "export_dxf": cmd_export_dxf, "close_document": cmd_close_document,
}

def handle_request(conn, addr):
    try:
        with conn.makefile("r") as f:
            for line in f:
                line = line.strip()
                if not line:
                    break
                try:
                    req = json.loads(line)
                    rid = req.get("id", 0)
                    fn = commands.get(req.get("cmd", ""))
                    if fn is None:
                        resp = {"id": rid, "ok": False, "error": "未知命令: %s" % req.get("cmd")}
                    else:
                        try:
                            resp = {"id": rid, "ok": True, "result": fn(req.get("params", {}))}
                        except ValueError as e:
                            resp = {"id": rid, "ok": False, "error": str(e)}
                        except Exception as e:
                            resp = {"id": rid, "ok": False, "error": "命令 %s 失败: %s" % (req.get("cmd"), e)}
                    conn.sendall((json.dumps(resp, ensure_ascii=False) + "\n").encode("utf-8"))
                except json.JSONDecodeError:
                    conn.sendall('{"id":0,"ok":false,"error":"JSON 解析失败"}\n'.encode("utf-8"))
    except Exception as e:
        print("handle_request 出错:", e)
    finally:
        try:
            conn.close()
        except Exception:
            pass

def run_server():
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind(("127.0.0.1", 8765))
    server.listen(5)
    print("FreeCAD bridge listening on 127.0.0.1:8765", flush=True)
    while True:
        conn, addr = server.accept()
        threading.Thread(target=handle_request, args=(conn, addr), daemon=True).start()

_started = False

def start_bridge():
    """启动 TCP 服务器线程（幂等）。直接脚本运行或 console exec() 都会调用它。"""
    global _started
    if _started:
        return
    _started = True
    t = threading.Thread(target=run_server, daemon=True)
    t.start()
    print("bridge 已启动，监听 127.0.0.1:8765", flush=True)

start_bridge()
