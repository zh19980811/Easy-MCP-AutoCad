from mcp.server.fastmcp import FastMCP, Context
from typing import Optional, List, Dict, Any
import win32com.client
import sqlite3
import json
import random
import re

# 创建服务器
mcp = FastMCP("AutoCAD-DB-Server")

# 初始化 SQLite 数据库
def init_db():
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        # 创建实体表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS cad_elements (
            id INTEGER PRIMARY KEY,
            handle TEXT UNIQUE,
            name TEXT NOT NULL,
            type TEXT NOT NULL,
            layer TEXT,
            properties TEXT
        )
        ''')
        # 创建文字内容统计表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS text_patterns (
            id INTEGER PRIMARY KEY,
            pattern TEXT UNIQUE,
            count INTEGER DEFAULT 0,
            drawing TEXT
        )
        ''')
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"数据库初始化失败: {str(e)}")
        return False

# 确保数据库初始化
init_db()

# ======= AutoCAD 基础工具 =======

@mcp.tool()
def create_new_drawing(ctx: Context, template: Optional[str] = None) -> str:
    """创建新的 AutoCAD 图纸"""
    try:
        # 尝试连接到 AutoCAD
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True
        
        # 创建新文档
        if template:
            doc = acad.Documents.Add(template)
        else:
            doc = acad.Documents.Add()
            
        return f"成功创建新图纸"
    except Exception as e:
        return f"创建图纸失败: {str(e)}"

@mcp.tool()
def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float, layer: Optional[str] = None) -> str:
    """在AutoCAD中绘制直线
    
    Args:
        start_x: 起点X坐标
        start_y: 起点Y坐标
        end_x: 终点X坐标
        end_y: 终点Y坐标
        layer: 可选的图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 如果指定了图层，先切换或创建图层
        if layer:
            try:
                # 尝试获取图层
                doc.Layers.Item(layer)
            except:
                # 图层不存在，创建新图层
                doc.Layers.Add(layer)
            
            # 设置当前图层
            doc.ActiveLayer = doc.Layers.Item(layer)
        
        # 创建直线
        line = model_space.AddLine(
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0]),
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        )
        
        # 将线条信息存入数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "start_point": [start_x, start_y, 0],
            "end_point": [end_x, end_y, 0]
        }
        cursor.execute(
            "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
            (line.Handle, "Line", "AcDbLine", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()
        
        return f"已创建直线，Handle: {line.Handle}, 图层: {doc.ActiveLayer.Name}"
    except Exception as e:
        return f"创建直线失败: {str(e)}"

@mcp.tool()
def draw_circle(ctx: Context, center_x: float, center_y: float, radius: float, layer: Optional[str] = None) -> str:
    """在AutoCAD中绘制圆
    
    Args:
        center_x: 圆心X坐标
        center_y: 圆心Y坐标  
        radius: 半径
        layer: 可选的图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 如果指定了图层，先切换或创建图层
        if layer:
            try:
                # 尝试获取图层
                doc.Layers.Item(layer)
            except:
                # 图层不存在，创建新图层
                doc.Layers.Add(layer)
            
            # 设置当前图层
            doc.ActiveLayer = doc.Layers.Item(layer)
        
        # 创建圆
        center_point = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [center_x, center_y, 0])
        circle = model_space.AddCircle(center_point, radius)
        
        # 将圆信息存入数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "center_point": [center_x, center_y, 0],
            "radius": radius
        }
        cursor.execute(
            "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
            (circle.Handle, "Circle", "AcDbCircle", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()
        
        return f"已创建圆，Handle: {circle.Handle}, 半径: {radius}, 图层: {doc.ActiveLayer.Name}"
    except Exception as e:
        return f"创建圆失败: {str(e)}"

# ======= 实体扫描和数据库交互 =======

@mcp.tool()
def scan_all_entities(ctx: Context) -> str:
    """扫描当前图纸中的所有实体并保存到数据库"""
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 连接数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        
        # 清空现有记录（可选）
        cursor.execute("DELETE FROM cad_elements")
        
        # 统计信息
        count = 0
        entity_types = {}
        
        # 遍历所有实体
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                entity_type = entity.ObjectName
                
                # 统计类型数量
                if entity_type in entity_types:
                    entity_types[entity_type] += 1
                else:
                    entity_types[entity_type] = 1
                
                # 获取基本属性
                properties = {}
                if entity_type == "AcDbLine":
                    properties = {
                        "start_point": [entity.StartPoint[0], entity.StartPoint[1], entity.StartPoint[2]],
                        "end_point": [entity.EndPoint[0], entity.EndPoint[1], entity.EndPoint[2]]
                    }
                elif entity_type == "AcDbCircle":
                    properties = {
                        "center": [entity.Center[0], entity.Center[1], entity.Center[2]],
                        "radius": entity.Radius
                    }
                elif entity_type == "AcDbText" or entity_type == "AcDbMText":
                    properties = {
                        "text": entity.TextString,
                        "position": [entity.InsertionPoint[0], entity.InsertionPoint[1], entity.InsertionPoint[2]] if hasattr(entity, "InsertionPoint") else None,
                        "height": entity.Height if hasattr(entity, "Height") else None
                    }
                
                # 将实体信息存入数据库
                cursor.execute(
                    "INSERT OR REPLACE INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
                    (entity.Handle, entity.ObjectName.replace("AcDb", ""), entity.ObjectName, entity.Layer, json.dumps(properties))
                )
                
                count += 1
            except Exception as e:
                print(f"处理实体 {i} 时出错: {str(e)}")
        
        conn.commit()
        conn.close()
        
        # 格式化类型统计
        type_summary = "\n".join([f"{t}: {c}" for t, c in entity_types.items()])
        
        return f"已扫描并保存 {count} 个实体到数据库。\n\n实体类型统计:\n{type_summary}"
    except Exception as e:
        return f"扫描实体失败: {str(e)}"

@mcp.tool()
def highlight_entity(ctx: Context, handle: str, color: int = 1) -> str:
    """通过Handle在AutoCAD中高亮显示指定实体
    
    Args:
        handle: 实体的Handle值
        color: 高亮颜色码（1=红色, 2=黄色, 3=绿色, 4=青色, 5=蓝色, 6=洋红色）
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        
        # 根据Handle直接获取实体
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"

        original_color = entity.Color
        entity.Color = color
        
        return f"已高亮实体 {handle}，颜色从 {original_color} 改为 {color}"
    except Exception as e:
        return f"高亮实体失败: {str(e)}"

# ======= 文本分析工具 =======

@mcp.tool()
def count_text_patterns(ctx: Context, pattern: str = "PMC-3M") -> str:
    """统计图纸中文本实体中特定模式的出现次数
    
    Args:
        pattern: 要搜索的文本模式，默认为"PMC-3M"
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        drawing_name = doc.Name
        
        # 连接数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        
        # 计数器
        count = 0
        matching_entities = []
        
        # 遍历所有实体
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                
                # 检查是否为文本实体
                if hasattr(entity, "TextString"):
                    text = entity.TextString
                    
                    # 搜索模式
                    if pattern in text:
                        count += 1
                        matching_entities.append({
                            "handle": entity.Handle,
                            "text": text,
                            "layer": entity.Layer,
                            "position": [entity.InsertionPoint[0], entity.InsertionPoint[1]] if hasattr(entity, "InsertionPoint") else None
                        })
            except Exception as e:
                print(f"处理文本实体 {i} 时出错: {str(e)}")
        
        # 保存统计结果到数据库
        cursor.execute(
            "INSERT OR REPLACE INTO text_patterns (pattern, count, drawing) VALUES (?, ?, ?)",
            (pattern, count, drawing_name)
        )
        conn.commit()
        conn.close()
        
        result = f"在图纸 '{drawing_name}' 中找到 {count} 处匹配模式 '{pattern}' 的文本。"
        
        # 如果有匹配项，显示详细信息
        if count > 0:
            details = "\n\n匹配详情："
            for i, match in enumerate(matching_entities[:10]):  # 限制显示前10个
                details += f"\n{i+1}. 文本: '{match['text']}', 图层: {match['layer']}, Handle: {match['handle']}"
            
            if len(matching_entities) > 10:
                details += f"\n... 以及其他 {len(matching_entities) - 10} 个匹配项"
                
            result += details
        
        return result
    except Exception as e:
        return f"统计文本模式失败: {str(e)}"

@mcp.tool()
def highlight_text_matches(ctx: Context, pattern: str = "PMC-3M", color: int = 1) -> str:
    """高亮显示包含指定文本模式的所有文本实体
    
    Args:
        pattern: 要搜索的文本模式，默认为"PMC-3M"
        color: 高亮颜色码（1=红色, 2=黄色, 3=绿色, 4=青色, 5=蓝色, 6=洋红色）
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 创建选择集
        try:
            # 尝试删除可能存在的选择集
            doc.SelectionSets.Item("TextMatches").Delete()
        except:
            pass
        
        selection = doc.SelectionSets.Add("TextMatches")
        
        # 计数器
        count = 0
        
        # 遍历所有实体
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                
                # 检查是否为文本实体
                if hasattr(entity, "TextString"):
                    text = entity.TextString
                    
                    # 搜索模式
                    if pattern in text:
                        # 保存原始颜色
                        original_color = entity.Color
                        
                        # 修改颜色
                        entity.Color = color
                        
                        # 添加到选择集
                        selection.AddItems([entity])
                        
                        count += 1
            except Exception as e:
                print(f"处理文本实体 {i} 时出错: {str(e)}")
        
        if count > 0:
            # 缩放到选择集
            doc.ActiveView.ZoomAll()
            return f"已高亮显示 {count} 个包含 '{pattern}' 的文本实体"
        else:
            selection.Delete()
            return f"未找到包含 '{pattern}' 的文本实体"
    except Exception as e:
        return f"高亮文本匹配失败: {str(e)}"

# ======= 数据库查询工具 =======

@mcp.tool()
def get_all_tables(ctx: Context) -> str:
    """获取数据库中的所有表"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        conn.close()
        
        table_list = [table[0] for table in tables]
        return json.dumps(table_list, indent=2)
    except Exception as e:
        return f"获取表列表失败: {str(e)}"

@mcp.tool()
def get_table_schema(ctx: Context, table_name: str) -> str:
    """获取指定表的结构信息"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA table_info({table_name});")
        columns = cursor.fetchall()
        conn.close()
        
        schema = []
        for col in columns:
            schema.append({
                "cid": col[0],
                "name": col[1],
                "type": col[2],
                "notnull": col[3],
                "default_value": col[4],
                "pk": col[5]
            })
        
        return json.dumps(schema, indent=2)
    except Exception as e:
        return f"获取表结构失败: {str(e)}"

@mcp.tool()
def execute_query(ctx: Context, query: str) -> str:
    """执行自定义数据库查询"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(query)
        
        # 如果是SELECT查询，获取结果
        if query.strip().upper().startswith("SELECT"):
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()
            
            # 将结果转换为字典列表
            result = []
            for row in rows:
                result.append(dict(zip(columns, row)))
                
            conn.commit()
            conn.close()
            return json.dumps(result, indent=2)
        else:
            # 非SELECT查询，返回影响的行数
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            return f"执行成功，影响了 {affected} 行"
    except Exception as e:
        return f"执行查询失败: {str(e)}"

@mcp.tool()
def query_and_highlight(ctx: Context, sql_query: str, highlight_color: int = 1) -> str:
    """根据SQL查询结果高亮显示AutoCAD实体
    
    Args:
        sql_query: 必须是返回handle列的SQL查询
        highlight_color: 高亮颜色码（1-255）
    """
    try:
        # 执行查询
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(sql_query)
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            return "查询未返回任何结果"
        
        # 获取列名
        column_names = [description[0] for description in cursor.description]
        
        # 查找handle列
        handle_index = -1
        for i, name in enumerate(column_names):
            if name.lower() == 'handle':
                handle_index = i
                break
        
        if handle_index == -1:
            return "查询结果中未找到handle列"
        
        # 提取所有handle
        handles = [row[handle_index] for row in rows]
        
        # 高亮实体
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        
        # 创建选择集
        try:
            doc.SelectionSets.Item("QueryResults").Delete()
        except:
            pass
        
        selection = doc.SelectionSets.Add("QueryResults")
        
        # 高亮找到的实体
        highlighted_count = 0
        entities_to_add = []
        
        for handle in handles:
            try:
                # 直接获取实体
                entity = doc.HandleToObject(handle)
                entity.Color = highlight_color
                entities_to_add.append(entity)
                highlighted_count += 1
            except Exception as e:
                print(f"处理实体 {handle} 时出错: {str(e)}")
        
        # 将实体添加到选择集（方便在CAD中查看属性等）
        if entities_to_add:
            try:
                # AddItems需要变体数组，comtypes会自动处理list，pywin32有时需要特别处理
                # 这里简单尝试，如果失败不影响高亮结果
                selection.AddItems(entities_to_add)
            except:
                pass
        
        if highlighted_count > 0:
            # 缩放到选择集
            doc.ActiveView.ZoomAll()
            return f"已高亮显示 {highlighted_count} 个实体（共 {len(handles)} 个结果）"
        else:
            return f"未能高亮任何实体"
    except Exception as e:
        return f"查询并高亮失败: {str(e)}"
# 添加到现有代码中

@mcp.tool()
@mcp.tool()
def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float, layer: Optional[str] = None) -> str:
    """在AutoCAD中绘制直线"""
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 处理图层...
        
        # 创建直线
        line = model_space.AddLine(
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0]),
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        )
        
        # 将线条信息存入数据库（不使用handle字段）
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "start_point": [start_x, start_y, 0],
            "end_point": [end_x, end_y, 0]
        }
        cursor.execute(
            "INSERT INTO cad_elements (name, type, layer, properties) VALUES (?, ?, ?, ?)",
            ("Line", "AcDbLine", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()
        
        return f"已创建直线，从 ({start_x}, {start_y}) 到 ({end_x}, {end_y})"
    except Exception as e:
        return f"创建直线失败: {str(e)}"

@mcp.tool()
def draw_polyline(ctx: Context, points: List[float], closed: bool = False, layer: Optional[str] = None) -> str:
    """绘制多段线 (Polyline)
    
    Args:
        points: 坐标列表 [x1, y1, x2, y2, ...]
        closed: 是否闭合
        layer: 图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档"
            
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 处理图层
        if layer:
            try:
                doc.Layers.Item(layer)
            except:
                doc.Layers.Add(layer)
            doc.ActiveLayer = doc.Layers.Item(layer)

        # 转换坐标点
        # LightweightPolyline 需要 2D 坐标数组 (double)
        pt_array = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, points)
        
        # 创建多段线
        pline = model_space.AddLightWeightPolyline(pt_array)
        pline.Closed = closed
        
        return f"已创建多段线，Handle: {pline.Handle}"
    except Exception as e:
        return f"创建多段线失败: {str(e)}"

@mcp.tool()
def draw_rectangle(ctx: Context, x1: float, y1: float, x2: float, y2: float, layer: Optional[str] = None) -> str:
    """绘制矩形
    
    Args:
        x1, y1: 角点1坐标
        x2, y2: 对角点2坐标
        layer: 图层名称
    """
    # 构造矩形的4个顶点坐标
    points = [x1, y1, x2, y1, x2, y2, x1, y2]
    return draw_polyline(ctx, points, closed=True, layer=layer)

@mcp.tool()
def draw_text(ctx: Context, text_string: str, insert_x: float, insert_y: float, height: float = 2.5, rotation: float = 0, layer: Optional[str] = None) -> str:
    """绘制单行文字
    
    Args:
        text_string: 文字内容
        insert_x, insert_y: 插入点坐标
        height: 文字高度
        rotation: 旋转角度 (度)
        layer: 图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档"
            
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 处理图层
        if layer:
            try:
                doc.Layers.Item(layer)
            except:
                doc.Layers.Add(layer)
            doc.ActiveLayer = doc.Layers.Item(layer)
            
        insert_pnt = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [insert_x, insert_y, 0])
        text_obj = model_space.AddText(text_string, insert_pnt, height)
        
        if rotation != 0:
            import math
            text_obj.Rotation = rotation * math.pi / 180
            
        return f"已创建文字 '{text_string}'，Handle: {text_obj.Handle}"
    except Exception as e:
        return f"创建文字失败: {str(e)}"

@mcp.tool()
def create_layer(ctx: Context, name: str, color_index: int = 7) -> str:
    """创建或修改图层
    
    Args:
        name: 图层名称
        color_index: 颜色索引 (1-255)
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        try:
            layer = doc.Layers.Item(name)
        except:
            layer = doc.Layers.Add(name)
            
        layer.Color = color_index
        return f"图层 '{name}' 已设置，颜色: {color_index}"
    except Exception as e:
        return f"设置图层失败: {str(e)}"

@mcp.tool()
def draw_device_connection(ctx: Context, start_device: str, end_device: str, start_x: Optional[float] = None, start_y: Optional[float] = None, end_x: Optional[float] = None, end_y: Optional[float] = None, layer: Optional[str] = None) -> str:
    """绘制设备之间的连接线
    
    Args:
        start_device: 起始设备标签，如"P14"
        end_device: 结束设备标签，如"P02"
        start_x: 可选的起始点X坐标（如果不提供则自动查找设备）
        start_y: 可选的起始点Y坐标
        end_x: 可选的结束点X坐标
        end_y: 可选的结束点Y坐标
        layer: 可选的图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 如果指定了图层，先切换或创建图层
        if layer:
            try:
                # 尝试获取图层
                doc.Layers.Item(layer)
            except:
                # 图层不存在，创建新图层
                doc.Layers.Add(layer)
            
            # 设置当前图层
            doc.ActiveLayer = doc.Layers.Item(layer)

        # 如果没有提供坐标，尝试从数据库中查找设备
        if start_x is None or start_y is None or end_x is None or end_y is None:
            conn = sqlite3.connect("autocad_data.db")
            cursor = conn.cursor()
            
            # 查找起始设备
            cursor.execute(
                "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                (start_device,)
            )
            start_result = cursor.fetchone()
            
            # 查找结束设备
            cursor.execute(
                "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                (end_device,)
            )
            end_result = cursor.fetchone()
            
            conn.close()
            
            if not start_result:
                return f"未找到标签为 {start_device} 的设备"
                
            if not end_result:
                return f"未找到标签为 {end_device} 的设备"
            
            # 解析设备位置和尺寸
            start_props = json.loads(start_result[0])
            end_props = json.loads(end_result[0])
            
            start_pos = start_props["position"]
            end_pos = end_props["position"]
            
            # 设置连接线起点和终点（设备的左侧连接点）
            start_x = start_pos[0] - 5  # 设备左侧
            start_y = start_pos[1]
            end_x = end_pos[0] - 5
            end_y = end_pos[1]
            
        # 创建连接线（水平线段 + 垂直线段 + 水平线段）
        # 首先创建起始水平线段
        line1_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0])
        line1_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, start_y, 0])
        line1 = model_space.AddLine(line1_start, line1_end)
        
        # 创建垂直连接线
        line2_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, start_y, 0])
        line2_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, end_y, 0])
        line2 = model_space.AddLine(line2_start, line2_end)
        
        # 创建结束水平线段
        line3_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, end_y, 0])
        line3_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        line3 = model_space.AddLine(line3_start, line3_end)
        
        return f"已创建从 {start_device} 到 {end_device} 的连接线"
    except Exception as e:
        return f"创建连接线失败: {str(e)}"
    
@mcp.tool()
def move_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
    """移动实体
    
    Args:
        handle: 实体句柄
        start_point: 基干点 [x, y, z]（通常都是[0,0,0]或者实体的某个点）
        end_point: 目标点 [x, y, z]
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"
             
        p1 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, start_point)
        p2 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, end_point)
        
        entity.Move(p1, p2)
        return f"已移动实体 {handle}"
    except Exception as e:
        return f"移动实体失败: {str(e)}"

@mcp.tool()
def rotate_entity(ctx: Context, handle: str, base_point: List[float], angle: float) -> str:
    """旋转实体
    
    Args:
        handle: 实体句柄
        base_point: 旋转中心点 [x, y, z]
        angle: 旋转角度（度）
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"
             
        p_base = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, base_point)
        
        import math
        angle_rad = angle * math.pi / 180
        
        entity.Rotate(p_base, angle_rad)
        return f"已旋转实体 {handle} {angle} 度"
    except Exception as e:
        return f"旋转实体失败: {str(e)}"

@mcp.tool()
def copy_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
    """复制实体
    
    Args:
        handle: 源实体句柄
        start_point: 基干点 [x, y, z]
        end_point: 目标点 [x, y, z]
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"
             
        # Copy()方法返回新对象
        new_entity = entity.Copy()
        
        p1 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, start_point)
        p2 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, end_point)
        
        new_entity.Move(p1, p2)
        return f"已复制实体，新Handle: {new_entity.Handle}"
    except Exception as e:
        return f"复制实体失败: {str(e)}"
# 启动服务器
if __name__ == "__main__":
    mcp.run()