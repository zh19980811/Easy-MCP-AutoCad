# FreeCAD Linux 移植版（Easy-MCP-AutoCad → FreeCAD）

原项目仅支持 Windows + AutoCAD COM，Linux 上无法运行。本移植把它的 19 个 MCP 工具 1:1 迁移到 FreeCAD 后端，并新增 `save_drawing` / `export_dxf`（共 21 个工具），让 LLM 通过 MCP 协议在 Ubuntu/Linux 上用自然语言完成绘图、扫描、查询、高亮、存盘。

## 目录内容

- `bridge.py` — FreeCAD 内部 TCP 桥接服务器（127.0.0.1:8765，JSON 行协议）
- `server_freecad.py` — MCP 服务器（stdio），连 bridge 执行几何操作 + 管理 SQLite 实体库
- `start_bridge_headless.sh` — 无头模式一键启动 bridge（自带进程清理）
- `templates/` — FCStd 保存时注入的视图状态模板（保证 GUI 打开对象可见）
- `使用说明.md` — 完整文档（架构、启动方式、21 工具清单、与原版差异、已知局限、测试记录）

## 快速开始（Ubuntu）

```bash
# 依赖：FreeCAD（含 freecadcmd）+ python3；venv 内安装请求库后：
./start_bridge_headless.sh                          # 终端 A：启动 bridge
python3 server_freecad.py                           # 终端 B：MCP 服务器（stdio）
```

MCP 客户端（如 Claude Desktop）配置示例见 `使用说明.md`。

## 与 AutoCAD 原版的差异

| 项 | 原版 | 本移植 |
|---|---|---|
| 后端 | AutoCAD COM (pywin32) | FreeCAD Part API |
| 句柄 | AutoCAD Handle | FreeCAD 对象 Name |
| 图层 | 原生图层 | 文档组 LAYER_&lt;名&gt; + 附加属性 |
| 存盘 | AutoCAD 自管 | 必须用 save_drawing / export_dxf |
| 数据库 type | AcDbLine 等 | FreeCAD TypeId（Part::Line 等，scan 与 draw 一致） |

## 测试

bridge 与 server 均已通过无头实测：7 类图形绘制、世界坐标扫描、文字统计、移动/旋转/复制、图层、FCStd/DXF 导出、缺参/未知命令错误路径、21 工具 MCP stdio 冒烟。
