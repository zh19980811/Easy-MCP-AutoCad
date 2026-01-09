# AutoCAD MCP 技术深度解析与 AI 绘图指南

本文档旨在揭示 `Easy-MCP-AutoCad` 的幕后工作原理，并以“导师”的视角，为您剖析 LLM（大语言模型）控制 AutoCAD 的技术路径。

---

## 🏗️ 1. 这个 MCP 是如何工作的？

MCP (Model Context Protocol) 是大模型与外部世界交互的**通用协议**。在这个项目中，它充当了 Claude/LLM 和笨重的工业软件 AutoCAD 之间的**翻译官**。

### 核心架构图

```mermaid
graph LR
    User[用户/LLM] -- 1. 自然语言指令 --> Client["MCP 客户端 (Claude Desktop)"]
    Client -- 2. JSON-RPC 请求 --> Server["MCP 服务器 (server.py)"]
    Server -- 3. COM 自动化调用 --> CAD[AutoCAD 应用程序]
    CAD -- 4. 执行结果/图纸数据 --> Server
    Server -- 5. 结构化响应 --> Client
    Client -- 6. 自然语言反馈 --> User
```

### 关键流程解析

1. **指令解析**：当你对其说“画个圆”，Claude 会分析这个意图，并匹配到 MCP 服务器提供的工具 `draw_circle`。
2. **工具调用**：Claude 向 `server.py` 发送一个 JSON 数据包，包含工具名和参数（如圆心坐标）。
3. **驱动 CAD**：`server.py` 收到请求后，不进行复杂的几何计算，而是直接通过 Windows 的 **COM 接口** 告诉 AutoCAD：“嘿，在你现在的模型空间里加一个圆对象”。
4. **即时反馈**：AutoCAD 完成绘制后，会生成一个唯一的 **Handle**（句柄/ID）。服务器捕捉这个 ID，反馈给 Claude，Claude 再告诉你“画好了”。

---

## 🛠️ 2. 技术栈揭秘

要实现这个魔法，我们用了几项关键技术：

| 技术组件 | 作用 | 为什么选它？ |
| :--- | :--- | :--- |
| **Python** | 核心编程语言 | 生态丰富，胶水语言，适合调用各种接口。 |
| **mcp (FastMCP)** | MCP 协议 SDK | 官方提供的 Python SDK，用极少的代码就能把 Python 函数变成 LLM 可调用的工具。 |
| **pywin32 (win32com)** | Windows COM 桥梁 | **最核心的技术**。它允许 Python 像 VBA 一样直接操作 Windows 上的应用程序（如 Office, CAD）。 |
| **SQLite** | 本地数据库 | 用来存储从 CAD 扫描下来的实体数据。把图形变成数据表，方便 LLM 进行复杂的统计和查询（因为 LLM 不擅长直接“看”图，但擅长读 SQL 结果）。 |

---

## 🤖 3. 深入：如何控制 AutoCAD？

AutoCAD 提供了一套名为 **ActiveX Automation (COM)** 的接口。这是一种标准，允许外部程序像内部脚本一样控制它。

### 3.1 建立连接

代码中的 `win32com.client.Dispatch("AutoCAD.Application")` 就是在寻找正在运行的 AutoCAD 进程。如果找到了，我们就拿到了这辆车的“方向盘”。

### 3.2 对象模型 (Object Model)

AutoCAD 的世界是层级化的，我们的 Python代码必须遵循这个层级：

1. **Application**: 整个 AutoCAD 程序。
2. **ActiveDocument**: 当前正在编辑的那张图纸。
3. **ModelSpace**: 模型空间（大家通常绘图的地方）。
4. **Entity**: 具体的实体（直线、圆、多段线）。

**举个栗子（Python 伪代码）：**

```python
# 拿到方向盘
app = GetAutoCAD()
# 拿到当前图纸
doc = app.ActiveDocument
# 拿到画布
space = doc.ModelSpace
# 命令画布：加一条线
line = space.AddLine(StartPoint, EndPoint)
# 修改那条线：旋转它
line.Rotate(BasePoint, 30_degrees)
```

### 3.3 数据获取与修改

* **获取**：我们遍历 `ModelSpace` 中的所有对象，读取它们的属性（如 `Layer`, `Color`, `Coordinates`）。
* **修改**：直接修改对象的属性（如 `entity.Color = 1`）或者调用对象的方法（如 `entity.Move()`）即可实时生效。

---

## 🔮 4. 扩展视野：AI 控制 AutoCAD 的其他技术路径

除了目前的 **Python + COM + MCP** 方案，让 AI 画图还有其他几条路，各有优劣：

### 方案 A：生成 AutoLISP / Visual LISP 代码 (极简流)

* **原理**：AutoLISP 是 CAD 的原生语言。LLM 非常擅长写 LISP 代码。
* **流程**：LLM 生成一段 `.lsp` 代码 -> 用户复制到 CAD 命令行执行。
* **优点**：无需安装任何 MCP 服务器，无需 Python 环境。
* **缺点**：**单向交互**。AI 只能“写”，很难“读”到你图纸里已有什么。适合一次性生成复杂图形。

### 方案 B：生成 DXF 文件 (文件流)

* **原理**：DXF 是 CAD 的文本交换格式。
* **流程**：LLM 直接写一个 `.dxf` 文件 -> 用户用 CAD 打开。
* **优点**：完全脱离 CAD 软件运行，可以在网页端生成。
* **缺点**：无法修改现有 DWG 图纸，只能生成新图。调试困难。

### 方案 C：.NET API (C#) + MCP (专业流)

* **原理**：AutoCAD 的 .NET API (C#) 比 COM 接口更强大、更现代化、性能更高。
* **流程**：编写一个 C# 的 MCP 服务器插件，直接嵌入 AutoCAD 内部运行 (In-Process)。
* **优点**：**性能最强**，功能覆盖最全（包括自定义实体、复杂的数据库操作）。
* **缺点**：开发难度大，需要编译，部署稍微麻烦（需要加载 DLL）。

### 方案 D：Python + ezdxf (无头流)

* **原理**：使用 `ezdxf` 库在没有安装 AutoCAD 的机器上直接读写 `.dxf` 文件。
* **优点**：适合在服务器端批量处理 CAD 数据，不需要买 CAD 许可证。
* **缺点**：对 DWG 格式支持有限，同样无法进行所见即所得的交互。

## 🏆 总结与建议

目前的 **Python + COM + MCP** 方案对于个人助手来说是**最佳平衡点**：

1. **开发快**：Python 简单易懂。
2. **交互好**：支持“读”和“写”，实现真正的对话式绘图。
3. **门槛低**：不需要精通 C# 或 LISP。

希望这份指南能帮您建立起对 AI CAD 开发的全局视角！如果您想尝试其他技术路径（比如让 LLM 写一段 Lisp 试试），随时告诉我。
